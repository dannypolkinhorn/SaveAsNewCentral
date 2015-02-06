Imports Autodesk.Revit
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.Utility
Imports System.IO
Imports xl = Microsoft.Office.Interop.Excel
Imports Autodesk.Revit.UI.Events

'AUTODESK DOES NOT GUARANTEE THAT YOU WILL BE ABLE TO SUCCESSFULLY DOWNLOAD OR IMPLEMENT ANY SAMPLE CODE.
'SAMPLE CODE IS SUBJECT TO CHANGE WITHOUT NOTICE TO YOU. AUTODESK PROVIDES SAMPLE CODE "AS IS" WITHOUT
'WARRANTY OF ANY KIND, WHETHER EXPRESS OR IMPLIED, INCLUDING WARRANTIES OF MERCHANTABILITY AND FITNESS
'FOR A PARTICULAR PURPOSE. IN NO EVENT SHALL AUTODESK OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER
'INCLUDING DIRECT, INDIRECT, INCIDENTAL, CONSEQUENTIAL, LOSS OF DATA, OR LOSS OF BUSINESS PROFITS OR
'SPECIAL DAMAGES, THAT MAY OCCUR AS A RESULT OF IMPLEMENTING OR USING ANY SAMPLE CODE, EVEN IF AUTODESK
'OR ITS SUPPLIERS HAVE BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. 


<Transaction(TransactionMode.Manual)> _
<Regeneration(RegenerationOption.Manual)> _
Public Class Commands
    Implements IExternalCommand
    'Open the files, and save them as new central files

    Public Function Execute(commandData As ExternalCommandData, ByRef message As String, elements As DB.ElementSet) As Result Implements IExternalCommand.Execute
        Dim count As Integer = 0
        Dim success As Integer = 0
        Dim wb As xl.Workbook = Nothing
        Dim files As List(Of RvtFile) = New List(Of RvtFile)

        'Clear the log
        Try
            If File.Exists(Path.GetTempPath & "SaveAsNewCentral.log") Then
                File.Delete(Path.GetTempPath & "SaveAsNewCentral.log")
            End If
        Catch ex As Exception
            'Swallow it
        End Try


        Try

            Dim app As Application = commandData.Application.Application

            LogThis("Starting Save As New Central")
            LogThis("Getting options and Excel data")

            Dim options As OptionsForm = New OptionsForm
            If options.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                Return Result.Cancelled
            End If

            LogThis("Options:" & vbCrLf & "  Excel File: " & options.ExcelFileName & vbCrLf & _
                    "  Sheet Name: " & options.SheetName & vbCrLf & _
                    "  Source Column: " & options.SourceColumn & vbCrLf & _
                    "  Destination Column: " & options.DestinationColumn & vbCrLf & _
                    "  Save to New Location: " & options.IsSavingToNewLocation.ToString & vbCrLf)

            LogThis("Opening Excel...")
            Dim objExcel As xl._Application = CreateObject("Excel.Application")
            'MsgBox(xlForm.FileName)
            wb = objExcel.Workbooks.Open(options.ExcelFileName)
            Dim sheet As xl.Worksheet = Nothing
            ' MsgBox(xlForm.SheetName)
            For Each sh In wb.Sheets
                If sh.Name = options.SheetName Then
                    sheet = sh
                    Exit For
                End If
            Next
            If sheet Is Nothing Then
                message = "Unable to access the correct sheet in Excel."
                wb.Close()
                Return Result.Failed
            End If

            LogThis("Getting file names from Excel...")
            Dim row As Integer = 2
            Dim emptyCellFound As Boolean = False
            Do Until emptyCellFound
                Try
                    'Get the source and destination filenames
                    Dim source As String = String.Empty
                    Dim dest As String = String.Empty
                    Dim rng As xl.Range = TryCast(sheet.Cells(row, options.SourceColumn), xl.Range)
                    source = rng.Value2
                    rng = TryCast(sheet.Cells(row, options.DestinationColumn), xl.Range)
                    dest = rng.Value2
                    If source = String.Empty Then
                        emptyCellFound = True
                        Exit Do
                    End If
                    files.Add(New RvtFile(source, dest, options.Suffix))
                    row += 1
                Catch ex As Exception
                    LogThis(vbCrLf & "Error getting the source and destination file names from the Excel file")
                    LogThis(vbCrLf & "Workbook name: " & options.ExcelFileName)
                    LogThis(vbCrLf & "Sheet name: " & options.SheetName)
                    LogThis(vbCrLf & "Row number: " & row)
                    LogThis(GetExceptionMessage(ex))
                End Try
            Loop
            wb.Close()

            For Each rvtFile As RvtFile In files
                count += 1
                LogThis(vbCrLf & "Starting file number " & count & ": " & rvtFile.Source)

                'Set Absolute Paths to Relative
                'If options.IsSettingAbsoluteToRelative Then
                '    LogThis("Getting linked files...")
                '    Dim linkCount As Integer = ListLinks(rvtFile.Source)
                '    'Set the options
                '    If linkCount Then
                '        LogThis(linkCount & " linked file(s) found, Absolute paths changed to Relative.")
                '    Else
                '        LogThis(" No linked files found.")
                '    End If
                'End If

                'Open the file or backup
                Dim openedDoc As Document

                If options.IsSavingToNewLocation Then
                    LogThis("Opening: " & rvtFile.Source)
                    openedDoc = OpenDetached(app, rvtFile.Source)
                Else
                    If BackupFile(rvtFile) Then
                        LogThis("Opening: " & rvtFile.BackupSource)
                        openedDoc = OpenDetached(app, rvtFile.BackupSource)
                    Else
                        Continue For
                    End If
                End If

                If openedDoc IsNot Nothing Then
                    'Save the document in a new location
                    Dim saveAsOptions As SaveAsOptions = New SaveAsOptions
                    saveAsOptions.OverwriteExistingFile = False
                    If openedDoc.IsWorkshared Then
                        Dim wsOptions As WorksharingSaveAsOptions = New WorksharingSaveAsOptions
                        wsOptions.SaveAsCentral = True
                        saveAsOptions.SetWorksharingOptions(wsOptions)

                        Try
                            If options.IsSavingToNewLocation Then
                                'Create the new destination folder, if it doesn't exist.
                                Dim di As DirectoryInfo = New DirectoryInfo(New FileInfo(rvtFile.Destination).DirectoryName)
                                If (Not di.Exists) Then
                                    LogThis("Creating folder that doesn't exist: " & di.FullName)
                                    di.Create()
                                End If
                                LogThis("Saving new central file to " & rvtFile.Destination)
                                openedDoc.SaveAs(rvtFile.Destination, saveAsOptions)
                            Else
                                LogThis("Saving new central file to " & rvtFile.Source)
                                openedDoc.SaveAs(rvtFile.Source, saveAsOptions)
                            End If
                            LogThis("Save successful. ")
                            success += 1
                        Catch ex As Exception
                            LogException(New Exception("Error saving " & rvtFile.Destination, ex))
                        End Try
                    Else
                        LogThis("Source file is NOT a Workshared document, skipping it. ")
                    End If

                    openedDoc.Close(False)

                End If

            Next

            LogThis(vbCrLf & success & " of " & count & " files saved to new locations.")
            MsgBox(success & " of " & count & " files saved to new locations." & vbCrLf & _
                   "Opening the log file...")
            System.Diagnostics.Process.Start("notepad.exe", Path.GetTempPath & "SaveAsNewCentral.log")

            Return Result.Succeeded
        Catch ex As Exception
            LogException(ex)
            message = "Unable to complete the task.  Opening the log file..."
            System.Diagnostics.Process.Start("notepad.exe", Path.GetTempPath & "SaveAsNewCentral.log")
            Return Result.Failed
        Finally
            LogThis("Complete.")
        End Try

    End Function

    Private Shared Sub LogThis(message As String)
        Dim temp As String = Path.GetTempPath
        Using outfile As New StreamWriter(temp & "SaveAsNewCentral.log", True)
            outfile.WriteLine(message)
        End Using
    End Sub

    Private Shared Sub LogException(exception As Exception)
        Dim log As String = "!! Exception !!"
        log += GetExceptionMessage(exception)
        LogThis(log)
    End Sub

    Private Shared Function GetExceptionMessage(ex As Exception) As String
        Dim log As String = vbCrLf & ex.Message
        If ex.InnerException IsNot Nothing Then
            log += GetExceptionMessage(ex.InnerException)
        End If
        Return log
    End Function

    Private Shared Function OpenDetached(application As Application, sourceFile As String) As Document
        Try
            If File.Exists(sourceFile) Then
                Dim sourcePath As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(sourceFile)
                Dim options As New OpenOptions()
                Dim wsoptions As New WorksetConfiguration
                'Don't open any worksets, so it opens faster, and prevents editing.
                wsoptions.CloseAll()

                'Open and Detach from Central so that it doesn't modify the file in any way.
                options.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
                options.SetOpenWorksetsConfiguration(wsoptions)
                Try
                    Dim openedDoc As Document = application.OpenDocumentFile(sourcePath, options)
                    Return openedDoc
                Catch ex As Exception
                    LogException(New Exception("Revit was not able to open the document", ex))
                End Try

            Else
                LogException(New Exception("File does not exist: " & sourceFile))
            End If
        Catch ex As Exception
            LogException(New Exception("Error opening " & sourceFile, ex))
        End Try

        Return Nothing

    End Function

    Private Function BackupFile(rvtFile As RvtFile) As Boolean
        ' Rename the file and delete the backup folder
        Try
            'Delete any files with the existing name.
            If File.Exists(rvtFile.BackupSource) Then
                LogThis("Deleting previous backup file: " & rvtFile.BackupSource)
                File.Delete(rvtFile.BackupSource)
            End If
            'rename the file
            LogThis("Renaming: " & rvtFile.Source & vbCrLf & "      to: " & rvtFile.BackupSource)
            File.Move(rvtFile.Source, rvtFile.BackupSource)

            'delete the backup folder
            If Directory.Exists(rvtFile.BackupSourceFolder) Then
                LogThis("Deleting backup folder: " & rvtFile.BackupSourceFolder)
                Directory.Delete(rvtFile.BackupSourceFolder, True)
            End If
            Return True
        Catch ex As Exception
            LogException(New Exception("Error backing up the original source file or deleting the backup folder for " & rvtFile.Source, ex))
            LogThis("File NOT saved.")
            Return False
        End Try
    End Function

    'Private Function ListLinks(location As String) As Integer

    '    'log += vbCrLf & "Getting Model Path for " & location.ToString
    '    Try
    '        Dim path As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(location)
    '        Dim linkedFiles As String = vbCrLf & vbCrLf & location
    '        Dim count As Integer = 0
    '        ' access transmission data in the given Revit file

    '        'log += vbCrLf & "Reading Transmission Data. "
    '        Dim transData As TransmissionData = TransmissionData.ReadTransmissionData(path)
    '        Dim externalReferences As ICollection(Of ElementId)

    '        If transData IsNot Nothing Then
    '            ' collect all (immediate) external references in the model

    '            externalReferences = transData.GetAllExternalFileReferenceIds()
    '            count = externalReferences.Count
    '            Dim abscount As Integer = 0

    '            If count > 0 Then
    '                'log += vbCrLf & "Linked files found.  Writing those with absolute file names to SaveAsNewCentral.LinkedFiles.csv"
    '                Dim isModified As Boolean = False
    '                For Each refId As ElementId In externalReferences
    '                    Dim extRef As ExternalFileReference = transData.GetLastSavedReferenceData(refId)
    '                    If extRef.IsValidObject Then
    '                        If extRef.ExternalFileReferenceType = ExternalFileReferenceType.CADLink Or extRef.ExternalFileReferenceType = ExternalFileReferenceType.DWFMarkup Or extRef.ExternalFileReferenceType = ExternalFileReferenceType.RevitLink Then
    '                            If extRef.PathType = PathType.Absolute Then
    '                                linkedFiles += vbTab & extRef.GetPath.ToString
    '                                'log += vbCrLf & "Linked file: " & extRef.GetPath.ToString

    '                                'TODO: Change any UNC paths to their new location.

    '                                Dim toLoad As Boolean = False
    '                                If extRef.GetLinkedFileStatus = LinkedFileStatus.Loaded Then
    '                                    toLoad = True
    '                                End If
    '                                transData.SetDesiredReferenceData(refId, extRef.GetAbsolutePath(), PathType.Relative, toLoad)
    '                                isModified = True
    '                                abscount += 1
    '                            End If
    '                        End If
    '                    End If
    '                Next

    '                'TODO: Check to see the effect on central files.

    '                If isModified Then
    '                    transData.IsTransmitted = True
    '                    TransmissionData.WriteTransmissionData(path, transData)
    '                    'log += vbCrLf & "Absolute paths have been set to Relative"
    '                End If

    '                Dim temp As String = System.IO.Path.GetTempPath
    '                Using outfile As New StreamWriter(temp & "SaveAsNewCentral.LinkedFiles.csv", True)
    '                    outfile.Write(linkedFiles)
    '                End Using
    '            Else
    '                'log += vbCrLf & "No external links found. "
    '            End If
    '        Else
    '            'log += vbCrLf & "No transmission data found. "
    '        End If

    '        Return count

    '    Catch ex As Exception
    '        'log += vbCrLf & "!! Exception !!"
    '        'log += vbCrLf & ex.Message
    '        If ex.InnerException IsNot Nothing Then
    '            GetExceptionMessage(ex.InnerException)
    '        End If
    '        Return 0
    '    Finally
    '        'log += vbCrLf & "End of LinkedFiles function. "
    '    End Try

    'End Function

End Class

<Transaction(TransactionMode.Manual)> _
<Regeneration(RegenerationOption.Manual)> _
Public Class UI
    Implements IExternalApplication

    Public Function OnShutdown(application As UIControlledApplication) As Result Implements IExternalApplication.OnShutdown
        'Clean up our dialog handler
        RemoveHandler application.DialogBoxShowing, AddressOf OnDialogBoxShowing

        Return Result.Succeeded
    End Function

    Public Function OnStartup(application As UIControlledApplication) As Result Implements IExternalApplication.OnStartup

        'Set up the button in the ribbon
        Dim path As String = System.Reflection.Assembly.GetExecutingAssembly().Location
        Dim caption As String = "Save As" & vbCrLf & "New Central"
        Dim d As New PushButtonData(caption, caption, path, "SaveAsNewCentral.Commands")
        d.AvailabilityClassName = "SaveAsNewCentral.Availability"
        Dim p As RibbonPanel = application.CreateRibbonPanel(caption)
        Dim b As PushButton = TryCast(p.AddItem(d), PushButton)
        b.ToolTip = "Saves a series of central files to a new location on the network"

        'Add an event handler for dialogs that pop up during opening of files.
        AddHandler application.DialogBoxShowing, AddressOf OnDialogBoxShowing

        Return Result.Succeeded

    End Function


    Private Sub OnDialogBoxShowing(sender As Object, e As DialogBoxShowingEventArgs)
        Dim e2 As TaskDialogShowingEventArgs = TryCast(e, TaskDialogShowingEventArgs)

        Select Case e2.DialogId

            'Let's only handle a few dialogs that show during the opening of a file.

            Case "TaskDialog_Loading_Transmitted_File"
                'This model has been transmitted from some other location. What do you want to do?
                e2.OverrideResult(1002)
                'Work with this model temporarily
            Case "TaskDialog_Audit_Warning"
                'Audit, This operation can take a long time...Do you want to continue?
                e2.OverrideResult(CInt(TaskDialogResult.Yes))
                'Yes
            Case "TaskDialog_Missing_Third_Party_Updater"
                'A revit extension is missing, what do you want to do?
                e2.OverrideResult(1001)
                'Continue working with the file
            Case "TaskDialog_Unresolved_References"
                'Revit could not find or read X number of referenced files. What do you want to do?
                e2.OverrideResult(1002)
                'Ignore and continue opening the project

        End Select

    End Sub

End Class


Public Class Availability
    Implements IExternalCommandAvailability
    Public Function IsCommandAvailable1(applicationData As UIApplication, selectedCategories As CategorySet) As Boolean Implements IExternalCommandAvailability.IsCommandAvailable
        Return True
    End Function
End Class



Public Class RvtFile

    Public Sub New(source As String, destination As String, suffix As String)
        Me.Source = source
        Me.Destination = destination
        Me.BackupSuffix = suffix
    End Sub

    Private _source As String
    Public Property Source() As String
        Get
            Return _source
        End Get
        Set(ByVal value As String)
            _source = value
            If _source.ToLower.EndsWith(".rvt") Then
                _backupSourceFolder = _source.Substring(0, _source.Length - 4) & "_backup"
            Else
                _backupSourceFolder = _source & "_backup"
            End If
        End Set
    End Property

    Private _destination As String
    Public Property Destination() As String
        Get
            Return _destination
        End Get
        Set(ByVal value As String)
            _destination = value
        End Set
    End Property

    Private _backupSource As String
    Public ReadOnly Property BackupSource() As String
        Get
            Return _backupSource
        End Get
    End Property

    Private _backupSourceFolder As String
    Public ReadOnly Property BackupSourceFolder() As String
        Get
            Return _backupSourceFolder
        End Get
    End Property

    Private _suffix As String
    Public WriteOnly Property BackupSuffix() As String
        Set(ByVal value As String)
            _suffix = value
            _backupSource = _source.Substring(0, _source.Length - 4) & _suffix & ".rvt"
        End Set
    End Property

End Class