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
        Dim log As String = ""
        Dim count As Integer = 0
        Dim success As Integer = 0
        Dim wb As xl.Workbook = Nothing

        Try

            Dim app As Application = commandData.Application.Application
            Dim source As String = String.Empty
            Dim dest As String = String.Empty

            log = "Getting Excel data"

            Dim xlForm As ExcelForm = New ExcelForm
            If xlForm.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                Return Result.Cancelled
            End If

            Dim objExcel As xl._Application = CreateObject("Excel.Application")
            'MsgBox(xlForm.FileName)
            wb = objExcel.Workbooks.Open(xlForm.FileName)
            Dim sheet As xl.Worksheet = Nothing
            ' MsgBox(xlForm.SheetName)
            For Each sh In wb.Sheets
                If sh.Name = xlForm.SheetName Then
                    sheet = sh
                    Exit For
                End If
            Next
            If sheet Is Nothing Then
                message = "Unable to access the correct sheet in Excel."
                wb.Close()
                Return Result.Failed
            End If

            Dim row As Integer = 2
            Dim emptyCellFound As Boolean = False
            log += vbCrLf & "Starting to save files in their new locations"
            Do Until emptyCellFound

                Try
                    'Get the source and destination filenames
                    Dim rng As xl.Range = TryCast(sheet.Cells(row, xlForm.SourceColumn), xl.Range)
                    source = rng.Value2
                    rng = TryCast(sheet.Cells(row, xlForm.DestinationColumn), xl.Range)
                    dest = rng.Value2
                    If source = String.Empty Or dest = String.Empty Then
                        emptyCellFound = True
                        Exit Do
                    End If
                    row += 1
                Catch ex As Exception
                    log += vbCrLf & "Error getting the source and destination file names from the Excel file"
                    log += vbCrLf & "Workbook name: " & xlForm.FileName
                    log += vbCrLf & "Sheet name: " & xlForm.SheetName
                    log += vbCrLf & "Row number: " & row
                    log += LogException(ex)
                End Try

                Try
                    log += vbCrLf & vbCrLf & "Opening file number " & count + 1 & ": " & source
                    Dim openedDoc As Document = OpenDetached(app, source)
                    If openedDoc IsNot Nothing Then
                        ''list the linked files
                        'log += vbCrLf & "Getting linked files..."
                        'Dim linkCount As Integer = ListLinks(ModelPathUtils.ConvertUserVisiblePathToModelPath(openedDoc.PathName), log)
                        'If linkCount Then
                        '    log += vbCrLf & linkCount & " linked file(s) with Absolute paths found.  They are listed in C:\Users\<UserName>\AppData\Local\Temp\SaveAsNewCentral.LinkedFiles.log"
                        'Else
                        '    log += " None found."
                        'End If


                        'Save the document in a new location
                        Dim saveAsOptions As SaveAsOptions = New SaveAsOptions
                        saveAsOptions.OverwriteExistingFile = False
                        If openedDoc.IsWorkshared Then
                            Dim wsOptions As WorksharingSaveAsOptions = New WorksharingSaveAsOptions
                            wsOptions.SaveAsCentral = True
                            saveAsOptions.SetWorksharingOptions(wsOptions)
                            log += vbCrLf & "Source file is a Workshared document. "
                        End If
                        Try
                            log += vbCrLf & "Saving to " & dest
                            openedDoc.SaveAs(dest, saveAsOptions)
                        Catch ex As Exception
                            Throw New Exception("Error saving " & dest, ex)
                        End Try
                        openedDoc.Close(False)
                    End If

                    success += 1



                Catch ex As Exception
                    Throw New Exception("Error opening or saving the Revit file", ex)
                End Try


                count += 1
            Loop
            log += vbCrLf & vbCrLf & success & " of " & count & " files saved to new locations."

            Return Result.Succeeded
        Catch ex As Exception
            log += vbCrLf & LogException(ex)
            message = "Unable to complete all files.  See the log file in C:\Users\<UserName>\AppData\Local\Temp\SaveAsNewCentral.log."
            Return Result.Failed
        Finally
            If wb IsNot Nothing Then
                wb.Close()
            End If
            Dim temp As String = Path.GetTempPath
            Using outfile As New StreamWriter(temp & "SaveAsNewCentral.log")
                outfile.Write(log)
            End Using
        End Try

    End Function

    Private Function LogException(ex As Exception)
        Dim log As String = ""
        log += vbCrLf & "!! Exception !!"
        log += vbCrLf & ex.Message
        If ex.InnerException IsNot Nothing Then
            LogException(ex.InnerException)
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
                    Throw New Exception("Revit was not able to open the document", ex)
                End Try

            Else
                Throw New Exception("File does not exist: " & sourceFile)
            End If
        Catch ex As Exception
            Throw New Exception("Error opening " & sourceFile, ex)
        End Try

        Return Nothing

    End Function

    'Private Function ListLinks(location As ModelPath, ByRef log As String) As Integer

    '    log += vbCrLf & "Getting Model Path for " & location.ToString
    '    Try
    '        Dim path As String = ModelPathUtils.ConvertModelPathToUserVisiblePath(location)
    '        Dim linkedFiles As String = vbCrLf & vbCrLf & path
    '        Dim count As Integer = 0
    '        ' access transmission data in the given Revit file

    '        log += vbCrLf & "Reading Transmission Data. "
    '        Dim transData As TransmissionData = TransmissionData.ReadTransmissionData(location)
    '        Dim externalReferences As ICollection(Of ElementId)

    '        If transData IsNot Nothing Then
    '            ' collect all (immediate) external references in the model

    '            externalReferences = transData.GetAllExternalFileReferenceIds()
    '            count = externalReferences.Count
    '            Dim abscount As Integer = 0

    '            If count > 0 Then
    '                log += vbCrLf & "Linked files found.  Writing those with absolute file names to SaveAsNewCentral.LinkedFiles.csv"
    '                For Each refId As ElementId In externalReferences
    '                    Dim extRef As ExternalFileReference = transData.GetLastSavedReferenceData(refId)
    '                    If extRef.IsValidObject Then
    '                        If extRef.PathType = PathType.Absolute Then
    '                            linkedFiles += vbTab & extRef.GetPath.ToString
    '                            log += vbCrLf & "Linked file: " & extRef.GetPath.ToString
    '                            abscount += 1
    '                        End If
    '                    End If
    '                Next

    '                Dim temp As String = System.IO.Path.GetTempPath
    '                Using outfile As New StreamWriter(temp & "SaveAsNewCentral.LinkedFiles.csv", True)
    '                    outfile.Write(linkedFiles)
    '                End Using
    '            Else
    '                log += vbCrLf & "No external links found. "
    '            End If
    '        Else
    '            log += vbCrLf & "No transmission data found. "
    '        End If

    '        Return count

    '    Catch ex As Exception
    '        log += vbCrLf & "!! Exception !!"
    '        log += vbCrLf & ex.Message
    '        If ex.InnerException IsNot Nothing Then
    '            LogException(ex.InnerException)
    '        End If
    '        Return 0
    '    Finally
    '        log += vbCrLf & "End of LinkedFiles function. "
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



