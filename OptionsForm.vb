Imports System.IO
Imports xl = Microsoft.Office.Interop.Excel



Public Class OptionsForm

    Public Property ExcelFileName As String
    Public Property SheetName As String
    Public Property SourceColumn As String
    Public Property DestinationColumn As String
    Public Property IsSavingToNewLocation As Boolean = True
    Public Property Suffix As String


    Private Sub BrowseButton_Click(sender As Object, e As EventArgs) Handles BrowseButton.Click
        OpenFileDialog.ShowDialog()
    End Sub

    Private Sub OpenFileDialog_FileOk(sender As Object, e As ComponentModel.CancelEventArgs) Handles OpenFileDialog.FileOk
        txtExcelFileName.Text = OpenFileDialog.FileName
        ExcelFileName = OpenFileDialog.FileName
        ToggleControls(File.Exists(ExcelFileName))
        Dim objExcel As xl._Application = CreateObject("Excel.Application")
        Dim wb As xl.Workbook = objExcel.Workbooks.Open(OpenFileDialog.FileName)
        Dim shNames As List(Of String) = New List(Of String)
        For Each sh In wb.Sheets
            shNames.Add(sh.Name)
        Next
        wb.Close()
        Sheets.DataSource = shNames
        Sheets.SelectedText = shNames(0)
        SheetName = shNames(0)
        SourceColumn = "A"
        DestinationColumn = "B"
        Sheets.Refresh()
    End Sub

    Private Sub ExcelFileName_TextChanged(sender As Object, e As EventArgs) Handles txtExcelFileName.TextChanged
        If File.Exists(txtExcelFileName.Text) Then
            ToggleControls(True)
        Else
            ToggleControls(False)
        End If
    End Sub

    Private Sub ToggleControls(Enabled As Boolean)
        LabelDestination.Enabled = Enabled
        LabelSheet.Enabled = Enabled
        LableSource.Enabled = Enabled
        Sheets.Enabled = Enabled
        SourceCol.Enabled = Enabled
        DestinationCol.Enabled = Enabled
        btnStart.Enabled = Enabled
        If Enabled Then
            Me.AcceptButton = btnStart
            btnStart.DialogResult = Windows.Forms.DialogResult.OK
        Else
            Me.AcceptButton = Nothing
            btnStart.DialogResult = Nothing
        End If

    End Sub

    Private Sub StartButton_Click(sender As Object, e As EventArgs)
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Hide()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Hide()
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        SourceCol.SelectedIndex = 0
        DestinationCol.SelectedIndex = 1
    End Sub

    Private Sub Sheets_SelectedValueChanged(sender As Object, e As EventArgs) Handles Sheets.SelectedValueChanged
        SheetName = Sheets.SelectedText
    End Sub

    Private Sub DestinationCol_SelectedValueChanged(sender As Object, e As EventArgs) Handles DestinationCol.SelectedValueChanged
        DestinationColumn = DestinationCol.SelectedText
    End Sub

    Private Sub SourceCol_SelectedValueChanged(sender As Object, e As EventArgs) Handles SourceCol.SelectedValueChanged
        SourceColumn = SourceCol.SelectedText
    End Sub

    Private Sub optOverwriteExisting_CheckedChanged(sender As Object, e As EventArgs) Handles optOverwriteExisting.CheckedChanged

        lblSuffix.Enabled = optOverwriteExisting.Checked
        txtSuffix.Enabled = optOverwriteExisting.Checked
        If File.Exists(txtExcelFileName.Text) Then
            LabelDestination.Enabled = Not optOverwriteExisting.Checked
            DestinationCol.Enabled = Not optOverwriteExisting.Checked
        End If

        '' Don't allow changing paths because it requires the file to be saved as a new name?
        'chkAbsToRel.Checked = Not optOverwriteExisting.Checked
        'chkAbsToRel.Enabled = Not optOverwriteExisting.Checked
    End Sub

    Private Sub txtSuffix_TextChanged(sender As Object, e As EventArgs) Handles txtSuffix.TextChanged
        If txtSuffix.Text = "_backup" Then
            MsgBox("Cannot use '_backup' because it conflicts with the existing backup folder name.", MsgBoxStyle.Information + vbOKOnly)
            btnStart.Enabled = False
        Else
            btnStart.Enabled = True
            Me.Suffix = txtSuffix.Text
        End If
    End Sub

    Private Sub optSaveToNewLocation_CheckedChanged(sender As Object, e As EventArgs) Handles optSaveToNewLocation.CheckedChanged
        IsSavingToNewLocation = optSaveToNewLocation.Checked
    End Sub

End Class