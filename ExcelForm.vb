Imports System.IO
Imports xl = Microsoft.Office.Interop.Excel



Public Class ExcelForm

    Public Property FileName As String
    Public Property SheetName As String
    Public Property SourceColumn As String
    Public Property DestinationColumn As String

    Private Sub BrowseButton_Click(sender As Object, e As EventArgs) Handles BrowseButton.Click
        OpenFileDialog.ShowDialog()
    End Sub

    Private Sub OpenFileDialog_FileOk(sender As Object, e As ComponentModel.CancelEventArgs) Handles OpenFileDialog.FileOk
        ExcelFileName.Text = OpenFileDialog.FileName
        FileName = OpenFileDialog.FileName

        Dim objExcel As xl._Application = CreateObject("Excel.Application")
        Dim wb As xl.Workbook = objExcel.Workbooks.Open(OpenFileDialog.FileName)
        Dim shNames As List(Of String) = New List(Of String)
        For Each sh In wb.Sheets
            shNames.Add(sh.Name)
        Next
        wb.Close()
        Sheets.DataSource = shNames
        Sheets.SelectedIndex = 0
        SheetName = shNames(0)
        SourceColumn = "A"
        DestinationColumn = "B"
        StartButton.Enabled = True

    End Sub

    Private Sub ExcelFileName_TextChanged(sender As Object, e As EventArgs) Handles ExcelFileName.TextChanged
        If File.Exists(ExcelFileName.Text) Then
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
        StartButton.Enabled = Enabled
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
End Class