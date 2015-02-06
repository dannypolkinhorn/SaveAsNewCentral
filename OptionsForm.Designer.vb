<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OptionsForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(OptionsForm))
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.LabelDestination = New System.Windows.Forms.Label()
        Me.DestinationCol = New System.Windows.Forms.ComboBox()
        Me.LableSource = New System.Windows.Forms.Label()
        Me.SourceCol = New System.Windows.Forms.ComboBox()
        Me.LabelSheet = New System.Windows.Forms.Label()
        Me.Sheets = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.BrowseButton = New System.Windows.Forms.Button()
        Me.txtExcelFileName = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lblSuffix = New System.Windows.Forms.Label()
        Me.txtSuffix = New System.Windows.Forms.TextBox()
        Me.lblOption2 = New System.Windows.Forms.Label()
        Me.lblOption1 = New System.Windows.Forms.Label()
        Me.optOverwriteExisting = New System.Windows.Forms.RadioButton()
        Me.optSaveToNewLocation = New System.Windows.Forms.RadioButton()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'OpenFileDialog
        '
        Me.OpenFileDialog.DefaultExt = "xlsx"
        Me.OpenFileDialog.Filter = "Excel files|*.xlsx;*.xls"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.LabelDestination)
        Me.GroupBox1.Controls.Add(Me.DestinationCol)
        Me.GroupBox1.Controls.Add(Me.LableSource)
        Me.GroupBox1.Controls.Add(Me.SourceCol)
        Me.GroupBox1.Controls.Add(Me.LabelSheet)
        Me.GroupBox1.Controls.Add(Me.Sheets)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.BrowseButton)
        Me.GroupBox1.Controls.Add(Me.txtExcelFileName)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(475, 142)
        Me.GroupBox1.TabIndex = 13
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Excel Options"
        '
        'LabelDestination
        '
        Me.LabelDestination.AutoSize = True
        Me.LabelDestination.Enabled = False
        Me.LabelDestination.Location = New System.Drawing.Point(13, 109)
        Me.LabelDestination.Name = "LabelDestination"
        Me.LabelDestination.Size = New System.Drawing.Size(111, 13)
        Me.LabelDestination.TabIndex = 17
        Me.LabelDestination.Text = "Destination column:"
        '
        'DestinationCol
        '
        Me.DestinationCol.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.DestinationCol.Enabled = False
        Me.DestinationCol.Items.AddRange(New Object() {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"})
        Me.DestinationCol.Location = New System.Drawing.Point(156, 106)
        Me.DestinationCol.MaxDropDownItems = 20
        Me.DestinationCol.Name = "DestinationCol"
        Me.DestinationCol.Size = New System.Drawing.Size(92, 21)
        Me.DestinationCol.Sorted = True
        Me.DestinationCol.TabIndex = 16
        '
        'LableSource
        '
        Me.LableSource.AutoSize = True
        Me.LableSource.Enabled = False
        Me.LableSource.Location = New System.Drawing.Point(13, 79)
        Me.LableSource.Name = "LableSource"
        Me.LableSource.Size = New System.Drawing.Size(86, 13)
        Me.LableSource.TabIndex = 15
        Me.LableSource.Text = "Source column:"
        '
        'SourceCol
        '
        Me.SourceCol.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.SourceCol.Enabled = False
        Me.SourceCol.FormattingEnabled = True
        Me.SourceCol.Items.AddRange(New Object() {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"})
        Me.SourceCol.Location = New System.Drawing.Point(156, 76)
        Me.SourceCol.Name = "SourceCol"
        Me.SourceCol.Size = New System.Drawing.Size(92, 21)
        Me.SourceCol.TabIndex = 14
        '
        'LabelSheet
        '
        Me.LabelSheet.AutoSize = True
        Me.LabelSheet.Enabled = False
        Me.LabelSheet.Location = New System.Drawing.Point(13, 52)
        Me.LabelSheet.Name = "LabelSheet"
        Me.LabelSheet.Size = New System.Drawing.Size(39, 13)
        Me.LabelSheet.TabIndex = 13
        Me.LabelSheet.Text = "Sheet:"
        '
        'Sheets
        '
        Me.Sheets.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Sheets.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Sheets.Enabled = False
        Me.Sheets.FormattingEnabled = True
        Me.Sheets.Location = New System.Drawing.Point(156, 49)
        Me.Sheets.Name = "Sheets"
        Me.Sheets.Size = New System.Drawing.Size(313, 21)
        Me.Sheets.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(54, 13)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Excel file:"
        '
        'BrowseButton
        '
        Me.BrowseButton.Location = New System.Drawing.Point(74, 20)
        Me.BrowseButton.Name = "BrowseButton"
        Me.BrowseButton.Size = New System.Drawing.Size(77, 23)
        Me.BrowseButton.TabIndex = 10
        Me.BrowseButton.Text = "Browse..."
        Me.BrowseButton.UseVisualStyleBackColor = True
        '
        'txtExcelFileName
        '
        Me.txtExcelFileName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtExcelFileName.Enabled = False
        Me.txtExcelFileName.Location = New System.Drawing.Point(156, 21)
        Me.txtExcelFileName.Name = "txtExcelFileName"
        Me.txtExcelFileName.Size = New System.Drawing.Size(313, 22)
        Me.txtExcelFileName.TabIndex = 9
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.lblSuffix)
        Me.GroupBox2.Controls.Add(Me.txtSuffix)
        Me.GroupBox2.Controls.Add(Me.lblOption2)
        Me.GroupBox2.Controls.Add(Me.lblOption1)
        Me.GroupBox2.Controls.Add(Me.optOverwriteExisting)
        Me.GroupBox2.Controls.Add(Me.optSaveToNewLocation)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 160)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(475, 214)
        Me.GroupBox2.TabIndex = 14
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Save Options"
        '
        'lblSuffix
        '
        Me.lblSuffix.AutoSize = True
        Me.lblSuffix.Enabled = False
        Me.lblSuffix.Location = New System.Drawing.Point(48, 148)
        Me.lblSuffix.Name = "lblSuffix"
        Me.lblSuffix.Size = New System.Drawing.Size(36, 13)
        Me.lblSuffix.TabIndex = 5
        Me.lblSuffix.Text = "Suffix"
        '
        'txtSuffix
        '
        Me.txtSuffix.Enabled = False
        Me.txtSuffix.Location = New System.Drawing.Point(94, 145)
        Me.txtSuffix.Name = "txtSuffix"
        Me.txtSuffix.Size = New System.Drawing.Size(100, 22)
        Me.txtSuffix.TabIndex = 4
        Me.txtSuffix.Text = "_migration"
        '
        'lblOption2
        '
        Me.lblOption2.Location = New System.Drawing.Point(210, 124)
        Me.lblOption2.Name = "lblOption2"
        Me.lblOption2.Size = New System.Drawing.Size(259, 87)
        Me.lblOption2.TabIndex = 3
        Me.lblOption2.Text = "Renames the existing Central file in the Source column with a suffix, deletes the" & _
    " Central file's backup folder and all backups, then saves the Central as the ori" & _
    "ginal name in the original location."
        '
        'lblOption1
        '
        Me.lblOption1.Location = New System.Drawing.Point(210, 23)
        Me.lblOption1.Name = "lblOption1"
        Me.lblOption1.Size = New System.Drawing.Size(259, 88)
        Me.lblOption1.TabIndex = 2
        Me.lblOption1.Text = resources.GetString("lblOption1.Text")
        '
        'optOverwriteExisting
        '
        Me.optOverwriteExisting.AutoSize = True
        Me.optOverwriteExisting.Location = New System.Drawing.Point(16, 122)
        Me.optOverwriteExisting.Name = "optOverwriteExisting"
        Me.optOverwriteExisting.Size = New System.Drawing.Size(178, 17)
        Me.optOverwriteExisting.TabIndex = 1
        Me.optOverwriteExisting.Text = "Overwrite in Existing Location"
        Me.optOverwriteExisting.UseVisualStyleBackColor = True
        '
        'optSaveToNewLocation
        '
        Me.optSaveToNewLocation.AutoSize = True
        Me.optSaveToNewLocation.Checked = True
        Me.optSaveToNewLocation.Location = New System.Drawing.Point(16, 21)
        Me.optSaveToNewLocation.Name = "optSaveToNewLocation"
        Me.optSaveToNewLocation.Size = New System.Drawing.Size(135, 17)
        Me.optSaveToNewLocation.TabIndex = 0
        Me.optSaveToNewLocation.TabStop = True
        Me.optSaveToNewLocation.Text = "Save to New Location"
        Me.optSaveToNewLocation.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Enabled = False
        Me.btnStart.Location = New System.Drawing.Point(374, 380)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(113, 28)
        Me.btnStart.TabIndex = 18
        Me.btnStart.Text = "Start >"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(296, 380)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(72, 28)
        Me.btnCancel.TabIndex = 17
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'OptionsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(499, 420)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(515, 458)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(515, 458)
        Me.Name = "OptionsForm"
        Me.Text = "Central File Save Options"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

End Sub
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents LabelDestination As System.Windows.Forms.Label
    Friend WithEvents DestinationCol As System.Windows.Forms.ComboBox
    Friend WithEvents LableSource As System.Windows.Forms.Label
    Friend WithEvents SourceCol As System.Windows.Forms.ComboBox
    Friend WithEvents LabelSheet As System.Windows.Forms.Label
    Friend WithEvents Sheets As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents BrowseButton As System.Windows.Forms.Button
    Friend WithEvents txtExcelFileName As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblSuffix As System.Windows.Forms.Label
    Friend WithEvents txtSuffix As System.Windows.Forms.TextBox
    Friend WithEvents lblOption2 As System.Windows.Forms.Label
    Friend WithEvents lblOption1 As System.Windows.Forms.Label
    Friend WithEvents optOverwriteExisting As System.Windows.Forms.RadioButton
    Friend WithEvents optSaveToNewLocation As System.Windows.Forms.RadioButton
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
End Class
