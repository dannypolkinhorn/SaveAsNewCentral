<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ExcelForm
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
        Me.ExcelFileName = New System.Windows.Forms.TextBox()
        Me.BrowseButton = New System.Windows.Forms.Button()
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Sheets = New System.Windows.Forms.ComboBox()
        Me.LabelSheet = New System.Windows.Forms.Label()
        Me.LableSource = New System.Windows.Forms.Label()
        Me.SourceCol = New System.Windows.Forms.ComboBox()
        Me.LabelDestination = New System.Windows.Forms.Label()
        Me.DestinationCol = New System.Windows.Forms.ComboBox()
        Me.CancelButton = New System.Windows.Forms.Button()
        Me.StartButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ExcelFileName
        '
        Me.ExcelFileName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ExcelFileName.Enabled = False
        Me.ExcelFileName.Location = New System.Drawing.Point(155, 7)
        Me.ExcelFileName.Name = "ExcelFileName"
        Me.ExcelFileName.Size = New System.Drawing.Size(340, 22)
        Me.ExcelFileName.TabIndex = 0
        '
        'BrowseButton
        '
        Me.BrowseButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BrowseButton.Location = New System.Drawing.Point(72, 7)
        Me.BrowseButton.Name = "BrowseButton"
        Me.BrowseButton.Size = New System.Drawing.Size(77, 23)
        Me.BrowseButton.TabIndex = 1
        Me.BrowseButton.Text = "Browse"
        Me.BrowseButton.UseVisualStyleBackColor = True
        '
        'OpenFileDialog
        '
        Me.OpenFileDialog.DefaultExt = "xlsx"
        Me.OpenFileDialog.Filter = "Excel files|*.xlsx;*.xls"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(54, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Excel file:"
        '
        'Sheets
        '
        Me.Sheets.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Sheets.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Sheets.Enabled = False
        Me.Sheets.FormattingEnabled = True
        Me.Sheets.Location = New System.Drawing.Point(155, 35)
        Me.Sheets.Name = "Sheets"
        Me.Sheets.Size = New System.Drawing.Size(340, 21)
        Me.Sheets.TabIndex = 3
        '
        'LabelSheet
        '
        Me.LabelSheet.AutoSize = True
        Me.LabelSheet.Enabled = False
        Me.LabelSheet.Location = New System.Drawing.Point(12, 38)
        Me.LabelSheet.Name = "LabelSheet"
        Me.LabelSheet.Size = New System.Drawing.Size(39, 13)
        Me.LabelSheet.TabIndex = 4
        Me.LabelSheet.Text = "Sheet:"
        '
        'LableSource
        '
        Me.LableSource.AutoSize = True
        Me.LableSource.Enabled = False
        Me.LableSource.Location = New System.Drawing.Point(12, 65)
        Me.LableSource.Name = "LableSource"
        Me.LableSource.Size = New System.Drawing.Size(86, 13)
        Me.LableSource.TabIndex = 6
        Me.LableSource.Text = "Source column:"
        '
        'SourceCol
        '
        Me.SourceCol.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.SourceCol.Enabled = False
        Me.SourceCol.FormattingEnabled = True
        Me.SourceCol.Items.AddRange(New Object() {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"})
        Me.SourceCol.Location = New System.Drawing.Point(155, 62)
        Me.SourceCol.Name = "SourceCol"
        Me.SourceCol.Size = New System.Drawing.Size(92, 21)
        Me.SourceCol.TabIndex = 5
        '
        'LabelDestination
        '
        Me.LabelDestination.AutoSize = True
        Me.LabelDestination.Enabled = False
        Me.LabelDestination.Location = New System.Drawing.Point(12, 95)
        Me.LabelDestination.Name = "LabelDestination"
        Me.LabelDestination.Size = New System.Drawing.Size(111, 13)
        Me.LabelDestination.TabIndex = 8
        Me.LabelDestination.Text = "Destination column:"
        '
        'DestinationCol
        '
        Me.DestinationCol.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.DestinationCol.Enabled = False
        Me.DestinationCol.Items.AddRange(New Object() {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"})
        Me.DestinationCol.Location = New System.Drawing.Point(155, 92)
        Me.DestinationCol.MaxDropDownItems = 20
        Me.DestinationCol.Name = "DestinationCol"
        Me.DestinationCol.Size = New System.Drawing.Size(92, 21)
        Me.DestinationCol.Sorted = True
        Me.DestinationCol.TabIndex = 7
        '
        'CancelButton
        '
        Me.CancelButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelButton.Location = New System.Drawing.Point(304, 119)
        Me.CancelButton.Name = "CancelButton"
        Me.CancelButton.Size = New System.Drawing.Size(72, 28)
        Me.CancelButton.TabIndex = 12
        Me.CancelButton.Text = "Cancel"
        Me.CancelButton.UseVisualStyleBackColor = True
        '
        'StartButton
        '
        Me.StartButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.StartButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.StartButton.Enabled = False
        Me.StartButton.Location = New System.Drawing.Point(382, 119)
        Me.StartButton.Name = "StartButton"
        Me.StartButton.Size = New System.Drawing.Size(113, 28)
        Me.StartButton.TabIndex = 11
        Me.StartButton.Text = "Start >"
        Me.StartButton.UseVisualStyleBackColor = True
        '
        'ExcelForm
        '
        Me.AcceptButton = Me.StartButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(507, 159)
        Me.ControlBox = False
        Me.Controls.Add(Me.CancelButton)
        Me.Controls.Add(Me.StartButton)
        Me.Controls.Add(Me.LabelDestination)
        Me.Controls.Add(Me.DestinationCol)
        Me.Controls.Add(Me.LableSource)
        Me.Controls.Add(Me.SourceCol)
        Me.Controls.Add(Me.LabelSheet)
        Me.Controls.Add(Me.Sheets)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.BrowseButton)
        Me.Controls.Add(Me.ExcelFileName)
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MinimumSize = New System.Drawing.Size(287, 197)
        Me.Name = "ExcelForm"
        Me.Text = "Source and Destination Data File"
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents ExcelFileName As System.Windows.Forms.TextBox
    Friend WithEvents BrowseButton As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Sheets As System.Windows.Forms.ComboBox
    Friend WithEvents LabelSheet As System.Windows.Forms.Label
    Friend WithEvents LableSource As System.Windows.Forms.Label
    Friend WithEvents SourceCol As System.Windows.Forms.ComboBox
    Friend WithEvents LabelDestination As System.Windows.Forms.Label
    Friend WithEvents DestinationCol As System.Windows.Forms.ComboBox
    Friend WithEvents CancelButton As System.Windows.Forms.Button
    Friend WithEvents StartButton As System.Windows.Forms.Button
End Class
