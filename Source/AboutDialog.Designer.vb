<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AboutDialog
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

    Friend WithEvents TableLayoutPanel As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents LabelProductName As System.Windows.Forms.Label
    Friend WithEvents LabelVersion As System.Windows.Forms.Label
    Friend WithEvents LabelCompanyName As System.Windows.Forms.Label
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents LabelCopyright As System.Windows.Forms.Label

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AboutDialog))
        TableLayoutPanel = New TableLayoutPanel()
        LabelProductName = New Label()
        LabelVersion = New Label()
        LabelCopyright = New Label()
        LabelCompanyName = New Label()
        OKButton = New Button()
        PictureBox1 = New PictureBox()
        Label1 = New Label()
        TextBoxDescription = New Label()
        CheckBox1 = New CheckBox()
        TableLayoutPanel.SuspendLayout()
        CType(PictureBox1, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' TableLayoutPanel
        ' 
        TableLayoutPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        TableLayoutPanel.ColumnCount = 1
        TableLayoutPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100F))
        TableLayoutPanel.Controls.Add(LabelProductName, 1, 0)
        TableLayoutPanel.Controls.Add(LabelVersion, 1, 1)
        TableLayoutPanel.Controls.Add(LabelCopyright, 1, 2)
        TableLayoutPanel.Controls.Add(LabelCompanyName, 1, 3)
        TableLayoutPanel.ForeColor = SystemColors.HighlightText
        TableLayoutPanel.Location = New Point(133, 10)
        TableLayoutPanel.Margin = New Padding(4, 3, 4, 3)
        TableLayoutPanel.Name = "TableLayoutPanel"
        TableLayoutPanel.RowCount = 4
        TableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 25F))
        TableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 25F))
        TableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 25F))
        TableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 25F))
        TableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 23F))
        TableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 23F))
        TableLayoutPanel.Size = New Size(304, 112)
        TableLayoutPanel.TabIndex = 0
        ' 
        ' LabelProductName
        ' 
        LabelProductName.Dock = DockStyle.Fill
        LabelProductName.Location = New Point(7, 0)
        LabelProductName.Margin = New Padding(7, 0, 4, 0)
        LabelProductName.MaximumSize = New Size(0, 20)
        LabelProductName.Name = "LabelProductName"
        LabelProductName.Size = New Size(293, 20)
        LabelProductName.TabIndex = 0
        LabelProductName.Text = "Product Name"
        LabelProductName.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' LabelVersion
        ' 
        LabelVersion.Dock = DockStyle.Fill
        LabelVersion.Location = New Point(7, 28)
        LabelVersion.Margin = New Padding(7, 0, 4, 0)
        LabelVersion.MaximumSize = New Size(0, 20)
        LabelVersion.Name = "LabelVersion"
        LabelVersion.Size = New Size(293, 20)
        LabelVersion.TabIndex = 0
        LabelVersion.Text = "Version"
        LabelVersion.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' LabelCopyright
        ' 
        LabelCopyright.Dock = DockStyle.Fill
        LabelCopyright.Location = New Point(7, 56)
        LabelCopyright.Margin = New Padding(7, 0, 4, 0)
        LabelCopyright.MaximumSize = New Size(0, 20)
        LabelCopyright.Name = "LabelCopyright"
        LabelCopyright.Size = New Size(293, 20)
        LabelCopyright.TabIndex = 0
        LabelCopyright.Text = "Copyright"
        LabelCopyright.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' LabelCompanyName
        ' 
        LabelCompanyName.Dock = DockStyle.Fill
        LabelCompanyName.Location = New Point(7, 84)
        LabelCompanyName.Margin = New Padding(7, 0, 4, 0)
        LabelCompanyName.MaximumSize = New Size(0, 20)
        LabelCompanyName.Name = "LabelCompanyName"
        LabelCompanyName.Size = New Size(293, 20)
        LabelCompanyName.TabIndex = 0
        LabelCompanyName.Text = "Company Name"
        LabelCompanyName.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' OKButton
        ' 
        OKButton.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        OKButton.BackColor = Color.Black
        OKButton.DialogResult = DialogResult.Cancel
        OKButton.FlatStyle = FlatStyle.Flat
        OKButton.ForeColor = SystemColors.HighlightText
        OKButton.Location = New Point(351, 278)
        OKButton.Margin = New Padding(4, 3, 4, 3)
        OKButton.Name = "OKButton"
        OKButton.Size = New Size(88, 27)
        OKButton.TabIndex = 0
        OKButton.Text = "&OK"
        OKButton.UseVisualStyleBackColor = False
        ' 
        ' PictureBox1
        ' 
        PictureBox1.Image = My.Resources.Resources.StartRight
        PictureBox1.Location = New Point(15, 10)
        PictureBox1.Margin = New Padding(4, 3, 4, 3)
        PictureBox1.Name = "PictureBox1"
        PictureBox1.Size = New Size(111, 112)
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        PictureBox1.TabIndex = 1
        PictureBox1.TabStop = False
        ' 
        ' Label1
        ' 
        Label1.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        Label1.AutoSize = True
        Label1.Location = New Point(14, 284)
        Label1.Margin = New Padding(4, 0, 4, 0)
        Label1.Name = "Label1"
        Label1.Size = New Size(155, 15)
        Label1.TabIndex = 2
        Label1.Text = "Made with <3 by @KRR1751"
        ' 
        ' TextBoxDescription
        ' 
        TextBoxDescription.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        TextBoxDescription.ForeColor = SystemColors.HighlightText
        TextBoxDescription.Location = New Point(15, 129)
        TextBoxDescription.Margin = New Padding(4, 0, 4, 0)
        TextBoxDescription.Name = "TextBoxDescription"
        TextBoxDescription.Size = New Size(422, 142)
        TextBoxDescription.TabIndex = 3
        TextBoxDescription.Text = resources.GetString("TextBoxDescription.Text")
        ' 
        ' CheckBox1
        ' 
        CheckBox1.AutoSize = True
        CheckBox1.ForeColor = SystemColors.HighlightText
        CheckBox1.Location = New Point(15, 283)
        CheckBox1.Margin = New Padding(4, 3, 4, 3)
        CheckBox1.Name = "CheckBox1"
        CheckBox1.Size = New Size(153, 19)
        CheckBox1.TabIndex = 4
        CheckBox1.Text = "Show normal BackColor"
        CheckBox1.UseVisualStyleBackColor = True
        ' 
        ' AboutDialog
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        BackColor = Color.Black
        CancelButton = OKButton
        ClientSize = New Size(450, 318)
        Controls.Add(CheckBox1)
        Controls.Add(TextBoxDescription)
        Controls.Add(Label1)
        Controls.Add(PictureBox1)
        Controls.Add(TableLayoutPanel)
        Controls.Add(OKButton)
        Font = New Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        ForeColor = SystemColors.HighlightText
        FormBorderStyle = FormBorderStyle.FixedDialog
        Margin = New Padding(4, 3, 4, 3)
        MaximizeBox = False
        MinimizeBox = False
        Name = "AboutDialog"
        Padding = New Padding(10)
        ShowInTaskbar = False
        StartPosition = FormStartPosition.CenterParent
        Text = "About Shell"
        TableLayoutPanel.ResumeLayout(False)
        CType(PictureBox1, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()

    End Sub

    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBoxDescription As Label
    Friend WithEvents CheckBox1 As CheckBox
End Class
