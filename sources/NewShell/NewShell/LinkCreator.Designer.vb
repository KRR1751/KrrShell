<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LinkCreator
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
        Me.components = New System.ComponentModel.Container
        Me.OK_Button = New System.Windows.Forms.Button
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.Label2 = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.cm = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.DirectoryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Label3 = New System.Windows.Forms.Label
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Controller = New System.Windows.Forms.Timer(Me.components)
        Me.cm.SuspendLayout()
        Me.SuspendLayout()
        '
        'OK_Button
        '
        Me.OK_Button.Enabled = False
        Me.OK_Button.Location = New System.Drawing.Point(180, 100)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(79, 23)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "Create Link"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(265, 100)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(161, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Path where Link will be created: "
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoEllipsis = True
        Me.LinkLabel1.Location = New System.Drawing.Point(161, 5)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(180, 13)
        Me.LinkLabel1.TabIndex = 3
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "(Path)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(262, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Pick a program path or a custom directory for the Link:"
        '
        'ComboBox1
        '
        Me.ComboBox1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystem
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(9, 34)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(289, 21)
        Me.ComboBox1.TabIndex = 5
        '
        'Button1
        '
        Me.Button1.ContextMenuStrip = Me.cm
        Me.Button1.Location = New System.Drawing.Point(299, 33)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(33, 23)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "..."
        Me.Button1.UseVisualStyleBackColor = True
        '
        'cm
        '
        Me.cm.AutoSize = False
        Me.cm.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DirectoryToolStripMenuItem, Me.FileToolStripMenuItem})
        Me.cm.Name = "cm"
        Me.cm.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.cm.ShowImageMargin = False
        Me.cm.Size = New System.Drawing.Size(100, 48)
        '
        'DirectoryToolStripMenuItem
        '
        Me.DirectoryToolStripMenuItem.Name = "DirectoryToolStripMenuItem"
        Me.DirectoryToolStripMenuItem.Size = New System.Drawing.Size(97, 22)
        Me.DirectoryToolStripMenuItem.Text = "Directory"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(97, 22)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Enabled = False
        Me.Label3.Location = New System.Drawing.Point(6, 58)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(59, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Link Label:"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Enabled = False
        Me.CheckBox1.Location = New System.Drawing.Point(71, 57)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(113, 17)
        Me.CheckBox1.TabIndex = 8
        Me.CheckBox1.Text = "Custom Link Label"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Enabled = False
        Me.TextBox1.Location = New System.Drawing.Point(9, 74)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(322, 20)
        Me.TextBox1.TabIndex = 9
        '
        'Controller
        '
        Me.Controller.Interval = 1
        '
        'LinkCreator
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(344, 131)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Cancel_Button)
        Me.Controls.Add(Me.OK_Button)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "LinkCreator"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "New Link"
        Me.cm.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents cm As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents DirectoryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Controller As System.Windows.Forms.Timer

End Class
