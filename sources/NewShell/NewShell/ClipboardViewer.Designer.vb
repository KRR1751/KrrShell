<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ClipboardViewer
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
        Me.components = New System.ComponentModel.Container()
        Me.Controller = New System.Windows.Forms.Timer(Me.components)
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ClipboardToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReloadToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ClearContentToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HistoryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StoreClipboardHistoryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RemoveHistoryAfterEndToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.WindowToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HideToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OnTopToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.TSSL1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.TSSL2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ReloadInfo = New System.Windows.Forms.Timer(Me.components)
        Me.MenuStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Controller
        '
        Me.Controller.Interval = 500
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClipboardToolStripMenuItem, Me.HistoryToolStripMenuItem, Me.WindowToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(587, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ClipboardToolStripMenuItem
        '
        Me.ClipboardToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReloadToolStripMenuItem, Me.ClearContentToolStripMenuItem})
        Me.ClipboardToolStripMenuItem.Name = "ClipboardToolStripMenuItem"
        Me.ClipboardToolStripMenuItem.Size = New System.Drawing.Size(71, 20)
        Me.ClipboardToolStripMenuItem.Text = "Clipboard"
        '
        'ReloadToolStripMenuItem
        '
        Me.ReloadToolStripMenuItem.Name = "ReloadToolStripMenuItem"
        Me.ReloadToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.ReloadToolStripMenuItem.Text = "Reload"
        '
        'ClearContentToolStripMenuItem
        '
        Me.ClearContentToolStripMenuItem.Name = "ClearContentToolStripMenuItem"
        Me.ClearContentToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.ClearContentToolStripMenuItem.Text = "Clear content"
        '
        'HistoryToolStripMenuItem
        '
        Me.HistoryToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StoreClipboardHistoryToolStripMenuItem, Me.RemoveHistoryAfterEndToolStripMenuItem})
        Me.HistoryToolStripMenuItem.Name = "HistoryToolStripMenuItem"
        Me.HistoryToolStripMenuItem.Size = New System.Drawing.Size(57, 20)
        Me.HistoryToolStripMenuItem.Text = "History"
        '
        'StoreClipboardHistoryToolStripMenuItem
        '
        Me.StoreClipboardHistoryToolStripMenuItem.CheckOnClick = True
        Me.StoreClipboardHistoryToolStripMenuItem.Enabled = False
        Me.StoreClipboardHistoryToolStripMenuItem.Name = "StoreClipboardHistoryToolStripMenuItem"
        Me.StoreClipboardHistoryToolStripMenuItem.Size = New System.Drawing.Size(215, 22)
        Me.StoreClipboardHistoryToolStripMenuItem.Text = "Store Clipboard History?"
        '
        'RemoveHistoryAfterEndToolStripMenuItem
        '
        Me.RemoveHistoryAfterEndToolStripMenuItem.Checked = True
        Me.RemoveHistoryAfterEndToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.RemoveHistoryAfterEndToolStripMenuItem.Enabled = False
        Me.RemoveHistoryAfterEndToolStripMenuItem.Name = "RemoveHistoryAfterEndToolStripMenuItem"
        Me.RemoveHistoryAfterEndToolStripMenuItem.Size = New System.Drawing.Size(215, 22)
        Me.RemoveHistoryAfterEndToolStripMenuItem.Text = "Remove History After End?"
        '
        'WindowToolStripMenuItem
        '
        Me.WindowToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.HideToolStripMenuItem, Me.OnTopToolStripMenuItem})
        Me.WindowToolStripMenuItem.Name = "WindowToolStripMenuItem"
        Me.WindowToolStripMenuItem.Size = New System.Drawing.Size(63, 20)
        Me.WindowToolStripMenuItem.Text = "Window"
        '
        'HideToolStripMenuItem
        '
        Me.HideToolStripMenuItem.Name = "HideToolStripMenuItem"
        Me.HideToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.F4), System.Windows.Forms.Keys)
        Me.HideToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.HideToolStripMenuItem.Text = "Hide"
        '
        'OnTopToolStripMenuItem
        '
        Me.OnTopToolStripMenuItem.Checked = True
        Me.OnTopToolStripMenuItem.CheckOnClick = True
        Me.OnTopToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.OnTopToolStripMenuItem.Name = "OnTopToolStripMenuItem"
        Me.OnTopToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.OnTopToolStripMenuItem.Text = "On Top?"
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBox1.Location = New System.Drawing.Point(0, 24)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(587, 316)
        Me.RichTextBox1.TabIndex = 1
        Me.RichTextBox1.Text = ""
        Me.RichTextBox1.Visible = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TSSL1, Me.TSSL2})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 340)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(587, 22)
        Me.StatusStrip1.TabIndex = 2
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'TSSL1
        '
        Me.TSSL1.Name = "TSSL1"
        Me.TSSL1.Size = New System.Drawing.Size(208, 17)
        Me.TSSL1.Text = "Clipboard has nothing now to display."
        '
        'TSSL2
        '
        Me.TSSL2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.TSSL2.Name = "TSSL2"
        Me.TSSL2.Size = New System.Drawing.Size(186, 17)
        Me.TSSL2.Text = "Clipboard Reloaded successfully!"
        Me.TSSL2.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PictureBox1.Location = New System.Drawing.Point(0, 24)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(587, 316)
        Me.PictureBox1.TabIndex = 3
        Me.PictureBox1.TabStop = False
        Me.PictureBox1.Visible = False
        '
        'ListBox1
        '
        Me.ListBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.IntegralHeight = False
        Me.ListBox1.Location = New System.Drawing.Point(0, 24)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(587, 316)
        Me.ListBox1.TabIndex = 4
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 24)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(587, 316)
        Me.Panel1.TabIndex = 5
        Me.Panel1.Visible = False
        '
        'Label1
        '
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Label1.Location = New System.Drawing.Point(0, 281)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(587, 35)
        Me.Label1.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Button1.Location = New System.Drawing.Point(245, 129)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(70, 42)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Play"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ReloadInfo
        '
        Me.ReloadInfo.Interval = 5000
        '
        'ClipboardViewer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(587, 362)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MinimizeBox = False
        Me.Name = "ClipboardViewer"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Clipboard Viewer"
        Me.TopMost = True
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Controller As System.Windows.Forms.Timer
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents TSSL1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents ClipboardToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ClearContentToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents WindowToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HideToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OnTopToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReloadToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TSSL2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ReloadInfo As System.Windows.Forms.Timer
    Friend WithEvents HistoryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StoreClipboardHistoryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RemoveHistoryAfterEndToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
