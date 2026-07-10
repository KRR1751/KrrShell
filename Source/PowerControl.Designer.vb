<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PowerControl
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
        Me.pnl_PWRP = New System.Windows.Forms.Panel()
        Me.lbl_PWRP = New System.Windows.Forms.Label()
        Me.pnl_Info = New System.Windows.Forms.Panel()
        Me.pnl_Links = New System.Windows.Forms.Panel()
        Me.lbl_Info = New System.Windows.Forms.Label()
        Me.pb_Icon = New System.Windows.Forms.PictureBox()
        Me.llbl_1 = New System.Windows.Forms.LinkLabel()
        Me.llbl_2 = New System.Windows.Forms.LinkLabel()
        Me.pnl_Info.SuspendLayout()
        Me.pnl_Links.SuspendLayout()
        CType(Me.pb_Icon, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnl_PWRP
        '
        Me.pnl_PWRP.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnl_PWRP.AutoScroll = True
        Me.pnl_PWRP.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.pnl_PWRP.Location = New System.Drawing.Point(3, 94)
        Me.pnl_PWRP.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.pnl_PWRP.Name = "pnl_PWRP"
        Me.pnl_PWRP.Size = New System.Drawing.Size(223, 90)
        Me.pnl_PWRP.TabIndex = 2
        '
        'lbl_PWRP
        '
        Me.lbl_PWRP.Dock = System.Windows.Forms.DockStyle.Top
        Me.lbl_PWRP.Location = New System.Drawing.Point(0, 73)
        Me.lbl_PWRP.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbl_PWRP.Name = "lbl_PWRP"
        Me.lbl_PWRP.Size = New System.Drawing.Size(228, 20)
        Me.lbl_PWRP.TabIndex = 1
        Me.lbl_PWRP.Text = "Select a power plan:"
        Me.lbl_PWRP.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'pnl_Info
        '
        Me.pnl_Info.BackColor = System.Drawing.SystemColors.Control
        Me.pnl_Info.Controls.Add(Me.lbl_Info)
        Me.pnl_Info.Controls.Add(Me.pb_Icon)
        Me.pnl_Info.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnl_Info.Location = New System.Drawing.Point(0, 0)
        Me.pnl_Info.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.pnl_Info.Name = "pnl_Info"
        Me.pnl_Info.Size = New System.Drawing.Size(228, 73)
        Me.pnl_Info.TabIndex = 0
        '
        'pnl_Links
        '
        Me.pnl_Links.BackColor = System.Drawing.SystemColors.ControlLight
        Me.pnl_Links.Controls.Add(Me.llbl_2)
        Me.pnl_Links.Controls.Add(Me.llbl_1)
        Me.pnl_Links.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnl_Links.Location = New System.Drawing.Point(0, 187)
        Me.pnl_Links.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.pnl_Links.Name = "pnl_Links"
        Me.pnl_Links.Size = New System.Drawing.Size(228, 63)
        Me.pnl_Links.TabIndex = 3
        '
        'lbl_Info
        '
        Me.lbl_Info.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbl_Info.AutoEllipsis = True
        Me.lbl_Info.Location = New System.Drawing.Point(70, 11)
        Me.lbl_Info.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbl_Info.Name = "lbl_Info"
        Me.lbl_Info.Size = New System.Drawing.Size(145, 50)
        Me.lbl_Info.TabIndex = 1
        Me.lbl_Info.Text = "No battery"
        '
        'pb_Icon
        '
        Me.pb_Icon.Location = New System.Drawing.Point(12, 11)
        Me.pb_Icon.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.pb_Icon.Name = "pb_Icon"
        Me.pb_Icon.Size = New System.Drawing.Size(50, 50)
        Me.pb_Icon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pb_Icon.TabIndex = 0
        Me.pb_Icon.TabStop = False
        '
        'llbl_1
        '
        Me.llbl_1.Dock = System.Windows.Forms.DockStyle.Top
        Me.llbl_1.Enabled = False
        Me.llbl_1.Location = New System.Drawing.Point(0, 0)
        Me.llbl_1.Name = "llbl_1"
        Me.llbl_1.Size = New System.Drawing.Size(228, 24)
        Me.llbl_1.TabIndex = 0
        Me.llbl_1.TabStop = True
        Me.llbl_1.Text = "Adjust screen Brightness"
        Me.llbl_1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'llbl_2
        '
        Me.llbl_2.Dock = System.Windows.Forms.DockStyle.Top
        Me.llbl_2.Location = New System.Drawing.Point(0, 24)
        Me.llbl_2.Name = "llbl_2"
        Me.llbl_2.Size = New System.Drawing.Size(228, 24)
        Me.llbl_2.TabIndex = 1
        Me.llbl_2.TabStop = True
        Me.llbl_2.Text = "More power Options"
        Me.llbl_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'PowerControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(228, 250)
        Me.ControlBox = False
        Me.Controls.Add(Me.pnl_Links)
        Me.Controls.Add(Me.lbl_PWRP)
        Me.Controls.Add(Me.pnl_PWRP)
        Me.Controls.Add(Me.pnl_Info)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Name = "PowerControl"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.TopMost = True
        Me.pnl_Info.ResumeLayout(False)
        Me.pnl_Links.ResumeLayout(False)
        CType(Me.pb_Icon, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pnl_PWRP As Panel
    Friend WithEvents lbl_PWRP As Label
    Friend WithEvents pnl_Info As Panel
    Friend WithEvents pnl_Links As Panel
    Friend WithEvents pb_Icon As PictureBox
    Friend WithEvents lbl_Info As Label
    Friend WithEvents llbl_2 As LinkLabel
    Friend WithEvents llbl_1 As LinkLabel
End Class
