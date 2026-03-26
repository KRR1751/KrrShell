<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SA
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SA))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.chkForceMode = New System.Windows.Forms.CheckBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkDisableWakeEvent = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pnlRadioActions = New System.Windows.Forms.Panel()
        Me.cmbPowerActions = New System.Windows.Forms.ComboBox()
        Me.chkShutdownHybrid = New System.Windows.Forms.CheckBox()
        Me.txtReason = New System.Windows.Forms.TextBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.chkBrutalForce = New System.Windows.Forms.CheckBox()
        Me.ToolTip2 = New System.Windows.Forms.ToolTip(Me.components)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(68, 231)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(88, 27)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(162, 231)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(88, 27)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBox1.Image = Global.NewShell.My.Resources.Resources.Clock
        Me.PictureBox1.Location = New System.Drawing.Point(14, 14)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(37, 37)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(64, 12)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(215, 15)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "What do you want the computer to do?"
        '
        'chkForceMode
        '
        Me.chkForceMode.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkForceMode.AutoSize = True
        Me.chkForceMode.Location = New System.Drawing.Point(0, 5)
        Me.chkForceMode.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.chkForceMode.Name = "chkForceMode"
        Me.chkForceMode.Size = New System.Drawing.Size(89, 19)
        Me.chkForceMode.TabIndex = 11
        Me.chkForceMode.Text = "Force Mode"
        Me.ToolTip1.SetToolTip(Me.chkForceMode, resources.GetString("chkForceMode.ToolTip"))
        Me.chkForceMode.UseVisualStyleBackColor = True
        '
        'chkDisableWakeEvent
        '
        Me.chkDisableWakeEvent.AutoSize = True
        Me.chkDisableWakeEvent.Location = New System.Drawing.Point(125, 3)
        Me.chkDisableWakeEvent.Name = "chkDisableWakeEvent"
        Me.chkDisableWakeEvent.Size = New System.Drawing.Size(128, 19)
        Me.chkDisableWakeEvent.TabIndex = 16
        Me.chkDisableWakeEvent.Text = "Disable Wake Event"
        Me.ToolTip1.SetToolTip(Me.chkDisableWakeEvent, "When checked, aplications like ""Alarms"" will not be able to wake up the system fr" &
        "om ""Sleep mode""" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "NOTE: Pressing anything on your keyboard of course will wake " &
        "it up instantly.")
        Me.chkDisableWakeEvent.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(-1, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(137, 15)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "Action reason: (Optimal)"
        Me.ToolTip1.SetToolTip(Me.Label1, "You can type a reason why you did this action. Similar to Windows Server")
        '
        'pnlRadioActions
        '
        Me.pnlRadioActions.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlRadioActions.Location = New System.Drawing.Point(0, 23)
        Me.pnlRadioActions.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.pnlRadioActions.Name = "pnlRadioActions"
        Me.pnlRadioActions.Size = New System.Drawing.Size(276, 90)
        Me.pnlRadioActions.TabIndex = 12
        '
        'cmbPowerActions
        '
        Me.cmbPowerActions.Dock = System.Windows.Forms.DockStyle.Top
        Me.cmbPowerActions.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPowerActions.FormattingEnabled = True
        Me.cmbPowerActions.Items.AddRange(New Object() {"Shutdown", "Shutdown with Fastboot", "Restart", "Sleep", "Hibernate"})
        Me.cmbPowerActions.Location = New System.Drawing.Point(0, 0)
        Me.cmbPowerActions.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cmbPowerActions.Name = "cmbPowerActions"
        Me.cmbPowerActions.Size = New System.Drawing.Size(276, 23)
        Me.cmbPowerActions.TabIndex = 13
        '
        'chkShutdownHybrid
        '
        Me.chkShutdownHybrid.AutoSize = True
        Me.chkShutdownHybrid.Location = New System.Drawing.Point(0, 3)
        Me.chkShutdownHybrid.Name = "chkShutdownHybrid"
        Me.chkShutdownHybrid.Size = New System.Drawing.Size(119, 19)
        Me.chkShutdownHybrid.TabIndex = 14
        Me.chkShutdownHybrid.Text = "Hybrid Shutdown"
        Me.chkShutdownHybrid.UseVisualStyleBackColor = True
        '
        'txtReason
        '
        Me.txtReason.Location = New System.Drawing.Point(0, 36)
        Me.txtReason.Name = "txtReason"
        Me.txtReason.Size = New System.Drawing.Size(276, 23)
        Me.txtReason.TabIndex = 15
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.txtReason)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.chkShutdownHybrid)
        Me.Panel2.Controls.Add(Me.chkDisableWakeEvent)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 134)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(276, 61)
        Me.Panel2.TabIndex = 16
        Me.Panel2.Visible = False
        '
        'CheckBox1
        '
        Me.CheckBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBox1.Appearance = System.Windows.Forms.Appearance.Button
        Me.CheckBox1.Location = New System.Drawing.Point(257, 231)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(87, 27)
        Me.CheckBox1.TabIndex = 17
        Me.CheckBox1.Text = "&Advanced"
        Me.CheckBox1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.pnlRadioActions)
        Me.Panel1.Controls.Add(Me.cmbPowerActions)
        Me.Panel1.Location = New System.Drawing.Point(67, 30)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(276, 195)
        Me.Panel1.TabIndex = 18
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.chkBrutalForce)
        Me.Panel3.Controls.Add(Me.chkForceMode)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 113)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(276, 21)
        Me.Panel3.TabIndex = 14
        '
        'chkBrutalForce
        '
        Me.chkBrutalForce.AutoSize = True
        Me.chkBrutalForce.Location = New System.Drawing.Point(125, 5)
        Me.chkBrutalForce.Name = "chkBrutalForce"
        Me.chkBrutalForce.Size = New System.Drawing.Size(87, 19)
        Me.chkBrutalForce.TabIndex = 12
        Me.chkBrutalForce.Text = "Brutal force"
        Me.ToolTip2.SetToolTip(Me.chkBrutalForce, resources.GetString("chkBrutalForce.ToolTip"))
        Me.chkBrutalForce.UseVisualStyleBackColor = True
        '
        'ToolTip2
        '
        Me.ToolTip2.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Warning
        Me.ToolTip2.ToolTipTitle = "Warning"
        '
        'SA
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(369, 271)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(375, 300)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(375, 150)
        Me.Name = "SA"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Shut Down PC"
        Me.TopMost = True
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkForceMode As CheckBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents pnlRadioActions As Panel
    Friend WithEvents cmbPowerActions As ComboBox
    Friend WithEvents chkShutdownHybrid As CheckBox
    Friend WithEvents txtReason As TextBox
    Friend WithEvents Panel2 As Panel
    Friend WithEvents chkDisableWakeEvent As CheckBox
    Friend WithEvents Label1 As Label
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents chkBrutalForce As CheckBox
    Friend WithEvents ToolTip2 As ToolTip
End Class
