<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NetworkControl
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
        Me.txt_SSID = New System.Windows.Forms.TextBox()
        Me.txt_Password = New System.Windows.Forms.TextBox()
        Me.btn_Refresh = New System.Windows.Forms.Button()
        Me.btn_Connect = New System.Windows.Forms.Button()
        Me.lv_Main = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.il_Main = New System.Windows.Forms.ImageList(Me.components)
        Me.chk_HideSSID = New System.Windows.Forms.CheckBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lbl_Tip = New System.Windows.Forms.Label()
        Me.lbl_WifiCount = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.chk_HidePassword = New System.Windows.Forms.CheckBox()
        Me.lbl_PassMan = New System.Windows.Forms.Label()
        Me.btn_Disconnect = New System.Windows.Forms.Button()
        Me.chk_AutoUpdate = New System.Windows.Forms.CheckBox()
        Me.Controller = New System.Windows.Forms.Timer(Me.components)
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txt_SSID
        '
        Me.txt_SSID.Location = New System.Drawing.Point(7, 37)
        Me.txt_SSID.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txt_SSID.Name = "txt_SSID"
        Me.txt_SSID.Size = New System.Drawing.Size(149, 23)
        Me.txt_SSID.TabIndex = 2
        '
        'txt_Password
        '
        Me.txt_Password.Location = New System.Drawing.Point(7, 84)
        Me.txt_Password.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txt_Password.Name = "txt_Password"
        Me.txt_Password.PasswordChar = Global.Microsoft.VisualBasic.ChrW(9679)
        Me.txt_Password.Size = New System.Drawing.Size(123, 23)
        Me.txt_Password.TabIndex = 3
        '
        'btn_Refresh
        '
        Me.btn_Refresh.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btn_Refresh.Location = New System.Drawing.Point(14, 303)
        Me.btn_Refresh.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btn_Refresh.Name = "btn_Refresh"
        Me.btn_Refresh.Size = New System.Drawing.Size(88, 27)
        Me.btn_Refresh.TabIndex = 1
        Me.btn_Refresh.Text = "Refresh"
        Me.btn_Refresh.UseVisualStyleBackColor = True
        '
        'btn_Connect
        '
        Me.btn_Connect.AutoEllipsis = True
        Me.btn_Connect.Enabled = False
        Me.btn_Connect.Location = New System.Drawing.Point(3, 127)
        Me.btn_Connect.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btn_Connect.Name = "btn_Connect"
        Me.btn_Connect.Size = New System.Drawing.Size(88, 27)
        Me.btn_Connect.TabIndex = 5
        Me.btn_Connect.Text = "Connect"
        Me.btn_Connect.UseVisualStyleBackColor = True
        '
        'lv_Main
        '
        Me.lv_Main.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lv_Main.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4})
        Me.lv_Main.FullRowSelect = True
        Me.lv_Main.HideSelection = False
        Me.lv_Main.LargeImageList = Me.il_Main
        Me.lv_Main.Location = New System.Drawing.Point(14, 14)
        Me.lv_Main.Name = "lv_Main"
        Me.lv_Main.Size = New System.Drawing.Size(337, 288)
        Me.lv_Main.SmallImageList = Me.il_Main
        Me.lv_Main.StateImageList = Me.il_Main
        Me.lv_Main.TabIndex = 0
        Me.lv_Main.UseCompatibleStateImageBehavior = False
        Me.lv_Main.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Wi-Fi name (SSID)"
        Me.ColumnHeader1.Width = 120
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Secured"
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "GHz"
        Me.ColumnHeader3.Width = 40
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Speed"
        '
        'il_Main
        '
        Me.il_Main.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit
        Me.il_Main.ImageSize = New System.Drawing.Size(16, 16)
        Me.il_Main.TransparentColor = System.Drawing.Color.Transparent
        '
        'chk_HideSSID
        '
        Me.chk_HideSSID.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chk_HideSSID.AutoSize = True
        Me.chk_HideSSID.Location = New System.Drawing.Point(163, 308)
        Me.chk_HideSSID.Name = "chk_HideSSID"
        Me.chk_HideSSID.Size = New System.Drawing.Size(176, 19)
        Me.chk_HideSSID.TabIndex = 3
        Me.chk_HideSSID.Text = "Hide Network names (SSIDs)"
        Me.chk_HideSSID.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.AutoScroll = True
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.lbl_Tip)
        Me.Panel1.Controls.Add(Me.lbl_WifiCount)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.btn_Disconnect)
        Me.Panel1.Controls.Add(Me.btn_Connect)
        Me.Panel1.Location = New System.Drawing.Point(357, 14)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(185, 288)
        Me.Panel1.TabIndex = 7
        '
        'lbl_Tip
        '
        Me.lbl_Tip.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbl_Tip.Enabled = False
        Me.lbl_Tip.Location = New System.Drawing.Point(3, 157)
        Me.lbl_Tip.Name = "lbl_Tip"
        Me.lbl_Tip.Size = New System.Drawing.Size(177, 109)
        Me.lbl_Tip.TabIndex = 8
        Me.lbl_Tip.Text = "NOTE: You don't need to set up each time a password when you already visited the " &
    "network once. Just hit ""Connect"" without a password!"
        Me.lbl_Tip.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lbl_WifiCount
        '
        Me.lbl_WifiCount.AutoEllipsis = True
        Me.lbl_WifiCount.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lbl_WifiCount.Location = New System.Drawing.Point(0, 265)
        Me.lbl_WifiCount.Name = "lbl_WifiCount"
        Me.lbl_WifiCount.Size = New System.Drawing.Size(183, 21)
        Me.lbl_WifiCount.TabIndex = 9
        Me.lbl_WifiCount.Text = "0 Networks are in this area."
        Me.lbl_WifiCount.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txt_SSID)
        Me.GroupBox1.Controls.Add(Me.txt_Password)
        Me.GroupBox1.Controls.Add(Me.chk_HidePassword)
        Me.GroupBox1.Controls.Add(Me.lbl_PassMan)
        Me.GroupBox1.Location = New System.Drawing.Point(2, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(178, 118)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Wi-Fi signin"
        '
        'chk_HidePassword
        '
        Me.chk_HidePassword.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.chk_HidePassword.Location = New System.Drawing.Point(130, 85)
        Me.chk_HidePassword.Name = "chk_HidePassword"
        Me.chk_HidePassword.Size = New System.Drawing.Size(26, 23)
        Me.chk_HidePassword.TabIndex = 4
        Me.chk_HidePassword.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.chk_HidePassword.UseVisualStyleBackColor = True
        '
        'lbl_PassMan
        '
        Me.lbl_PassMan.AutoSize = True
        Me.lbl_PassMan.Location = New System.Drawing.Point(4, 21)
        Me.lbl_PassMan.Name = "lbl_PassMan"
        Me.lbl_PassMan.Size = New System.Drawing.Size(122, 60)
        Me.lbl_PassMan.TabIndex = 1
        Me.lbl_PassMan.Text = "Network name (SSID):" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Password:"
        '
        'btn_Disconnect
        '
        Me.btn_Disconnect.AutoEllipsis = True
        Me.btn_Disconnect.Enabled = False
        Me.btn_Disconnect.Location = New System.Drawing.Point(92, 127)
        Me.btn_Disconnect.Name = "btn_Disconnect"
        Me.btn_Disconnect.Size = New System.Drawing.Size(88, 27)
        Me.btn_Disconnect.TabIndex = 6
        Me.btn_Disconnect.Text = "Disconnect"
        Me.btn_Disconnect.UseVisualStyleBackColor = True
        '
        'chk_AutoUpdate
        '
        Me.chk_AutoUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chk_AutoUpdate.AutoSize = True
        Me.chk_AutoUpdate.Location = New System.Drawing.Point(14, 336)
        Me.chk_AutoUpdate.Name = "chk_AutoUpdate"
        Me.chk_AutoUpdate.Size = New System.Drawing.Size(93, 19)
        Me.chk_AutoUpdate.TabIndex = 2
        Me.chk_AutoUpdate.Text = "Auto Update"
        Me.chk_AutoUpdate.UseVisualStyleBackColor = True
        '
        'Controller
        '
        Me.Controller.Interval = 5000
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OK_Button.Location = New System.Drawing.Point(454, 328)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(88, 27)
        Me.OK_Button.TabIndex = 9
        Me.OK_Button.Text = "OK"
        Me.OK_Button.UseVisualStyleBackColor = True
        '
        'NetworkControl
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(554, 365)
        Me.Controls.Add(Me.OK_Button)
        Me.Controls.Add(Me.chk_AutoUpdate)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.chk_HideSSID)
        Me.Controls.Add(Me.lv_Main)
        Me.Controls.Add(Me.btn_Refresh)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "NetworkControl"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Network Control"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txt_SSID As TextBox
    Friend WithEvents txt_Password As TextBox
    Friend WithEvents btn_Refresh As Button
    Friend WithEvents btn_Connect As Button
    Friend WithEvents lv_Main As ListView
    Friend WithEvents ColumnHeader1 As ColumnHeader
    Friend WithEvents ColumnHeader2 As ColumnHeader
    Friend WithEvents ColumnHeader3 As ColumnHeader
    Friend WithEvents ColumnHeader4 As ColumnHeader
    Friend WithEvents chk_HideSSID As CheckBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents il_Main As ImageList
    Friend WithEvents chk_HidePassword As CheckBox
    Friend WithEvents btn_Disconnect As Button
    Friend WithEvents lbl_WifiCount As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents lbl_PassMan As Label
    Friend WithEvents lbl_Tip As Label
    Friend WithEvents chk_AutoUpdate As CheckBox
    Friend WithEvents Controller As Timer
    Friend WithEvents OK_Button As Button
End Class
