Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports NewShell.AppBar

Public Class AppbarProperties
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindow(
     ByVal lpClassName As String,
     ByVal lpWindowName As String) As IntPtr
    End Function
    <DllImport("user32.dll")>
    Private Shared Function GetWindowPlacement(ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function
    <DllImport("user32.dll")>
    Private Shared Function SetWindowPlacement(ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function

    Private Structure WINDOWPLACEMENT
        Public Length As Integer
        Public flags As Integer
        Public showCmd As ShowWindowCommands
        Public ptMinPosition As POINTAPI
        Public ptMaxPosition As POINTAPI
        Public rcNormalPosition As RECT
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure DWMCOLORIZATIONPARAMS
        Public clrColor As UInteger
        Public clrAfterGlow As UInteger
        Public nIntensity As UInteger
        Public clrAfterGlowBalance As UInteger
        Public clrBlurBalance As UInteger
        Public nColorizationGlassAttribute As UInteger
        Public nColorizationOpaqueBlend As UInteger
    End Structure

    <DllImport("dwmapi.dll", EntryPoint:="#127")>
    Private Shared Function DwmGetColorizationParameters(ByRef parameters As DWMCOLORIZATIONPARAMS) As Integer
    End Function

    <DllImport("dwmapi.dll")>
    Private Shared Function DwmGetWindowAttribute(hWnd As IntPtr, dwAttribute As UInteger, ByRef pvAttribute As Boolean, cbAttribute As UInteger) As Integer
    End Function

    Private Const DWMWA_USE_HOST_BACKDROP_BRUSH As UInteger = 19
    Private Const DWMWA_CAPTION_COLOR As UInteger = 35

    Public Shared Function GetAccentColor() As Color
        Dim Params As New DWMCOLORIZATIONPARAMS()
        Dim result As Integer = DwmGetColorizationParameters(Params)

        If result = 0 Then
            Dim colorValue As UInteger = Params.clrColor

            Dim A As Integer = CInt((colorValue >> 24) And &HFF) ' Alpha
            Dim R As Integer = CInt((colorValue >> 16) And &HFF) ' Red
            Dim G As Integer = CInt((colorValue >> 8) And &HFF)  ' Green
            Dim B As Integer = CInt(colorValue And &HFF)        ' Blue

            Return Color.FromArgb(R, G, B)
        Else
            Return Color.FromArgb(0, 120, 215) ' Classic Blue from Windows
        End If
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        AppBar.LoadApps()

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub AppbarProperties_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckBox1.Checked = AppBar.AutoHideToolStripMenuItem.Checked
        CheckBox2.Checked = AppBar.StartButtonToolStripMenuItem.Checked
        CheckBox3.Checked = AppBar.PinnedBarToolStripMenuItem.Checked
        CheckBox4.Checked = AppBar.MenuBarToolStripMenuItem.Checked
        CheckBox5.Checked = AppBar.Button2.Visible
        CheckBox6.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "Enabled", "0")
        CheckBox7.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowBackImage", False)
        CheckBox8.Checked = AppBar.TopMost
        CheckBox9.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "SpaceForEmergeTray", False)

        CheckBox12.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "Style", False)

        GroupBox3.Enabled = CheckBox6.Checked

        NumericUpDown2.Maximum = SystemInformation.PrimaryMonitorSize.Width
        Try
            NumericUpDown2.Value = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "CustomWidth", 900)
        Catch ex As Exception
            NumericUpDown2.Value = 900
        End Try

        ComboBox5.SelectedIndex = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\", "CombineMode", "0")

        Try
            TrackBar1.Value = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "TransparencyLevel", "90")
            NumericUpDown3.Value = TrackBar1.Value
        Catch ex As Exception
            TrackBar1.Value = 90
            NumericUpDown3.Value = 90
        End Try

        Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layout", 2)
            Case 0
                Button16.Enabled = True
                Label12.Enabled = True

                AppBar.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", "#B77358"))
                AppBar.TransparencyKey = Nothing

                AppBar.Opacity = 1

            Case 1
                Button16.Enabled = True
                Label12.Enabled = True

                AppBar.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", "#B77358"))
                AppBar.TransparencyKey = Nothing

                Label6.Enabled = True
                TrackBar1.Enabled = True
                NumericUpDown3.Enabled = True

                AppBar.Opacity = TrackBar1.Value / 100

            Case 2
                Button16.Enabled = True
                Label12.Enabled = True

                AppBar.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", "#B77358"))
                AppBar.TransparencyKey = Nothing

                Label6.Enabled = False
                TrackBar1.Enabled = False
                NumericUpDown3.Enabled = False

                AppBar.Opacity = 1

        End Select

        Try
            CheckBox11.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "UseSystemColor", 1)
        Catch ex As Exception
            CheckBox11.Checked = True
        End Try

        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "KillMethod", "0") = 0 Then
            RadioButton4.Checked = True
            RadioButton5.Checked = False
        Else
            RadioButton5.Checked = True
            RadioButton4.Checked = False
        End If

        Blacklist.Items.Clear()
        For Each i As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\HiddenProcesses").GetValueNames
            Blacklist.Items.Add(i)
        Next

        Blocklist.Items.Clear()
        For Each i As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\BlockedProcesses\.\").GetValueNames
            Blocklist.Items.Add(i)
        Next

        PictureBox1.BackColor = AppBar.BackColor
        Label4.BackColor = AppBar.Splitter1.BackColor
        ComboBox3.SelectedIndex = AppBar.BackgroundImageLayout
        PictureBox1.BackgroundImageLayout = ComboBox3.SelectedIndex

        If CheckBox7.Checked = False Then
            PictureBox1.Enabled = False
            ComboBox2.Enabled = False
            ComboBox3.Enabled = False
            Button11.Enabled = False
        End If

        If AppBar.Button1.BackgroundImageLayout = ImageLayout.Stretch Then
            ComboBox4.SelectedIndex = 0
        Else
            ComboBox4.SelectedIndex = 1
        End If

        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 0 Then
            RadioButton8.Checked = True
            RadioButton9.Checked = False
        ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 1 Then
            RadioButton8.Checked = False
            RadioButton9.Checked = True
        ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 2 Then
            'RadioButton10.Checked = True
        End If

        Try
            PictureBox4.Image = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Normal", ""))
        Catch ex As Exception
            PictureBox4.Image = My.Resources.StartRight
        End Try
        Try
            PictureBox5.Image = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Hover", ""))
        Catch ex As Exception
            PictureBox5.Image = My.Resources.StartRight
        End Try
        Try
            PictureBox6.Image = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Pressed", ""))
        Catch ex As Exception
            PictureBox6.Image = My.Resources.StartLeft
        End Try
        Try
            PictureBox7.Image = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "ORB", My.Resources.StartRight))
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        ' Disallowing this because of not knowing anything xdd

        AppBar.AutoHideToolStripMenuItem.Checked = CheckBox1.Checked
        AppBar.fBarRegistered = AppBar.AutoHideToolStripMenuItem.Checked
        AppBar.RegisterBar()
        If AppBar.AutoHideToolStripMenuItem.Checked = True Then
            AppBar.Hidden = True
            Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - 2)
        Else
            AppBar.Hidden = False
            Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - Me.Height)
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        AppBar.StartButtonToolStripMenuItem.Checked = CheckBox2.Checked
        AppBar.Button1.Visible = AppBar.StartButtonToolStripMenuItem.Checked
        If AppBar.LockAppbarToolStripMenuItem.Checked = False Then
            AppBar.Splitter1.Visible = AppBar.StartButtonToolStripMenuItem.Checked
        End If
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\", "StartButton", CheckBox2.Checked, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        AppBar.PinnedBarToolStripMenuItem.Checked = CheckBox3.Checked
        AppBar.Panel1.Visible = AppBar.PinnedBarToolStripMenuItem.Checked
        If AppBar.LockAppbarToolStripMenuItem.Checked = False Then
            AppBar.Splitter2.Visible = AppBar.PinnedBarToolStripMenuItem.Checked
        End If
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\", "PinnedBar", CheckBox3.Checked, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        AppBar.MenuBarToolStripMenuItem.Checked = CheckBox4.Checked
        AppBar.Panel4.Visible = AppBar.MenuBarToolStripMenuItem.Checked
        If AppBar.LockAppbarToolStripMenuItem.Checked = False Then
            AppBar.Splitter3.Visible = AppBar.MenuBarToolStripMenuItem.Checked
        End If
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\", "MenuBar", CheckBox4.Checked, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        AppBar.Button2.Visible = CheckBox5.Checked
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\", "ShowDesktopButton", CheckBox5.Checked, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'On Error Resume Next
        Dim Output As String = InputBox("Type a process name do you want to Hide from the Appbar...", "Input Box", "")
        If Output IsNot String.Empty Then
            Try
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\HiddenProcesses\", Output, "")
                Blacklist.Items.Add(Output)
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'On Error Resume Next
        If MessageBox.Show("Are you sure do you want to remove those selected processes from your Blacklist?", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            Shell("reg.exe delete HKEY_CURRENT_USER\Software\Shell\Appbar\HiddenProcesses\ /v " & Blacklist.SelectedItem & " /f", AppWinStyle.NormalFocus, True)
            Blacklist.Items.Clear()
            For Each ii As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\HiddenProcesses").GetValueNames
                Blacklist.Items.Add(ii)
            Next

            AppBar.LoadApps()
        End If
    End Sub

    Private Sub Controller_Tick(sender As Object, e As EventArgs) Handles Controller.Tick
        If Blacklist.SelectedItems.Count = 0 Then
            Button2.Enabled = False
            Button3.Enabled = True
        Else
            Button3.Enabled = False
            Button2.Enabled = True
        End If

        If Blocklist.SelectedItems.Count = 0 Then
            Button5.Enabled = False
            Button4.Enabled = True
        Else
            Button4.Enabled = False
            Button5.Enabled = True
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        On Error Resume Next
        If MessageBox.Show("Are you sure do you want to remove All the processes from your Blacklist?", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            For i As Integer = 0 To Blacklist.Items.Count - 1
                Shell("reg.exe delete HKEY_CURRENT_USER\Software\Shell\Appbar\HiddenProcesses\ /v " & Blacklist.SelectedItem & " /f", AppWinStyle.NormalFocus, True)
            Next
            Blacklist.Items.Clear()

            AppBar.LoadApps()
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        GroupBox3.Enabled = CheckBox6.Checked
        AppBar.BlockingProcesses.Enabled = CheckBox6.Checked
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "Enabled", CheckBox6.Checked, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim Output As String = InputBox("Type a process name do you want to Block entirely...", "Input Box", "")
        If Output IsNot String.Empty Then
            If Output = "dwm" OrElse Output = "winlogon" OrElse Output = "ntoskrnl" OrElse Output = "svchost" OrElse Output = "wininit" OrElse Output = "lsass" OrElse Output = "smss" OrElse Output = "rundll32" OrElse Output = "taskhostw" OrElse Output = "wudfhost" OrElse Output = "sppsvc" OrElse Output = "services" OrElse Output = "csrss" OrElse Output = "fontdrvhost" OrElse Output = "conhost" OrElse Output = "krrshell" OrElse Output = "newshell" Then
                MessageBox.Show("I'm sorry but this process is NOT ALLOWED to be blocked. Since """ & Output & """ is a Critical Windows process that cannot be easily terminated." & Environment.NewLine & "And for your safety, we will not include this process to the Blocklist to keep your computer safe!", "This process cannot be ""Blocked""", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Else
                Try
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses\.\", Output, "")
                    Blocklist.Items.Add(Output)
                Catch ex As Exception

                End Try
            End If
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        On Error Resume Next
        If MessageBox.Show("Are you sure do you want to remove those selected processes from your Blocklist?", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            For i As Integer = 0 To Blocklist.SelectedItems.Count - 1
                'My.Computer.Registry.CurrentUser.OpenSubKey("\Software\Shell\Appbar\BlockedProcesses\.\", True, True).DeleteValue(Blocklist.SelectedItems.Item(i))
                Shell("reg.exe delete HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses\.\ /v " & Blocklist.SelectedItem & " /f", AppWinStyle.NormalFocus, True)
                Blocklist.Items.Clear()
                For Each ii As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\BlockedProcesses\.\").GetValueNames
                    Blocklist.Items.Add(ii)
                Next
            Next
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        On Error Resume Next
        If MessageBox.Show("Are you sure do you want to remove All the processes from your Blocklist?", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            For i As Integer = 0 To Blocklist.Items.Count - 1
                Shell("reg.exe delete HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses\.\ /v " & Blocklist.SelectedItem & " /f", AppWinStyle.NormalFocus, True)
            Next
            Blocklist.Items.Clear()
        End If
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        AppBar.BlockingProcesses.Interval = NumericUpDown1.Value
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "Ticks", NumericUpDown1.Value, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses\", "KillMethod", "0", Microsoft.Win32.RegistryValueKind.DWord)
        Else
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses\", "KillMethod", "1", Microsoft.Win32.RegistryValueKind.DWord)
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        AboutDialog.ShowDialog(Me)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If MessageBox.Show("Are you sure do you want end this shell? This means that you will no longer have any Taskbar, Desktop or Startmenu. (To revert it back, please press CTRL+SHIFT+ESC to start a new shell process.)", "Confirm box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            AppBar.CanClose = True
            Desktop.CanClose = True
            SA.SaveSettings()
            End
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        AppBar.CanClose = True
        Desktop.CanClose = True
        Application.Restart()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Dim CD As New ColorDialog
        CD.AllowFullOpen = True
        CD.AnyColor = True
        CD.FullOpen = True
        CD.Color = AppBar.BackColor
        If CD.ShowDialog(Me) = DialogResult.OK Then
            Label12.BackColor = CD.Color
            PictureBox1.BackColor = CD.Color
            AppBar.BackColor = CD.Color
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", ColorTranslator.ToHtml(CD.Color))
        End If
    End Sub

    Private Sub TabControl1_TabIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.TabIndexChanged
        If TabControl1.SelectedTab.TabIndex = 1 Then
            Button14.Enabled = True
            OK_Button.Enabled = False
        Else
            Button14.Enabled = False
            OK_Button.Enabled = True
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "OnTop", CheckBox8.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        AppBar.TopMost = CheckBox8.Checked
    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        Try
            Dim side As AppBar.side = side
            side.Left = -1
            side.Right = -1
            side.Top = -1
            side.Bottom = -1
            Dim result As Integer = AppBar.DwmExtendFrameIntoClientArea(AppBar.Handle, side)
        Catch ex As Exception
            Dim side As AppBar.side = side
            side.Left = 0
            side.Right = 0
            side.Top = 0
            side.Bottom = 0
            Dim result As Integer = AppBar.DwmExtendFrameIntoClientArea(AppBar.Handle, side)
        End Try

        Button16.Enabled = True
        Label12.Enabled = True

        AppBar.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", "#B77358"))
        AppBar.TransparencyKey = Nothing

        AppBar.Opacity = 1

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layout", "0", Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) ' Old
        ' When I'll come with an alternative to display an transparent window
        ' I'll replace this

        Dim side As AppBar.side = side
        side.Left = 0
        side.Right = 0
        side.Top = 0
        side.Bottom = 0
        Dim result As Integer = AppBar.DwmExtendFrameIntoClientArea(AppBar.Handle, side)

        Button16.Enabled = False
        Label12.Enabled = False

        AppBar.BackColor = Color.Black
        AppBar.TransparencyKey = Color.Black

        AppBar.Opacity = 1

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layout", "1", Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged, RadioButton1.CheckedChanged
        Dim side As AppBar.side = side
        side.Left = 0
        side.Right = 0
        side.Top = 0
        side.Bottom = 0
        Dim result As Integer = AppBar.DwmExtendFrameIntoClientArea(AppBar.Handle, side)

        Button16.Enabled = True
        Label12.Enabled = True

        AppBar.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", "#B77358"))
        AppBar.TransparencyKey = Nothing

        If RadioButton1.Checked = True Then
            Label6.Enabled = True
            TrackBar1.Enabled = True
            NumericUpDown3.Enabled = True

            AppBar.Opacity = TrackBar1.Value

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layout", "1", Microsoft.Win32.RegistryValueKind.DWord)
        Else
            Label6.Enabled = False
            TrackBar1.Enabled = False
            NumericUpDown3.Enabled = False

            AppBar.Opacity = 1

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layout", "2", Microsoft.Win32.RegistryValueKind.DWord)
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim CD As New ColorDialog
        CD.AllowFullOpen = True
        CD.AnyColor = True
        CD.FullOpen = True
        CD.Color = AppBar.Splitter1.BackColor
        If CD.ShowDialog(Me) = DialogResult.OK Then
            Label4.BackColor = CD.Color
            AppBar.Splitter1.BackColor = CD.Color
            AppBar.Splitter2.BackColor = CD.Color
            AppBar.Splitter3.BackColor = CD.Color
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "SplitterColor", ColorTranslator.ToHtml(CD.Color))
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowBackImage", CheckBox7.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        If CheckBox7.Checked = True Then
            PictureBox1.Enabled = True
            ComboBox2.Enabled = True
            Button11.Enabled = True
            ComboBox3.Enabled = True
            If ComboBox2.Text IsNot Nothing AndAlso File.Exists(ComboBox2.Text) Then
                AppBar.BackgroundImage = Image.FromFile(ComboBox2.Text)
            Else
                AppBar.BackgroundImage = My.Resources.AppBarMainTransparent1
            End If

        Else
            PictureBox1.Enabled = False
            ComboBox2.Enabled = False
            Button11.Enabled = False
            ComboBox3.Enabled = False
            AppBar.BackgroundImage = Nothing
        End If
    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        If ComboBox2.Enabled = True Then
            Try
                If IO.File.Exists(ComboBox2.Text) = True Then
                    Dim fi As New FileInfo(ComboBox2.Text)
                    If fi.Extension = ".png" OrElse fi.Extension = ".jpg" OrElse fi.Extension = ".bmp" OrElse fi.Extension = ".gif" Then
                        PictureBox1.Image = Image.FromFile(fi.FullName)
                        AppBar.BackgroundImage = Image.FromFile(fi.FullName)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImage", ComboBox2.Text)
                    End If
                ElseIf ComboBox2.Text = "" Or Not ComboBox2.Text Then
                    PictureBox1.Image = My.Resources.AppBarMain
                    AppBar.BackgroundImage = My.Resources.AppBarMain
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImage", "")
                End If
            Catch ex As Exception
                PictureBox1.Image = My.Resources.AppBarMain
                AppBar.BackgroundImage = My.Resources.AppBarMain
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImage", "")
            End Try
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim ofd As New OpenFileDialog()
        ofd.Title = "Open a ""Appbar"" image for to show in the Appbar"
        ofd.CheckFileExists = True
        ofd.Multiselect = False
        If IO.File.Exists(ComboBox2.Text) = True Then
            Dim fi As New FileInfo(ComboBox2.Text)
            ofd.InitialDirectory = fi.DirectoryName
            ofd.FileName = fi.Name
        ElseIf IO.Directory.Exists(ComboBox2.Text) = True Then
            ofd.InitialDirectory = ComboBox2.Text
            ofd.FileName = ""
        Else
            ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
            ofd.FileName = ""
        End If
        ofd.Filter = "Pictures (*.png;*.jpg;*.bmp;*.gif)|*.png;*.jpg;*.bmp;*.gif"
        If ofd.ShowDialog(Me) = DialogResult.OK Then
            ComboBox2.Text = ofd.FileName
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        AppBar.BackgroundImageLayout = ComboBox3.SelectedIndex
        PictureBox1.BackgroundImageLayout = ComboBox3.SelectedIndex
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImageLayout", ComboBox3.SelectedIndex, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long,
                                                        ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long,
                                                        ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOZORDER As UInteger = &H4
    Private Const SWP_SHOWWINDOW As UInteger = &H40

    Private Const HWND_TOPMOST = -1 '-- Bring to top and stay there
    Private Sub MoveAndResizeWindow(ByVal hWnd As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer)
        SetWindowPos(hWnd, IntPtr.Zero, X, Y, Width, Height, SWP_NOZORDER Or SWP_SHOWWINDOW)
    End Sub
    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        If CheckBox9.Checked = True Then
            Panel1.Enabled = True
            Dim prFound As Boolean = False
            For Each pr As Process In Process.GetProcesses
                If pr.ProcessName = "emergeTray" Then
                    prFound = True
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "SpaceForEmergeTray", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    AppBar.Panel4.Visible = False
                    AppBar.Width = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "CustomWidth", SystemInformation.PrimaryMonitorSize.Width / 2 + SystemInformation.PrimaryMonitorSize.Width / 4)
                    ClockTray.Show()
                    Try
                        Dim hWnd As IntPtr = FindWindow("EmergeDesktopApplet", vbNullString)
                        If hWnd <> IntPtr.Zero Then
                            Dim newX As Integer = AppBar.Width
                            Dim newY As Integer = AppBar.Location.Y
                            Dim newWidth As Integer = SystemInformation.PrimaryMonitorSize.Width - AppBar.Width - ClockTray.Width
                            Dim newHeight As Integer = AppBar.Height

                            MoveAndResizeWindow(hWnd, newX, newY, newWidth, newHeight)
                        Else
                            MessageBox.Show("Main window for moving cannot be found.", "Info")
                        End If
                    Catch ex As Exception

                    End Try
                End If
            Next
            If prFound = False Then
                MessageBox.Show("KrrShell doesn't find any ""emergeTray"" process running! This only works when that process is running.", "KrrShell Could not find emergeTray", MessageBoxButtons.OK, MessageBoxIcon.Hand)
                CheckBox9.Checked = False
            End If
        Else
            Panel1.Enabled = False
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "SpaceForEmergeTray", "0", Microsoft.Win32.RegistryValueKind.DWord)
            AppBar.Panel4.Visible = True
            AppBar.Width = SystemInformation.PrimaryMonitorSize.Width
            ClockTray.Hide()
            Try
                Dim hWnd As Long
                hWnd = FindWindow("EmergeDesktopApplet", vbNullString)
                Dim wp As WINDOWPLACEMENT
                wp.Length = Marshal.SizeOf(wp)
                GetWindowPlacement(FindWindow("EmergeDesktopApplet", ""), wp)

                Dim wp2 As WINDOWPLACEMENT
                wp2.showCmd = ShowWindowCommands.ShowNoActivate
                wp2.ptMinPosition = wp.ptMinPosition
                wp2.ptMaxPosition = New POINTAPI(0, 0)
                wp2.rcNormalPosition = New RECT(AppBar.Width - AppBar.Panel4.Width, AppBar.Location.Y - 1, AppBar.Width - AppBar.Button2.Width - AppBar.DayLabel.Width - AppBar.ToolStripButton1.Width - AppBar.ToolStripButton3.Width - 3, AppBar.Location.Y + AppBar.Height)

                wp2.flags = wp.flags
                wp2.Length = Marshal.SizeOf(wp2)
                SetWindowPlacement(FindWindow("EmergeDesktopApplet", ""), wp2)
                SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton8.CheckedChanged, RadioButton9.CheckedChanged, RadioButton10.CheckedChanged
        If RadioButton8.Checked = True Then
            SB2.Enabled = False
            SB3.Enabled = False
            AppBar.Button1.BackgroundImage = My.Resources.StartRight
            AppBar.Button1.BackgroundImageLayout = ImageLayout.Stretch
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0", Microsoft.Win32.RegistryValueKind.DWord)

        ElseIf RadioButton9.Checked = True Then
            SB2.Enabled = True
            SB3.Enabled = False
            AppBar.Button1.BackgroundImage = PictureBox4.Image
            If ComboBox4.SelectedIndex = 0 Then
                AppBar.Button1.BackgroundImageLayout = ImageLayout.Stretch
            Else
                AppBar.Button1.BackgroundImageLayout = ImageLayout.Center
            End If
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "1", Microsoft.Win32.RegistryValueKind.DWord)

        ElseIf RadioButton10.Checked = True Then
            SB2.Enabled = False
            SB3.Enabled = True
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "2", Microsoft.Win32.RegistryValueKind.DWord)
        End If
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        On Error Resume Next
        Dim OFD As New OpenFileDialog
        OFD.Title = "Select a Custom Image for ""Default"" button state"
        OFD.CheckFileExists = True
        OFD.FileName = ""
        OFD.Filter = "Image Support Formats (*.png;*.jpg*;*.gif;*.bmp)|*.png;*.jpg*;*.gif;*.bmp|All files (*.*)|*.*"
        If OFD.ShowDialog = DialogResult.OK Then
            PictureBox4.Image = Image.FromFile(OFD.FileName)
            AppBar.Button1.BackgroundImage = Image.FromFile(OFD.FileName)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Normal", OFD.FileName)
        End If
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        On Error Resume Next
        Dim OFD As New OpenFileDialog
        OFD.Title = "Select a Custom Image for ""Hover"" button state"
        OFD.CheckFileExists = True
        OFD.FileName = ""
        OFD.Filter = "Image Support Formats (*.png;*.jpg*;*.gif;*.bmp)|*.png;*.jpg*;*.gif;*.bmp|All files (*.*)|*.*"
        If OFD.ShowDialog = DialogResult.OK Then
            PictureBox5.Image = Image.FromFile(OFD.FileName)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Hover", OFD.FileName)
        End If
    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        On Error Resume Next
        Dim OFD As New OpenFileDialog
        OFD.Title = "Select a Custom Image for ""Pressed"" button state"
        OFD.CheckFileExists = True
        OFD.FileName = ""
        OFD.Filter = "Image Support Formats (*.png;*.jpg*;*.gif;*.bmp)|*.png;*.jpg*;*.gif;*.bmp|All files (*.*)|*.*"
        If OFD.ShowDialog = DialogResult.OK Then
            PictureBox6.Image = Image.FromFile(OFD.FileName)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Pressed", OFD.FileName)
        End If
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked = True Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "Enabled", "1", Microsoft.Win32.RegistryValueKind.DWord)
        Else
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "Enabled", "0", Microsoft.Win32.RegistryValueKind.DWord)
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        If ComboBox4.SelectedIndex = 0 Then
            AppBar.Button1.BackgroundImageLayout = ImageLayout.Stretch
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Layout", "0", Microsoft.Win32.RegistryValueKind.DWord)
        Else
            AppBar.Button1.BackgroundImageLayout = ImageLayout.Center
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Layout", "1", Microsoft.Win32.RegistryValueKind.DWord)
        End If
    End Sub

    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click
        AppBar.SizeLocked = False
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "CustomWidth", NumericUpDown2.Value, Microsoft.Win32.RegistryValueKind.DWord)
        AppBar.Width = NumericUpDown2.Value
        Try
            Dim hWnd As IntPtr = FindWindow("EmergeDesktopApplet", vbNullString)
            If hWnd <> IntPtr.Zero Then
                Dim newX As Integer = AppBar.Width
                Dim newY As Integer = AppBar.Location.Y
                Dim newWidth As Integer = SystemInformation.PrimaryMonitorSize.Width - AppBar.Width - ClockTray.Width
                Dim newHeight As Integer = AppBar.Height

                MoveAndResizeWindow(hWnd, newX, newY, newWidth, newHeight)
            Else
            End If
        Catch ex As Exception
        End Try
        AppBar.SizeLocked = True
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\", "CombineMode", ComboBox5.SelectedIndex, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        If RadioButton1.Checked = True Then AppBar.Opacity = TrackBar1.Value / 100
        NumericUpDown3.Value = TrackBar1.Value
    End Sub

    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged
        If RadioButton1.Checked = True Then AppBar.Opacity = NumericUpDown3.Value / 100
        TrackBar1.Value = NumericUpDown3.Value
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        If CheckBox11.Checked = True Then
            Dim accentColor As Color = GetAccentColor()

            Label12.BackColor = accentColor
            PictureBox1.BackColor = accentColor
            AppBar.BackColor = accentColor

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", ColorTranslator.ToHtml(accentColor))
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "UseSystemColor", 1, Microsoft.Win32.RegistryValueKind.DWord)
        Else
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "UseSystemColor", 0, Microsoft.Win32.RegistryValueKind.DWord)
        End If
    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "Style", CheckBox12.Checked, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub
End Class