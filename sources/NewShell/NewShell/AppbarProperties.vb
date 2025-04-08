Imports System.IO
Imports System.Windows.Forms

Public Class AppbarProperties

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
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
        CheckBox8.Checked = AppBar.TopMost

        GroupBox3.Enabled = CheckBox6.Checked

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

        Label1.BackColor = AppBar.BackColor
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
            'MsgBox("cmd.exe /k C:\Windows\System32\reg.exe delete HKEY_CURRENT_USER\Software\Shell\Appbar\HiddenProcesses\ /v " & Blocklist.SelectedItem)
            Shell("reg.exe delete HKEY_CURRENT_USER\Software\Shell\Appbar\HiddenProcesses\ /v " & Blacklist.SelectedItem & " /f", AppWinStyle.NormalFocus, True)
            Blacklist.Items.Clear()
            For Each ii As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\HiddenProcesses").GetValueNames
                Blacklist.Items.Add(ii)
            Next
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
                'Dim ItemDelete As New ProcessStartInfo("reg.exe")
                'ItemDelete.CreateNoWindow = True
                'ItemDelete.FileName = "reg.exe"
                'ItemDelete.Arguments = "delete HKEY_CURRENT_USER\Software\Shell\Appbar\HiddenProcesses\. /v " & Blocklist.SelectedItems.Item(i) & " /f"
                'Process.Start(ItemDelete)
                Shell("reg.exe delete HKEY_CURRENT_USER\Software\Shell\Appbar\HiddenProcesses\ /v " & Blacklist.SelectedItem & " /f", AppWinStyle.NormalFocus, True)
            Next
            Blacklist.Items.Clear()
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
            Try
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses\.\", Output, "")
                Blocklist.Items.Add(Output)
            Catch ex As Exception

            End Try
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

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim CD As New ColorDialog
        CD.AllowFullOpen = True
        CD.AnyColor = True
        CD.FullOpen = True
        CD.Color = AppBar.BackColor
        If CD.ShowDialog(Me) = DialogResult.OK Then
            Label1.BackColor = CD.Color
            AppBar.BackColor = CD.Color
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
    End Sub
End Class
