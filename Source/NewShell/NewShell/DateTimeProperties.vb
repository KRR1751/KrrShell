Imports System.Windows.Forms

Public Class DateTimeProperties

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSecondsInSystemClock", CheckBox1.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\DateTime", "ShowTime", CheckBox2.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\DateTime", "ShowDay", CheckBox4.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\DateTime", "ShowDate", CheckBox3.Checked, Microsoft.Win32.RegistryValueKind.DWord)

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        AppBar.SC_isSecondsEnabled = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSecondsInSystemClock", False)
        AppBar.SC_ShowTime = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\DateTime", "ShowTime", True)
        AppBar.SC_ShowDay = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\DateTime", "ShowDay", True)
        AppBar.SC_ShowDate = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\DateTime", "ShowDate", True)

        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DateTimeProperties_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckBox1.Checked = AppBar.SC_isSecondsEnabled
        CheckBox2.Checked = AppBar.SC_ShowTime
        CheckBox4.Checked = AppBar.SC_ShowDay
        CheckBox3.Checked = AppBar.SC_ShowDate
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        AppBar.SC_isSecondsEnabled = CheckBox1.Checked
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        AppBar.SC_ShowTime = CheckBox2.Checked
        CheckBox1.Enabled = CheckBox2.Checked
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        AppBar.SC_ShowDate = CheckBox3.Checked
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        AppBar.SC_ShowDay = CheckBox4.Checked
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        On Error Resume Next
        Process.Start(New ProcessStartInfo(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "TimeDateProperties", Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\timedate.cpl")) With {.UseShellExecute = True})
    End Sub
End Class
