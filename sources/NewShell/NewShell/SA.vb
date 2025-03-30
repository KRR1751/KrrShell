Public Class SA

    Private Sub SA_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Topper.Enabled = False
    End Sub

    Private Sub SA_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate, Me.LostFocus
        Me.Activate()
        Me.BringToFront()
    End Sub

    Private Sub SA_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        BlackScreen.Close()
        Topper.Enabled = False
    End Sub

    Private Sub SA_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        BlackScreen.Show()
        BlackScreen.BringToFront()
        Topper.Enabled = True
        Me.Focus()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        MessageBox.Show("Simply select what action the computer will do by your selection. ;)", "Help", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Public Sub SaveSettings()
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "AutoHide", AppBar.AutoHideToolStripMenuItem.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer1", AppBar.Button1.Width, Microsoft.Win32.RegistryValueKind.String)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer2", AppBar.Panel1.Width, Microsoft.Win32.RegistryValueKind.String)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer3", AppBar.Panel4.Width, Microsoft.Win32.RegistryValueKind.String)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", ColorTranslator.ToHtml(AppBar.BackColor), Microsoft.Win32.RegistryValueKind.String)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        On Error GoTo 0
        Dim F As String = String.Empty
        AppBar.CanClose = True
        Desktop.CanClose = True
        If CheckBox1.Checked = True Then
            F = "/f "
        End If
        If RadioButton1.Checked = True Then
            SaveSettings()
            Shell("shutdown.exe " & F & "/s /t 0", AppWinStyle.Hide)
            Me.Close()
        ElseIf RadioButton2.Checked = True Then
            SaveSettings()
            Shell("shutdown.exe " & F & "/s /hybrid /t 0", AppWinStyle.Hide)
            Me.Close()
        ElseIf RadioButton3.Checked = True Then
            SaveSettings()
            Shell("shutdown.exe " & F & "/r /t 0", AppWinStyle.Hide)
            Me.Close()
        ElseIf RadioButton4.Checked = True Then
            SaveSettings()
            Shell("RunDll32.exe powrprof.dll,SetSuspendState", AppWinStyle.Hide)
            Me.Close()
        ElseIf RadioButton5.Checked = True Then
            SaveSettings()
            Shell("shutdown.exe " & F & "/h", AppWinStyle.Hide)
            Me.Close()
        Else
0:
            MessageBox.Show("Failed to process this action, Please try that again later.", "Microsoft Windows", MessageBoxButtons.OK)
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Topper_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Topper.Tick
        Me.BringToFront()
    End Sub
End Class