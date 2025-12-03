Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class VolumeControl
    Public IsMuted As Boolean = False
    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        Label2.Text = TrackBar1.Value & "%"
        DelayVolSet.Enabled = True
        If TrackBar1.Value = 0 Then
            AppBar.ToolStripButton1.BackgroundImage = My.Resources._121
        ElseIf TrackBar1.Value < 50 AndAlso Not TrackBar1.Value = 0 Then
            AppBar.ToolStripButton1.BackgroundImage = My.Resources._122
        ElseIf TrackBar1.Value >= 50 AndAlso Not TrackBar1.Value = 100 Then
            AppBar.ToolStripButton1.BackgroundImage = My.Resources._123
        ElseIf TrackBar1.Value = 100 Then
            AppBar.ToolStripButton1.BackgroundImage = My.Resources._124
        End If
    End Sub

    'GetVolume missing idk how right now.
    Private Sub VolumeControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Location = New Point(MousePosition.X - Me.Width / 2, MousePosition.Y - Me.Height + 20)
        TrackBar1.Value = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Volume", "LastVol", Nothing)
        CheckBox1.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Volume", "IsMuted", Nothing)
        IsMuted = CheckBox1.Checked
        Label2.Text = TrackBar1.Value & "%"
    End Sub

    Private Sub VolumeControl_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate, Me.LostFocus
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Volume", "LastVol", TrackBar1.Value, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Volume", "IsMuted", CheckBox1.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        Me.Close()
    End Sub

    Private Sub TrackBar1_MouseUp(sender As Object, e As MouseEventArgs) Handles TrackBar1.MouseUp
        Shell("C:\Windows\Setvol.exe " & TrackBar1.Value, AppWinStyle.Hide)
        'AppBar.TextBox1.Text = TrackBar1.Value
        'VNUP.Value = TrackBar1.Value
        Label2.Text = TrackBar1.Value & "%"
        If TrackBar1.Value = 0 Then
            AppBar.ToolStripButton1.BackgroundImage = My.Resources._121
        ElseIf TrackBar1.Value < 50 AndAlso Not TrackBar1.Value = 0 Then
            AppBar.ToolStripButton1.BackgroundImage = My.Resources._122
        ElseIf TrackBar1.Value >= 50 AndAlso Not TrackBar1.Value = 100 Then
            AppBar.ToolStripButton1.BackgroundImage = My.Resources._123
        ElseIf TrackBar1.Value = 100 Then
            AppBar.ToolStripButton1.BackgroundImage = My.Resources._124
        End If
    End Sub

    Private Sub SizeFix_Tick(sender As Object, e As EventArgs) Handles DelayVolSet.Tick
        Shell("C:\Windows\Setvol.exe " & TrackBar1.Value, AppWinStyle.Hide)
        DelayVolSet.Enabled = False
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Shell("C:\Windows\Setvol.exe mute", AppWinStyle.Hide)
            IsMuted = True
        Else
            Shell("C:\Windows\Setvol.exe unmute", AppWinStyle.Hide)
            IsMuted = False
        End If
    End Sub
End Class