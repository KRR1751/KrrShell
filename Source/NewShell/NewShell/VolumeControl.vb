Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class VolumeControl
    Public IsMuted As Boolean = False
    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        Label2.Text = TrackBar1.Value & "%"

        If CheckBox2.Checked = True Then
            DelayVolSet.Enabled = True

            If TrackBar1.Value = 0 Then
                AppBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeZero

            ElseIf TrackBar1.Value < 50 AndAlso Not TrackBar1.Value = 0 Then
                AppBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeBelow50

            ElseIf TrackBar1.Value >= 50 AndAlso Not TrackBar1.Value = 100 Then
                AppBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeAbove50.ToBitmap

            ElseIf TrackBar1.Value = 100 Then
                AppBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeFull

            End If
        End If
    End Sub
    Dim SetVolPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\SetVol.exe"
    Private Sub VolumeControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim success As Boolean = False
        If File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\SetVol.exe") Then
            success = True
            SetVolPath = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\SetVol.exe"

        ElseIf File.Exists(Application.StartupPath & "\SetVol\SetVol.exe") Then
            success = True
            SetVolPath = Application.StartupPath & "\SetVol\SetVol.exe"
            Console.WriteLine("File not detected in %windir%, the shell may have problems with SetVol as it is recommended to move it there.")

        Else
            success = False
            MsgBox($"Failed to locate ""{SetVolPath}"". Please check if you have SetVol.exe installed. Otherwise you cannot change any System Volume from this shell.", MsgBoxStyle.Critical, "File for Volume Changing is missing")
        End If

        If success = True Then
            Me.Location = New Point(MousePosition.X - Me.Width / 2, MousePosition.Y - Me.Height + 20)
            Me.BringToFront()
            TrackBar1.Value = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Volume", "LastVol", 50)
            CheckBox1.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Volume", "IsMuted", 0)
            CheckBox2.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Volume", "RealTimeChanging", 1)
            IsMuted = CheckBox1.Checked
            Label2.Text = TrackBar1.Value & "%"
        Else
            Me.Close()
        End If
    End Sub

    Private Sub VolumeControl_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Volume", "LastVol", TrackBar1.Value, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Volume", "IsMuted", CheckBox1.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Volume", "RealTimeChanging", CheckBox2.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        Me.Close()
    End Sub

    Private Sub TrackBar1_MouseUp(sender As Object, e As MouseEventArgs) Handles TrackBar1.MouseUp
        Dim PSI As New ProcessStartInfo With {
        .FileName = SetVolPath,
        .Arguments = TrackBar1.Value,
        .CreateNoWindow = True
        }

        Process.Start(PSI)
        Label2.Text = TrackBar1.Value & "%"

        If TrackBar1.Value = 0 Then
            AppBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeZero

        ElseIf TrackBar1.Value < 50 AndAlso Not TrackBar1.Value = 0 Then
            AppBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeBelow50

        ElseIf TrackBar1.Value >= 50 AndAlso Not TrackBar1.Value = 100 Then
            AppBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeAbove50.ToBitmap

        ElseIf TrackBar1.Value = 100 Then
            AppBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeFull

        End If
    End Sub

    Private Sub SizeFix_Tick(sender As Object, e As EventArgs) Handles DelayVolSet.Tick
        Dim PSI As New ProcessStartInfo With {
        .FileName = SetVolPath,
        .Arguments = TrackBar1.Value,
        .CreateNoWindow = True
        }

        Process.Start(PSI)
        DelayVolSet.Enabled = False
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Dim PSI As New ProcessStartInfo With {
        .FileName = SetVolPath,
        .CreateNoWindow = True
        }

        If CheckBox1.Checked = True Then
            PSI.Arguments = "mute"
            IsMuted = True
        Else
            PSI.Arguments = "unmute"
            IsMuted = False
        End If

        Process.Start(PSI)

        If IsMuted = True Then
            AppBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeMute
        Else
            If TrackBar1.Value = 0 Then
                AppBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeZero

            ElseIf TrackBar1.Value < 50 AndAlso Not TrackBar1.Value = 0 Then
                AppBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeBelow50

            ElseIf TrackBar1.Value >= 50 AndAlso Not TrackBar1.Value = 100 Then
                AppBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeAbove50.ToBitmap

            ElseIf TrackBar1.Value = 100 Then
                AppBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeFull

            End If
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Button1.Enabled = False
        Else
            Button1.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim PSI As New ProcessStartInfo With {
        .FileName = SetVolPath,
        .Arguments = TrackBar1.Value,
        .CreateNoWindow = True
        }

        Process.Start(PSI)
    End Sub
End Class