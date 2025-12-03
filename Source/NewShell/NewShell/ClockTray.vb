Imports System.Reflection.Emit
Imports System.Threading

Public Class ClockTray
    Private Sub StartClockUpdate()
        Task.Run(Sub()
                     While True
                         Dim currentTime As String = DateTime.Now.ToString("HH:mm:ss")
                         Dim currentDay As String = DateTime.Now.DayOfWeek.ToString
                         Dim currentDate As String = DateTime.Now.ToString("dd. MM. yyyy")

                         If Me.InvokeRequired Then
                             Me.Invoke(Sub()
                                           TimeLabel.Text = currentTime
                                           DayLabel.Text = currentDay
                                           DateLabel.Text = currentDate
                                       End Sub)
                         Else
                             TimeLabel.Text = currentTime
                             DayLabel.Text = currentDay
                             DateLabel.Text = currentDate
                         End If

                         Thread.Sleep(1000)
                     End While
                 End Sub)
    End Sub
    Private Sub Controller_Tick(sender As Object, e As EventArgs) ' Handles Controller.Tick
        TimeLabel.Text = DateTime.Now.ToString("HH:mm:ss")
        DayLabel.Text = DateTime.Now.DayOfWeek.ToString
        DateLabel.Text = DateTime.Now.ToString("dd. MM. yyyy")
    End Sub

    Private Sub ClockTray_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Location = New Point(SystemInformation.PrimaryMonitorSize.Width - Me.Width, SystemInformation.PrimaryMonitorSize.Height - Me.Height)
        StartClockUpdate()
    End Sub

    Private Sub ClockTray_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
        If Not Me.Location = New Point(SystemInformation.PrimaryMonitorSize.Width - Me.Width, SystemInformation.PrimaryMonitorSize.Height - Me.Height) Then
            Me.Location = New Point(SystemInformation.PrimaryMonitorSize.Width - Me.Width, SystemInformation.PrimaryMonitorSize.Height - Me.Height)
        End If
    End Sub

    Private Sub ClockTray_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If AppBar.CanClose = False Then
            e.Cancel = True
            SA.ShowDialog()
        Else
            If GlobalKeyboardHook.Unhook() Then
                Debug.WriteLine("Hook ended.")
            Else
                Debug.WriteLine("Hook ended failed.")
            End If
        End If
    End Sub
End Class