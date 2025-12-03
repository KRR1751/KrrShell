Imports System.Reflection.Emit
Imports System.Threading
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar

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

        SyncAppearance()

        StartClockUpdate()
    End Sub

    Public Sub SyncAppearance()

        Me.BackColor = AppBar.BackColor
        Me.ForeColor = AppBar.ForeColor
        Me.Font = AppBar.Font

        If AppbarProperties.ComboBox2.Text IsNot Nothing AndAlso IO.File.Exists(AppbarProperties.ComboBox2.Text) Then
            Me.BackgroundImage = Image.FromFile(AppbarProperties.ComboBox2.Text)
        Else
            Me.BackgroundImage = My.Resources.AppBarMainTransparent1
        End If

        Me.BackgroundImageLayout = AppBar.BackgroundImageLayout

        Me.Opacity = AppBar.Opacity
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Desktop.BringToFront()
    End Sub

    Private Sub ClockTray_BackColorChanged(sender As Object, e As EventArgs) Handles Me.BackColorChanged
        If Me.BackColor.R >= 150 OrElse Me.BackColor.G >= 130 AndAlso Me.BackColor.B >= 130 Then
            Me.ForeColor = Color.Black
            TimeLabel.ForeColor = Color.Black
            DayLabel.ForeColor = Color.Black
            DateLabel.ForeColor = Color.Black
            Button1.FlatAppearance.BorderColor = Color.White
        Else
            Me.ForeColor = Color.White
            TimeLabel.ForeColor = Color.White
            DayLabel.ForeColor = Color.White
            DateLabel.ForeColor = Color.White
            Button1.FlatAppearance.BorderColor = Color.Black
        End If
    End Sub
End Class