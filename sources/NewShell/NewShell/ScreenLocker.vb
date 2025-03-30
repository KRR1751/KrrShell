Public Class ScreenLocker
    Dim IsBClose As Boolean = False
    Private Sub ScreenLocker_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click
        Me.BringToFront()
    End Sub

    Private Sub ScreenLocker_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If IsBClose = False Then
            e.Cancel = True
        Else
            IsBClose = False
            e.Cancel = False
        End If
    End Sub

    Private Sub ScreenLocker_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Location = New Point(0, 0)
        Me.Size = New Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
        'Me.BackgroundImage = New Bitmap(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
        Me.BringToFront()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Cursor.Show()
        IsBClose = True
        Me.Close()
    End Sub

    Private Sub ScreenLocker_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        If e.Location = New Point(0, 0) Then
            Button1.Visible = True
        Else
            Button1.Visible = False
        End If
    End Sub
End Class