Public Class AeroPeekScreen

    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            ' WS_EX_TRANSPARENT (&H20)
            ' WS_EX_LAYERED (&H80000)
            cp.ExStyle = cp.ExStyle Or &H20 Or &H80000 Or &H80L
            Return cp
        End Get
    End Property

    Public Shadows Sub Show()
        Me.Opacity = 0
        MyBase.Show()
        FadeTimer.Start()
    End Sub

    Public Shadows Sub Hide()
        FadeTimer.Stop()
        Me.Opacity = 0
        MyBase.Hide()
    End Sub

    Private Sub FadeTimer_Tick(sender As Object, e As EventArgs) Handles FadeTimer.Tick
        If Me.Opacity < 1.0 Then
            Me.Opacity += 0.1
        Else
            FadeTimer.Stop()
        End If
    End Sub

    Private Sub AeroPeekScreen_Click(sender As Object, e As EventArgs) Handles Me.Click
        Me.Hide()
    End Sub
End Class