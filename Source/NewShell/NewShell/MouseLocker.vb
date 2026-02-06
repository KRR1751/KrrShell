Public Class MLf
    Dim c As Boolean = False
    Private Sub MLf_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If c = False Then
            e.Cancel = True
        Else
            c = False
            e.Cancel = False
        End If

    End Sub

    Private Sub MLf_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = "c"c Then
            TimeDate.MouseLocker.Enabled = False
            TimeDate.TPauseB.Enabled = False
            TimeDate.TPauseB.Text = "Pause"
            TimeDate.TsT.Enabled = False
            TimeDate.Td = 0
            TimeDate.Th = 0
            TimeDate.Tm = 0
            TimeDate.Ts = 0
            TimeDate.GroupBox4.Enabled = True
            If TimeDate.CheckBox15.Checked = True Then
                TimeDate.GroupBox7.Enabled = True
            End If
            TimeDate.CheckBox15.Enabled = True
            TimeDate.alarm.Stop()
            TimeDate.TStartB.Text = "Start"
            c = True
            Me.Close()
        End If
    End Sub

    Private Sub MouseLocker_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Size = Cursor.Size
        Me.Location = New Point(Cursor.Position.X - Me.Width / 2, Cursor.Position.Y - Me.Height / 2)
        Me.BringToFront()
    End Sub
    Dim clc As Integer = 0

    Private Sub MouseLocker_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LostFocus
        Me.BringToFront()
    End Sub
    Private Sub MouseLocker_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Right Then
            If clc = 3 Then
                TimeDate.MouseLocker.Enabled = False
                TimeDate.TPauseB.Enabled = False
                TimeDate.TPauseB.Text = "Pause"
                TimeDate.TsT.Enabled = False
                TimeDate.Td = 0
                TimeDate.Th = 0
                TimeDate.Tm = 0
                TimeDate.Ts = 0
                TimeDate.GroupBox4.Enabled = True
                If TimeDate.CheckBox15.Checked = True Then
                    TimeDate.GroupBox7.Enabled = True
                End If
                TimeDate.CheckBox15.Enabled = True
                TimeDate.alarm.Stop()
                TimeDate.TStartB.Text = "Start"
                Me.Close()
            Else
                Reseter.Enabled = True
                clc += 1
            End If
        End If
    End Sub

    Private Sub Reseter_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Reseter.Tick
        clc = 0
        Reseter.Enabled = False
    End Sub
End Class