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
            Dialog3.MouseLocker.Enabled = False
            Dialog3.TPauseB.Enabled = False
            Dialog3.TPauseB.Text = "Pause"
            Dialog3.TsT.Enabled = False
            Dialog3.Td = 0
            Dialog3.Th = 0
            Dialog3.Tm = 0
            Dialog3.Ts = 0
            Dialog3.GroupBox4.Enabled = True
            If Dialog3.CheckBox15.Checked = True Then
                Dialog3.GroupBox7.Enabled = True
            End If
            Dialog3.CheckBox15.Enabled = True
            Dialog3.alarm.Stop()
            Dialog3.TStartB.Text = "Start"
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
                Dialog3.MouseLocker.Enabled = False
                Dialog3.TPauseB.Enabled = False
                Dialog3.TPauseB.Text = "Pause"
                Dialog3.TsT.Enabled = False
                Dialog3.Td = 0
                Dialog3.Th = 0
                Dialog3.Tm = 0
                Dialog3.Ts = 0
                Dialog3.GroupBox4.Enabled = True
                If Dialog3.CheckBox15.Checked = True Then
                    Dialog3.GroupBox7.Enabled = True
                End If
                Dialog3.CheckBox15.Enabled = True
                Dialog3.alarm.Stop()
                Dialog3.TStartB.Text = "Start"
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