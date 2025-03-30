Public Class NW
    Public MODE As Integer = 3
    Dim c As Boolean = False
    Private Sub NW_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If c = False Then
            e.Cancel = True
        Else
            c = False
            e.Cancel = False
        End If
    End Sub
    Dim alarm As New Media.SoundPlayer
    Private Sub NW_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If MODE = 1 Then
            Me.Height = 215
            B1.Visible = True
            B1.Enabled = True
            B2.Visible = False
            B2.Enabled = False
            B3.Visible = False
            B3.Enabled = False
        ElseIf MODE = 2 Then
            Me.Height = 215
            B1.Visible = False
            B1.Enabled = False
            B2.Visible = True
            B2.Enabled = True
            B2.Location = New Point(12, 156)
            B3.Visible = True
            B3.Enabled = True
            B3.Location = New Point(184, 156)
        ElseIf MODE = 3 Then
            Me.Height = 245
            B1.Visible = True
            B1.Enabled = True
            B2.Visible = True
            B2.Enabled = True
            B2.Location = New Point(12, 185)
            B3.Visible = True
            B3.Enabled = True
            B3.Location = New Point(184, 185)
        End If
        If NotifyWindowD.ComboBox1.SelectedIndex = 1 Then
            Me.Text = NotifyWindowD.TextBox1.Text
        End If
        If NotifyWindowD.ComboBox2.SelectedIndex = 0 Then
            'Me.Icon = New Icon("C:\Windows\NotificationWindow\NotifyIcon.ico")
        Else
            If NotifyWindowD.ComboBox2.Text.Length = 0 Then
                'Me.Icon = New Icon("C:\Windows\NotificationWindow\NotifyIcon.ico")
            Else
                Me.Icon = New Icon(NotifyWindowD.ComboBox3.Text)
            End If
        End If
        If NotifyWindowD.CheckBox1.Checked = True Then
            If NotifyWindowD.RadioButton1.Checked = True Then
                If NotifyWindowD.CheckBox2.Checked = True Then
                    B1.BackColor = NotifyWindowD.CL.BackColor
                    B2.BackColor = NotifyWindowD.CL.BackColor
                    B3.BackColor = NotifyWindowD.CL.BackColor
                    TL.BackColor = NotifyWindowD.CL.BackColor
                    Me.BackColor = NotifyWindowD.CL.BackColor
                Else
                    B1.BackColor = SystemColors.Control
                    B2.BackColor = SystemColors.Control
                    B3.BackColor = SystemColors.Control
                    TL.BackColor = SystemColors.Control
                    Me.BackColor = NotifyWindowD.CL.BackColor
                End If
            Else
                If NotifyWindowD.CheckBox2.Checked = True Then
                    B1.BackColor = Color.LightBlue
                    B2.BackColor = Color.LightBlue
                    B3.BackColor = Color.LightBlue
                    TL.BackColor = Color.LightBlue
                    Me.BackColor = Color.LightBlue
                Else
                    B1.BackColor = SystemColors.Control
                    B2.BackColor = SystemColors.Control
                    B3.BackColor = SystemColors.Control
                    TL.BackColor = SystemColors.Control
                    Me.BackColor = Color.LightBlue
                End If
            End If
        End If
        If NotifyWindowD.ComboBox4.SelectedIndex = 0 Then
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
        Else
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedToolWindow
        End If
        Me.TopMost = NotifyWindowD.CheckBox3.Checked
        Me.ControlBox = NotifyWindowD.CheckBox4.Checked
        If NotifyWindowD.CheckBox5.Checked = True Then

        End If
        If NotifyWindowD.CheckBox7.Checked = True Then
            B1.ForeColor = NotifyWindowD.CL2.BackColor
            B2.ForeColor = NotifyWindowD.CL2.BackColor
            B3.ForeColor = NotifyWindowD.CL2.BackColor
            TL.ForeColor = NotifyWindowD.CL2.BackColor
            Me.ForeColor = NotifyWindowD.CL2.BackColor
        Else
            B1.ForeColor = SystemColors.ControlText
            B2.ForeColor = SystemColors.ControlText
            B3.ForeColor = SystemColors.ControlText
            TL.ForeColor = SystemColors.ControlText
            Me.ForeColor = SystemColors.ControlText
        End If
        If Dialog3.ComboBox5.SelectedIndex = 0 Then
            Dim RV16 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "EnableRinging", Nothing)
            Dim RV17 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingMode", Nothing)
            Dim RV18 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingALIST", Nothing)
            Dim RV19 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingPath", Nothing)
            Dim RV1a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingSYS", Nothing)
            Dim RV1b = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingRepeat", Nothing)
            Dim RV1c = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingTimesRepeat", Nothing)
            Dim RV1d = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingVol", Nothing)
            If RV16 = "1" Then
                If RV17 = "0" Then
                    If RV1a = "0" Then
                        Media.SystemSounds.Asterisk.Play()
                    ElseIf RV1a = "1" Then
                        Media.SystemSounds.Beep.Play()
                    ElseIf RV1a = "2" Then
                        Media.SystemSounds.Exclamation.Play()
                    ElseIf RV1a = "3" Then
                        Media.SystemSounds.Hand.Play()
                    ElseIf RV1a = "4" Then
                        Media.SystemSounds.Question.Play()
                    End If
                ElseIf RV17 = "1" Then
                    AlCol.Items.Clear()
                    Dim i As String
                    If My.Computer.FileSystem.DirectoryExists("C:\Windows\Media\Alarms\") Then
                        For Each i In My.Computer.FileSystem.GetFiles("C:\Windows\Media\Alarms\")
                            AlCol.Items.Add(i.Substring(i.LastIndexOf("\") + 1))
                        Next
                    End If
                    Dim alrpath As String = "C:\Windows\Media\Alarms\" & AlCol.Items.Item(RV18)
                    alarm = New Media.SoundPlayer(alrpath)
                    If RV1b = "1" Then
                        alarm.PlayLooping()
                    Else
                        'if RingTimesRepeat
                        alarm.Play()
                    End If
                ElseIf RV17 = "2" Then
                    If IO.File.Exists(RV19) = True Then
                        Dim alrpath As String = RV19
                        alarm = New Media.SoundPlayer(alrpath)
                        If RV1b = "1" Then
                            alarm.PlayLooping()
                        Else
                            'if RingTimesRepeat
                            alarm.Play()
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub B1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles B1.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        alarm.Stop()
        c = True
        'Me.Close()
    End Sub

    Private Sub B2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles B2.Click
        Me.DialogResult = Windows.Forms.DialogResult.Yes
        alarm.Stop()
        c = True
        'Me.Close()
    End Sub

    Private Sub B3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles B3.Click
        Me.DialogResult = Windows.Forms.DialogResult.No
        alarm.Stop()
        c = True
        'Me.Close()
    End Sub
End Class