Imports System.Windows.Forms

Public Class RunDialog

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Run()
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            If Me.Location.Y + 335 > SystemInformation.WorkingArea.Height Then
                Dim FixY As Integer = Me.Location.Y + 335 - SystemInformation.WorkingArea.Height
                Me.Location = New Point(Me.Location.X, Me.Location.Y - FixY)
                Me.Size = New Size(404, 335)
            Else
                Me.Size = New Size(404, 335)
            End If
        Else
            Me.Size = New Size(404, 158)
        End If
    End Sub

    Private Sub RunDialog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Location = New Point(0, SystemInformation.WorkingArea.Height - 158)
        ComboBox2.SelectedIndex = 1
        ComboBox3.SelectedIndex = 0
        If CheckBox1.Checked = True Then
            Me.Size = New Size(404, 335)
        Else
            Me.Size = New Size(404, 158)
        End If

        'To first select the Textbox
        ComboBox1.Focus()
        ComboBox1.Select()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If OFD.ShowDialog(Me) = DialogResult.OK Then
            ComboBox1.Text = OFD.FileName
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If FBD.ShowDialog(Me) = DialogResult.OK Then
            ComboBox1.Text = FBD.SelectedPath
        End If
    End Sub

    Public Sub Run()
        If RadioButton1.Checked = True Then
            Try
                If CheckBox1.Checked = True Then
                    Dim SP As New ProcessStartInfo(ComboBox1.Text)

                    'RunAs
                    If CheckBox4.Checked = True Then
                        SP.UserName = TextBox6.Text
                        SP.PasswordInClearText = MaskedTextBox1.Text
                    End If

                    'Basic
                    SP.Arguments = TextBox5.Text
                    SP.WindowStyle = ComboBox3.SelectedIndex
                    SP.CreateNoWindow = CheckBox7.Checked
                    SP.UseShellExecute = CheckBox8.Checked

                    'Debuging - Maybe
                    SP.ErrorDialog = CheckBox12.Checked
                    SP.ErrorDialogParentHandle = CheckBox13.Checked
                    SP.RedirectStandardError = CheckBox14.Checked
                    SP.RedirectStandardInput = CheckBox15.Checked
                    SP.RedirectStandardOutput = CheckBox17.Checked

                    'Working Custom Directory

                    If CheckBox10.Checked = True Then
                        SP.WorkingDirectory = TextBox7.Text
                    End If

                    'Custom "Domain"
                    If CheckBox11.Checked = True Then
                        SP.Domain = TextBox8.Text
                    End If

                    'Actual Starting the Process
                    Process.Start(SP)
                Else
                    Process.Start(ComboBox1.Text)
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error while Starting a new Process")
            End Try
        Else
            Try
                If CheckBox1.Checked = True Then
                    Shell(ComboBox1.Text, ComboBox2.SelectedIndex, CheckBox2.Checked, NUD1.Value)
                Else
                    Shell(ComboBox1.Text)
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error while Starting a Program")
            End Try
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        P1.Visible = RadioButton1.Checked
        Dim si As New ProcessStartInfo
        ' si.WindowStyle.  
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        P2.Visible = RadioButton2.Checked
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        NUD1.Enabled = CheckBox2.Checked
        If CheckBox2.Checked = False Then
            NUD1.Value = -1
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Run()
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        TextBox6.Enabled = CheckBox4.Checked
        MaskedTextBox1.Enabled = CheckBox4.Checked
        CheckBox5.Enabled = CheckBox4.Checked
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            MaskedTextBox1.PasswordChar = ""
        Else
            MaskedTextBox1.PasswordChar = "●"
        End If
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        TextBox7.Enabled = CheckBox10.Checked
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        TextBox8.Enabled = CheckBox11.Checked
    End Sub
End Class
