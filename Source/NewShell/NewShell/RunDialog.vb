Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Class RunDialog

    <DllImport("shell32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function ExtractIconEx(
    ByVal lpszFile As String,
    ByVal nIconIndex As Integer,
    ByVal phiconLarge() As IntPtr,
    ByVal phiconSmall() As IntPtr,
    ByVal nIcons As UInteger
) As UInteger
    End Function

    <DllImport("user32.dll")>
    Public Shared Function DestroyIcon(ByVal hIcon As IntPtr) As Boolean
    End Function

    Public Shared Function GetIcon(ByVal filePath As String, ByVal iconIndex As Integer, ByVal isSmallIcon As Boolean) As Icon
        Dim hIconLarge(0) As IntPtr
        Dim hIconSmall(0) As IntPtr
        Dim iconToReturn As Icon = Nothing

        Try
            Dim result As UInteger = ExtractIconEx(filePath, iconIndex, hIconLarge, hIconSmall, 1)

            If result > 0 Then
                Dim hIcon As IntPtr = If(isSmallIcon, hIconSmall(0), hIconLarge(0))

                If hIcon <> IntPtr.Zero Then
                    iconToReturn = Icon.FromHandle(hIcon).Clone()

                    DestroyIcon(hIcon)
                End If
            End If

        Catch ex As Exception
            Return Nothing
        End Try

        Return iconToReturn
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Run(True)
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        TabControl1.Visible = CheckBox1.Checked
        If CheckBox1.Checked = True Then
            If Me.Location.Y - Me.MaximumSize.Height < 0 Then
                Me.Size = Me.MaximumSize
            Else
                Me.Size = Me.MaximumSize
                Me.Location = New Point(Me.Location.X, Me.Location.Y - Me.MaximumSize.Height + Me.MinimumSize.Height)

            End If
        Else
            If Me.Location.Y - Me.MaximumSize.Height < 0 Then
                Me.Size = Me.MinimumSize
            Else
                Me.Size = Me.MinimumSize
                Me.Location = New Point(Me.Location.X, Me.Location.Y + Me.MaximumSize.Height - Me.MinimumSize.Height)
            End If
        End If
    End Sub

    Public Function IsProgramRunningAsAdmin() As Boolean
        Return My.User.IsInRole(ApplicationServices.BuiltInRole.Administrator)
    End Function

    Private Sub RunDialog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Location = New Point(0, SystemInformation.WorkingArea.Height - Me.MinimumSize.Height)
        Me.Activate()

        ComboBox3.SelectedIndex = 0

        If IsProgramRunningAsAdmin() = True Then
            CheckBox6.Visible = False
            Label3.Visible = True

            Label3.Image = GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\system32\imageres.dll", 73, True).ToBitmap
        Else
            CheckBox6.Visible = True
            Label3.Visible = False
        End If

        Try
            PictureBox1.Image = GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\system32\imageres.dll", 95, False).ToBitmap
            Me.Icon = GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\system32\imageres.dll", 95, False)
        Catch ex As Exception
            PictureBox1.Image = My.Resources.RunDialogIcon
        End Try

        If CheckBox1.Checked = True Then
            Me.Size = Me.MaximumSize
        Else
            Me.Size = Me.MinimumSize
        End If

        ' To first select the Textbox
        ComboBox1.Focus()
        ComboBox1.Select()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If OFD.ShowDialog(Me) = DialogResult.OK Then
            ComboBox1.Text = $"""{OFD.FileName.Trim}"""
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If FBD.ShowDialog(Me) = DialogResult.OK Then
            ComboBox1.Text = $"""{FBD.SelectedPath.Trim}"""
        End If
    End Sub

    Public Sub Run(ByVal WithClose As Boolean)

        ' The Run starts here!

        Try
            Dim FirstSpaceIndex As Integer = ComboBox1.Text.Trim.IndexOf(" ")

            If CheckBox1.Checked = True Then
                Dim SP As New ProcessStartInfo(ComboBox1.Text.Trim)

                SP.UseShellExecute = True

                'RunAs
                If CheckBox4.Checked = True Then
                    SP.UserName = TextBox6.Text
                    SP.PasswordInClearText = MaskedTextBox1.Text
                End If

                'Basic
                Dim ProgramPath As String
                Dim Arguments As String

                If FirstSpaceIndex > -1 Then
                    ProgramPath = ComboBox1.Text.Trim.Substring(0, FirstSpaceIndex)
                    Arguments = ComboBox1.Text.Trim.Substring(FirstSpaceIndex + 1).TrimStart()
                Else
                    ProgramPath = ComboBox1.Text.Trim
                    Arguments = String.Empty
                End If

                SP.FileName = ProgramPath
                SP.Arguments = Arguments

                If CheckBox6.Checked = True Then SP.Verb = "runas"
                'Options
                SP.WindowStyle = ComboBox3.SelectedIndex
                SP.CreateNoWindow = CheckBox7.Checked
                SP.LoadUserProfile = CheckBox3.Checked

                'Debuging - Maybe
                SP.ErrorDialog = CheckBox12.Checked
                SP.ErrorDialogParentHandle = CheckBox13.Checked
                SP.RedirectStandardError = CheckBox14.Checked
                SP.RedirectStandardInput = CheckBox15.Checked
                SP.RedirectStandardOutput = CheckBox17.Checked

                'Custom "Working Directory"

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
                Dim ProgramPath As String
                Dim Arguments As String

                If FirstSpaceIndex > -1 Then
                    ProgramPath = ComboBox1.Text.Trim.Substring(0, FirstSpaceIndex)
                    Arguments = ComboBox1.Text.Trim.Substring(FirstSpaceIndex + 1).TrimStart()
                Else
                    ProgramPath = ComboBox1.Text.Trim
                    Arguments = String.Empty
                End If

                Process.Start(ProgramPath, Arguments)
                If WithClose = True Then Me.Close()

            End If
        Catch ex As Exception
            MsgBox($"Failed to execute: ""{ComboBox1.Text.Trim}"". " & ex.Message, MsgBoxStyle.Critical, "Error while Starting a new Process")

            ComboBox1.Focus()
            ComboBox1.Select()
            ComboBox1.SelectAll()
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Run(False)
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        GroupBox1.Enabled = CheckBox4.Checked
        If CheckBox4.Checked = True Then
            CheckBox6.Checked = False
            CheckBox6.Enabled = False
        Else
            CheckBox6.Enabled = True
        End If
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

    Private Sub RRemover_Tick(sender As Object, e As EventArgs) Handles RRemover.Tick
        ' Removes the "R" at start

        ComboBox1.Text = ""
        RRemover.Enabled = False
    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            If OK_Button.Enabled = True Then
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Run(True)
            End If
        End If
    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        If Not ComboBox1.Text = String.Empty Then
            OK_Button.Enabled = True
            Button1.Enabled = True
        Else
            OK_Button.Enabled = False
            Button1.Enabled = False
        End If
    End Sub
End Class
