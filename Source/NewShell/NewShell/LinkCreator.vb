Imports System.Windows.Forms

Public Class LinkCreator
    Public PATH As String = String.Empty
    Public Result As String = String.Empty

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim targetFile As String = ComboBox1.Text.Trim()
        Dim shortcutPath As String = PATH.Trim & "\" & TextBox1.Text

        If String.IsNullOrWhiteSpace(targetFile) OrElse String.IsNullOrWhiteSpace(shortcutPath) Then
            MessageBox.Show("You didn't entered any program path!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If Not System.IO.File.Exists(targetFile) Then
            MessageBox.Show("Target File not found. Please check the file if it exists.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If Not shortcutPath.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase) Then
            shortcutPath &= ".lnk"
        End If

        Try
            Dim wshShell As Object = CreateObject("WScript.Shell")

            Dim shortcut As Object = wshShell.CreateShortcut(shortcutPath)

            shortcut.TargetPath = targetFile

            If CheckBox2.Checked = True Then

                If IO.Directory.Exists(TextBox2.Text) Then shortcut.WorkingDirectory = TextBox2.Text Else shortcut.WorkingDirectory = System.IO.Path.GetDirectoryName(targetFile)
                If PictureBox1.Tag IsNot Nothing Then shortcut.IconLocation = PictureBox1.Tag

                shortcut.Description = TextBox3.Text
            End If

            shortcut.Save()

            System.Runtime.InteropServices.Marshal.ReleaseComObject(shortcut)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wshShell)
            shortcut = Nothing
            wshShell = Nothing

            MessageBox.Show($"Link has been created successfully: {shortcutPath}", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MessageBox.Show($"Error while creating Link: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Public Sub OldLinkCreator()
        If IO.File.Exists(ComboBox1.Text) = True Then
            Shell("cmd.exe /c mklink " & """" & PATH & TextBox1.Text & """ """ & ComboBox1.Text & """", AppWinStyle.Hide, True)
            Result = TextBox1.Text
        ElseIf IO.Directory.Exists(ComboBox1.Text) = True Then
            Shell("cmd.exe /c mklink /D " & """" & PATH & TextBox1.Text & """ """ & ComboBox1.Text & """", AppWinStyle.Hide, True)
            Result = TextBox1.Text
        End If
        'Shell("mklink ")

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub ComboBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.TextChanged, ComboBox2.TextChanged, TextBox1.TextChanged, CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            If TextBox1.Text.Contains("*") OrElse TextBox1.Text.Contains("?") OrElse TextBox1.Text.Contains("""") OrElse TextBox1.Text.Contains(":") OrElse TextBox1.Text.Contains("/") OrElse TextBox1.Text.Contains("\") OrElse TextBox1.Text.Contains("<") OrElse TextBox1.Text.Contains(">") OrElse TextBox1.Text.Contains("|") Then
                OK_Button.Enabled = False
            Else
                If IO.File.Exists(ComboBox1.Text) = True Then
                    OK_Button.Enabled = True
                    CheckBox1.Enabled = True
                ElseIf IO.Directory.Exists(ComboBox1.Text) = True Then
                    OK_Button.Enabled = True
                    CheckBox1.Enabled = True
                Else
                    OK_Button.Enabled = False
                    TextBox1.Enabled = False
                    CheckBox1.Enabled = False
                    CheckBox1.Checked = False
                End If
            End If
        Else
            If IO.File.Exists(ComboBox1.Text) = True Then
                OK_Button.Enabled = True
                CheckBox1.Enabled = True
            ElseIf IO.Directory.Exists(ComboBox1.Text) = True Then
                OK_Button.Enabled = True
                CheckBox1.Enabled = True
            Else
                OK_Button.Enabled = False
                TextBox1.Enabled = False
                CheckBox1.Enabled = False
                CheckBox1.Checked = False
            End If
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox1.Enabled = True
        Else
            TextBox1.Enabled = False
            If IO.File.Exists(ComboBox1.Text) = True Then
                Dim FI As New IO.FileInfo(ComboBox1.Text)
                TextBox1.Text = FI.Name
            ElseIf IO.Directory.Exists(ComboBox1.Text) = True Then
                Dim DI As New IO.DirectoryInfo(ComboBox1.Text)
                TextBox1.Text = DI.Name
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        cm.Show(Control.MousePosition)
    End Sub

    Private Sub LinkCreator_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

    End Sub

    Private Sub LinkCreator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LinkLabel1.Text = PATH

        If PATH = String.Empty Then
            OK_Button.Enabled = False
        Else
            If CheckBox1.Checked = True Then
                If TextBox1.Text.Contains("*") OrElse TextBox1.Text.Contains("?") OrElse TextBox1.Text.Contains("""") OrElse TextBox1.Text.Contains(":") OrElse TextBox1.Text.Contains("/") OrElse TextBox1.Text.Contains("\") OrElse TextBox1.Text.Contains("<") OrElse TextBox1.Text.Contains(">") OrElse TextBox1.Text.Contains("|") Then
                    OK_Button.Enabled = False
                Else
                    If IO.File.Exists(ComboBox1.Text) = True Then
                        OK_Button.Enabled = True
                        CheckBox1.Enabled = True
                    ElseIf IO.Directory.Exists(ComboBox1.Text) = True Then
                        OK_Button.Enabled = True
                        CheckBox1.Enabled = True
                    Else
                        OK_Button.Enabled = False
                        TextBox1.Enabled = False
                        CheckBox1.Enabled = False
                        CheckBox1.Checked = False
                    End If
                End If
            Else
                If IO.File.Exists(ComboBox1.Text) = True Then
                    OK_Button.Enabled = True
                    CheckBox1.Enabled = True
                ElseIf IO.Directory.Exists(ComboBox1.Text) = True Then
                    OK_Button.Enabled = True
                    CheckBox1.Enabled = True
                Else
                    OK_Button.Enabled = False
                    TextBox1.Enabled = False
                    CheckBox1.Enabled = False
                    CheckBox1.Checked = False
                End If
            End If
        End If

        Select Case CheckBox2.Checked
            Case True
                Me.Size = New Size(360, 365)
            Case False
                Me.Size = New Size(360, 188)
        End Select
    End Sub

    Private Sub Controller_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        LinkLabel1.Visible = False

        If IO.Directory.Exists(LinkLabel1.Text) Then
            ComboBox2.Text = LinkLabel1.Text.Trim
        End If

        ComboBox2.Visible = True
        Button2.Visible = True

        ComboBox2.Select()
        ComboBox2.SelectAll()
    End Sub

    Private Sub LinkLabel1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel1.TextChanged
        If PATH = String.Empty Then
            LinkLabel1.Text = "(Path)"
        End If
    End Sub

    Private Sub DirectoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DirectoryToolStripMenuItem.Click
        Dim D As New FolderBrowserDialog()
        D.RootFolder = Environment.SpecialFolder.Desktop
        D.ShowNewFolderButton = False
        D.Description = "Select a Directory do you want to set up for your new Link file:"
        If IO.File.Exists(ComboBox1.Text) = True Then
            Dim FI As New IO.FileInfo(ComboBox1.Text)
            D.SelectedPath = FI.DirectoryName
        ElseIf IO.Directory.Exists(ComboBox1.Text) = True Then
            D.SelectedPath = ComboBox1.Text
        Else
            D.SelectedPath = ""
        End If

        If D.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            Dim DI As New IO.DirectoryInfo(D.SelectedPath)
            ComboBox1.Text = D.SelectedPath
            TextBox1.Text = DI.Name

        End If
    End Sub

    Private Sub FileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FileToolStripMenuItem.Click
        Dim D As New OpenFileDialog()
        D.CheckFileExists = True
        D.DereferenceLinks = True
        D.Filter = "All files (*.*)|*.*"
        D.Multiselect = False
        D.Title = "Select a File do you want to set up for your new Link file"
        D.SupportMultiDottedExtensions = True
        If IO.File.Exists(ComboBox1.Text) = True Then
            Dim FI As New IO.FileInfo(ComboBox1.Text)
            D.InitialDirectory = FI.DirectoryName
            D.FileName = FI.Name
        ElseIf IO.Directory.Exists(ComboBox1.Text) = True Then
            D.InitialDirectory = ComboBox1.Text
        Else
            D.InitialDirectory = Environment.SpecialFolder.Desktop
        End If

        If D.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            Dim FI As New IO.FileInfo(D.FileName)
            ComboBox1.Text = D.FileName
            TextBox1.Text = FI.Name
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Select Case CheckBox2.Checked
            Case True
                Me.Size = New Size(360, 365)
            Case False
                Me.Size = New Size(360, 188)
        End Select

        GroupBox1.Enabled = CheckBox2.Checked
    End Sub

    Private Sub ComboBox2_LostFocus(sender As Object, e As EventArgs) Handles ComboBox2.LostFocus, Button2.LostFocus
        If Button2.Focused = False Then
            LinkLabel1.Visible = True

            If IO.Directory.Exists(ComboBox2.Text) Then
                LinkLabel1.Text = ComboBox2.Text.Trim
            End If

            ComboBox2.Visible = False
            Button2.Visible = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim D As New FolderBrowserDialog()
        D.RootFolder = Environment.SpecialFolder.Desktop
        D.ShowNewFolderButton = False
        D.Description = "Select a Directory where the new Link file will be located:"
        If IO.Directory.Exists(ComboBox1.Text) = True Then
            D.SelectedPath = ComboBox1.Text.Trim
        Else
            D.SelectedPath = ""
        End If
        If D.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            ComboBox2.Text = D.SelectedPath
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim D As New OpenFileDialog()
        D.CheckFileExists = True
        D.DereferenceLinks = True
        D.Filter = "Icon file (*.ico)|*.ico"
        D.Multiselect = False
        D.Title = "Select your custom Icon for your Link"
        D.SupportMultiDottedExtensions = True
        If IO.File.Exists(PictureBox1.Tag) = True Then
            Dim FI As New IO.FileInfo(ComboBox1.Text)
            D.InitialDirectory = FI.DirectoryName
            D.FileName = FI.Name
        Else
            D.InitialDirectory = Environment.SpecialFolder.Desktop
        End If

        If D.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            PictureBox1.Tag = D.FileName
            PictureBox1.Image = Image.FromFile(D.FileName)
        End If
    End Sub
End Class
