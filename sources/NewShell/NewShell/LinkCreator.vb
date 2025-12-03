Imports System.Windows.Forms

Public Class LinkCreator
    Public PATH As String = String.Empty
    Public Result As String = String.Empty
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If IO.File.Exists(ComboBox1.Text) = True Then
            Shell("cmd.exe /c mklink " & """" & PATH & TextBox1.Text & """ """ & ComboBox1.Text & """", AppWinStyle.Hide, True)
            Result = TextBox1.Text
        ElseIf IO.Directory.Exists(ComboBox1.Text) = True Then
            Shell("cmd.exe /c mklink /D " & """" & PATH & TextBox1.Text & """ """ & ComboBox1.Text & """", AppWinStyle.Hide, True)
            Result = TextBox1.Text
        End If
        'Shell("mklink ")
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub ComboBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.TextChanged
        If CheckBox1.Checked = False Then
            If IO.File.Exists(ComboBox1.Text) = True Then
                Dim FI As New IO.FileInfo(ComboBox1.Text)
                TextBox1.Text = FI.Name
            ElseIf IO.Directory.Exists(ComboBox1.Text) = True Then
                Dim DI As New IO.DirectoryInfo(ComboBox1.Text)
                TextBox1.Text = DI.Name
            Else
                TextBox1.Text = ""
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
        Controller.Enabled = False
    End Sub

    Private Sub LinkCreator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Controller.Enabled = True
        LinkLabel1.Text = PATH
    End Sub

    Private Sub Controller_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Controller.Tick
        If PATH = String.Empty Then
            OK_Button.Enabled = False
        Else
            If CheckBox1.Checked = True Then
                If TextBox1.Text.Contains("*") OrElse TextBox1.Text.Contains("?") OrElse TextBox1.Text.Contains("""") OrElse TextBox1.Text.Contains(":") OrElse TextBox1.Text.Contains("/") OrElse TextBox1.Text.Contains("\") OrElse TextBox1.Text.Contains("<") OrElse TextBox1.Text.Contains(">") OrElse TextBox1.Text.Contains("|") Then
                    OK_Button.Enabled = False
                Else
                    If IO.File.Exists(ComboBox1.Text) = True Then
                        OK_Button.Enabled = True
                        Label3.Enabled = True
                        CheckBox1.Enabled = True
                    ElseIf IO.Directory.Exists(ComboBox1.Text) = True Then
                        OK_Button.Enabled = True
                        Label3.Enabled = True
                        CheckBox1.Enabled = True
                    Else
                        OK_Button.Enabled = False
                        TextBox1.Enabled = False
                        Label3.Enabled = False
                        CheckBox1.Enabled = False
                        CheckBox1.Checked = False
                    End If
                End If
            Else
                If IO.File.Exists(ComboBox1.Text) = True Then
                    OK_Button.Enabled = True
                    Label3.Enabled = True
                    CheckBox1.Enabled = True
                ElseIf IO.Directory.Exists(ComboBox1.Text) = True Then
                    OK_Button.Enabled = True
                    Label3.Enabled = True
                    CheckBox1.Enabled = True
                Else
                    OK_Button.Enabled = False
                    TextBox1.Enabled = False
                    Label3.Enabled = False
                    CheckBox1.Enabled = False
                    CheckBox1.Checked = False
                End If
            End If
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        If Not PATH = String.Empty Then
            System.Diagnostics.Process.Start(PATH)
        End If
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
            ComboBox1.Text = D.SelectedPath
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
            ComboBox1.Text = D.FileName
        End If
    End Sub
End Class
