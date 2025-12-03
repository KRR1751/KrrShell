Imports System.IO
Imports System.Windows.Forms

Public Class TileCreator

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Try
            Dim FI As New FileInfo(ComboBox1.Text)
            IO.Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & FI.Name)
            My.Computer.FileSystem.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & FI.Name & "\Program", FI.FullName, False)

            If CheckBox1.Checked = True Then
                My.Computer.FileSystem.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & FI.Name & "\Arguments", ComboBox2.Text, False)
            End If

            If RadioButton1.Checked = True Then
                My.Computer.FileSystem.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & FI.Name & "\SizeType", "1", False)
            ElseIf RadioButton2.Checked = True Then
                My.Computer.FileSystem.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & FI.Name & "\SizeType", "2", False)
            ElseIf RadioButton3.Checked = True Then
                My.Computer.FileSystem.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & FI.Name & "\SizeType", "3", False)
            End If

            Startmenu.ReloadTiles()
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        If IO.File.Exists(ComboBox1.Text) Then
            OK_Button.Enabled = True
        Else
            OK_Button.Enabled = False
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ofd As New OpenFileDialog
        ofd.Title = "Select a program to pin into the Tiles..."
        ofd.Multiselect = False
        ofd.AutoUpgradeEnabled = False
        ofd.CheckFileExists = True
        If IO.File.Exists(ComboBox1.Text) Then
            Dim fi As New FileInfo(ComboBox1.Text)
            ofd.InitialDirectory = fi.DirectoryName
            ofd.FileName = fi.Name
        ElseIf io.Directory.Exists(ComboBox1.Text) Then
            ofd.InitialDirectory = ComboBox1.Text
            ofd.FileName = ""
        Else
            ofd.FileName = ""
        End If
        ofd.Filter = "Programs (*.exe;*.pif;*.cmd;*.bat)|*.exe;*.pif;*.cmd;*.bat;*.lnk|All files (*.*)|*.*"
        If ofd.ShowDialog(Me) = DialogResult.OK Then
            ComboBox1.Text = ofd.FileName
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        ComboBox2.Enabled = CheckBox1.Checked
    End Sub

    Private Sub TileCreator_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        Label7.Text = "Location: " & Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned"
    End Sub
End Class
