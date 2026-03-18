Imports System.IO
Imports System.Windows.Forms

Public Class TileCreator

    Private Sub OK_ButtonOld_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles OK_Button.Click
        Try
            Dim FI As New FileInfo(ComboBox1.Text)
            IO.Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & FI.Name)
            My.Computer.FileSystem.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & FI.Name & "\Program", FI.FullName, False)

            If CheckBox1.Checked = True Then
                My.Computer.FileSystem.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & FI.Name & "\Arguments", ComboBox2.Text, False)
            End If

            'If RadioButton1.Checked = True Then
            My.Computer.FileSystem.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & FI.Name & "\SizeType", "1", False)
            'ElseIf RadioButton2.Checked = True Then
            My.Computer.FileSystem.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & FI.Name & "\SizeType", "2", False)
            'ElseIf RadioButton3.Checked = True Then
            My.Computer.FileSystem.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & FI.Name & "\SizeType", "3", False)
            'End If

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

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged, TextBox1.TextChanged
        If IO.File.Exists(ComboBox1.Text) AndAlso Not String.IsNullOrWhiteSpace(TextBox1.Text) Then
            OK_Button.Enabled = True

            If CheckBox2.Checked = True Then
                Dim bmp As Bitmap = AppBar.GetFileIcon(ComboBox1.Text, True).ToBitmap

                If bmp IsNot Nothing Then
                    Label6.Image = bmp
                Else
                    Label6.Image = My.Resources.ProgramMedium
                End If
            End If
        Else
            OK_Button.Enabled = False
            Label6.Image = My.Resources.ProgramMedium
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ofd As New OpenFileDialog With {
            .AutoUpgradeEnabled = AppBar.UseExplorerFP,
            .Title = "Select a program to pin as a Tile...",
            .Multiselect = False,
            .CheckFileExists = True,
            .Filter = "Programs (*.exe;*.pif;*.cmd;*.bat)|*.exe;*.pif;*.cmd;*.bat;*.lnk|All files (*.*)|*.*"}

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

        If ofd.ShowDialog(Me) = DialogResult.OK Then
            ComboBox1.Text = ofd.FileName
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        ComboBox2.Enabled = CheckBox1.Checked
    End Sub

    Private Sub TileCreator_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next

        Width_ComboBox.SelectedIndex = 1
        Height_ComboBox.SelectedIndex = 1

        TextBox1.Focus()
        TextBox1.Select()
        'Label7.Text = "Location: " & Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned"
    End Sub

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click

        Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles", True)

        If myKey IsNot Nothing Then
            Try
                myKey.CreateSubKey(TextBox1.Text)
                Dim myCreatedKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles\" & TextBox1.Text, True)

                If myCreatedKey IsNot Nothing Then

                    ' Basic
                    If File.Exists(ComboBox1.Text) Then myCreatedKey.SetValue("Program", ComboBox1.Text)
                    If CheckBox1.Checked = True Then myCreatedKey.SetValue("Arguments", ComboBox1.Text) Else myCreatedKey.SetValue("Arguments", "")

                    ' Style
                    myCreatedKey.SetValue("Order", 0, Microsoft.Win32.RegistryValueKind.DWord)
                    If Not Label6.BackColor = Color.Transparent Then myCreatedKey.SetValue("BackColor", ColorTranslator.ToHtml(Label6.BackColor))

                    ' Sizing
                    myCreatedKey.SetValue("SizeTypeW", Width_ComboBox.SelectedIndex, Microsoft.Win32.RegistryValueKind.DWord)
                    myCreatedKey.SetValue("SizeTypeH", Height_ComboBox.SelectedIndex, Microsoft.Win32.RegistryValueKind.DWord)

                    ' Finishing...
                    myCreatedKey.Close()
                Else
                    MsgBox("Failed to create a Tile. Registry path doesn't exists.")
                End If

            Catch ex As Exception
                MsgBox("Failed to create a Tile. " & ex.Message)

            Finally
                myKey.Close()
                Startmenu.ReloadTiles()

                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Me.Close()
            End Try
        End If
    End Sub

    Dim SMextraSmall As Integer = 48
    Dim SMSmall As Integer = 96
    Dim SMMedium As Integer = 144
    Dim SMLarge As Integer = 192

    Private Sub Width_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Width_ComboBox.SelectedIndexChanged
        Select Case Width_ComboBox.SelectedIndex ' Width

            Case 0 ' Extra Small
                Label6.Width = SMextraSmall
            Case 1 ' Small
                Label6.Width = SMSmall
            Case 2 ' Medium
                Label6.Width = SMMedium
            Case 3 ' Large
                Label6.Width = SMLarge

        End Select
    End Sub

    Private Sub Height_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Height_ComboBox.SelectedIndexChanged
        Select Case Height_ComboBox.SelectedIndex ' Height

            Case 0 ' Extra Small
                Label6.Height = SMextraSmall
            Case 1 ' Small
                Label6.Height = SMSmall
            Case 2 ' Medium
                Label6.Height = SMMedium
            Case 3 ' Large
                Label6.Height = SMLarge

        End Select
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = False Then
            Label6.Image = My.Resources.ProgramMedium
            Exit Sub
        End If

        If File.Exists(ComboBox1.Text) Then

            Dim bmp As Bitmap = AppBar.GetFileIcon(ComboBox1.Text, True).ToBitmap

            If bmp IsNot Nothing Then
                Label6.Image = bmp
            Else
                Label6.Image = My.Resources.ProgramMedium
            End If
        Else
            Label6.Image = My.Resources.ProgramMedium
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Label6.BackColor = Color.Transparent
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim CD As New ColorDialog With {
            .AllowFullOpen = True,
            .AnyColor = True,
            .FullOpen = True
        }

        If Not Label6.BackColor = Color.Transparent Then
            CD.Color = Label6.BackColor
        End If

        If CD.ShowDialog(Me) = DialogResult.OK Then
            Label6.BackColor = CD.Color
        End If
    End Sub
End Class
