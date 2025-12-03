Imports System.Windows.Forms

Public Class NotifyWindowD
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If CheckBox6.Checked = True Then
            If ComboBox1.SelectedIndex = 0 Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "EnableCustomTitle", "0", Microsoft.Win32.RegistryValueKind.DWord)
            ElseIf ComboBox1.SelectedIndex = 1 Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "EnableCustomTitle", "1", Microsoft.Win32.RegistryValueKind.DWord)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "CustomTitle", TextBox1.Text, Microsoft.Win32.RegistryValueKind.String)
            End If
            If ComboBox2.SelectedIndex = 0 Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "EnableCustomIcon", "0", Microsoft.Win32.RegistryValueKind.DWord)
            Else
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "EnableCustomIcon", "1", Microsoft.Win32.RegistryValueKind.DWord)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "CustomIcon", ComboBox3.Text, Microsoft.Win32.RegistryValueKind.String)
            End If
            If CheckBox1.Checked = True Then
                If RadioButton1.Checked = True Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "EnableCustomBackColor", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "CustomBR", CL.BackColor.R, Microsoft.Win32.RegistryValueKind.DWord)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "CustomBG", CL.BackColor.G, Microsoft.Win32.RegistryValueKind.DWord)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "CustomBB", CL.BackColor.B, Microsoft.Win32.RegistryValueKind.DWord)
                    If CheckBox2.Checked = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "CAffectControls", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "CAffectControls", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                Else
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "EnableCustomBackColor", "2", Microsoft.Win32.RegistryValueKind.DWord)
                End If
            Else
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "EnableCustomBackColor", "0", Microsoft.Win32.RegistryValueKind.DWord)
            End If
            If ComboBox4.SelectedIndex = 0 Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "BorderStyle", "0", Microsoft.Win32.RegistryValueKind.DWord)
            Else
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "BorderStyle", "1", Microsoft.Win32.RegistryValueKind.DWord)
            End If
            If CheckBox3.Checked = False Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "OnTop", "0", Microsoft.Win32.RegistryValueKind.DWord)
            Else
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "OnTop", "1", Microsoft.Win32.RegistryValueKind.DWord)
            End If
            If CheckBox4.Checked = False Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "CloseButton", "0", Microsoft.Win32.RegistryValueKind.DWord)
            Else
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "CloseButton", "1", Microsoft.Win32.RegistryValueKind.DWord)
            End If
            If CheckBox5.Checked = False Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "FocusOnActivation", "0", Microsoft.Win32.RegistryValueKind.DWord)
            Else
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "FocusOnActivation", "1", Microsoft.Win32.RegistryValueKind.DWord)
            End If
            If CheckBox6.Checked = False Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "SaveBTNChecked", "0", Microsoft.Win32.RegistryValueKind.DWord)
            Else
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "SaveBTNChecked", "1", Microsoft.Win32.RegistryValueKind.DWord)
            End If
            If CheckBox6.Checked = False Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "EnableCustomTextColor", "0", Microsoft.Win32.RegistryValueKind.DWord)
            Else
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "EnableCustomTextColor", "1", Microsoft.Win32.RegistryValueKind.DWord)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "CustomTR", CL2.BackColor.R, Microsoft.Win32.RegistryValueKind.DWord)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "CustomTG", CL2.BackColor.G, Microsoft.Win32.RegistryValueKind.DWord)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\", "CustomTB", CL2.BackColor.B, Microsoft.Win32.RegistryValueKind.DWord)
            End If
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        Else
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub NotifyWindowD_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Controller.Enabled = False
    End Sub

    Private Sub NotifyWindowD_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox4.SelectedIndex = 0
        Controller.Enabled = True
        AppBar.NWSet()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ContextMenuStrip1.Show(Control.MousePosition)
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            RadioButton1.Enabled = True
            If RadioButton1.Checked = True Then
                CL.Enabled = True
            End If
            RadioButton2.Enabled = True
            CheckBox2.Enabled = True
        Else
            RadioButton1.Enabled = False
            CL.Enabled = False
            RadioButton2.Enabled = False
            CheckBox2.Enabled = False
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        CL2.Enabled = CheckBox7.Checked
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex = 0 OrElse ComboBox2.SelectedIndex = 2 Then
            ComboBox3.Enabled = False
            Button1.Enabled = False
            PictureBox1.Enabled = False
        ElseIf ComboBox2.SelectedIndex = 1 Then
            ComboBox3.Enabled = True
            Button1.Enabled = True
            PictureBox1.Enabled = True
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 0 Then
            TextBox1.Enabled = False
        Else
            TextBox1.Enabled = True
        End If
    End Sub

    Private Sub OneButtonToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OneButtonToolStripMenuItem.Click
        NW.MODE = 1
        NW.B1.Text = "B1_TEXT"
        NW.B2.Text = "B2_TEXT"
        NW.B3.Text = "B3_TEXT"
        NW.TL.Text = "TITLE_TEXT"
        NW.IP.Text = "DESCRIPTION"
        If NW.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            NW.Close()
        End If
    End Sub

    Private Sub TwoButtonsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TwoButtonsToolStripMenuItem.Click
        NW.MODE = 2
        NW.B1.Text = "B1_TEXT"
        NW.B2.Text = "B2_TEXT"
        NW.B3.Text = "B3_TEXT"
        NW.TL.Text = "TITLE_TEXT"
        NW.IP.Text = "DESCRIPTION"
        If NW.ShowDialog(Me) = Windows.Forms.DialogResult.OK OrElse NW.ShowDialog(Me) = Windows.Forms.DialogResult.No OrElse NW.ShowDialog(Me) = Windows.Forms.DialogResult.Yes Then
            NW.Close()
        End If
    End Sub

    Private Sub ThreeButtonsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ThreeButtonsToolStripMenuItem.Click
        NW.MODE = 3
        NW.B1.Text = "B1_TEXT"
        NW.B2.Text = "B2_TEXT"
        NW.B3.Text = "B3_TEXT"
        NW.TL.Text = "TITLE_TEXT"
        NW.IP.Text = "DESCRIPTION"
        If NW.ShowDialog(Me) = Windows.Forms.DialogResult.OK OrElse NW.ShowDialog(Me) = Windows.Forms.DialogResult.No OrElse NW.ShowDialog(Me) = Windows.Forms.DialogResult.Yes Then
            NW.Close()
        End If
    End Sub

    Private Sub CL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CL.Click
        CD.Color = CL.BackColor
        If CD.ShowDialog = Windows.Forms.DialogResult.OK Then
            CL.BackColor = CD.Color
        End If
    End Sub

    Private Sub CL2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CL2.Click
        CD.Color = CL2.BackColor
        If CD.ShowDialog = Windows.Forms.DialogResult.OK Then
            CL2.BackColor = CD.Color
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ofd As New OpenFileDialog
        ofd.CheckPathExists = True
        ofd.Title = "Open File Dialog - Custom Icon for Notify Window"
        ofd.Filter = "Icon (*.ico)|*.ico"
        If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
            ComboBox3.Text = ofd.FileName
            If IO.File.Exists(ComboBox3.Text) = True Then
                Dim FI As New IO.FileInfo(ComboBox3.Text)
                ofd.InitialDirectory = FI.DirectoryName
                ofd.FileName = ComboBox3.Text
            ElseIf IO.Directory.Exists(ComboBox3.Text) = True Then
                ofd.InitialDirectory = ComboBox3.Text
            Else
                ofd.InitialDirectory = "C:\Users\Kristian\Pictures"
            End If
        End If
    End Sub

    Private Sub Controller_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Controller.Tick
        If ComboBox2.SelectedIndex = 1 Then
            If IO.File.Exists(ComboBox3.Text) = True Then
                Dim FI As New IO.FileInfo(ComboBox3.Text)
                If FI.Extension = ".ico" Then
                    OK_Button.Enabled = True
                Else
                    OK_Button.Enabled = False
                End If
            Else
                OK_Button.Enabled = False
            End If
        Else
            OK_Button.Enabled = True
        End If
    End Sub
End Class
