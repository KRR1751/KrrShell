Imports System.Windows.Forms

Public Class StartmenuProperties

    <Runtime.InteropServices.DllImport("dwmapi.dll")> Public Shared Function DwmExtendFrameIntoClientArea(ByVal hWnd As IntPtr, ByRef pMarinset As side) As Integer
    End Function

    <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)> Public Structure side
        Public Left As Integer
        Public Right As Integer
        Public Top As Integer
        Public Bottom As Integer
    End Structure

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If RadioButton1.Checked = True Then
            Try
                Startmenu.BackColor = Color.Black
                Startmenu.Panel2.BackColor = Color.Transparent
                Startmenu.TreeView2.BackColor = Color.Black

                Startmenu.Opacity = 1

                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "CustomTransparency", CheckBox1.Checked, Microsoft.Win32.RegistryValueKind.DWord)

                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "UseSystemColor", CheckBox2.Checked, Microsoft.Win32.RegistryValueKind.DWord)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "TransparencyLevel", NumericUpDown1.Value, Microsoft.Win32.RegistryValueKind.DWord)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "Mode", "0", Microsoft.Win32.RegistryValueKind.DWord)

                Startmenu.Invalidate()
                Startmenu.ReloadTiles()

                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Me.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        ElseIf RadioButton3.Checked = True Then
            Try
                Startmenu.BackColor = SystemColors.Control
                Startmenu.Panel2.BackColor = SystemColors.ControlLight

                If CheckBox1.Checked = True Then
                    Startmenu.Opacity = NumericUpDown1.Value / 100
                Else
                    Startmenu.Opacity = 1
                End If

                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "CustomTransparency", CheckBox1.Checked, Microsoft.Win32.RegistryValueKind.DWord)

                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "UseSystemColor", CheckBox2.Checked, Microsoft.Win32.RegistryValueKind.DWord)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "TransparencyLevel", NumericUpDown1.Value, Microsoft.Win32.RegistryValueKind.DWord)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "Mode", "1", Microsoft.Win32.RegistryValueKind.DWord)

                Startmenu.Invalidate()
                Startmenu.ReloadTiles()
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Me.Close()
            Catch ex As Exception

            End Try
        ElseIf RadioButton2.Checked = True Then
            Try
                Startmenu.startColor = Label1.BackColor
                Startmenu.endColor = Label3.BackColor
                Startmenu.Panel2.BackColor = Color.Transparent

                If CheckBox1.Checked = True Then
                    Startmenu.Opacity = NumericUpDown1.Value / 100
                Else
                    Startmenu.Opacity = 1
                End If

                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "CustomTransparency", CheckBox1.Checked, Microsoft.Win32.RegistryValueKind.DWord)

                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "UseSystemColor", CheckBox2.Checked, Microsoft.Win32.RegistryValueKind.DWord)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "TransparencyLevel", NumericUpDown1.Value, Microsoft.Win32.RegistryValueKind.DWord)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "Mode", "2", Microsoft.Win32.RegistryValueKind.DWord)

                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "ColorPos", ComboBox1.SelectedIndex, Microsoft.Win32.RegistryValueKind.DWord)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "ColorPos2", ComboBox2.SelectedIndex, Microsoft.Win32.RegistryValueKind.DWord)

                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "BackColor", ColorTranslator.ToHtml(Startmenu.startColor), Microsoft.Win32.RegistryValueKind.String)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "BackColor2", ColorTranslator.ToHtml(Startmenu.endColor), Microsoft.Win32.RegistryValueKind.String)

                Dim side As Startmenu.side = side
                side.Left = 0
                side.Right = 0
                side.Top = 0
                side.Bottom = 0
                Dim result As Integer = Startmenu.DwmExtendFrameIntoClientArea(Startmenu.Handle, side)

                Startmenu.Invalidate()
                Startmenu.ReloadTiles()
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Me.Close()
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cd As New ColorDialog
        cd.AllowFullOpen = True
        cd.FullOpen = True
        cd.AnyColor = True
        cd.Color = Label1.BackColor
        If cd.ShowDialog(Me) = DialogResult.OK Then
            Label1.BackColor = cd.Color
        End If
    End Sub

    Private Sub StartmenuProperties_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim mode As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "Mode", 2)
        If mode = 0 Then
            RadioButton1.Checked = True
        ElseIf mode = 1 Then
            RadioButton3.Checked = True
        ElseIf mode = 2 Then
            RadioButton2.Checked = True
        End If
        Label1.BackColor = Startmenu.startColor
        Label3.BackColor = Startmenu.endColor

        ComboBox1.SelectedIndex = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "ColorPos", 0)
        ComboBox2.SelectedIndex = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "ColorPos2", 0)

        CheckBox1.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "CustomTransparency", 1)
        CheckBox2.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "UseSystemColor", 1)

        NumericUpDown1.Value = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "TransparencyLevel", 90)
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged, RadioButton1.CheckedChanged, RadioButton3.CheckedChanged
        If RadioButton2.Checked = True Then
            Label1.Enabled = True
            Label2.Enabled = True
            Label3.Enabled = True
            Button1.Enabled = True
            Button2.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
        Else
            Label1.Enabled = False
            Label2.Enabled = False
            Label3.Enabled = False
            Button1.Enabled = False
            Button2.Enabled = False
            ComboBox1.Enabled = False
            ComboBox2.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim cd As New ColorDialog
        cd.AllowFullOpen = True
        cd.FullOpen = True
        cd.AnyColor = True
        cd.Color = Label3.BackColor
        If cd.ShowDialog(Me) = DialogResult.OK Then
            Label3.BackColor = cd.Color
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        MsgBox("This option will affect your ClientArea into the full ""Start Menu"". Or in another words, your entire Window GUI will be extended to the full window.

Which I highly don't recommend if you have Windows 10/11 and Title bar color set to 'White' or 'Dark' to have this enabled. Because then your entire Start menu could be not text displayed correctly etc. If you have Windows 7 with Aero or just have Aero in total, it will be really nice and clean. But if you don't know what are u doing, just enable this by your own risk.", MsgBoxStyle.Information, "About this feature")
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        NumericUpDown1.Enabled = CheckBox1.Checked
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim accentColor As Color = WindowsColorSettings.GetAccentColor

        Label1.BackColor = accentColor
        Label3.BackColor = accentColor
    End Sub
End Class
