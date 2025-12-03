Imports System.Windows.Forms

Public Class APD
    Public MODE As Integer = 0
    REM 1
    Public remTXT1 As String = String.Empty
    Public remTXT2 As String = String.Empty
    Public remTXT3 As String = String.Empty
    Public remTXT4 As String = String.Empty
    Public remCHB1 As Boolean = False
    Public remCHB2 As Boolean = True
    Public remCHB3 As Boolean = True
    Public remDUD1 As Integer = 1
    Public remDUD2 As Integer = 0
    Public remDUD3 As Integer = 0
    Public remDUD4 As Integer = 0
    Public remDUD5 As Integer = 0
    Public remCB1 As Integer = 1
    Public remCB2 As Integer = 0
    Public remCB3 As Integer = 0
    Public remCB4 As String = String.Empty
    Public remNUPX As Integer = 0
    Public remNUPY As Integer = 0
    Public remNUP1 As Integer = 50
    Public remNUP2 As Integer = 1
    REM 2
    Public remCHB11 As Boolean = True
    Public remCB11 As Integer = 0
    Public remCB12 As Integer = 0
    Public remCB13 As String = String.Empty
    Public remDUD11 As Integer = 0
    Public remNUP11 As Integer = 50
    Public remNUP12 As Integer = 1
    REM 3
    Public remCB21 As String = String.Empty
    Public remRB21 As Boolean = True
    REM 3.1
    Public remTXT211 As String = String.Empty
    Public remTXT212 As String = String.Empty
    Public remTXT213 As String = String.Empty
    Public remTXT214 As String = String.Empty
    Public remTXT215 As String = String.Empty
    Public remCB211 As Integer = 0
    Public remCHB211 As Boolean = False
    Public remCHB212 As Boolean = False
    Public remCHB213 As Boolean = False
    Public remCHB214 As Boolean = False
    Public remCHB215 As Boolean = False
    Public remCHB216 As Boolean = False
    Public remCHB217 As Boolean = True
    Public remCHB218 As Boolean = True
    Public remCHB219 As Boolean = True
    Public remCHB21a As Boolean = True
    Public remCHB21b As Boolean = True
    REM 3.2
    Public remCB221 As Integer = 0
    Public remCHB221 As Boolean = False
    Public remNUP221 As Integer = -1
    REM GO
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If MODE = 1 Then
            remTXT1 = TextBox1.Text
            remTXT2 = TextBox4.Text
            remTXT3 = TextBox3.Text
            remTXT4 = TextBox2.Text
            remCHB1 = CheckBox1.Checked
            remCHB2 = CheckBox16.Checked
            remCHB3 = CheckBox2.Checked
            remDUD1 = DomainUpDown1.SelectedIndex
            remDUD2 = DomainUpDown3.SelectedIndex
            remDUD3 = DomainUpDown2.SelectedIndex
            remDUD4 = DomainUpDown4.SelectedIndex
            remDUD5 = DomainUpDown5.SelectedIndex
            remCB1 = ComboBox1.SelectedIndex
            remCB2 = ComboBox2.SelectedIndex
            remCB3 = ComboBox3.SelectedIndex
            remCB4 = ComboBox4.Text
            remNUPX = NumericUpDown1.Value
            remNUPY = NumericUpDown2.Value
            remNUP1 = NumericUpDown12.Value
            remNUP2 = ThmN.Value
        ElseIf MODE = 2 Then
            remCHB11 = CheckBox3.Checked
            remCB11 = ComboBox6.SelectedIndex
            remCB13 = ComboBox5.Text
            remCB12 = ComboBox7.SelectedIndex
            remDUD11 = DomainUpDown6.SelectedIndex
            remNUP11 = NumericUpDown4.Value
            remNUP12 = NumericUpDown3.Value
        ElseIf MODE = 3 Then
            remCB21 = ComboBox8.Text
            remRB21 = RadioButton3.Checked
            CheckBox5.Checked = False
            If RadioButton3.Checked = True Then
                remTXT211 = TextBox5.Text
                remTXT212 = TextBox6.Text
                remTXT213 = MaskedTextBox1.Text
                remTXT214 = TextBox7.Text
                remTXT215 = TextBox8.Text
                remCB211 = ComboBox10.SelectedIndex
                remCHB211 = CheckBox7.Checked
                remCHB212 = CheckBox8.Checked
                remCHB213 = CheckBox4.Checked
                remCHB214 = CheckBox9.Checked
                remCHB215 = CheckBox10.Checked
                remCHB216 = CheckBox11.Checked
                remCHB217 = CheckBox12.Checked
                remCHB218 = CheckBox13.Checked
                remCHB219 = CheckBox14.Checked
                remCHB21a = CheckBox15.Checked
                remCHB21b = CheckBox17.Checked
            Else
                remCB221 = ComboBox9.SelectedIndex
                remCHB221 = CheckBox6.Checked
                remNUP221 = NumericUpDown5.Value
            End If
        End If
        If Dialog3.CheckBox25.Checked = True Then
            If SaveBTN.Checked = True Then
                If Dialog3.ComboBox5.SelectedIndex = 0 Then
                    If remCB1 = 9 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "X", remNUPX, Microsoft.Win32.RegistryValueKind.String)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "Y", remNUPY, Microsoft.Win32.RegistryValueKind.String)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "Position", remCB1, Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "Position", remCB1, Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If remCHB1 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableRinging", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingMode", remDUD1, Microsoft.Win32.RegistryValueKind.DWord)
                        If remDUD1 = 0 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingSYS", remCB2, Microsoft.Win32.RegistryValueKind.DWord)
                        ElseIf remDUD1 = 1 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingALIST", remCB3, Microsoft.Win32.RegistryValueKind.String)
                        ElseIf remDUD1 = 2 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingPath", remCB4, Microsoft.Win32.RegistryValueKind.String)
                        End If
                        If remCHB2 = False Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingRepeat", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingTimesRepeat", remNUP2, Microsoft.Win32.RegistryValueKind.String)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingRepeat", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingVol", remNUP1, Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableRinging", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If remCHB3 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableB2", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        If remDUD5 = 0 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomB2", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomB2", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "CustomB2", remTXT4, Microsoft.Win32.RegistryValueKind.String)
                        End If
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableB2", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If remDUD2 = 0 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomTitle", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomTitle", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "CustomTitle", remTXT1, Microsoft.Win32.RegistryValueKind.String)
                    End If
                    If remDUD3 = 0 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomPicture", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    ElseIf remDUD3 = 1 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomPicture", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    ElseIf remDUD3 = 2 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomPicture", "2", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "CustomPicturePath", remTXT2, Microsoft.Win32.RegistryValueKind.String)
                    End If
                    If remDUD4 = 0 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomB1", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomB1", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "CustomB1", remTXT3, Microsoft.Win32.RegistryValueKind.String)
                    End If
                ElseIf Dialog3.ComboBox5.SelectedIndex = 1 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingMode", remDUD11, Microsoft.Win32.RegistryValueKind.DWord)
                    If remDUD11 = 0 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingSYS", remCB11, Microsoft.Win32.RegistryValueKind.DWord)
                    ElseIf remDUD11 = 1 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingALIST", remCB12, Microsoft.Win32.RegistryValueKind.String)
                    ElseIf remDUD11 = 2 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingPath", remCB13, Microsoft.Win32.RegistryValueKind.String)
                    End If
                    If remCHB11 = False Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingRepeat", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingTimesRepeat", remNUP12, Microsoft.Win32.RegistryValueKind.String)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingRepeat", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingVol", remNUP11, Microsoft.Win32.RegistryValueKind.DWord)
                ElseIf Dialog3.ComboBox5.SelectedIndex = 2 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\", "Path", remCB21, Microsoft.Win32.RegistryValueKind.String)
                    If remRB21 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\", "Mode", "1", Microsoft.Win32.RegistryValueKind.DWord)

                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "Arguments", remTXT211, Microsoft.Win32.RegistryValueKind.String)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "AppWinStyle", remCB211, Microsoft.Win32.RegistryValueKind.DWord)
                        If remCHB211 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "CreateNoWindow", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "CreateNoWindow", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If remCHB212 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "UseShellExecute", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "UseShellExecute", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If remCHB213 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RunAs", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "Username", remTXT212, Microsoft.Win32.RegistryValueKind.String)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RunAs", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If remCHB214 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "LoadUserProfile", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "LoadUserProfile", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If remCHB215 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "EnableCustomWorkingDir", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "CustomWorkDir", remTXT212, Microsoft.Win32.RegistryValueKind.String)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "EnableCustomWorkingDir", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If remCHB216 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "EnableDomain", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "Domain", remTXT215, Microsoft.Win32.RegistryValueKind.String)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "EnableDomain", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If remCHB217 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "ErrorDialog", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "ErrorDialog", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If remCHB218 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "ErrorDialogParentHandle", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "ErrorDialogParentHandle", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If remCHB219 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RSError", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RSError", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If remCHB21a = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RSInput", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RSInput", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If remCHB21b = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RSOutput", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RSOutput", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\", "Mode", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\2\", "AppWinStyle", remCB221, Microsoft.Win32.RegistryValueKind.DWord)
                        If remCHB221 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\2\", "Wait", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\2\", "Timeout", remNUP221, Microsoft.Win32.RegistryValueKind.String)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\2\", "Wait", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\2\", "Timeout", "-1", Microsoft.Win32.RegistryValueKind.String)
                        End If
                    End If
                End If
            End If
        Else
            If Dialog3.ComboBox5.SelectedIndex = 0 Then
                If remCB1 = 9 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "X", remNUPX, Microsoft.Win32.RegistryValueKind.String)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "Y", remNUPY, Microsoft.Win32.RegistryValueKind.String)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "Position", remCB1, Microsoft.Win32.RegistryValueKind.DWord)
                Else
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "Position", remCB1, Microsoft.Win32.RegistryValueKind.DWord)
                End If
                If remCHB1 = True Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "EnableRinging", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingMode", remDUD1, Microsoft.Win32.RegistryValueKind.DWord)
                    If remDUD1 = 0 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingSYS", remCB2, Microsoft.Win32.RegistryValueKind.DWord)
                    ElseIf remDUD1 = 1 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingALIST", remCB3, Microsoft.Win32.RegistryValueKind.String)
                    ElseIf remDUD1 = 2 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingPath", remCB4, Microsoft.Win32.RegistryValueKind.String)
                    End If
                    If remCHB2 = False Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingRepeat", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingTimesRepeat", remNUP2, Microsoft.Win32.RegistryValueKind.String)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingRepeat", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingVol", remNUP1, Microsoft.Win32.RegistryValueKind.DWord)
                Else
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "EnableRinging", "0", Microsoft.Win32.RegistryValueKind.DWord)
                End If
                If remCHB3 = True Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "EnableB2", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    If remDUD5 = 0 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "EnableCustomB2", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "EnableCustomB2", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "CustomB2", remTXT4, Microsoft.Win32.RegistryValueKind.String)
                    End If
                Else
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "EnableB2", "0", Microsoft.Win32.RegistryValueKind.DWord)
                End If
                If remDUD2 = 0 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "EnableCustomTitle", "0", Microsoft.Win32.RegistryValueKind.DWord)
                Else
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "EnableCustomTitle", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "CustomTitle", remTXT1, Microsoft.Win32.RegistryValueKind.String)
                End If
                If remDUD3 = 0 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "EnableCustomPicture", "0", Microsoft.Win32.RegistryValueKind.DWord)
                ElseIf remDUD3 = 1 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "EnableCustomPicture", "1", Microsoft.Win32.RegistryValueKind.DWord)
                ElseIf remDUD3 = 2 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "EnableCustomPicture", "2", Microsoft.Win32.RegistryValueKind.DWord)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "CustomPicturePath", remTXT2, Microsoft.Win32.RegistryValueKind.String)
                End If
                If remDUD4 = 0 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "EnableCustomB1", "0", Microsoft.Win32.RegistryValueKind.DWord)
                Else
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "EnableCustomB1", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "CustomB1", remTXT3, Microsoft.Win32.RegistryValueKind.String)
                End If
            ElseIf Dialog3.ComboBox5.SelectedIndex = 1 Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\2\", "RingMode", remDUD11, Microsoft.Win32.RegistryValueKind.DWord)

                If remDUD11 = 0 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\2\", "RingSYS", remCB11, Microsoft.Win32.RegistryValueKind.DWord)
                ElseIf remDUD11 = 1 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\2\", "RingALIST", remCB12, Microsoft.Win32.RegistryValueKind.String)
                ElseIf remDUD11 = 2 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\2\", "RingPath", remCB13, Microsoft.Win32.RegistryValueKind.String)
                End If
                If remCHB11 = False Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\2\", "RingRepeat", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\2\", "RingTimesRepeat", remNUP12, Microsoft.Win32.RegistryValueKind.String)
                Else
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\2\", "RingRepeat", "1", Microsoft.Win32.RegistryValueKind.DWord)
                End If
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\2\", "RingVol", remNUP11, Microsoft.Win32.RegistryValueKind.DWord)
            ElseIf Dialog3.ComboBox5.SelectedIndex = 2 Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\", "Path", remCB21, Microsoft.Win32.RegistryValueKind.String)
                If remRB21 = True Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\", "Mode", "1", Microsoft.Win32.RegistryValueKind.DWord)

                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "Arguments", remTXT211, Microsoft.Win32.RegistryValueKind.String)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "AppWinStyle", remCB211, Microsoft.Win32.RegistryValueKind.DWord)
                    If remCHB211 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "CreateNoWindow", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "CreateNoWindow", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If remCHB212 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "UseShellExecute", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "UseShellExecute", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If remCHB213 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "RunAs", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "Username", remTXT212, Microsoft.Win32.RegistryValueKind.String)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "RunAs", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If remCHB214 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "LoadUserProfile", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "LoadUserProfile", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If remCHB215 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "EnableCustomWorkingDir", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "CustomWorkDir", remTXT212, Microsoft.Win32.RegistryValueKind.String)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "EnableCustomWorkingDir", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If remCHB216 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "EnableDomain", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "Domain", remTXT215, Microsoft.Win32.RegistryValueKind.String)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "EnableDomain", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If remCHB217 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "ErrorDialog", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "ErrorDialog", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If remCHB218 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "ErrorDialogParentHandle", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "ErrorDialogParentHandle", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If remCHB219 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "RSError", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "RSError", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If remCHB21a = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "RSInput", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "RSInput", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If remCHB21b = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "RSOutput", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\1\", "RSOutput", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                Else
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\", "Mode", "0", Microsoft.Win32.RegistryValueKind.DWord)

                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\2\", "AppWinStyle", remCB221, Microsoft.Win32.RegistryValueKind.DWord)
                    If remCHB221 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\2\", "Wait", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\2\", "Timeout", remNUP221, Microsoft.Win32.RegistryValueKind.String)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\2\", "Wait", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\2\", "Timeout", "-1", Microsoft.Win32.RegistryValueKind.String)
                    End If
                End If
            End If
        End If
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    REM PANEL1 HEREEEEEEEEEEEEEEEEEE

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        NotifyWindowD.ShowDialog(Me)
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

    End Sub
    Dim remSEMI3 As Integer = 0
    Private Sub ComboBox3_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.DropDown
        On Error Resume Next
        ComboBox3.Items.Clear()
        Dim i As String
        Dim CTSPath As String = "C:\Windows\Media\Alarms"
        If My.Computer.FileSystem.DirectoryExists(CTSPath) Then
            For Each i In My.Computer.FileSystem.GetFiles(CTSPath)
                ComboBox3.Items.Add(i.Substring(i.LastIndexOf("\") + 1))
            Next
        End If
        ComboBox3.SelectedIndex = remSEMI3
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        remSEMI3 = ComboBox3.SelectedIndex
    End Sub

    Private Sub APD_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Controller.Enabled = False
        MODE = 0
    End Sub
    Dim RV17
    Dim RV18
    Dim RV19
    Private Sub APD_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next
        Controller.Enabled = True
        If Dialog3.CheckBox25.Checked = True Then
            RV17 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingMode", Nothing)
            RV19 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\", "Mode", Nothing)
        Else
            RV17 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingMode", Nothing)
            RV19 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\3\", "Mode", Nothing)
        End If
        If RV17 = "0" Then
            remDUD1 = 0
            DomainUpDown1.SelectedIndex = 0
        ElseIf RV17 = "1" Then
            remDUD1 = 1
            DomainUpDown1.SelectedIndex = 1
        Else
            remDUD1 = 2
            DomainUpDown1.SelectedIndex = 2
        End If
        If Dialog3.CheckBox25.Checked = True Then
            RV18 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingMode", Nothing)
        Else
            RV18 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\2\", "RingMode", Nothing)
        End If
        If RV18 = "0" Then
            remDUD11 = 0
            DomainUpDown6.SelectedIndex = 0
        ElseIf RV18 = "1" Then
            remDUD11 = 1
            DomainUpDown6.SelectedIndex = 1
        Else
            remDUD11 = 2
            DomainUpDown6.SelectedIndex = 2
        End If
        If RV19 = "1" Then
            remRB21 = True
            RadioButton2.Checked = False
            RadioButton3.Checked = True
        Else
            remRB21 = False
            RadioButton2.Checked = True
            RadioButton3.Checked = False
        End If
        REM MAIN
        If MODE = 0 Then
            Panel1.Visible = False
            Panel2.Visible = False
            Panel3.Visible = False
        ElseIf MODE = 1 Then
            Panel1.Visible = True
            Panel2.Visible = False
            Panel3.Visible = False
            Me.Size = New Size(388, Panel1.Height + ControlP.Height + 33)
        ElseIf MODE = 2 Then
            Panel1.Visible = False
            Panel2.Visible = True
            Panel3.Visible = False
            Me.Size = New Size(388, Panel2.Height + ControlP.Height + 33)
        ElseIf MODE = 3 Then
            Panel1.Visible = False
            Panel2.Visible = False
            Panel3.Visible = True
            Me.Size = New Size(388, Panel3.Height + ControlP.Height + 33)
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        GroupBox2.Enabled = CheckBox1.Checked
    End Sub

    Private Sub DomainUpDown1_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DomainUpDown1.SelectedItemChanged
        If DomainUpDown1.SelectedIndex = 0 Then
            ComboBox3.Visible = False
            ComboBox4.Visible = False
            Button16.Visible = False
            ComboBox2.Visible = True
            ComboBox2.SelectedIndex = remCB2
        ElseIf DomainUpDown1.SelectedIndex = 1 Then
            Dim i As String
            Dim CTSPath As String = "C:\Windows\Media\Alarms"
            If My.Computer.FileSystem.DirectoryExists(CTSPath) Then
                For Each i In My.Computer.FileSystem.GetFiles(CTSPath)
                    ComboBox3.Items.Add(i.Substring(i.LastIndexOf("\") + 1))
                Next
            End If
            ComboBox2.Visible = False
            ComboBox4.Visible = False
            Button16.Visible = False
            ComboBox3.Visible = True
            ComboBox3.SelectedIndex = remCB3
        ElseIf DomainUpDown1.SelectedIndex = 2 Then
            ComboBox2.Visible = False
            ComboBox3.Visible = False
            Button16.Visible = True
            ComboBox4.Visible = True
            ComboBox4.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\1\", "RingPath", Nothing)
        End If
    End Sub

    Private Sub CheckBox16_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox16.CheckedChanged
        If CheckBox16.Checked = True Then
            Label27.Enabled = False
            ThmN.Enabled = False
        Else
            Label27.Enabled = True
            ThmN.Enabled = True
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 9 Then
            Label3.Enabled = True
            NumericUpDown1.Enabled = True
            NumericUpDown2.Enabled = True
        Else
            Label3.Enabled = False
            NumericUpDown1.Enabled = False
            NumericUpDown2.Enabled = False
        End If
    End Sub

    Private Sub DomainUpDown3_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DomainUpDown3.SelectedItemChanged
        If DomainUpDown3.SelectedIndex = 1 Then
            TextBox1.Enabled = True
        Else
            TextBox1.Enabled = False
        End If
    End Sub

    Private Sub DomainUpDown2_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DomainUpDown2.SelectedItemChanged
        If DomainUpDown2.SelectedIndex = 2 Then
            TextBox4.Enabled = True
            Button2.Enabled = True
        Else
            TextBox4.Enabled = False
            Button2.Enabled = False
        End If
    End Sub

    Private Sub DomainUpDown4_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DomainUpDown4.SelectedItemChanged
        If DomainUpDown4.SelectedIndex = 1 Then
            TextBox3.Enabled = True
        Else
            TextBox3.Enabled = False
        End If
    End Sub

    Private Sub DomainUpDown5_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DomainUpDown5.SelectedItemChanged
        If DomainUpDown5.SelectedIndex = 1 Then
            TextBox2.Enabled = True
        Else
            TextBox2.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Label8.Enabled = True
            DomainUpDown5.Enabled = True
            If DomainUpDown5.SelectedIndex = 1 Then
                TextBox2.Enabled = True
            End If
        Else
            Label8.Enabled = False
            DomainUpDown5.Enabled = False
            DomainUpDown5.SelectedIndex = 0
            TextBox2.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim ofd As New OpenFileDialog
        ofd.CheckPathExists = True
        ofd.Title = "Open File Dialog - Custom Picture for NotifyWindow"
        ofd.Filter = "Supported Picture Formats (*.png;*.jpg;*.bmp;*.ico)|*.png;*.jpeg;*.jpg;*.bmp;*.ico|All Files (*.*)|*.*"
        If IO.File.Exists(TextBox4.Text) = True Then
            Dim FI As New IO.FileInfo(TextBox4.Text)
            ofd.InitialDirectory = FI.DirectoryName
            ofd.FileName = TextBox4.Text
        ElseIf IO.Directory.Exists(TextBox4.Text) = True Then
            ofd.InitialDirectory = TextBox4.Text
        Else
            ofd.InitialDirectory = "C:\Users\Kristian\Pictures"
        End If
        If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
            TextBox4.Text = ofd.FileName
        End If
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Dim ofd As New OpenFileDialog
        ofd.CheckPathExists = True
        ofd.Title = "Open File Dialog - Custom Sound for Alarm"
        ofd.Filter = "Supported Sound Format (*.wav)|*.wav;*.wave"
        If IO.File.Exists(ComboBox4.Text) = True Then
            Dim FI As New IO.FileInfo(ComboBox4.Text)
            ofd.InitialDirectory = FI.DirectoryName
            ofd.FileName = ComboBox4.Text
        ElseIf IO.Directory.Exists(ComboBox4.Text) = True Then
            ofd.InitialDirectory = ComboBox4.Text
        Else
            ofd.InitialDirectory = "C:\Users\Kristian\Music"
        End If
        If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
            ComboBox4.Text = ofd.FileName
        End If
    End Sub

    REM PANEL 2 HEROOOOOOO

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            Label9.Enabled = False
            NumericUpDown3.Enabled = False
        Else
            Label9.Enabled = True
            NumericUpDown3.Enabled = True
        End If
    End Sub

    Private Sub DomainUpDown6_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DomainUpDown6.SelectedItemChanged
        If DomainUpDown6.SelectedIndex = 0 Then
            ComboBox5.Visible = False
            ComboBox7.Visible = False
            Button3.Visible = False
            ComboBox6.Visible = True
            ComboBox6.SelectedIndex = remCB11
        ElseIf DomainUpDown6.SelectedIndex = 1 Then
            Dim i As String
            Dim CTSPath As String = "C:\Windows\Media\Alarms"
            If My.Computer.FileSystem.DirectoryExists(CTSPath) Then
                For Each i In My.Computer.FileSystem.GetFiles(CTSPath)
                    ComboBox7.Items.Add(i.Substring(i.LastIndexOf("\") + 1))
                Next
            End If
            ComboBox6.Visible = False
            ComboBox7.Visible = True
            Button3.Visible = False
            ComboBox5.Visible = False
            ComboBox7.SelectedIndex = remCB12
        ElseIf DomainUpDown6.SelectedIndex = 2 Then
            ComboBox5.Visible = True
            ComboBox6.Visible = False
            Button3.Visible = True
            ComboBox7.Visible = False
            ComboBox5.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & Dialog3.AlarmList.SelectedItem & "\Mode\2\", "RingPath", Nothing)
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim ofd As New OpenFileDialog
        ofd.CheckPathExists = True
        ofd.Title = "Open File Dialog - Custom Sound for Alarm"
        ofd.Filter = "Supported Sound Format (*.wav)|*.wav;*.wave"
        If IO.File.Exists(ComboBox5.Text) = True Then
            Dim FI As New IO.FileInfo(ComboBox5.Text)
            ofd.InitialDirectory = FI.DirectoryName
            ofd.FileName = ComboBox5.Text
        ElseIf IO.Directory.Exists(ComboBox5.Text) = True Then
            ofd.InitialDirectory = ComboBox5.Text
        Else
            ofd.InitialDirectory = "C:\Users\Kristian\Music"
        End If
        If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
            ComboBox5.Text = ofd.FileName
        End If
    End Sub

    REM PANEL3

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            Panel31.Visible = True
            Panel32.Visible = False
        Else
            Panel32.Visible = True
            Panel31.Visible = False
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim ofd As New OpenFileDialog
        ofd.CheckPathExists = True
        ofd.Title = "Open File Dialog - Path program"
        ofd.Filter = "Applications (*.exe;*.scr;*.pif;*.com;*.lnk;*.bat;*.cmd;*.vbs)|*.exe;*.scr;*.pif;*.com;*.lnk;*.bat;*.cmd;*.vbs"
        If IO.File.Exists(ComboBox8.Text) = True Then
            Dim FI As New IO.FileInfo(ComboBox8.Text)
            ofd.InitialDirectory = FI.DirectoryName
            ofd.FileName = ComboBox8.Text
        ElseIf IO.Directory.Exists(ComboBox8.Text) = True Then
            ofd.InitialDirectory = ComboBox8.Text
        Else
            ofd.InitialDirectory = "C:\Users\Kristian\Music"
        End If
        If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
            ComboBox8.Text = ofd.FileName
        End If
    End Sub

    Private Sub Panel32_TabIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel32.TabIndexChanged
        CheckBox5.Checked = False
    End Sub

    Private Sub ComboBox8_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox8.TextChanged
        If IO.File.Exists(ComboBox8.Text) = True Then
            Panel32.Enabled = True
            Panel31.Enabled = True
        Else
            Panel32.Enabled = False
            Panel31.Enabled = False
        End If
    End Sub

    Private Sub Controller_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Controller.Tick
        If MODE = 1 Then
            If CheckBox1.Checked = True Then
                If DomainUpDown1.SelectedIndex = 2 Then
                    If IO.File.Exists(ComboBox4.Text) Then
                        Dim FI As New IO.FileInfo(ComboBox4.Text)
                        If Not FI.Extension = ".wav" OrElse FI.Extension = ".wave" Then
                            OK_Button.Enabled = False
                        Else
                            If DomainUpDown2.SelectedIndex = 2 Then
                                If IO.File.Exists(TextBox4.Text) Then
                                    OK_Button.Enabled = True
                                Else
                                    OK_Button.Enabled = False
                                End If
                            Else
                                OK_Button.Enabled = True
                            End If
                        End If
                    Else
                            OK_Button.Enabled = False
                    End If
                Else
                    If DomainUpDown2.SelectedIndex = 2 Then
                        If IO.File.Exists(TextBox4.Text) Then
                            OK_Button.Enabled = True
                        Else
                            OK_Button.Enabled = False
                        End If
                    Else
                        OK_Button.Enabled = True
                    End If
                End If
            Else
                If DomainUpDown2.SelectedIndex = 2 Then
                    If IO.File.Exists(TextBox4.Text) Then
                        OK_Button.Enabled = True
                    Else
                        OK_Button.Enabled = False
                    End If
                Else
                    OK_Button.Enabled = True
                End If
            End If
        ElseIf MODE = 2 Then
            If DomainUpDown6.SelectedIndex = 2 Then
                If IO.File.Exists(ComboBox5.Text) Then
                    Dim FI As New IO.FileInfo(ComboBox5.Text)
                    If Not FI.Extension = ".wav" OrElse FI.Extension = ".wave" Then
                        OK_Button.Enabled = False
                    Else
                        OK_Button.Enabled = True
                    End If
                Else
                    OK_Button.Enabled = False
                End If
            Else
                OK_Button.Enabled = True
            End If
        End If
0:
    End Sub

    Private Sub CheckBox10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox10.CheckedChanged
        TextBox7.Enabled = CheckBox10.Checked
    End Sub

    Private Sub CheckBox11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox11.CheckedChanged
        TextBox8.Enabled = CheckBox11.Checked
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            MaskedTextBox1.PasswordChar = ""
        Else
            MaskedTextBox1.PasswordChar = "*"
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        GroupBox5.Enabled = CheckBox4.Checked
    End Sub
    Dim remSEMI7 As Integer = 0
    Private Sub ComboBox7_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox7.DropDown
        On Error Resume Next
        ComboBox7.Items.Clear()
        Dim i As String
        Dim CTSPath As String = "C:\Windows\Media\Alarms"
        If My.Computer.FileSystem.DirectoryExists(CTSPath) Then
            For Each i In My.Computer.FileSystem.GetFiles(CTSPath)
                ComboBox7.Items.Add(i.Substring(i.LastIndexOf("\") + 1))
            Next
        End If
        ComboBox7.SelectedIndex = remSEMI7
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox7.SelectedIndexChanged
        remSEMI7 = ComboBox7.SelectedIndex
    End Sub

End Class
