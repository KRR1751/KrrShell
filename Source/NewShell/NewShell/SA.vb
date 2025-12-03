Imports System.Runtime.InteropServices

Public Class SA
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

    Private Sub SA_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        SaveSettings()
    End Sub

    Private Sub SA_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate, Me.LostFocus
        Me.Activate()
        Me.BringToFront()
    End Sub

    Private Sub SA_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        BlackScreen.Close()
    End Sub

    Private Sub SA_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        BlackScreen.Show()
        BlackScreen.BringToFront()
        Me.Focus()

        ComboBox1.SelectedIndex = 0

        Try
            PictureBox1.Image = GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\system32\imageres.dll", 104, False).ToBitmap
        Catch ex As Exception
            PictureBox1.Image = My.Resources.Clock
        End Try

        Try
            CheckBox1.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "ForceMode", 0)
        Catch ex As Exception
            CheckBox1.Checked = False
        End Try

        Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "Style", 0)
            Case 0
                ComboBox1.Visible = True
                Panel1.Visible = False

                Me.Size = Me.MinimumSize
            Case Else
                ComboBox1.Visible = False
                Panel1.Visible = True

                Me.Size = Me.MaximumSize
        End Select

        Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "LastAction", 0)
            Case 0
                RadioButton1.Checked = True

            Case 1
                RadioButton2.Checked = True

            Case 2
                RadioButton3.Checked = True

            Case 3
                RadioButton4.Checked = True

            Case 4
                RadioButton5.Checked = True

            Case Else

        End Select

        Try
            ComboBox1.SelectedIndex = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "LastAction", 0)
        Catch ex As Exception
            ComboBox1.SelectedIndex = 0
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        MessageBox.Show("Simply select what action the computer will do. This button is here just for astetic :D", "Help", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Public Sub SaveSettings()
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "AutoHide", AppBar.AutoHideToolStripMenuItem.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer1", AppBar.Button1.Width, Microsoft.Win32.RegistryValueKind.String)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer2", AppBar.Panel1.Width, Microsoft.Win32.RegistryValueKind.String)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer3", AppBar.Panel4.Width, Microsoft.Win32.RegistryValueKind.String)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", ColorTranslator.ToHtml(AppBar.BackColor), Microsoft.Win32.RegistryValueKind.String)

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\", "StartLastWidth", Startmenu.Width, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\", "StartLastHeight", Startmenu.Height, Microsoft.Win32.RegistryValueKind.DWord)

        If RadioButton1.Checked = True Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "LastAction", 0, Microsoft.Win32.RegistryValueKind.DWord)

        ElseIf RadioButton2.Checked = True Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "LastAction", 1, Microsoft.Win32.RegistryValueKind.DWord)

        ElseIf RadioButton3.Checked = True Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "LastAction", 2, Microsoft.Win32.RegistryValueKind.DWord)

        ElseIf RadioButton4.Checked = True Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "LastAction", 3, Microsoft.Win32.RegistryValueKind.DWord)

        ElseIf RadioButton5.Checked = True Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "LastAction", 4, Microsoft.Win32.RegistryValueKind.DWord)

        Else
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "LastAction", 0, Microsoft.Win32.RegistryValueKind.DWord)

        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        On Error GoTo 0
        Dim F As String = String.Empty

        AppBar.CanClose = True
        Desktop.CanClose = True

        If CheckBox1.Checked = True Then
            '---Force Mode---
            F = "/f "
        End If

        If Panel1.Visible = True Then
            If RadioButton1.Checked = True Then
                'Shutdown

                SaveSettings()
                Shell("shutdown.exe " & F & "/s /t 0", AppWinStyle.Hide, True)
                Me.Close()

            ElseIf RadioButton2.Checked = True Then
                'Shutdown with Fastboot

                SaveSettings()
                Shell("shutdown.exe " & F & "/s /hybrid /t 0", AppWinStyle.Hide, True)
                Me.Close()

            ElseIf RadioButton3.Checked = True Then
                'Restart

                SaveSettings()
                Shell("shutdown.exe " & F & "/r /t 0", AppWinStyle.Hide, True)
                Me.Close()

            ElseIf RadioButton4.Checked = True Then
                'Sleep

                SaveSettings()
                Shell("RunDll32.exe powrprof.dll,SetSuspendState", AppWinStyle.Hide, True)
                Me.Close()

            ElseIf RadioButton5.Checked = True Then
                'Hibernate

                SaveSettings()
                Shell("shutdown.exe /h", AppWinStyle.Hide, True)
                Me.Close()
            Else
0:              'Other

                AppBar.CanClose = False
                Desktop.CanClose = False
                MessageBox.Show("Failed to process this action. Please try that again later.", "Microsoft Windows", MessageBoxButtons.OK)
            End If
        ElseIf ComboBox1.Visible = True Then
            Select Case ComboBox1.SelectedIndex
                Case 0
                    'Shutdown

                    SaveSettings()
                    Shell("shutdown.exe " & F & "/s /t 0", AppWinStyle.Hide, True)
                    Me.Close()
                Case 1
                    'Shutdown with Fastboot

                    SaveSettings()
                    Shell("shutdown.exe " & F & "/s /hybrid /t 0", AppWinStyle.Hide, True)
                    Me.Close()
                Case 2
                    'Restart

                    SaveSettings()
                    Shell("shutdown.exe " & F & "/r /t 0", AppWinStyle.Hide, True)
                    Me.Close()
                Case 3
                    'Sleep

                    SaveSettings()
                    Shell("RunDll32.exe powrprof.dll,SetSuspendState", AppWinStyle.Hide, True)
                    Me.Close()
                Case 4
                    'Hibernate

                    SaveSettings()
                    Shell("shutdown.exe /h", AppWinStyle.Hide, True)
                    Me.Close()
                Case Else
                    'Other

                    AppBar.CanClose = False
                    Desktop.CanClose = False
                    MessageBox.Show("Failed to process this action. Please try that again later.", "Microsoft Windows", MessageBoxButtons.OK)
            End Select


        Else
            Me.Close()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        SaveSettings()
        Me.Close()
    End Sub

    Private Sub Topper_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.BringToFront()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "ForceMode", CheckBox1.Checked, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "LastAction", ComboBox1.SelectedIndex, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub
End Class