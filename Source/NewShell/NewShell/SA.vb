Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Public Class SA
    Public Const WS_EX_TOOLWINDOW As Long = &H80L
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            ' Apply the ToolWindow style before the window is actually created
            cp.ExStyle = cp.ExStyle Or WS_EX_TOOLWINDOW
            Return cp
        End Get
    End Property

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
        AppBar.CanClose = False
        Desktop.CanClose = False

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

        Try
            PictureBox1.Image = GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\system32\imageres.dll", 104, False).ToBitmap
        Catch ex As Exception
            PictureBox1.Image = My.Resources.Clock
        End Try

        Try
            chkForceMode.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "ForceMode", 0)
        Catch ex As Exception
            chkForceMode.Checked = False
        End Try

        cmbPowerActions.DataSource = Nothing

        Dim actions = PowerManager.GetAvailablePowerActions()

        Dim defaultActionId = GetDefaultPowerAction()

        Dim style = Microsoft.Win32.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "Style", 0)

        If Convert.ToInt32(style) = 1 Then
            cmbPowerActions.Visible = False
            pnlRadioActions.Visible = True

            Me.MinimumSize = New Size(Me.MinimumSize.Width, lastMinFormHeight + radioHeightMode)
            Me.MaximumSize = New Size(Me.MinimumSize.Width, lastMaxFormHeight + radioHeightMode)

            GenerateRadioButtons(actions)
        Else
            cmbPowerActions.Visible = True
            pnlRadioActions.Visible = False

            Me.MinimumSize = New Size(Me.MinimumSize.Width, lastMinFormHeight)
            Me.MaximumSize = New Size(Me.MinimumSize.Width, lastMaxFormHeight)

            cmbPowerActions.DataSource = actions

            Select Case defaultActionId
                Case 2 : SetComboTo(cmbPowerActions, "Shut Down")
                Case 4 : SetComboTo(cmbPowerActions, "Restart")
                Case 1 : SetComboTo(cmbPowerActions, "Log Off")
                Case 64 : SetComboTo(cmbPowerActions, "Sleep")
            End Select
        End If

        Me.Size = New Size(Me.Width, Me.MinimumSize.Height)

        chkShutdownHybrid.Checked = PowerManager.IsFastStartupEnabled()
    End Sub

    Private Sub ExecuteAction()
        Dim selectedAction As String = String.Empty
        Dim userReason As UInt32 = 0

        If Not String.IsNullOrWhiteSpace(txtReason.Text) Then
            Try
                userReason = Convert.ToUInt32(txtReason.Text.Replace("&H", ""), 16)
            Catch
                userReason = 0
            End Try
        End If

        If cmbPowerActions.Visible Then
            selectedAction = cmbPowerActions.SelectedItem.ToString()
        Else
            Dim checkedRb = pnlRadioActions.Controls.OfType(Of RadioButton)().FirstOrDefault(Function(r) r.Checked)
            If checkedRb IsNot Nothing Then
                selectedAction = checkedRb.Tag.ToString()
            End If
        End If

        If Not String.IsNullOrEmpty(selectedAction) Then
            PowerManager.ExecuteAction(selectedAction, chkForceMode.Checked, userReason, chkShutdownHybrid.Checked, chkDisableWakeEvent.Checked, chkBrutalForce.Checked)
        End If
    End Sub

    Public Function GetDefaultPowerAction() As Integer
        Try
            Dim val = Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_PowerButtonAction", 2)
            Return DirectCast(val, Integer)
        Catch
            Return 2 ' Default to Shutdown
        End Try
    End Function
    Dim lastMinFormHeight As Integer = 150
    Dim lastMaxFormHeight As Integer = 270
    Dim radioHeightMode As Integer = 65
    Private Sub GenerateRadioButtons(actions As List(Of String))
        pnlRadioActions.Controls.Clear()
        pnlRadioActions.AutoScroll = True

        Dim startX As Integer = 0
        Dim startY As Integer = 2
        Dim currentY As Integer = startY
        Dim columnWidth As Integer = 160

        Dim defaultActionId = GetDefaultPowerAction()
        Dim defaultText As String = "Shut Down"

        Select Case defaultActionId
            Case 2 : defaultText = "Shut Down"
            Case 4 : defaultText = "Restart"
            Case 1 : defaultText = "Log Off"
            Case 64 : defaultText = "Sleep"
        End Select

        For Each action In actions
            If currentY + 25 > pnlRadioActions.Height - 10 Then
                currentY = startY
                startX += columnWidth
            End If

            Dim rb As New RadioButton With {
                .Text = action,
                .Tag = action,
                .AutoSize = True,
                .Location = New Point(startX, currentY),
                .FlatStyle = FlatStyle.System
            }

            If action.Contains(defaultText) Then
                rb.Checked = True
            End If

            pnlRadioActions.Controls.Add(rb)
            currentY += 25
        Next
    End Sub

    Private Sub SetComboTo(cb As ComboBox, text As String)
        Dim index = cb.FindStringExact(text)
        If index <> -1 Then cb.SelectedIndex = index
    End Sub

    Public Sub SaveSettings()
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "AutoHide", AppBar.AutoHideToolStripMenuItem.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer1", AppBar.Start.Width, Microsoft.Win32.RegistryValueKind.String)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer2", AppBar.Panel1.Width, Microsoft.Win32.RegistryValueKind.String)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer3", AppBar.Panel4.Width, Microsoft.Win32.RegistryValueKind.String)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", ColorTranslator.ToHtml(AppBar.BackColor), Microsoft.Win32.RegistryValueKind.String)

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\", "StartLastWidth", Startmenu.Width, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\", "StartLastHeight", Startmenu.Height, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        On Error GoTo 0
        Dim F As String = String.Empty

        AppBar.CanClose = True
        Desktop.CanClose = True

        SaveSettings()
        ExecuteAction()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        AppBar.CanClose = False
        Desktop.CanClose = False

        SaveSettings()
        Me.Close()
    End Sub

    Private Sub Topper_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.BringToFront()
    End Sub

    Private Sub ChkForceMode_CheckedChanged(sender As Object, e As EventArgs) Handles chkForceMode.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "ForceMode", chkForceMode.Checked, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPowerActions.SelectedIndexChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "LastAction", cmbPowerActions.SelectedIndex, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Panel2.Visible = CheckBox1.Checked

        If CheckBox1.Checked = True Then
            Me.Height = Me.MaximumSize.Height
        Else
            Me.Height = Me.MinimumSize.Height
        End If
    End Sub

    Private Sub chkBrutalForce_CheckedChanged(sender As Object, e As EventArgs) Handles chkBrutalForce.CheckedChanged
        If chkBrutalForce.Checked = True Then
            chkForceMode.Checked = False
            chkForceMode.Enabled = False
        Else
            chkForceMode.Enabled = True
        End If
    End Sub
End Class

Public Class PowerManager
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function ExitWindowsEx(uFlags As UInt32, dwReason As UInt32) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Sub LockWorkStation()
    End Sub

    <DllImport("wtsapi32.dll", SetLastError:=True)>
    Private Shared Function WTSDisconnectSession(hServer As IntPtr, SessionId As Integer, bWait As Boolean) As Boolean
    End Function

    <DllImport("PowrProf.dll", SetLastError:=True)>
    Private Shared Function GetPwrCapabilities(ByRef lpspc As SYSTEM_POWER_CAPABILITIES) As Boolean
    End Function

    <DllImport("PowrProf.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)>
    Public Shared Function SetSuspendState(hibernate As Boolean, forceCritical As Boolean, disableWakeEvent As Boolean) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Private Structure SYSTEM_POWER_CAPABILITIES
        <MarshalAs(UnmanagedType.U1)> Public PowerButtonPresent As Boolean
        <MarshalAs(UnmanagedType.U1)> Public SleepButtonPresent As Boolean
        <MarshalAs(UnmanagedType.U1)> Public LidPresent As Boolean
        <MarshalAs(UnmanagedType.U1)> Public SystemS1 As Boolean
        <MarshalAs(UnmanagedType.U1)> Public SystemS2 As Boolean
        <MarshalAs(UnmanagedType.U1)> Public SystemS3 As Boolean ' Sleep
        <MarshalAs(UnmanagedType.U1)> Public SystemS4 As Boolean ' Hibernate
        <MarshalAs(UnmanagedType.U1)> Public SystemS5 As Boolean ' Shutdown
        <MarshalAs(UnmanagedType.U1)> Public HiberFilePresent As Boolean
        ' ...
    End Structure

    Private Const EWX_LOGOFF As UInt32 = &H0
    Private Const EWX_SHUTDOWN As UInt32 = &H1
    Private Const EWX_REBOOT As UInt32 = &H2
    Private Const EWX_FORCE As UInt32 = &H4
    Private Const EWX_POWEROFF As UInt32 = &H8
    Private Const EWX_FORCEIFHUNG As UInt32 = &H10
    Private Const EWX_RESTARTAPPS As UInt32 = &H40
    Private Const EWX_HYBRID_SHUTDOWN As UInt32 = &H200000
    Private Const EWX_ARSO As UInt32 = &H4000000

    Public Shared Function GetAvailablePowerActions() As List(Of String)
        Dim actions As New List(Of String)

        Dim caps As New SYSTEM_POWER_CAPABILITIES

        Try
            GetPwrCapabilities(caps)
        Catch
        End Try

        actions.Add("Log Off")

        If Not IsShutdownRestricted() Then
            Dim updateStatus = GetUpdateStatus()

            If updateStatus.UpdateRestart Then actions.Add("Update and Restart")
            actions.Add("Restart")

            If IsWinREEnabled() Then
                actions.Add("Restart to WinRE")
            End If

            If IsUefiSupported() Then
                actions.Add("Restart to UEFI (BIOS)")
            End If

            If updateStatus.UpdateShutdown Then actions.Add("Update and Shutdown")
            actions.Add("Shut Down")

            If caps.SystemS3 Then actions.Add("Sleep")
            If caps.SystemS4 Then actions.Add("Hibernate")
        End If

        Return actions
    End Function

    Public Shared Function IsWinREEnabled() As Boolean
        Try
            Dim startInfo As New ProcessStartInfo("reagentc.exe", "/info")
            startInfo.RedirectStandardOutput = True
            startInfo.UseShellExecute = False
            startInfo.CreateNoWindow = True

            Using process As Process = Process.Start(startInfo)
                Dim output As String = process.StandardOutput.ReadToEnd()
                process.WaitForExit()

                Return output.Contains("Enabled") OrElse output.Contains("Povoleno")
            End Using
        Catch
            Return False
        End Try
    End Function

    Public Shared Sub RebootToWinRE()
        Try
            Dim startInfo As New ProcessStartInfo("reagentc.exe", "/boottore") With {
                .WindowStyle = ProcessWindowStyle.Hidden,
                .CreateNoWindow = True,
                .UseShellExecute = True,
                .Verb = "runas"
            }

            Dim p = Process.Start(startInfo)
            p.WaitForExit()

            ExecuteAction("Restart", False, False, &H80000000UI Or &H30000)
        Catch ex As Exception
            MessageBox.Show("Failed to perform Windows Recovery Environment: " & ex.Message, "Error")
        End Try
    End Sub

    Public Shared Function IsUefiSupported() As Boolean
        Try
            GetFirmwareEnvironmentVariable("", "{00000000-0000-0000-0000-000000000000}", IntPtr.Zero, 0)
            Dim lastError = Marshal.GetLastWin32Error()
            Return lastError <> 1
        Catch
            Return False
        End Try
    End Function

    <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function GetFirmwareEnvironmentVariable(lpName As String, lpGuid As String, pBuffer As IntPtr, nSize As UInt32) As UInt32
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Private Shared Function OpenProcessToken(ProcessHandle As IntPtr, DesiredAccess As Integer, ByRef TokenHandle As IntPtr) As Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function LookupPrivilegeValue(lpSystemName As String, lpName As String, ByRef lpLuid As Long) As Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Private Shared Function AdjustTokenPrivileges(TokenHandle As IntPtr, DisableAllPrivileges As Boolean, ByRef NewState As TOKEN_PRIVILEGES, BufferLength As Integer, PreviousState As IntPtr, ReturnLength As IntPtr) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential, Pack:=4)>
    Private Structure TOKEN_PRIVILEGES
        Public PrivilegeCount As Integer
        Public Luid As Long
        Public Attributes As Integer
    End Structure

    Private Const SE_PRIVILEGE_ENABLED As Integer = &H2
    Private Const TOKEN_QUERY As Integer = &H8
    Private Const TOKEN_ADJUST_PRIVILEGES As Integer = &H20
    Private Const SE_SHUTDOWN_NAME As String = "SeShutdownPrivilege"

    Private Shared Sub GrantShutdownPrivilege()
        Dim hToken As IntPtr
        Dim luid As Long
        Dim tp As New TOKEN_PRIVILEGES

        If OpenProcessToken(Process.GetCurrentProcess().Handle, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) Then
            LookupPrivilegeValue(Nothing, "SeShutdownPrivilege", luid)
            tp.PrivilegeCount = 1
            tp.Luid = luid
            tp.Attributes = SE_PRIVILEGE_ENABLED
            AdjustTokenPrivileges(hToken, False, tp, 0, IntPtr.Zero, IntPtr.Zero)

            LookupPrivilegeValue(Nothing, "SeSystemEnvironmentPrivilege", luid)
            tp.Luid = luid
            AdjustTokenPrivileges(hToken, False, tp, 0, IntPtr.Zero, IntPtr.Zero)
        End If
    End Sub

    Private Const SHTDN_REASON_MAJOR_SOFTWARE As UInt32 = &H30000
    Private Const SHTDN_REASON_MINOR_RECONFIG As UInt32 = &H4
    Private Const SHTDN_REASON_FLAG_PLANNED As UInt32 = &H80000000UI
    Private Const EWX_REBOOT_REASON_FIRMWARE As UInt32 = &H800000

    Public Shared Sub RebootToUefi()
        GrantShutdownPrivilege()

        Dim flags As UInt32 = &H2 Or &H10 Or &H800000UI
        Dim reason As UInt32 = EWX_REBOOT_REASON_FIRMWARE Or SHTDN_REASON_MAJOR_SOFTWARE Or SHTDN_REASON_MINOR_RECONFIG Or SHTDN_REASON_FLAG_PLANNED

        If Not ExitWindowsEx(flags, reason) Then
            Dim PSI As New ProcessStartInfo("shutdown.exe") With {
            .CreateNoWindow = True,
            .Arguments = "/r /fw /t 0",
            .WindowStyle = ProcessWindowStyle.Hidden,
            .ErrorDialog = True}

            Process.Start(PSI)
        End If
    End Sub

    Public Shared Sub ExecuteSessionAction(actionName As String)
        Select Case actionName
            Case "Lock"
                LockWorkStation()
            Case "Switch User"
                WTSDisconnectSession(IntPtr.Zero, -1, False)
            Case "Log Off"
                GrantShutdownPrivilege()
                ExitWindowsEx(EWX_LOGOFF, &H0)
        End Select
    End Sub

    Public Shared Sub ExecuteAction(actionName As String,
                                    Optional forceMode As Boolean = False,
                                    Optional customReason As UInt32 = 0,
                                    Optional hybridShutdown As Boolean = False,
                                    Optional disableWakeEvent As Boolean = False,
                                    Optional forceBrutal As Boolean = False)
        GrantShutdownPrivilege()

        Dim finalReason As UInt32
        Try
            If customReason = Nothing OrElse customReason.ToString() = "0" Then
                finalReason = &H40000UI Or &H80000000UI
            Else
                finalReason = Convert.ToUInt32(customReason)
            End If
        Catch
            finalReason = &H40000UI Or &H80000000UI
        End Try

        Dim forceFlag As UInt32 = 0
        If forceBrutal Then
            forceFlag = EWX_FORCE ' Brutal force
        ElseIf forceMode Then
            forceFlag = EWX_FORCEIFHUNG ' Calm force
        End If

        Select Case actionName
            Case "Log Off"
                ExitWindowsEx(EWX_LOGOFF Or forceFlag, finalReason)

            Case "Restart", "Update and Restart"
                ExitWindowsEx(EWX_REBOOT Or forceFlag, finalReason)

            Case "Shut Down", "Update and Shutdown"
                Dim shutdownType As UInt32 = If(IsLegacyShutdownEnforced(), EWX_SHUTDOWN, EWX_POWEROFF)
                Dim flags As UInt32 = shutdownType Or forceFlag

                If shutdownType = EWX_POWEROFF AndAlso hybridShutdown AndAlso Not forceBrutal Then
                    flags = flags Or EWX_HYBRID_SHUTDOWN
                End If
                ExitWindowsEx(flags, finalReason)

            Case "Sleep"
                SetSuspendState(False, forceMode Or forceBrutal, disableWakeEvent)

            Case "Hibernate"
                SetSuspendState(True, forceMode Or forceBrutal, disableWakeEvent)

            Case "Restart to UEFI (BIOS)"
                RebootToUefi()

            Case "Restart to WinRE"
                RebootToWinRE()

        End Select
    End Sub

    Private Shared Function IsLegacyShutdownEnforced() As Boolean
        Try
            Dim val = Registry.GetValue("HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\System", "ShutdownBehavior", 0)

            Return If(val IsNot Nothing, Convert.ToInt32(val) = 1, False)
        Catch
            Return False
        End Try
    End Function

    Public Shared Function IsFastStartupEnabled() As Boolean
        Try
            Dim val = Registry.GetValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power", "HiberbootEnabled", 0)
            Return Convert.ToInt32(val) = 1
        Catch
            Return False
        End Try
    End Function

    Private Shared Function GetUpdateStatus() As (UpdateShutdown As Boolean, UpdateRestart As Boolean)
        Try
            Dim val = Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\WindowsUpdate\Orchestrator", "ShutdownFlyoutOptions", 0)

            If val IsNot Nothing Then
                Dim bitField As Integer = Convert.ToInt32(val)

                ' Bit 1: Update and Shutdown
                ' Bit 3: Update and Restart
                Return ((bitField And 8) <> 0, (bitField And 2) <> 0)
            End If
        Catch
        End Try
        Return (False, False)
    End Function

    Private Shared Function IsShutdownRestricted() As Boolean
        Try
            Dim val = Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose", 0)
            If val IsNot Nothing Then
                Return Convert.ToInt32(val) = 1
            End If
        Catch
        End Try
        Return False
    End Function
End Class