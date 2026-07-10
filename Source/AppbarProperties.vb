Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Security.Principal
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports KrrShell.AppBar

Public Class AppbarProperties
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindow(
     ByVal lpClassName As String,
     ByVal lpWindowName As String) As IntPtr
    End Function
    <DllImport("user32.dll")>
    Private Shared Function GetWindowPlacement(ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function
    <DllImport("user32.dll")>
    Private Shared Function SetWindowPlacement(ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function

    Private Structure WINDOWPLACEMENT
        Public Length As Integer
        Public flags As Integer
        Public showCmd As ShowWindowCommands
        Public ptMinPosition As POINTAPI
        Public ptMaxPosition As POINTAPI
        Public rcNormalPosition As RECT
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure DWMCOLORIZATIONPARAMS
        Public clrColor As UInteger
        Public clrAfterGlow As UInteger
        Public nIntensity As UInteger
        Public clrAfterGlowBalance As UInteger
        Public clrBlurBalance As UInteger
        Public nColorizationGlassAttribute As UInteger
        Public nColorizationOpaqueBlend As UInteger
    End Structure

    <DllImport("dwmapi.dll", EntryPoint:="#127")>
    Private Shared Function DwmGetColorizationParameters(ByRef parameters As DWMCOLORIZATIONPARAMS) As Integer
    End Function

    <DllImport("dwmapi.dll")>
    Private Shared Function DwmGetWindowAttribute(hWnd As IntPtr, dwAttribute As UInteger, ByRef pvAttribute As Boolean, cbAttribute As UInteger) As Integer
    End Function

    Private Const DWMWA_USE_HOST_BACKDROP_BRUSH As UInteger = 19
    Private Const DWMWA_CAPTION_COLOR As UInteger = 35

    Public Shared Function GetAccentColor() As Color
        Dim Params As New DWMCOLORIZATIONPARAMS()
        Dim result As Integer = DwmGetColorizationParameters(Params)

        If result = 0 Then
            Dim colorValue As UInteger = Params.clrColor

            Dim A As Integer = CInt((colorValue >> 24) And &HFF) ' Alpha
            Dim R As Integer = CInt((colorValue >> 16) And &HFF) ' Red
            Dim G As Integer = CInt((colorValue >> 8) And &HFF)  ' Green
            Dim B As Integer = CInt(colorValue And &HFF)        ' Blue

            Return Color.FromArgb(R, G, B)
        Else
            Return Color.FromArgb(0, 120, 215) ' Classic Blue from Windows
        End If
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        ApplySettings()

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Apply_Button_Click(sender As Object, e As EventArgs) Handles Apply_Button.Click
        ApplySettings()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        CancelAction()

        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub AppbarProperties_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Owner = AppBar

        LoadSettingsIntoAppbarProperties()
    End Sub

    Public Sub LoadSettingsIntoAppbarProperties()
        Dim liveBar = Application.OpenForms.OfType(Of AppBar)().FirstOrDefault()


        ' ---Appbar Main---
        ' Aplets

        chk_Aplet1.Checked = liveBar.StartButtonToolStripMenuItem.Checked
        chk_Aplet2.Checked = liveBar.PinnedBarToolStripMenuItem.Checked
        chk_Aplet3.Checked = liveBar.MenuBarToolStripMenuItem.Checked
        chk_Aplet4.Checked = liveBar.btn_LanguageBar.Visible
        chk_Aplet5.Checked = liveBar.btn_ShowDesktop.Visible
        chk_Aplet6.Checked = liveBar.ShowKeyboardButton
        chk_Aplet7.Checked = liveBar.ToolStripButton4.Visible
        chk_Aplet8.Checked = liveBar.ToolStripButton5.Visible

        ' Alignment

        chk_SmallIcons.Checked = isSmallIcons

        cmb_Alignment.Items.Clear()
        For Each alitem As String In [Enum].GetNames(GetType(AppBar.DockPosition))
            cmb_Alignment.Items.Add(alitem)
        Next

        cmb_Alignment.SelectedIndex = currentDockPos

        nud_SmallIcons.Value = defSmallSize
        nud_LargeIcons.Value = defSize

        ' Customization

        Try
            chk_DWMColor.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "UseSystemColor", 1)
        Catch ex As Exception
            chk_DWMColor.Checked = True
        End Try

        Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layout", 2)
            Case 0 ' DWM Transparent
                rb_ABCdwm.Checked = True

                btn_PickerBC.Enabled = True
                btn_PickerSC.Enabled = True

            Case 1 ' Default Transparent
                rb_ABCdefault.Checked = True

                btn_PickerBC.Enabled = True
                btn_PickerSC.Enabled = True

                lbl_ABCTip.Enabled = True
                tb_ABCtransparent.Enabled = True
                nud_ABCtransparent.Enabled = True

            Case 2 ' Opaque
                rb_ABCopaque.Checked = True

                btn_PickerBC.Enabled = True
                btn_PickerSC.Enabled = True

                lbl_ABCTip.Enabled = False
                tb_ABCtransparent.Enabled = False
                nud_ABCtransparent.Enabled = False
        End Select

        btn_PickerBC.BackColor = liveBar.BackColor
        btn_PickerSC.BackColor = liveBar.Splitter1.BackColor

        btn_PickerBC.ForeColor = liveBar.primaryForeColor
        btn_PickerSC.ForeColor = liveBar.primaryForeColor

        Try
            tb_ABCtransparent.Value = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "TransparencyLevel", "90")
            nud_ABCtransparent.Value = tb_ABCtransparent.Value
        Catch ex As Exception
            tb_ABCtransparent.Value = 90
            nud_ABCtransparent.Value = 90
        End Try

        ' Options

        chk_AutoHide.Checked = liveBar.AutoHideToolStripMenuItem.Checked
        chk_OnTop.Checked = liveBar.TopMost

        cmb_BtnTexts.SelectedIndex = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\", "CombineMode", "0")
        cmb_MiddleActions.SelectedIndex = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "MiddleAction", 0)

        chk_SysVolume.Checked = liveBar.UseLegacyVolumeSlider
        chk_SysRunDialog.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\", "UseSystemRunDialog", False)
        chk_AeroPeek.Checked = liveBar.isAeroPeekEnabled

        ' Hidden Processes

        Blacklist.Items.Clear()
        For Each i As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\HiddenProcesses").GetValueNames
            Blacklist.Items.Add(i)
        Next

        ' Custom Background

        chk_CustomImage.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowBackImage", False)

        If chk_CustomImage.Checked = False Then
            pb_CIPreview.Enabled = False
            cmb_CIPathToImage.Enabled = False
            cmb_CIImageLayout.Enabled = False
            btn_CIPathToImage.Enabled = False
        Else
            pb_CIPreview.Enabled = True
            cmb_CIPathToImage.Enabled = True
            cmb_CIImageLayout.Enabled = True
            btn_CIPathToImage.Enabled = True
        End If

        cmb_CIPathToImage.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImage", "")
        cmb_CIImageLayout.SelectedIndex = liveBar.BackgroundImageLayout

        pb_CIPreview.BackColor = liveBar.BackColor
        pb_CIPreview.BackgroundImageLayout = cmb_CIImageLayout.SelectedIndex


        '---Start-menu---
        ' Start Button

        Try : pb_SBDefault.Image = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Normal", "")) : Catch ex As Exception : pb_SBDefault.Image = My.Resources.StartRight : End Try
        Try : rb_SBHover.Image = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Hover", "")) : Catch ex As Exception : rb_SBHover.Image = My.Resources.StartRight : End Try
        Try : pb_SBPressed.Image = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Pressed", "")) : Catch ex As Exception : pb_SBPressed.Image = My.Resources.StartLeft : End Try

        cmb_SBPathToORB.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "CustomORB", "")

        Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0")
            Case 0 ' Default
                rb_SBCustomImage.Checked = False
                rb_SBCustomORB.Checked = False
                rb_SBDefault.Checked = True

            Case 1 ' Custom Images
                rb_SBDefault.Checked = False
                rb_SBCustomORB.Checked = False
                rb_SBCustomImage.Checked = True

            Case 2 ' Custom ORB
                rb_SBDefault.Checked = False
                rb_SBCustomImage.Checked = False
                rb_SBCustomORB.Checked = True

                If Not String.IsNullOrEmpty(cmb_SBPathToORB.Text) AndAlso File.Exists(cmb_SBPathToORB.Text) Then
                    pb_SBORB.Image = Image.FromFile(cmb_SBPathToORB.Text)
                End If

        End Select

        cmb_SBImageLayout.SelectedIndex = If(liveBar.Start.BackgroundImageLayout = ImageLayout.Stretch, 0, 1)


        '---Shutdown-Dialog---
        ' Main

        chk_SDrbbtns.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "Style", False)
        chk_SDback.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "Background", False)

        '---Hotkeys-(page)---

        chk_HKconfig.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "UseHotkeysForPinned", True)
        chk_pausetowindowscreen.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "PauseKeyToCaptureWindow", False)

        For i As Integer = 0 To 9
            AddHandler hotkeyPanel.Controls.Item(i).MouseUp, AddressOf HotkeyBtnRClick

            Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, i, "")

            If File.Exists(executePath) Then
                Try : hotkeyPanel.Controls.Item(i).BackgroundImage = Icon.ExtractAssociatedIcon(executePath).ToBitmap : Catch ex As Exception : hotkeyPanel.Controls.Item(i).BackgroundImage = My.Resources.ProgramSmall : End Try
                hotkeyPanel.Controls.Item(i).Text = String.Empty

                hotkeyPanel.Controls.Item(i).Tag = executePath
            Else
                hotkeyPanel.Controls.Item(i).BackgroundImage = Nothing
                hotkeyPanel.Controls.Item(i).Text = i

                hotkeyPanel.Controls.Item(i).Tag = String.Empty
            End If
        Next

        cmb_HKpreferences.Items.Clear()
        For Each v As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\CustomPaths\Hotkeys").GetSubKeyNames()
            cmb_HKpreferences.Items.Add(v)
        Next

        If cmb_HKpreferences.Items.Count = 0 Then
            cmb_HKpreferences.Enabled = False
            btn_HKRemovePR.Enabled = False
        Else
            cmb_HKpreferences.Enabled = True
            btn_HKRemovePR.Enabled = True

            Dim curHotkeyPreset As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths\Hotkeys", "CurrentPreset", String.Empty)
            Try
                cmb_HKpreferences.SelectedIndex = cmb_HKpreferences.FindString(curHotkeyPreset)
            Catch ex As Exception
                cmb_HKpreferences.SelectedIndex = 0
            End Try
        End If

        cmb_HKsearch.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "SearchApp", Nothing)
        cmb_HKcopilot.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "Assistant", Nothing)

        chk_textStylizer.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "TextStylizer", False)

        '---WorkingArea-(Page)---
        ' Main
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "Enabled", False) = 1 Then
            chk_WorkingAreaMain.Checked = True
        End If

        nud_WATop.Maximum = SystemInformation.PrimaryMonitorSize.Height
        nud_WABottom.Maximum = SystemInformation.PrimaryMonitorSize.Height
        nud_WALeft.Maximum = SystemInformation.PrimaryMonitorSize.Width
        nud_WARight.Maximum = SystemInformation.PrimaryMonitorSize.Width

        Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "Type", 0)

            Case 0 ' Explorer
                rb_WAexplorer.Checked = True

            Case 1 ' User32.dll
                rb_WAuser32.Checked = True

        End Select

        If My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\WorkingArea\Presets") IsNot Nothing Then

            cmb_WApreset.Items.Clear()

            ' Gets the list of presets into the Combobox.
            For Each i As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\WorkingArea\Presets").GetValueNames
                cmb_WApreset.Items.Add(i)
            Next
        End If

        Dim targetPreset As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "CurrentPreset", "Default")
        If targetPreset IsNot Nothing Then

            Dim realPreset As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea\Presets", targetPreset, "Default")
            If realPreset IsNot Nothing Then

                Dim parts As String() = realPreset.Split(" "c)

                If parts.Length >= 4 Then
                    nud_WATop.Value = parts(0)
                    nud_WABottom.Value = parts(1)
                    nud_WALeft.Value = parts(2)
                    nud_WARight.Value = parts(3)
                End If

                ' Selects an selected preset in Combobox.
                If cmb_WApreset.Items.Count > 0 Then
                    Try
                        cmb_WApreset.SelectedIndex = cmb_WApreset.FindStringExact(targetPreset)
                        'TriggerWorkingArea()
                    Catch ex As Exception
                        cmb_WApreset.SelectedIndex = 0
                    End Try
                End If
            Else ' Load defaults if no presets will be found.
                nud_WATop.Value = defWorkTop
                nud_WABottom.Value = defWorkBottom
                nud_WALeft.Value = defWorkLeft
                nud_WARight.Value = defWorkRight
                'TriggerWorkingArea()
            End If
        End If

        chk_WAAutoUpdate.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "AutoUpdate", True)
        chk_WAAppbarAuto.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "AppbarAlign", True)
        chk_WAturnoff.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "ResetAfterShutdown", True)


        '---Advanced---
        ' Custom File Explorer

        Dim usecustomfm As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "UseCustomFileManager", 0)
        Select Case usecustomfm
            Case 0
                rb_DFEclassic.Checked = True

            Case 1
                rb_DFEcustom.Checked = True
        End Select

        Dim customfm As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "CustomFileManager", String.Empty)
        If customfm IsNot Nothing AndAlso File.Exists(customfm) Then
            If New FileInfo(customfm).Extension.ToLower = ".exe" Then
                cmb_DFEPathToExplorer.Text = customfm
            End If
        End If

        Try
            btn_DFEnoCLSIDs.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "DisableCLSIDsWarnings", False)
        Catch ex As Exception
            btn_DFEnoCLSIDs.Checked = False
        End Try

        Try
            btn_DFEnoFilePicker.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "AvoidExplorerFileDialog", False)
        Catch ex As Exception
            btn_DFEnoFilePicker.Checked = False
        End Try

        ' Blocking Processes

        btn_BlockingProcesses.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "Enabled", False)

        If rb_BPOclose.Checked Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "KillMethod", 0, Microsoft.Win32.RegistryValueKind.DWord)
        ElseIf rb_BPOkill.Checked Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "KillMethod", 1, Microsoft.Win32.RegistryValueKind.DWord)
        End If

        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "KillMethod", "0") = 0 Then
            rb_BPOkill.Checked = True
            rb_BPOclose.Checked = False
        Else
            rb_BPOclose.Checked = True
            rb_BPOkill.Checked = False
        End If

        Blocklist.Items.Clear()
        For Each i As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\BlockedProcesses\.\").GetValueNames
            Blocklist.Items.Add(i)
        Next


        ' Alt Tab
        chk_ATenabled.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "Enabled", False)
        chk_ATAffectDWM.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "DWMAffect", False)
        chk_ATclassicmode.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "ClassicMode", False)
        nud_ATdelay.Value = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "ShowDelay", 0)


        ' ---Advanced-2---
        ' Custom paths
        cmb_CPmain.Items.Clear()
        For Each str As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\CustomPaths").GetValueNames
            cmb_CPmain.Items.Add(str)
        Next

        cmb_CPmain.SelectedIndex = 0


        ' Experimental Features

        chk_Balloons.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Windows\Explorer", "EnableLegacyBalloonNotifications", True)

    End Sub

    Public Sub ApplySettings()
        Dim liveBar = Application.OpenForms.OfType(Of AppBar)().FirstOrDefault()


        '---Appbar---
        ' Aplets
        '   Start
        liveBar.StartButtonToolStripMenuItem.Checked = chk_Aplet1.Checked
        liveBar.Start.Visible = liveBar.StartButtonToolStripMenuItem.Checked

        '   Pinned Bar
        liveBar.PinnedBarToolStripMenuItem.Checked = chk_Aplet2.Checked
        liveBar.Panel1.Visible = liveBar.PinnedBarToolStripMenuItem.Checked

        '   Menu Bar
        liveBar.MenuBarToolStripMenuItem.Checked = chk_Aplet3.Checked
        liveBar.Panel4.Visible = liveBar.MenuBarToolStripMenuItem.Checked

        '   Language List
        liveBar.btn_LanguageBar.Visible = chk_Aplet4.Checked

        '   Show Desktop
        liveBar.btn_ShowDesktop.Visible = chk_Aplet5.Checked

        '   Keyboard
        liveBar.ShowKeyboardButton = chk_Aplet6.Checked

        '  Saving...
        If liveBar.LockAppbarToolStripMenuItem.Checked = False Then
            liveBar.Splitter1.Visible = liveBar.StartButtonToolStripMenuItem.Checked
            liveBar.Splitter2.Visible = liveBar.PinnedBarToolStripMenuItem.Checked
            liveBar.Splitter3.Visible = liveBar.MenuBarToolStripMenuItem.Checked
        End If

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\", "StartButton", chk_Aplet1.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\", "PinnedBar", chk_Aplet2.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\", "MenuBar", chk_Aplet3.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "LanguageList", chk_Aplet4.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\", "ShowDesktopButton", chk_Aplet5.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowKeyboard", chk_Aplet6.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowPower", chk_Aplet7.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowNetwork", chk_Aplet8.Checked, Microsoft.Win32.RegistryValueKind.DWord)

        liveBar?.UpdatePowerUI()
        liveBar?.UpdateNetworkUI()

        ' Alignment

        defSize = nud_LargeIcons.Value
        defSmallSize = nud_SmallIcons.Value

        isSmallIcons = chk_SmallIcons.Checked
        currentDockPos = cmb_Alignment.SelectedIndex

        liveBar.Location = New Point(liveBar.Location.X - 1, liveBar.Location.Y - 1)
        liveBar.Size = New Size(liveBar.Width - 1, liveBar.Height - 1)

        '  Saving...

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\Alignment", "SmallIcons", chk_SmallIcons.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\Alignment", "Dock", cmb_Alignment.SelectedIndex, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\Alignment", "DefSize", nud_LargeIcons.Value, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\Alignment", "DefSmallSize", nud_SmallIcons.Value, Microsoft.Win32.RegistryValueKind.DWord)


        ' Customization
        Dim accentColor As Color = GetAccentColor()

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "ColorPrevalence", chk_DWMColor.Checked, Microsoft.Win32.RegistryValueKind.DWord)

        If chk_DWMColor.Checked = True Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", ColorTranslator.ToHtml(accentColor))
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "UseSystemColor", 1, Microsoft.Win32.RegistryValueKind.DWord)
        Else
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "UseSystemColor", 0, Microsoft.Win32.RegistryValueKind.DWord)
        End If

        'liveBar.BackColor = btn_PickerBC.BackColor
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", ColorTranslator.ToHtml(btn_PickerBC.BackColor))

        liveBar.Splitter1.BackColor = btn_PickerSC.BackColor
        liveBar.Splitter2.BackColor = btn_PickerSC.BackColor
        liveBar.Splitter3.BackColor = btn_PickerSC.BackColor
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "SplitterColor", ColorTranslator.ToHtml(btn_PickerSC.BackColor))

        If rb_ABCdwm.Checked = True Then
            Try
                Dim side As AppBar.Side = side
                side.Left = -1
                side.Right = -1
                side.Top = -1
                side.Bottom = -1
                Dim result As Integer = AppBar.DwmExtendFrameIntoClientArea(liveBar.Handle, side)
            Catch ex As Exception
                Dim side As AppBar.Side = side
                side.Left = 0
                side.Right = 0
                side.Top = 0
                side.Bottom = 0
                Dim result As Integer = AppBar.DwmExtendFrameIntoClientArea(liveBar.Handle, side)
            End Try

            liveBar.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", "#B77358"))
            liveBar.TransparencyKey = Nothing

            liveBar.Opacity = 1

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layout", 0, Microsoft.Win32.RegistryValueKind.DWord)

        ElseIf rb_ABCdefault.Checked = True Then
            Dim side As AppBar.Side = side
            side.Left = 0
            side.Right = 0
            side.Top = 0
            side.Bottom = 0
            Dim result As Integer = AppBar.DwmExtendFrameIntoClientArea(liveBar.Handle, side)

            liveBar.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", "#B77358"))
            liveBar.TransparencyKey = Nothing

            liveBar.Opacity = nud_ABCtransparent.Value / 100

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layout", 1, Microsoft.Win32.RegistryValueKind.DWord)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "TransparencyLevel", nud_ABCtransparent.Value)

        ElseIf rb_ABCopaque.Checked = True Then
            Dim side As AppBar.Side = side
            side.Left = 0
            side.Right = 0
            side.Top = 0
            side.Bottom = 0
            Dim result As Integer = AppBar.DwmExtendFrameIntoClientArea(liveBar.Handle, side)

            liveBar.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", "#B77358"))
            liveBar.TransparencyKey = Nothing

            liveBar.Opacity = 1

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layout", 2, Microsoft.Win32.RegistryValueKind.DWord)
        End If

        ' Options

        liveBar.AutoHideToolStripMenuItem.Checked = chk_AutoHide.Checked
        liveBar?.AutoHideOnOff()

        liveBar.TopMost = chk_OnTop.Checked

        liveBar.UseLegacyVolumeSlider = chk_SysVolume.Checked

        liveBar.IsAeroPeekEnabled = chk_AeroPeek.Checked

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "AutoHide", chk_AutoHide.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "OnTop", chk_OnTop.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Volume", "UseLegacy", chk_SysVolume.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\DWM", "EnableAeroPeek", chk_AeroPeek.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\", "CombineMode", cmb_BtnTexts.SelectedIndex, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "MiddleAction", cmb_MiddleActions.SelectedIndex, Microsoft.Win32.RegistryValueKind.DWord)

        RefreshHiddenProcesses()
        liveBar?.LoadApps()

        ' Custom BackImages
        Dim finalBackImage As String

        If chk_CustomImage.Checked = True Then
            If Not String.IsNullOrEmpty(cmb_CIPathToImage.Text) AndAlso File.Exists(cmb_CIPathToImage.Text) Then
                liveBar.BackgroundImage = Image.FromFile(cmb_CIPathToImage.Text)
                finalBackImage = cmb_CIPathToImage.Text.Trim
            Else
                liveBar.BackgroundImage = My.Resources.AppBarMain
                finalBackImage = ""
            End If
        Else
            liveBar.BackgroundImage = Nothing
            finalBackImage = ""
        End If

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowBackImage", chk_CustomImage.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImage", finalBackImage)


        '---Start-menu---
        ' Start Button

        If rb_SBDefault.Checked = True Then
            liveBar.Start.BackgroundImage = My.Resources.StartRight
            liveBar.Start.BackgroundImageLayout = ImageLayout.Stretch

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0", Microsoft.Win32.RegistryValueKind.DWord)

        ElseIf rb_SBCustomImage.Checked = True Then
            liveBar.Start.BackgroundImage = pb_SBDefault.Image
            liveBar.Start.BackgroundImageLayout = If(cmb_SBImageLayout.SelectedIndex = 0, ImageLayout.Stretch, ImageLayout.Center)

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "1", Microsoft.Win32.RegistryValueKind.DWord)

        ElseIf rb_SBCustomORB.Checked = True Then
            If Not String.IsNullOrEmpty(cmb_SBPathToORB.Text) AndAlso File.Exists(cmb_SBPathToORB.Text) Then
                liveBar?.LoadAndSplitOrb(cmb_SBPathToORB.Text)
            End If

            liveBar.Start.BackgroundImage = liveBar.OrbNormal
            liveBar.Start.BackgroundImageLayout = If(cmb_SBImageLayout.SelectedIndex = 0, ImageLayout.Stretch, ImageLayout.Center)

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "2", Microsoft.Win32.RegistryValueKind.DWord)
        End If

        If cmb_SBImageLayout.SelectedIndex = 0 Then
            liveBar.Start.BackgroundImageLayout = ImageLayout.Stretch
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Layout", "0", Microsoft.Win32.RegistryValueKind.DWord)
        Else
            liveBar.Start.BackgroundImageLayout = ImageLayout.Center
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Layout", "1", Microsoft.Win32.RegistryValueKind.DWord)
        End If

        '---Hotkeys---
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell", "UseHotkeysForPinned", chk_HKconfig.Checked, Microsoft.Win32.RegistryValueKind.DWord)

        My.Computer.Registry.SetValue(hotkeyStart, "1", If(hotkeyBTN1.Tag.ToString, ""))
        My.Computer.Registry.SetValue(hotkeyStart, "2", If(hotkeyBTN2.Tag.ToString, ""))
        My.Computer.Registry.SetValue(hotkeyStart, "3", If(hotkeyBTN3.Tag.ToString, ""))
        My.Computer.Registry.SetValue(hotkeyStart, "4", If(hotkeyBTN4.Tag.ToString, ""))
        My.Computer.Registry.SetValue(hotkeyStart, "5", If(hotkeyBTN5.Tag.ToString, ""))
        My.Computer.Registry.SetValue(hotkeyStart, "6", If(hotkeyBTN6.Tag.ToString, ""))
        My.Computer.Registry.SetValue(hotkeyStart, "7", If(hotkeyBTN7.Tag.ToString, ""))
        My.Computer.Registry.SetValue(hotkeyStart, "8", If(hotkeyBTN8.Tag.ToString, ""))
        My.Computer.Registry.SetValue(hotkeyStart, "9", If(hotkeyBTN9.Tag.ToString, ""))
        My.Computer.Registry.SetValue(hotkeyStart, "0", If(hotkeyBTN0.Tag.ToString, ""))

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell", "TextStylizer", chk_textStylizer.Checked, Microsoft.Win32.RegistryValueKind.DWord)

        '---Working-Area---
        'Main

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "Enabled", chk_WorkingAreaMain.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "AutoUpdate", chk_WAAutoUpdate.Checked, Microsoft.Win32.RegistryValueKind.DWord)

        AlocateAppbarWorkingArea = chk_WAAppbarAuto.Checked
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "AppbarAlign", chk_WAAppbarAuto.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "ResetAfterShutdown", chk_WAturnoff.Checked, Microsoft.Win32.RegistryValueKind.DWord)

        If rb_WAexplorer.Checked = True Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "Type", 0, Microsoft.Win32.RegistryValueKind.DWord)
            TriggerWorkingArea()

        ElseIf rb_WAuser32.Checked = True Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "Type", 1, Microsoft.Win32.RegistryValueKind.DWord)
            TriggerWorkingArea()

        End If


        '---Advanced---
        ' File Explorer
        If Not String.IsNullOrWhiteSpace(cmb_DFEPathToExplorer.Text) AndAlso File.Exists(cmb_DFEPathToExplorer.Text) Then
            If New FileInfo(cmb_DFEPathToExplorer.Text).Extension.ToLower = ".exe" Then
                liveBar.CustomFMPath = cmb_DFEPathToExplorer.Text

                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell", "CustomFileManager", cmb_DFEPathToExplorer.Text)
            End If
        End If

        If rb_DFEclassic.Checked = True Then

            liveBar.UseExplorerFM = True
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell", "UseCustomFileManager", False, Microsoft.Win32.RegistryValueKind.DWord)

        ElseIf rb_DFEcustom.Checked = True Then

            liveBar.UseExplorerFM = False
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell", "UseCustomFileManager", True, Microsoft.Win32.RegistryValueKind.DWord)

            If File.Exists(cmb_DFEPathToExplorer.Text) Then
                If New FileInfo(cmb_DFEPathToExplorer.Text).Extension.ToLower = ".exe" Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell", "CustomFileManager", cmb_DFEPathToExplorer.Text)
                End If
            End If
        End If

        liveBar.UseExplorerFP = btn_DFEnoFilePicker.Checked <> True

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell", "AvoidExplorerFileDialog", btn_DFEnoFilePicker.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell", "DisableCLSIDsWarnings", btn_DFEnoCLSIDs.Checked, Microsoft.Win32.RegistryValueKind.DWord)


        ' Blocking Processes
        liveBar?.ToggleProcessMonitoring(btn_BlockingProcesses.Checked)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "Enabled", btn_BlockingProcesses.Checked, Microsoft.Win32.RegistryValueKind.DWord)

        If rb_BPOclose.Checked Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "KillMethod", 0, Microsoft.Win32.RegistryValueKind.DWord)
        ElseIf rb_BPOkill.Checked Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "KillMethod", 1, Microsoft.Win32.RegistryValueKind.DWord)
        End If

        RefreshBlockedProcesses()


        ' Alt Tab
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "Enabled", chk_ATenabled.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "DWMAffect", chk_ATAffectDWM.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "ClassicMode", chk_ATclassicmode.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "ShowDelay", nud_ATdelay.Value, Microsoft.Win32.RegistryValueKind.DWord)


        ' Experimental Features
        ' ...
    End Sub

    Public Sub CancelAction()

    End Sub

    Dim hotkeyStart As String = "HKEY_CURRENT_USER\Software\Shell\CustomPaths\Hotkeys"

    Public defWorkTop As Integer = 0
    Public defWorkBottom As Integer = 0
    Public defWorkLeft As Integer = 0
    Public defWorkRight As Integer = 0

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btn_HPAddProcess.Click
        'On Error Resume Next
        Dim Output As String = InputBox("Type a process name do you want to Hide from the Appbar...", "Input Box", "")
        If Output IsNot String.Empty Then
            Try
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\HiddenProcesses\", Output, "")
                Blacklist.Items.Add(Output)
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btn_HPRemoveProcess.Click
        'On Error Resume Next
        If MessageBox.Show("Are you sure do you want to remove those selected processes from your Blacklist?", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            Shell("reg.exe delete HKEY_CURRENT_USER\Software\Shell\Appbar\HiddenProcesses\ /v " & Blacklist.SelectedItem & " /f", AppWinStyle.Hide, True)
            Blacklist.Items.Clear()
            For Each ii As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\HiddenProcesses").GetValueNames
                Blacklist.Items.Add(ii)
            Next

            AppBar.LoadApps()
        End If
    End Sub

    Private Sub Controller_Tick(sender As Object, e As EventArgs) Handles Controller.Tick
        If Blacklist.SelectedItems.Count = 0 Then
            btn_HPRemoveProcess.Enabled = False
            btn_HPClearProcesses.Enabled = True
        Else
            btn_HPClearProcesses.Enabled = False
            btn_HPRemoveProcess.Enabled = True
        End If

        If Blocklist.SelectedItems.Count = 0 Then
            btn_BPRemovePr.Enabled = False
            btn_BPClearPr.Enabled = True
        Else
            btn_BPClearPr.Enabled = False
            btn_BPRemovePr.Enabled = True
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles btn_HPClearProcesses.Click
        On Error Resume Next
        If MessageBox.Show("Are you sure do you want to remove All the processes from your Blacklist?", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            For i As Integer = 0 To Blacklist.Items.Count - 1
                Shell("reg.exe delete HKEY_CURRENT_USER\Software\Shell\Appbar\HiddenProcesses\ /v " & Blacklist.SelectedItem & " /f", AppWinStyle.Hide, True)
            Next
            Blacklist.Items.Clear()

            AppBar.LoadApps()
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles btn_BlockingProcesses.CheckedChanged
        grp_BlockingProcesses.Enabled = btn_BlockingProcesses.Checked
    End Sub

    Public Enum CriticalProcesses
        winlogon
        dwm
        ntoskrnl
        csrss
        smss
        svchost
        lsass
        wininit
        rundll32
        taskhostw
        wudfhost
        sppsvc
        services
        fontdrvhost
        conhost
        krrshell
        newshell
    End Enum

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles btn_BPAddPr.Click
        Dim Output As String = InputBox("Type a process name Do you want To Block entirely...", "Input Box", "")
        If Not String.IsNullOrWhiteSpace(Output) Then

            For Each str As String In [Enum].GetNames(GetType(CriticalProcesses))
                If Output.ToLower = str.ToLower Then
                    MessageBox.Show("This process is NOT ALLOWED to be blocked. Since """ & Output & """ is a Critical Windows process that cannot be easily terminated." & Environment.NewLine & "And for your safety, it will not be included to the Blocklist!", "This process cannot be ""Blocked""", MessageBoxButtons.OK, MessageBoxIcon.Hand)
                    Exit Sub
                End If
            Next

            Try
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses\.\", Output, "")
                Blocklist.Items.Add(Output)
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles btn_BPRemovePr.Click
        On Error Resume Next
        If MessageBox.Show("Are you sure do you want to remove those selected processes from your Blocklist?", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            For i As Integer = 0 To Blocklist.SelectedItems.Count - 1
                'My.Computer.Registry.CurrentUser.OpenSubKey("\Software\Shell\Appbar\BlockedProcesses\.\", True, True).DeleteValue(Blocklist.SelectedItems.Item(i))
                Shell("reg.exe delete HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses\.\ /v " & Blocklist.SelectedItem & " /f", AppWinStyle.Hide, True)
                Blocklist.Items.Clear()
                For Each ii As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\BlockedProcesses\.\").GetValueNames
                    Blocklist.Items.Add(ii)
                Next
            Next
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles btn_BPClearPr.Click
        On Error Resume Next
        If MessageBox.Show("Are you sure do you want to remove All the processes from your Blocklist?", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            For i As Integer = 0 To Blocklist.Items.Count - 1
                Shell("reg.exe delete HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses\.\ /v " & Blocklist.SelectedItem & " /f", AppWinStyle.Hide, True)
            Next
            Blocklist.Items.Clear()
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles rb_BPOkill.CheckedChanged
        If rb_BPOkill.Checked = True Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses\", "KillMethod", "0", Microsoft.Win32.RegistryValueKind.DWord)
        Else
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses\", "KillMethod", "1", Microsoft.Win32.RegistryValueKind.DWord)
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles llbl_About.LinkClicked
        On Error Resume Next
        AboutDialog.Show(Me)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles btn_PickerBC.Click
        Dim CD As New ColorDialog With {
            .AllowFullOpen = True,
            .AnyColor = True,
            .FullOpen = True,
            .Color = AppBar.BackColor
        }

        If CD.ShowDialog(Me) = DialogResult.OK Then
            btn_PickerBC.BackColor = CD.Color
            pb_CIPreview.BackColor = CD.Color
        End If
    End Sub

    Private Sub RadioButton6_Click(sender As Object, e As EventArgs) Handles rb_ABCdwm.Click
        If rb_ABCdwm.Checked = True Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layout", 0, Microsoft.Win32.RegistryValueKind.DWord)
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) ' Old
        ' When I'll come with an alternative to display an transparent window
        ' I'll replace this

        Dim side As AppBar.Side = side
        side.Left = 0
        side.Right = 0
        side.Top = 0
        side.Bottom = 0
        Dim result As Integer = AppBar.DwmExtendFrameIntoClientArea(AppBar.Handle, side)

        btn_PickerBC.Enabled = False
        btn_PickerSC.Enabled = False

        AppBar.BackColor = Color.Black
        AppBar.TransparencyKey = Color.Black

        AppBar.Opacity = 1

        'My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layout", 1, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles rb_ABCopaque.Click, rb_ABCdefault.Click
        Dim side As AppBar.Side = side
        side.Left = 0
        side.Right = 0
        side.Top = 0
        side.Bottom = 0
        Dim result As Integer = AppBar.DwmExtendFrameIntoClientArea(AppBar.Handle, side)

        btn_PickerBC.Enabled = True
        btn_PickerSC.Enabled = True

        AppBar.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", "#B77358"))
        AppBar.TransparencyKey = Nothing

        If rb_ABCdefault.Checked = True Then
            lbl_ABCTip.Enabled = True
            tb_ABCtransparent.Enabled = True
            nud_ABCtransparent.Enabled = True

            AppBar.Opacity = tb_ABCtransparent.Value

            'My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layout", 1, Microsoft.Win32.RegistryValueKind.DWord)
        Else
            lbl_ABCTip.Enabled = False
            tb_ABCtransparent.Enabled = False
            nud_ABCtransparent.Enabled = False

            AppBar.Opacity = 1

            'My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layout", 2, Microsoft.Win32.RegistryValueKind.DWord)
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles btn_PickerSC.Click
        Dim CD As New ColorDialog With {
            .AllowFullOpen = True,
            .AnyColor = True,
            .FullOpen = True,
            .Color = AppBar.Splitter1.BackColor
        }

        If CD.ShowDialog(Me) = DialogResult.OK Then
            btn_PickerSC.BackColor = CD.Color
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles chk_CustomImage.CheckedChanged
        If chk_CustomImage.Checked = True Then
            grp_CustomImage.Enabled = True
        Else
            grp_CustomImage.Enabled = False
        End If
    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles cmb_CIPathToImage.TextChanged
        If cmb_CIPathToImage.Enabled = True Then
            Try
                If IO.File.Exists(cmb_CIPathToImage.Text) = True Then
                    Dim fi As New FileInfo(cmb_CIPathToImage.Text)
                    If fi.Extension = ".png" OrElse fi.Extension = ".jpg" OrElse fi.Extension = ".bmp" OrElse fi.Extension = ".gif" Then
                        pb_CIPreview.Image = Image.FromFile(fi.FullName)
                        AppBar.BackgroundImage = Image.FromFile(fi.FullName)
                    End If

                ElseIf cmb_CIPathToImage.Text = "" Then
                    pb_CIPreview.Image = My.Resources.AppBarMain
                    AppBar.BackgroundImage = My.Resources.AppBarMain
                End If

            Catch ex As Exception
                pb_CIPreview.Image = My.Resources.AppBarMain
                AppBar.BackgroundImage = My.Resources.AppBarMain

            End Try
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles btn_CIPathToImage.Click
        Dim ofd As New OpenFileDialog With {
            .AutoUpgradeEnabled = AppBar.UseExplorerFP,
            .Title = "Open a ""Appbar"" image for to show in the Appbar",
            .Filter = "Pictures (*.png;*.jpg;*.bmp;*.gif)|*.png;*.jpg;*.bmp;*.gif",
            .CheckFileExists = True,
            .Multiselect = False
        }
        If IO.File.Exists(cmb_CIPathToImage.Text) = True Then
            Dim fi As New FileInfo(cmb_CIPathToImage.Text)
            ofd.InitialDirectory = fi.DirectoryName
            ofd.FileName = fi.Name
        ElseIf IO.Directory.Exists(cmb_CIPathToImage.Text) = True Then
            ofd.InitialDirectory = cmb_CIPathToImage.Text
            ofd.FileName = ""
        Else
            ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
            ofd.FileName = ""
        End If

        If ofd.ShowDialog(Me) = DialogResult.OK Then
            cmb_CIPathToImage.Text = ofd.FileName
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_CIImageLayout.SelectedIndexChanged
        AppBar.BackgroundImageLayout = cmb_CIImageLayout.SelectedIndex
        pb_CIPreview.BackgroundImageLayout = cmb_CIImageLayout.SelectedIndex
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImageLayout", cmb_CIImageLayout.SelectedIndex, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long,
                                                        ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long,
                                                        ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOZORDER As UInteger = &H4
    Private Const SWP_SHOWWINDOW As UInteger = &H40

    Private Const HWND_TOPMOST = -1 '-- Bring to top and stay there
    Private Sub MoveAndResizeWindow(ByVal hWnd As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer)
        SetWindowPos(hWnd, IntPtr.Zero, X, Y, Width, Height, SWP_NOZORDER Or SWP_SHOWWINDOW)
    End Sub

    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles rb_SBDefault.CheckedChanged, rb_SBCustomImage.CheckedChanged, rb_SBCustomORB.CheckedChanged
        If rb_SBDefault.Checked = True Then
            pnl_SB2.Enabled = False
            pnl_SB3.Enabled = False

        ElseIf rb_SBCustomImage.Checked = True Then
            pnl_SB2.Enabled = True
            pnl_SB3.Enabled = False

        ElseIf rb_SBCustomORB.Checked = True Then
            pnl_SB2.Enabled = False
            pnl_SB3.Enabled = True

        End If
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles pb_SBDefault.Click
        On Error Resume Next
        Dim OFD As New OpenFileDialog With {
            .AutoUpgradeEnabled = AppBar.UseExplorerFP,
            .Title = "Select a Custom Image for ""Default"" button state",
            .CheckFileExists = True,
            .FileName = "",
            .Filter = "Image Support Formats (*.png;*.jpg*;*.gif;*.bmp)|*.png;*.jpg*;*.gif;*.bmp|All files (*.*)|*.*"
        }

        If OFD.ShowDialog = DialogResult.OK Then
            pb_SBDefault.Image = Image.FromFile(OFD.FileName)
            AppBar.Start.BackgroundImage = Image.FromFile(OFD.FileName)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Normal", OFD.FileName)
        End If
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles rb_SBHover.Click
        On Error Resume Next
        Dim OFD As New OpenFileDialog With {
            .AutoUpgradeEnabled = AppBar.UseExplorerFP,
            .Title = "Select a Custom Image for ""Hover"" button state",
            .CheckFileExists = True,
            .FileName = "",
            .Filter = "Image Support Formats (*.png;*.jpg*;*.gif;*.bmp)|*.png;*.jpg*;*.gif;*.bmp|All files (*.*)|*.*"
        }

        If OFD.ShowDialog = DialogResult.OK Then
            rb_SBHover.Image = Image.FromFile(OFD.FileName)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Hover", OFD.FileName)
        End If
    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles pb_SBPressed.Click
        On Error Resume Next
        Dim OFD As New OpenFileDialog With {
            .AutoUpgradeEnabled = AppBar.UseExplorerFP,
            .Title = "Select a Custom Image for ""Pressed"" button state",
            .CheckFileExists = True,
            .FileName = "",
            .Filter = "Image Support Formats (*.png;*.jpg*;*.gif;*.bmp)|*.png;*.jpg*;*.gif;*.bmp|All files (*.*)|*.*"
        }

        If OFD.ShowDialog = DialogResult.OK Then
            pb_SBPressed.Image = Image.FromFile(OFD.FileName)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Pressed", OFD.FileName)
        End If
    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles tb_ABCtransparent.Scroll
        nud_ABCtransparent.Value = tb_ABCtransparent.Value
    End Sub

    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles nud_ABCtransparent.ValueChanged
        tb_ABCtransparent.Value = nud_ABCtransparent.Value
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles chk_DWMColor.CheckedChanged
        Dim accentColor As Color = GetAccentColor()

        If chk_DWMColor.Checked = True Then
            btn_PickerBC.BackColor = accentColor
            pb_CIPreview.BackColor = accentColor
            AppBar.BackColor = accentColor
        End If
    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles chk_SDrbbtns.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "Style", chk_SDrbbtns.Checked, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles btn_SBPathToORB.Click
        On Error Resume Next
        Dim OFD As New OpenFileDialog With {
            .AutoUpgradeEnabled = AppBar.UseExplorerFP,
            .Title = "Select a Custom ORB for your Start Menu",
            .CheckFileExists = True,
            .FileName = "",
            .Filter = "Image Support Formats (*.png;*.jpg*;*.gif;*.bmp)|*.png;*.jpg*;*.gif;*.bmp|All files (*.*)|*.*"
        }

        If OFD.ShowDialog = DialogResult.OK Then
            pb_SBORB.Image = Image.FromFile(OFD.FileName)
            cmb_SBPathToORB.Text = OFD.FileName

            AppBar.LoadAndSplitOrb(OFD.FileName)
            AppBar.Start.BackgroundImage = AppBar.OrbNormal
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "CustomORB", OFD.FileName)

        End If
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles btn_WASetWorkingArea.Click, rb_WAexplorer.Click, rb_WAuser32.Click
        TriggerWorkingArea()
    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles chk_WorkingAreaMain.CheckedChanged
        pnl_WorkingAreaMain.Enabled = chk_WorkingAreaMain.Checked
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles rb_WAexplorer.CheckedChanged, rb_WAuser32.CheckedChanged
        If rb_WAexplorer.Checked = True Then
            pnl_WAmain.Enabled = False

            btn_WAregister.Enabled = True
            btn_WAunregister.Enabled = True

        ElseIf rb_WAuser32.Checked = True Then
            btn_WAregister.Enabled = False
            btn_WAunregister.Enabled = False

            pnl_WAmain.Enabled = True
        End If
    End Sub

    ' Working Area Updates
    Public Sub TriggerWorkingArea()
        Dim liveBar = Application.OpenForms.OfType(Of AppBar)().FirstOrDefault()

        If rb_WAexplorer.Checked = True Then
            ' On Explorer

            liveBar.fBarRegistered = False
            liveBar?.RegisterBar()

        ElseIf rb_WAuser32.Checked = True Then
            ' On User32.dll

            UpdateWorkingArea()

        End If
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_WApreset.SelectedIndexChanged
        Dim selectedPreset As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea\Presets", cmb_WApreset.SelectedItem, "Default")
        If selectedPreset IsNot Nothing Then

            If cmb_WApreset.SelectedItem = "Default" Then btn_WARemovePr.Enabled = False Else btn_WARemovePr.Enabled = True

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "CurrentPreset", cmb_WApreset.SelectedItem)

            Try
                Dim parts As String() = selectedPreset.Split(" "c)

                If parts.Length >= 4 Then
                    nud_WATop.Value = parts(0)
                    nud_WABottom.Value = parts(1)
                    nud_WALeft.Value = parts(2)
                    nud_WARight.Value = parts(3)
                End If

            Catch ex As Exception
                nud_WATop.Value = defWorkTop
                nud_WABottom.Value = defWorkBottom
                nud_WALeft.Value = defWorkLeft
                nud_WARight.Value = defWorkRight
            End Try
        End If

        If chk_WAAutoUpdate.Checked = True Then TriggerWorkingArea()
    End Sub

    Private Sub CheckBox14_CheckedChanged(sender As Object, e As EventArgs) Handles chk_WAAutoUpdate.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "AutoUpdate", chk_WAAutoUpdate.Checked, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles btn_WARemovePr.Click

        Dim selectedPreset As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea\Presets", cmb_WApreset.SelectedItem, "Default")
        If selectedPreset IsNot Nothing AndAlso cmb_WApreset.SelectedItem <> "Default" Then

            If MessageBox.Show($"Do you really want to remove this preset? [{cmb_WApreset.SelectedItem}]", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\WorkingArea\Presets", True)

                If myKey IsNot Nothing Then
                    Try
                        myKey.DeleteValue(cmb_WApreset.SelectedItem)
                    Finally
                        myKey.Close()
                    End Try
                End If

                cmb_WApreset.Items.RemoveAt(cmb_WApreset.FindStringExact(cmb_WApreset.SelectedItem))
                cmb_WApreset.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles btn_WASave.Click
        Dim selectedPreset As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea\Presets", cmb_WApreset.SelectedItem, "Default")
        If selectedPreset IsNot Nothing Then

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea\Presets\", cmb_WApreset.SelectedItem,
                                          $"{nud_WATop.Value} {nud_WABottom.Value} {nud_WALeft.Value} {nud_WARight.Value}")

        End If
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles btn_WAAddPr.Click
        Dim createdPreset As String = InputBox("Type a name for your made preset how it will be called.", "Saving a new Preset...", "MyCustom")
        If createdPreset IsNot Nothing Then

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea\Presets\", createdPreset,
                                          $"{nud_WATop.Value} {nud_WABottom.Value} {nud_WALeft.Value} {nud_WARight.Value}")

            cmb_WApreset.Items.Add(createdPreset)
            cmb_WApreset.SelectedIndex = cmb_WApreset.FindStringExact(createdPreset)

        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles btn_WAregister.Click
        AppBar.fBarRegistered = False
        AppBar.RegisterBar()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles btn_WAunregister.Click
        AppBar.fBarRegistered = True
        AppBar.RegisterBar()
    End Sub

    Private Sub RadioButton11_CheckedChanged(sender As Object, e As EventArgs) Handles rb_DFEclassic.CheckedChanged, rb_DFEcustom.CheckedChanged
        If rb_DFEclassic.Checked = True Then

            cmb_DFEPathToExplorer.Enabled = False
            btn_DFEPathToExplorer.Enabled = False

        ElseIf rb_DFEcustom.Checked = True Then

            cmb_DFEPathToExplorer.Enabled = True
            btn_DFEPathToExplorer.Enabled = True

        End If
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles btn_DFEPathToExplorer.Click
        Dim ofd As New OpenFileDialog With {
            .AutoUpgradeEnabled = AppBar.UseExplorerFP,
            .ShowHelp = True,
            .CheckFileExists = True,
            .Title = "Select a custom File Manager for this shell...",
            .Filter = "Executables (*.exe)|*.exe"
        }

        If File.Exists(cmb_DFEPathToExplorer.Text) Then
            Dim FI As New FileInfo(cmb_DFEPathToExplorer.Text)

            ofd.InitialDirectory = FI.DirectoryName
            ofd.FileName = FI.Name
        ElseIf Directory.Exists(cmb_DFEPathToExplorer.Text) Then
            ofd.InitialDirectory = cmb_DFEPathToExplorer.Text
        Else
            ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
        End If

        AddHandler ofd.HelpRequest, Sub()
                                        MessageBox.Show("Select a program (any *.exe) that is a File Manager, can execute command args for exp. ""C:\YourExplorer.exe C:\""" & Environment.NewLine & Environment.NewLine & "NOTE: If a program you pick doesn't support Command Args or a different program that is not a File Manager, errors may occur when something will be selected.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    End Sub

        Select Case ofd.ShowDialog
            Case DialogResult.OK
                cmb_DFEPathToExplorer.Text = ofd.FileName
        End Select
    End Sub

    Private Sub HotkeyBtnClick(btn As System.Windows.Forms.Button)
        Dim ofd As New OpenFileDialog With {
            .AutoUpgradeEnabled = AppBar.UseExplorerFP,
            .CheckFileExists = True,
            .Title = "Pick a program/file you want to execute after hotkey pressed.",
            .Filter = "Executable (*.exe;*.pif)|*.exe;*.pif|All files (*.*)|*.*",
            .Multiselect = False,
            .SupportMultiDottedExtensions = True,
            .ValidateNames = True}

        If btn.Tag IsNot Nothing Then
            If File.Exists(btn.Tag.ToString) Then
                Dim fi As New FileInfo(btn.Tag.ToString)
                ofd.InitialDirectory = fi.DirectoryName
                ofd.FileName = fi.Name
            End If
        Else
            ofd.FileName = String.Empty
        End If

        'Hotkey setup
        If ofd.ShowDialog(Me) = DialogResult.OK Then

            btn.Tag = ofd.FileName

            Try : btn.BackgroundImage = Icon.ExtractAssociatedIcon(ofd.FileName).ToBitmap : Catch ex As Exception : btn.BackgroundImage = My.Resources.ProgramSmall : End Try

            btn.Text = String.Empty
        End If
    End Sub

    Private Sub HotkeyBtnRClick(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            If MessageBox.Show("Are you sure do you want remove this program?", "Confirm box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then

                Dim btn As Control = CType(sender, Control)

                My.Computer.Registry.SetValue(hotkeyStart, hotkeyPanel.Controls.GetChildIndex(btn), "")

                btn.BackgroundImage = Nothing
                btn.Text = hotkeyPanel.Controls.GetChildIndex(btn)
                btn.Tag = Nothing

            End If
        End If
    End Sub

    Private Sub HotkeyBTN1_Click(sender As Object, e As EventArgs) Handles hotkeyBTN1.Click
        HotkeyBtnClick(hotkeyBTN1)
    End Sub

    Private Sub HotkeyBTN2_Click(sender As Object, e As EventArgs) Handles hotkeyBTN2.Click
        HotkeyBtnClick(hotkeyBTN2)
    End Sub

    Private Sub HotkeyBTN3_Click(sender As Object, e As EventArgs) Handles hotkeyBTN3.Click
        HotkeyBtnClick(hotkeyBTN3)
    End Sub

    Private Sub HotkeyBTN4_Click(sender As Object, e As EventArgs) Handles hotkeyBTN4.Click
        HotkeyBtnClick(hotkeyBTN4)
    End Sub

    Private Sub HotkeyBTN5_Click(sender As Object, e As EventArgs) Handles hotkeyBTN5.Click
        HotkeyBtnClick(hotkeyBTN5)
    End Sub

    Private Sub HotkeyBTN6_Click(sender As Object, e As EventArgs) Handles hotkeyBTN6.Click
        HotkeyBtnClick(hotkeyBTN6)
    End Sub

    Private Sub HotkeyBTN7_Click(sender As Object, e As EventArgs) Handles hotkeyBTN7.Click
        HotkeyBtnClick(hotkeyBTN7)
    End Sub

    Private Sub HotkeyBTN8_Click(sender As Object, e As EventArgs) Handles hotkeyBTN8.Click
        HotkeyBtnClick(hotkeyBTN8)
    End Sub

    Private Sub HotkeyBTN9_Click(sender As Object, e As EventArgs) Handles hotkeyBTN9.Click
        HotkeyBtnClick(hotkeyBTN9)
    End Sub

    Private Sub HotkeyBTN0_Click(sender As Object, e As EventArgs) Handles hotkeyBTN0.Click
        HotkeyBtnClick(hotkeyBTN0)
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ActionCM.Opening
        RestartAndRunAsAdministratorToolStripMenuItem.Image = SystemIcons.Shield.ToBitmap
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles llbl_Actions.LinkClicked
        ActionCM.Show(MousePosition)
    End Sub

    Private Sub EndShellToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EndShellToolStripMenuItem.Click
        If MessageBox.Show("Are you sure do you want end this shell? This means that you will no longer have any Taskbar, Desktop or Startmenu. (To revert it back, please press CTRL+SHIFT+ESC to start a new shell process.)", "Confirm box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            Dim workingAreaReset As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "ResetAfterShutdown", True)
            Dim workspace As New WorkspaceManager

            If workingAreaReset Then workspace.UpdateWorkingArea(0, 0, 0, 0)

            AppBar.CanClose = True
            Desktop.CanClose = True

            Application.Exit()
            End
        End If
    End Sub

    Private Sub RestartShellToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestartShellToolStripMenuItem.Click
        AppBar.CanClose = True
        Desktop.CanClose = True

        Dim cmdPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\cmd.exe"

        If Not File.Exists(cmdPath) Then Application.Restart()

        Dim psi As New ProcessStartInfo(cmdPath) With {
            .CreateNoWindow = True,
            .UseShellExecute = True,
            .ErrorDialog = True,
            .WindowStyle = ProcessWindowStyle.Hidden,
            .Arguments = "/c timeout 1 && start " & Application.ExecutablePath & " && exit"}

        Dim p As Process = Process.Start(psi)

        Task.Delay(20).ContinueWith(Sub() End)
    End Sub

    Private Sub RestartAndRunAsAdministratorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestartAndRunAsAdministratorToolStripMenuItem.Click
        If Not IsUserAnAdmin() Then
            Dim result As DialogResult = MessageBox.Show(
                "Do you want to restart this program with an administrator privilages?",
                "Restart As Administrator",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                Try
                    AppBar.CanClose = True
                    Desktop.CanClose = True

                    Dim cmdPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\cmd.exe"

                    If Not File.Exists(cmdPath) Then Application.Restart()

                    Dim psi As New ProcessStartInfo(cmdPath) With {
                    .CreateNoWindow = True,
                    .UseShellExecute = True,
                    .ErrorDialog = True,
                    .WindowStyle = ProcessWindowStyle.Hidden,
                    .Arguments = "/c timeout 1 && start " & Application.ExecutablePath & " && exit",
                    .Verb = "runas"}

                    Process.Start(psi)

                    Dim p As Process = Process.Start(psi)

                    Task.Delay(20).ContinueWith(Sub() End)
                Catch ex As Exception
                    MessageBox.Show("Failed to gain Admin permissions " & ex.Message)
                End Try
            Else
                ' Shell runs as normal User
            End If
        Else
            If MessageBox.Show("Shell already runs as Administrator, restart anyway?", "Admin permissions already gained", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                AppBar.CanClose = True
                Desktop.CanClose = True

                Dim cmdPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\cmd.exe"

                If Not File.Exists(cmdPath) Then Application.Restart()

                Dim psi As New ProcessStartInfo(cmdPath) With {
                    .CreateNoWindow = True,
                    .UseShellExecute = True,
                    .ErrorDialog = True,
                    .WindowStyle = ProcessWindowStyle.Hidden,
                    .Arguments = "/c timeout 1 && start " & Application.ExecutablePath & " && exit",
                    .Verb = "runas"}

                Process.Start(psi)

                Task.Delay(20).ContinueWith(Sub() Application.Exit())
            End If
        End If
    End Sub

    Private Function IsUserAnAdmin() As Boolean
        Dim identity = WindowsIdentity.GetCurrent()
        Dim principal = New WindowsPrincipal(identity)
        Return principal.IsInRole(WindowsBuiltInRole.Administrator)
    End Function

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles btn_MoreStartProperties.Click
        StartmenuProperties.Show()
    End Sub

    Private Sub CheckBox20_CheckedChanged(sender As Object, e As EventArgs) Handles chk_HKconfig.CheckedChanged
        pnl_HKmain.Enabled = chk_HKconfig.Checked <> True
    End Sub

    Private Sub ComboBox11_TextChanged(sender As Object, e As EventArgs) Handles cmb_HKsearch.TextChanged
        Dim decodedText As String = AppBar.EncodeFilePath(cmb_HKsearch.Text)

        If File.Exists(decodedText) Then
            Dim fi As New FileInfo(decodedText)
            If fi.Extension.ToLower = ".exe" Then My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "SearchApp", cmb_HKsearch.Text) Else Exit Sub
        End If
    End Sub

    Private Sub ComboBox12_TextChanged(sender As Object, e As EventArgs) Handles cmb_HKcopilot.TextChanged
        Dim decodedText As String = AppBar.EncodeFilePath(cmb_HKcopilot.Text)

        If File.Exists(decodedText) Then
            Dim fi As New FileInfo(decodedText)
            If fi.Extension.ToLower = ".exe" Then My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "Assistant", cmb_HKcopilot.Text) Else Exit Sub
        End If
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles btn_HKsearch.Click
        Dim ofd As New OpenFileDialog With {
    .AutoUpgradeEnabled = AppBar.UseExplorerFP,
    .CheckFileExists = True,
    .Title = "Pick a program/file you want to execute after hotkey pressed.",
    .Filter = "Executable (*.exe;*.pif)|*.exe;*.pif|All files (*.*)|*.*",
    .Multiselect = False,
    .SupportMultiDottedExtensions = True,
    .ValidateNames = True
        }

        If File.Exists(cmb_HKsearch.Text) Then
            Dim fi As New FileInfo(cmb_HKsearch.Text)
            ofd.InitialDirectory = fi.DirectoryName
            ofd.FileName = fi.Name
        Else
            ofd.FileName = String.Empty
        End If

        'Hotkey setup
        If ofd.ShowDialog(Me) = DialogResult.OK Then

            cmb_HKsearch.Text = ofd.FileName

        End If
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles btn_HKcopilot.Click
        Dim ofd As New OpenFileDialog With {
.AutoUpgradeEnabled = AppBar.UseExplorerFP,
.CheckFileExists = True,
.Title = "Pick a program/file you want to execute after hotkey pressed.",
.Filter = "Executable (*.exe;*.pif)|*.exe;*.pif|All files (*.*)|*.*",
.Multiselect = False,
.SupportMultiDottedExtensions = True,
.ValidateNames = True}


        If File.Exists(cmb_HKcopilot.Text) Then
            Dim fi As New FileInfo(cmb_HKcopilot.Text)
            ofd.InitialDirectory = fi.DirectoryName
            ofd.FileName = fi.Name
        Else
            ofd.FileName = String.Empty
        End If

        'Hotkey setup
        If ofd.ShowDialog(Me) = DialogResult.OK Then

            cmb_HKcopilot.Text = ofd.FileName

        End If
    End Sub

    Private Sub chk_SysRunDialog_CheckedChanged(sender As Object, e As EventArgs) Handles chk_SysRunDialog.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "UseSystemRunDialog", chk_SysRunDialog.Checked, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub CheckBox1_CheckedChanged_1(sender As Object, e As EventArgs) Handles chk_Balloons.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Windows\Explorer", "EnableLegacyBalloonNotifications", chk_Balloons.Checked, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub chk_pausetowindowscreen_CheckedChanged(sender As Object, e As EventArgs) Handles chk_pausetowindowscreen.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell", "PauseKeyToCaptureWindow", chk_pausetowindowscreen.Checked, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub btn_HKAddPr_Click(sender As Object, e As EventArgs) Handles btn_HKAddPr.Click
        Dim str As String = InputBox("Type a name of your Hotkey preset to set up", "Adding Preset", "")

        If Not String.IsNullOrWhiteSpace(str) Then
            cmb_HKpreferences.Items.Add(str)

            My.Computer.Registry.SetValue(hotkeyStart, "CurrentPreset", str)
            My.Computer.Registry.CurrentUser.CreateSubKey("Software\Shell\CustomPaths\Hotkeys\" & str, True)

            cmb_HKpreferences.SelectedItem = str
        End If
    End Sub

    Private Sub btn_HKRemovePR_Click(sender As Object, e As EventArgs) Handles btn_HKRemovePR.Click
        If MessageBox.Show("Are you sure do you want remove this preset?", "Confirm box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            My.Computer.Registry.CurrentUser.DeleteSubKey("Software\Shell\CustomPaths\Hotkeys\" & cmb_HKpreferences.SelectedItem, False)

            cmb_HKpreferences.Items.RemoveAt(cmb_HKpreferences.SelectedIndex)
        End If
    End Sub

    Private Sub cmb_HKpreferences_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_HKpreferences.SelectedIndexChanged
        If cmb_HKpreferences.Items.Count = 0 Then
            cmb_HKpreferences.Enabled = False
            btn_HKRemovePR.Enabled = False
            Exit Sub
        Else
            cmb_HKpreferences.Enabled = True
            btn_HKRemovePR.Enabled = True
            My.Computer.Registry.SetValue(hotkeyStart, "CurrentPreset", cmb_HKpreferences.SelectedItem)
        End If

        For i As Integer = 0 To 9
            Dim curHotkeyPreset As String = My.Computer.Registry.GetValue(hotkeyStart, "CurrentPreset", String.Empty)
            Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart & "\" & curHotkeyPreset, i, "")

            If File.Exists(executePath) Then
                Try : hotkeyPanel.Controls.Item(i).BackgroundImage = Icon.ExtractAssociatedIcon(executePath).ToBitmap : Catch ex As Exception : hotkeyPanel.Controls.Item(i).BackgroundImage = My.Resources.ProgramSmall : End Try
                hotkeyPanel.Controls.Item(i).Text = String.Empty

                hotkeyPanel.Controls.Item(i).Tag = executePath
            Else
                hotkeyPanel.Controls.Item(i).BackgroundImage = Nothing
                hotkeyPanel.Controls.Item(i).Text = i

                hotkeyPanel.Controls.Item(i).Tag = String.Empty
            End If
        Next
    End Sub

    Private Sub chk_SDback_CheckedChanged(sender As Object, e As EventArgs) Handles chk_SDback.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\ShutdownDialog", "Background", chk_SDback.Checked, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub cmb_CPmain_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_CPmain.SelectedIndexChanged
        Dim value As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", cmb_CPmain.SelectedItem, Nothing)

        If value IsNot Nothing Then
            cmb_CPpath.Text = value
        End If
    End Sub

    Private Sub cmb_CPpath_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_CPpath.SelectedIndexChanged
        If Not String.IsNullOrWhiteSpace(cmb_CPpath.Text) AndAlso File.Exists(AppBar.EncodeFilePath(cmb_CPpath.Text)) Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", cmb_CPmain.SelectedItem, cmb_CPpath.Text, Microsoft.Win32.RegistryValueKind.String)
        End If
    End Sub

    Private Sub btn_CPpath_Click(sender As Object, e As EventArgs) Handles btn_CPpath.Click
        Dim ofd As New OpenFileDialog With {
    .AutoUpgradeEnabled = AppBar.UseExplorerFP,
    .ShowHelp = False,
    .CheckFileExists = True,
    .Title = $"Select a custom path to a default program ({cmb_CPmain.Text})",
    .Filter = "Executables (*.exe)|*.exe"
}
        Dim sortedPath As String = AppBar.EncodeFilePath(cmb_CPpath.Text)

        If File.Exists(sortedPath) Then
            Dim FI As New FileInfo(sortedPath)

            ofd.InitialDirectory = FI.DirectoryName
            ofd.FileName = FI.Name
        ElseIf Directory.Exists(sortedPath) Then
            ofd.InitialDirectory = sortedPath
        Else
            ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
        End If

        Select Case ofd.ShowDialog
            Case DialogResult.OK
                cmb_CPpath.Text = ofd.FileName
        End Select
    End Sub

    Private Sub chk_ATclassicmode_CheckedChanged(sender As Object, e As EventArgs) Handles chk_ATclassicmode.CheckedChanged
        chk_ATAffectDWM.Enabled = Not chk_ATclassicmode.Checked
    End Sub
End Class

