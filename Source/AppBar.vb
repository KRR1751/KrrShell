Imports System.ComponentModel
Imports System.Configuration.Assemblies
Imports System.Data.OleDb
Imports System.Drawing.Drawing2D
Imports System.Globalization
Imports System.IO
Imports System.Management
Imports System.Net.Security
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading
Imports System.Windows
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports KrrShell.AppbarProperties
Imports KrrShell.GlobalKeyboardHook
Imports KrrShell.NetworkControl
Imports Microsoft.VisualBasic.Devices
Imports Microsoft.Win32
Imports NAudio.CoreAudioApi

Public Class AppBar

    Private Const WM_MOUSEACTIVATE As Integer = &H21
    Private Const MA_NOACTIVATE As Integer = 3


    <DllImport("user32.dll")>
    Private Shared Function SetWindowLong(ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    End Function

    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            ' Apply the ToolWindow style before the window is actually created
            cp.ExStyle = cp.ExStyle Or WS_EX_TOOLWINDOW
            Return cp
        End Get
    End Property

    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_SETTINGCHANGE As Integer = &H1A
        Const SPI_SETCOLORS As Integer = &H2

        Const WM_POWERBROADCAST As Integer = &H218
        Const PBT_APMPOWERSTATUSCHANGE As Integer = &HA

        ' When system settings change
        If m.Msg = WM_SETTINGCHANGE Then
            Dim wParamInt As Integer = m.WParam.ToInt32()
            Dim isColorChange As Boolean = (wParamInt = SPI_SETCOLORS)

            If Not isColorChange AndAlso m.LParam <> IntPtr.Zero Then
                Dim lParamStr As String = Marshal.PtrToStringAuto(m.LParam)
                If lParamStr IsNot Nothing AndAlso
               (lParamStr.Contains("Color") OrElse lParamStr.Contains("Theme")) Then
                    isColorChange = True
                End If
            End If

            If isColorChange Then
                Task.Delay(300).ContinueWith(Sub()
                                                 Me.Invoke(Sub()
                                                               UpdateAppbarAccent()
                                                           End Sub)
                                             End Sub)
            End If

        End If

        ' When Windows sends a power update change
        If m.Msg = WM_POWERBROADCAST Then
            If m.WParam.ToInt32() = PBT_APMPOWERSTATUSCHANGE Then
                UpdatePowerUI()
            End If
        End If

        MyBase.WndProc(m)
    End Sub

    Public Sub UpdatePowerUI()
        Dim isPowerVisible As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowPower", CapacityManager.HasBattery)
        ToolStripButton4.Visible = isPowerVisible

        If isPowerVisible = False Then Exit Sub

        Dim currentBatteryPercent As Integer = CapacityManager.GetBatteryPercentage()
        Dim currentIcon As Bitmap = CapacityManager.GetPowerIcon()

        ' Icon
        ToolStripButton4.BackgroundImage = currentIcon

        ' Text
        Dim finalString As String

        If CapacityManager.HasBattery Then
            Dim badBatery As String = ", Bad battery"

            If CapacityManager.IsCharging Then
                finalString = "Charging, " & currentBatteryPercent & "%"
            Else
                finalString = "Not charging, " & currentBatteryPercent & "%"
            End If

            finalString &= If(CapacityManager.IsBatteryBad, badBatery, String.Empty)
        Else
            finalString = "No battery"
        End If

        ToolStripButton4.Text = finalString
    End Sub

    Public currentNet As NetworkControl.WifiNetwork
    Public Sub UpdateNetworkUI()
        Dim isNetworkVisible As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowNetwork", True)
        ToolStripButton5.Visible = isNetworkVisible

        If isNetworkVisible = False Then Exit Sub

        ' Check drivers first
        If Not NetworkControl.IsWifiDriverPresent AndAlso Not NetworkControl.IsEthernetDriverPresent Then
            ToolStripButton5.Text = "Drivers not loaded."

            ToolStripButton5.BackgroundImage = My.Resources.NetworkError
            Exit Sub
        End If

        ' When past, set up icons
        currentNet = NetworkControl.GetActiveNetworkLive()

        If currentNet IsNot Nothing Then
            ' Connected!
            If currentNet.IsLan Then
                ToolStripButton5.BackgroundImage = My.Resources.NetworkLAN
                ToolStripButton5.ToolTipText = "Connected (LAN)"
            Else
                Dim signal As UInteger = currentNet.SignalQuality
                ToolStripButton5.ToolTipText = $"Connected. ({signal}%)"

                If signal > 80 Then
                    ToolStripButton5.BackgroundImage = My.Resources.NetworkMax
                ElseIf signal > 55 Then
                    ToolStripButton5.BackgroundImage = My.Resources.NetworkMid
                ElseIf signal > 30 Then
                    ToolStripButton5.BackgroundImage = My.Resources.NetworkLow
                ElseIf signal > 5 Then
                    ToolStripButton5.BackgroundImage = My.Resources.NetworkNone
                Else
                    ToolStripButton5.BackgroundImage = My.Resources.Network
                End If
            End If
        Else
            If NetworkControl.IsWifiDriverPresent Then
                ToolStripButton5.BackgroundImage = My.Resources.Network
                ToolStripButton5.ToolTipText = "Networks are avalible to connect"
            Else
                ToolStripButton5.BackgroundImage = My.Resources.Network
                ToolStripButton5.ToolTipText = "Not connected"
            End If
        End If
    End Sub

    Private isAutoRefreshing As Boolean = False
    Public Async Sub StartNetworkTimer(ByVal intervalSeconds As Integer)
        If isAutoRefreshing Then Return
        isAutoRefreshing = True

        If intervalSeconds <= 0 Then intervalSeconds = 5

        While isAutoRefreshing
            UpdateNetworkUI()

            Await Task.Delay(intervalSeconds * 1000)

            If Me.IsDisposed OrElse Not Me.IsHandleCreated Then Exit While
        End While
    End Sub

    Public Sub StopNetworkTimer()
        isAutoRefreshing = False
    End Sub

    <DllImport("dwmapi.dll")>
    Public Shared Function DwmExtendFrameIntoClientArea(ByVal hWnd As IntPtr, ByRef pMarinset As Side) As Integer
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure Side
        Public Left As Integer
        Public Right As Integer
        Public Top As Integer
        Public Bottom As Integer
    End Structure

    Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

    'Window Capturing
    <DllImport("user32.dll", SetLastError:=True)>
    Friend Shared Function PrintWindow(hwnd As IntPtr, hDC As IntPtr, nFlags As UInteger) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetClientRect(hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function

    <DllImport("dwmapi.dll")>
    Public Shared Function DwmSetWindowAttribute(hwnd As IntPtr, attr As Integer, ByRef attrValue As Integer, attrSize As Integer) As Integer
    End Function

    <DllImport("dwmapi.dll")>
    Private Shared Function DwmGetWindowAttribute(hwnd As IntPtr, dwAttribute As Integer, ByRef pvAttribute As RECT, cbAttribute As Integer) As Integer
    End Function

    <DllImport("gdi32.dll")>
    Private Shared Function BitBlt(hdcDest As IntPtr, nXDest As Integer, nYDest As Integer, nWidth As Integer, nHeight As Integer, hdcSrc As IntPtr, nXSrc As Integer, nYSrc As Integer, dwRop As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetDCEx(ByVal hWnd As IntPtr, ByVal hrgnClip As IntPtr, ByVal flags As Integer) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindowDC(hWnd As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function ReleaseDC(hWnd As IntPtr, hDC As IntPtr) As Integer
    End Function

    Public Enum DWMWINDOWATTRIBUTE As UInteger
        DWMWA_NCRENDERING_ENABLED = 1      ' [get] Is non-client rendering enabled/disabled
        DWMWA_NCRENDERING_POLICY           ' [set] DWMNCRENDERINGPOLICY - Non-client rendering policy
        DWMWA_TRANSITIONS_FORCEDISABLED    ' [set] Potentially enable/forcibly disable transitions
        DWMWA_ALLOW_NCPAINT                ' [set] Allow contents rendered In the non-client area To be visible On the DWM-drawn frame.
        DWMWA_CAPTION_BUTTON_BOUNDS        ' [get] Bounds Of the caption button area In window-relative space.
        DWMWA_NONCLIENT_RTL_LAYOUT         ' [set] Is non-client content RTL mirrored
        DWMWA_FORCE_ICONIC_REPRESENTATION  ' [set] Force this window To display iconic thumbnails.
        DWMWA_FLIP3D_POLICY                ' [set] Designates how Flip3D will treat the window.
        DWMWA_EXTENDED_FRAME_BOUNDS        ' [get] Gets the extended frame bounds rectangle In screen space
        DWMWA_HAS_ICONIC_BITMAP            ' [set] Indicates an available bitmap When there Is no better thumbnail representation.
        DWMWA_DISALLOW_PEEK                ' [set] Don't invoke Peek on the window.
        DWMWA_EXCLUDED_FROM_PEEK           ' [set] LivePreview exclusion information
        DWMWA_CLOAK                        ' [set] Cloak Or uncloak the window
        DWMWA_CLOAKED                      ' [get] Gets the cloaked state Of the window
        DWMWA_FREEZE_REPRESENTATION        ' [set] BOOL, Force this window To freeze the thumbnail without live update
        DWMWA_LAST
    End Enum

    Public Enum PrintWindowOptions As UInteger
        PW_DEFAULT = 0
        PW_CLIENTONLY = 1
        PW_RENDERFULLCONTENT = 2
    End Enum

    Public Enum DockPosition
        Bottom = 0
        Top = 1
        Left = 2
        Right = 3
    End Enum

    Public Function RenderWindow(hWnd As IntPtr, clientAreaOnly As Boolean, Optional dwmRender As Boolean = False) As Bitmap
        Dim windowRect As New RECT()
        DwmGetWindowAttribute(hWnd, 9, windowRect, Marshal.SizeOf(Of RECT)()) ' 9 = DWMWA_EXTENDED_FRAME_BOUNDS

        If windowRect.Width <= 0 OrElse windowRect.Height <= 0 Then
            Return GeneratePlaceholder(hWnd)
        End If

        Dim bmp As New Bitmap(windowRect.Width, windowRect.Height)

        Using g = Graphics.FromImage(bmp)
            Dim hdcDest = g.GetHdc()

            If dwmRender Then
                Dim hdcSrc As IntPtr = GetDCEx(hWnd, IntPtr.Zero, &H2 Or &H1)
                If hdcSrc <> IntPtr.Zero Then
                    BitBlt(hdcDest, 0, 0, windowRect.Width, windowRect.Height, hdcSrc, 0, 0, &HCC0020) ' SRCCOPY
                    ReleaseDC(hWnd, hdcSrc)
                Else
                    PrintWindow(hWnd, hdcDest, If(clientAreaOnly, &H1, 0))
                End If
            Else
                Dim success As Boolean = PrintWindow(hWnd, hdcDest, &H2)
                If Not success OrElse IsResultBlack(bmp) Then
                    PrintWindow(hWnd, hdcDest, If(clientAreaOnly, &H1, 0))
                End If
            End If

            g.ReleaseHdc(hdcDest)
        End Using

        If IsResultBlack(bmp) Then
            Return GeneratePlaceholder(hWnd)
        End If

        Return bmp
    End Function

    Private Shared Function IsResultBlack(bmp As Bitmap) As Boolean
        Try
            Dim rect As New Rectangle(0, 0, bmp.Width, bmp.Height)
            Dim bmpData As System.Drawing.Imaging.BitmapData = bmp.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, bmp.PixelFormat)

            Dim bytesPerPixel As Integer = Image.GetPixelFormatSize(bmp.PixelFormat) \ 8
            Dim middleX As Integer = bmp.Width \ 2
            Dim middleY As Integer = bmp.Height \ 2
            Dim scan0 As IntPtr = bmpData.Scan0

            Dim offset As Integer = (middleY * bmpData.Stride) + (middleX * bytesPerPixel)

            Dim r As Byte = Marshal.ReadByte(scan0, offset + 2)
            Dim g As Byte = Marshal.ReadByte(scan0, offset + 1)
            Dim b As Byte = Marshal.ReadByte(scan0, offset)
            Dim a As Byte = If(bytesPerPixel = 4, Marshal.ReadByte(scan0, offset + 3), 255)

            bmp.UnlockBits(bmpData)

            Return r = 0 AndAlso g = 0 AndAlso b = 0 AndAlso a = 255
        Catch
            Return True
        End Try
    End Function

    Private Shared Function GeneratePlaceholder(hWnd As IntPtr) As Bitmap
        Dim w As Integer = 178
        Dim h As Integer = 140
        Dim bmp As New Bitmap(w, h)

        Try
            Dim exePath As String = GetProcessPath(hWnd)
            Dim appIcon As Icon = If(Not String.IsNullOrEmpty(exePath), Icon.ExtractAssociatedIcon(exePath), SystemIcons.Application)

            Using g = Graphics.FromImage(bmp)
                g.Clear(GetAccentColor())

                g.SmoothingMode = SmoothingMode.HighQuality
                g.InterpolationMode = InterpolationMode.HighQualityBicubic

                If appIcon IsNot Nothing Then
                    g.DrawIcon(appIcon, New Rectangle((w \ 2) - 16, (h \ 2) - 30, 32, 32))
                End If

                Using font As New Font(SystemFonts.CaptionFont.FontFamily, 9, FontStyle.Regular)
                    Using brush As New SolidBrush(Color.White)
                        Dim sf As New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                        g.DrawString("Minimized", font, brush, New RectangleF(0, h - 40, w, 20), sf)
                    End Using
                End Using
            End Using
        Catch
        End Try

        Return bmp
    End Function

    Private Shared Function GetProcessPath(hWnd As IntPtr) As String
        Try
            Dim pid As Integer
            GetWindowThreadProcessId(hWnd, pid)
            Return Process.GetProcessById(pid).MainModule.FileName
        Catch
            Return ""
        End Try
    End Function

#Region " Process System"

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer,
    ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Public Shared Function DestroyIcon(ByVal hIcon As IntPtr) As Boolean
    End Function


    Public Declare Function EnumWindows Lib "user32" (
    ByVal lpEnumFunc As EnumWindowsProc,
    ByVal lParam As IntPtr) As Boolean

    Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (
    ByVal hwnd As IntPtr) As Integer

    Public Declare Function IsWindowVisible Lib "user32" (
    ByVal hwnd As IntPtr) As Boolean

    Public Declare Function GetAncestor Lib "user32" (
    ByVal hwnd As IntPtr,
    ByVal gaFlags As Integer) As IntPtr

    Public Declare Function GetParent Lib "user32" (
    ByVal hwnd As IntPtr) As IntPtr

    <DllImport("user32.dll")>
    Private Shared Function MonitorFromWindow(hwnd As IntPtr, dwFlags As UInteger) As IntPtr
    End Function

    Public Const GA_ROOTOWNER As Integer = 3
    Public Const GA_ROOT As Integer = 2

    Public Const GWL_EXSTYLE As Integer = -20

    Public Const WS_EX_TOOLWINDOW As Long = &H80L
    Public Const WS_EX_APPWINDOW As Long = &H40000L

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetWindowLong(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Long
    End Function

    <DllImport("uxtheme.dll", ExactSpelling:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function SetWindowTheme(
        ByVal hWnd As IntPtr,
        ByVal pszSubAppName As String,
        ByVal pszSubIdList As String
    ) As Integer
    End Function

    Public Delegate Function EnumWindowsProc(ByVal hwnd As IntPtr, ByVal lParam As IntPtr) As Boolean

    Public Class TaskbarWindowInfo
        Public Property Handle As IntPtr
        Public Property Title As String
        Public Property PID As Integer
    End Class

    Private ReadOnly BlacklistedProcesses As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly BlocklistedProcesses As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    Private Sub LoadBlocklistFromRegistry()
        BlocklistedProcesses.Clear()

        Try
            Dim registryPath As String = "Software\Shell\Appbar\BlockedProcesses\."

            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(registryPath)
                If key IsNot Nothing Then
                    For Each valueName As String In key.GetValueNames()
                        Dim procName As String = key.GetValue(valueName).ToString()
                        If Not String.IsNullOrEmpty(procName) Then
                            BlocklistedProcesses.Add(procName)
                        End If
                    Next
                End If
            End Using
        Catch ex As Exception
            Debug.WriteLine("Error loading BlackList processes from Registry: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadBlacklistFromRegistry()
        BlacklistedProcesses.Clear()

        Try
            Dim registryPath As String = "Software\Shell\Appbar\HiddenProcesses"

            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(registryPath)
                If key IsNot Nothing Then
                    For Each valueName As String In key.GetValueNames()
                        Dim procName As String = key.GetValue(valueName).ToString()
                        If Not String.IsNullOrEmpty(procName) Then
                            BlacklistedProcesses.Add(procName)
                        End If
                    Next
                End If
            End Using
        Catch ex As Exception
            Debug.WriteLine("Error loading BlackList processes from Registry: " & ex.Message)
        End Try
    End Sub

    Private Function EnumWindowCallBack(hWnd As IntPtr, lParam As IntPtr) As Boolean
        If IsTaskbarWindow(hWnd) Then
            AddToolStripButton(hWnd, GetWindowTitleSafe(hWnd))
        End If
        Return True
    End Function
    Private Const MONITOR_DEFAULTTONULL As UInteger = 0
    Private Const MONITOR_DEFAULTTOPRIMARY As UInteger = 1
    Private Sub CleanupDeadWindows()
        For i As Integer = ProcessStrip.Items.Count - 1 To 0 Step -1
            Dim item = ProcessStrip.Items(i)

            If item.Tag IsNot Nothing AndAlso TypeOf item.Tag Is IntPtr Then
                Dim hwnd As IntPtr = CType(item.Tag, IntPtr)

                If Not IsWindow(hwnd) Then
                    StopAlert(hwnd)

                    ProcessStrip.Items.RemoveAt(i)
                    item.Dispose()

                    Debug.WriteLine($"Non-existent Window has been removed: {hwnd}")
                End If
            End If
        Next
    End Sub

    Private Sub RefreshTaskbarForMonitor()
        If Not OnlyShowWindowsOnCurrentMonitor Then Return

        Dim myMonitor As IntPtr = MonitorFromWindow(Me.Handle, MONITOR_DEFAULTTOPRIMARY)

        For i As Integer = ProcessStrip.Items.Count - 1 To 0 Step -1
            Dim item = ProcessStrip.Items(i)
            If item.Tag IsNot Nothing AndAlso TypeOf item.Tag Is IntPtr Then
                Dim targetHwnd As IntPtr = CType(item.Tag, IntPtr)
                Dim targetMonitor As IntPtr = MonitorFromWindow(targetHwnd, MONITOR_DEFAULTTONULL)

                If targetMonitor <> myMonitor Then
                    ProcessStrip.Items.RemoveAt(i)
                    item.Dispose()
                End If
            End If
        Next

        EnumWindows(AddressOf EnumWindowsCallback, IntPtr.Zero)
    End Sub

    Private Sub CheckIfShouldFlash(hwnd As IntPtr)

        If hwnd = GetForegroundWindow() Then
            StopAlert(hwnd)
            Return
        End If

        FlashTaskbarButton(hwnd)
    End Sub
    Private AlertTimers As New Dictionary(Of IntPtr, System.Windows.Forms.Timer)
    Private Sub FlashTaskbarButton(hwnd As IntPtr)
        Dim btn = ProcessStrip.Items.Cast(Of ToolStripItem).FirstOrDefault(Function(x) x.Tag.Equals(hwnd))
        If btn Is Nothing OrElse AlertTimers.ContainsKey(hwnd) Then Return

        Dim t As New System.Windows.Forms.Timer With {
            .Interval = 500
        }

        AddHandler t.Tick, Sub(sender, e)
                               If hwnd = GetForegroundWindow() Then
                                   StopAlert(hwnd)
                                   Return
                               End If

                               If btn.BackColor = Color.Red Then
                                   btn.BackColor = Control.DefaultBackColor
                                   btn.ForeColor = primaryForeColor
                               Else
                                   btn.BackColor = Color.Red
                                   btn.ForeColor = Color.White
                               End If

                               ProcessStrip.Invalidate(btn.Bounds)
                           End Sub

        AlertTimers.Add(hwnd, t)
        t.Start()

        Task.Delay(10000).ContinueWith(Sub()
                                           If Me.IsHandleCreated Then Me.Invoke(Sub() StopAlert(hwnd))
                                       End Sub)
    End Sub

    Private Sub StopAlert(hwnd As IntPtr)
        If AlertTimers.ContainsKey(hwnd) Then
            AlertTimers(hwnd).Stop()
            AlertTimers(hwnd).Dispose()
            AlertTimers.Remove(hwnd)

            Dim btn = ProcessStrip.Items.Cast(Of ToolStripItem).FirstOrDefault(Function(x) x.Tag.Equals(hwnd))
            If btn IsNot Nothing Then
                btn.BackColor = Color.Transparent
                btn.ForeColor = Control.DefaultForeColor
            End If
        End If
    End Sub

    Private Sub UpdateWindowInfo(hwnd As IntPtr)
        Dim btn = ProcessStrip.Items.Cast(Of ToolStripItem).FirstOrDefault(Function(x) x.Tag.Equals(hwnd))

        If btn IsNot Nothing Then
            Dim newTitle As String = GetWindowTitleSafe(hwnd)

            If String.IsNullOrEmpty(newTitle) OrElse Not IsTaskbarWindow(hwnd) Then
                RemoveToolStripButton(hwnd)
                Return
            End If

            btn.Text = newTitle

            Dim newIcon As Icon = GetWindowIcon(hwnd)
            If newIcon IsNot Nothing Then
                If btn.Image IsNot Nothing Then btn.Image.Dispose()
                btn.Image = newIcon.ToBitmap()
            End If
        Else
            If IsTaskbarWindow(hwnd) Then
                AddToolStripButton(hwnd, GetWindowTitleSafe(hwnd))
            End If
        End If
    End Sub

    Private Function IsTaskbarWindow(ByVal hwnd As IntPtr) As Boolean
        If Not IsWindowVisible(hwnd) Then Return False
        If GetWindowTextLength(hwnd) = 0 Then Return False

        Dim windowPid As Integer = 0
        GetWindowThreadProcessId(hwnd, windowPid)
        If windowPid > 0 Then
            Try
                Dim procName As String = Process.GetProcessById(windowPid).ProcessName
                If BlacklistedProcesses.Contains(procName) Then Return False
            Catch
                Return False
            End Try
        End If

        Dim myPid As Integer = Process.GetCurrentProcess().Id

        If GetParent(hwnd) <> IntPtr.Zero Then Return False

        If windowPid = myPid Then
            If hwnd = Me.Handle Then Return False
        End If

        If OnlyShowWindowsOnCurrentMonitor Then
            Dim myMonitor As IntPtr = MonitorFromWindow(Me.Handle, MONITOR_DEFAULTTOPRIMARY)
            Dim targetMonitor As IntPtr = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONULL)
            If targetMonitor <> myMonitor Then Return False
        End If

        Dim exStyle As Long = GetWindowLong(hwnd, GWL_EXSTYLE)
        Dim cloaked As Integer = 0
        DwmGetWindowAttribute(hwnd, DWMWA_CLOAKED, cloaked, Marshal.SizeOf(cloaked))
        If cloaked <> 0 Then Return False

        Dim isToolWindow As Boolean = (exStyle And WS_EX_TOOLWINDOW) <> 0
        Dim isAppWindow As Boolean = (exStyle And WS_EX_APPWINDOW) <> 0
        If isToolWindow AndAlso Not isAppWindow Then Return False

        Dim owner As IntPtr = GetWindow(hwnd, GW_OWNER)

        If windowPid = myPid Then
            Return Not isToolWindow
        End If

        If owner <> IntPtr.Zero AndAlso Not isAppWindow Then Return False

        Return True
    End Function

    Public Function GetWindowTitleSafe(hwnd As IntPtr) As String
        Dim sb As New System.Text.StringBuilder(256)
        InternalGetWindowText(hwnd, sb, sb.Capacity)
        Return sb.ToString()
    End Function

    Public Const WM_GETICON As Integer = &H7F
    Public Const ICON_SMALL As Integer = 0
    Public Const ICON_BIG As Integer = 1
    Public Const GCL_HICON As Integer = -14
    Public Const GCL_HICONSM As Integer = -34

    Private Const GWL_STYLE As Integer = -16

    Private Const DWMWA_CLOAKED As Integer = 14

    Private Const WS_OVERLAPPEDWINDOW As Long = &HCF0000L

    Private Sub DisposeIconCache()
        If ToolStripIconCache IsNot Nothing Then
            For Each img In ToolStripIconCache.Values
                img.Dispose()
            Next
            ToolStripIconCache.Clear()
        End If
    End Sub

    Public Enum IconSize
        Small = 0
        Large = 1
    End Enum

    Public Function GetWindowIcon(ByVal hWnd As IntPtr, Optional ByVal preferredSize As IconSize = IconSize.Large) As Icon
        Dim hIcon As IntPtr = IntPtr.Zero

        Dim msgParam As Integer = If(preferredSize = IconSize.Large, ICON_BIG, ICON_SMALL)
        hIcon = SendMessage(hWnd, WM_GETICON, New IntPtr(msgParam), IntPtr.Zero)

        If hIcon = IntPtr.Zero Then
            Dim altParam As Integer = If(preferredSize = IconSize.Large, ICON_SMALL, ICON_BIG)
            hIcon = SendMessage(hWnd, WM_GETICON, New IntPtr(altParam), IntPtr.Zero)
        End If

        If hIcon = IntPtr.Zero Then
            Dim classParam As Integer = If(preferredSize = IconSize.Large, GCL_HICON, GCL_HICONSM)
            hIcon = GetClassLongPtr(hWnd, classParam)
        End If

        If hIcon = IntPtr.Zero Then
            Dim classAltParam As Integer = If(preferredSize = IconSize.Large, GCL_HICONSM, GCL_HICON)
            hIcon = GetClassLongPtr(hWnd, classAltParam)
        End If

        If hIcon <> IntPtr.Zero Then
            Try
                Return Icon.FromHandle(hIcon).Clone()
            Catch
                Return Nothing
            End Try
        End If

        Return Nothing
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function InternalGetWindowText(
    ByVal hWnd As IntPtr,
    ByVal lpString As System.Text.StringBuilder,
    ByVal nMaxCount As Integer) As Integer
    End Function

    Private Const WM_GETTEXT As UInteger = &HD
    Private Const SMTO_ABORTIFHUNG As UInteger = &H2
    Private Const SMTO_BLOCK As UInteger = &H1

    Private WindowList As New List(Of TaskbarWindowInfo)

    Private Function EnumWindowsCallback(ByVal hwnd As IntPtr, ByVal lParam As IntPtr) As Boolean
        If IsTaskbarWindow(hwnd) Then
            AddToolStripButton(hwnd, GetWindowTitleSafe(hwnd))
        End If
        Return True
    End Function

    Public Function ColorToAbgr(ByVal c As Color) As Integer
        Return (CInt(c.A) << 24) Or (CInt(c.B) << 16) Or (CInt(c.G) << 8) Or c.R
    End Function

    Public Function GetProcessFromWindowHandle(ByVal hWnd As IntPtr) As Process
        Dim processId As Integer = 0

        Dim threadId As Integer = GetWindowThreadProcessId(hWnd, processId)

        If processId > 0 Then
            Try
                Return Process.GetProcessById(processId)
            Catch ex As ArgumentException
                Return Nothing
            End Try
        End If

        Return Nothing
    End Function

    Public Function GetTaskbarApplications() As List(Of TaskbarWindowInfo)
        WindowList.Clear()
        EnumWindows(AddressOf EnumWindowsCallback, IntPtr.Zero)
        Return WindowList
    End Function

    Private ReadOnly ToolStripIconCache As New Dictionary(Of IntPtr, Image)

    Private Function GetProcessIcon(ByVal handle As IntPtr, ByVal defaultImage As Image) As Image
        If ToolStripIconCache.ContainsKey(handle) Then
            Return ToolStripIconCache(handle)
        End If

        Dim windowImage As Image = Nothing
        Dim winIcon As Icon = GetWindowIcon(handle)

        If winIcon IsNot Nothing Then
            windowImage = winIcon.ToBitmap()
            ToolStripIconCache.Add(handle, windowImage)
        Else
            windowImage = defaultImage
        End If

        Return windowImage
    End Function

    Private Function GetMainWindowHandleByProcessId(ByVal processId As String) As IntPtr
        Dim process As Process = Process.GetProcessById(processId)
        If process.Id > 0 Then
            Return process.MainWindowHandle
        Else
            Return IntPtr.Zero
        End If
    End Function

    Public ActiveWindowHandle As IntPtr = IntPtr.Zero

    Private Const EVENT_OBJECT_CREATE As UInt32 = &H8000
    Private Const EVENT_OBJECT_DESTROY As UInt32 = &H8001
    Private Const EVENT_OBJECT_NAMECHANGE As UInt32 = &H800C
    Private Const EVENT_OBJECT_LOCATIONCHANGE As UInteger = &H800B

    Private Const EVENT_SYSTEM_ALERT As UInteger = &H2
    Private Const EVENT_SYSTEM_FOREGROUND As UInteger = &H3

    Private Const WINEVENT_OUTOFCONTEXT As UInt32 = 0

    Private Const GCLP_HICON As Integer = -14

    Public OnlyShowWindowsOnCurrentMonitor As Boolean = False
    Private Const OBJID_WINDOW As Integer = 0

    Public ProcessStripSelectedItem As New ToolStripButton
    Private Sub AddToolStripButton(hwnd As IntPtr, title As String)
        If ProcessStrip.Items.Cast(Of ToolStripItem).Any(Function(x) x.Tag.Equals(hwnd)) Then Return

        Dim btn As New ToolStripButton(title) With {
            .Tag = hwnd,
            .DisplayStyle = ToolStripItemDisplayStyle.Image
        }

        Dim icoSize As IconSize = If(isSmallIcons, IconSize.Small, IconSize.Large)

        If icoSize = IconSize.Large Then
            ProcessStrip.ImageScalingSize = New Size(32, 32)
        Else
            ProcessStrip.ImageScalingSize = New Size(16, 16)
        End If

        Dim windowImage As Image = GetWindowIcon(hwnd, icoSize)?.ToBitmap()
        If windowImage IsNot Nothing Then
            btn.Image = windowImage
        Else
            Dim defProgramIco As Icon = GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\shell32.dll", 2, False)

            If defProgramIco IsNot Nothing Then
                btn.Image = defProgramIco?.ToBitmap
            Else
                btn.Image = My.Resources.ProgramMedium
            End If
        End If

        btn.AutoSize = False
        btn.AutoToolTip = True

        btn.ImageScaling = ToolStripItemImageScaling.SizeToFit

        Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\", "CombineMode", "0")
            Case 0
                btn.DisplayStyle = ToolStripItemDisplayStyle.Image

                Dim finalWidth As Integer = If(isSmallIcons, 24, 48)

                Select Case currentDockPos
                    Case DockPosition.Bottom, DockPosition.Top
                        btn.Size = New Size(finalWidth, Me.Height)
                    Case DockPosition.Left, DockPosition.Right
                        btn.Size = New Size(Me.Width, finalWidth)
                End Select

            Case 1
                btn.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText

                Select Case currentDockPos
                    Case DockPosition.Bottom, DockPosition.Top
                        btn.Size = New Size(150, Me.Height)
                    Case DockPosition.Left, DockPosition.Right
                        btn.Size = New Size(Me.Width, 150)
                End Select

                Dim maxWidth As Integer = 20
                If btn.Text.Length > maxWidth Then
                    btn.Text = btn.Text.Substring(0, maxWidth - 3) & "..."
                End If

                Select Case currentDockPos
                    Case DockPosition.Top, DockPosition.Bottom
                        btn.TextDirection = ToolStripTextDirection.Horizontal

                        btn.TextImageRelation = TextImageRelation.ImageBeforeText
                        btn.TextAlign = ContentAlignment.MiddleLeft
                        btn.ImageAlign = ContentAlignment.MiddleLeft
                    Case DockPosition.Left, DockPosition.Right
                        btn.TextDirection = ToolStripTextDirection.Vertical90

                        btn.TextImageRelation = TextImageRelation.ImageAboveText
                        btn.TextAlign = ContentAlignment.MiddleCenter
                        btn.ImageAlign = ContentAlignment.TopCenter
                End Select

                If Me.BackColor.R >= 150 OrElse Me.BackColor.G >= 130 AndAlso Me.BackColor.B >= 130 Then
                    btn.ForeColor = Color.Black
                Else
                    btn.ForeColor = Color.White
                End If

        End Select

        Dim currentPid As Integer = 0
        GetWindowThreadProcessId(hwnd, currentPid)

        Dim lastIndex As Integer = -1


        For i As Integer = 0 To ProcessStrip.Items.Count - 1
            Dim itemHwnd As IntPtr = CType(ProcessStrip.Items(i).Tag, IntPtr)
            Dim itemPid As Integer = 0
            GetWindowThreadProcessId(itemHwnd, itemPid)

            If itemPid = currentPid Then
                lastIndex = i
            End If
        Next

        AddHandler btn.MouseMove, Sub(sender As Object, e As MouseEventArgs)
                                      ProcessStrip.Focus()
                                      btn.Select()

                                      ProcessStripSelectedItem = btn
                                  End Sub

        AddHandler btn.MouseMove, AddressOf ProcessStrip_MouseMove
        AddHandler btn.MouseUp, AddressOf ProcessStripItemMouseUp

        If lastIndex <> -1 Then
            ProcessStrip.Items.Insert(lastIndex + 1, btn)
        Else
            ProcessStrip.Items.Add(btn)
        End If
    End Sub

    Private Sub RemoveToolStripButton(hwnd As IntPtr)
        Dim toRemove As ToolStripItem = Nothing
        For Each item As ToolStripItem In ProcessStrip.Items
            If item.Tag IsNot Nothing AndAlso item.Tag.Equals(hwnd) Then
                toRemove = item
                Exit For
            End If
        Next

        If toRemove IsNot Nothing Then
            ProcessStrip.Items.Remove(toRemove)
            toRemove.Dispose()
        End If
    End Sub

    Private Sub ProcessStripItemMouseUp(sender As Object, e As MouseEventArgs)
        Dim action As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "MiddleAction", 0)
        Dim item As ToolStripButton = CType(sender, ToolStripButton)

        If item Is Nothing Then Exit Sub
        If TypeOf item.Tag IsNot IntPtr Then Exit Sub
        Dim hWnd As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)

        If e.Button = MouseButtons.Middle Then
            Select Case action
                Case 0 ' Close
                    SendMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero)

                Case 1 ' Duplicate
                    Dim pr As Process = GetProcessFromWindowHandle(hWnd)

                    If pr IsNot Nothing Then
                        Process.Start(New ProcessStartInfo(pr.MainModule.FileName) With {.UseShellExecute = True})
                    End If

                    LoadApps()

                Case 2 ' System menu
                    SendMessage(hWnd, WM_POPUPSYSTEMMENU, IntPtr.Zero, IntPtr.Zero)

            End Select
        End If
    End Sub

    Public Sub MoveProcessStripItem(item As ToolStripButton, orientation As Integer)
        If item Is Nothing OrElse item.Owner Is Nothing Then Exit Sub

        Dim ts As ToolStrip = item.Owner
        Dim currentIndex As Integer = ts.Items.IndexOf(item)
        Dim newIndex As Integer = currentIndex

        Select Case orientation
            Case 0 'Left
                newIndex = currentIndex - 1
            Case 1 'Right
                newIndex = currentIndex + 1
        End Select

        If newIndex >= 0 AndAlso newIndex < ts.Items.Count Then
            ts.Items.Remove(item)
            ts.Items.Insert(newIndex, item)
        End If
    End Sub
#End Region

#Region " AppBar "
    <StructLayout(LayoutKind.Sequential)> Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer

        Public ReadOnly Property Width As Integer
            Get
                Return right - left
            End Get
        End Property
        Public ReadOnly Property Height As Integer
            Get
                Return bottom - top
            End Get
        End Property

        Public Sub New(ByVal X As Integer, ByVal Y As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
            Me.left = X
            Me.top = Y
            Me.right = X2
            Me.bottom = Y2
        End Sub
    End Structure
    <StructLayout(LayoutKind.Sequential)> Structure APPBARDATA
        Public cbSize As Integer
        Public hWnd As IntPtr
        Public uCallbackMessage As Integer
        Public uEdge As Integer
        Public rc As RECT
        Public lParam As IntPtr
    End Structure
    Enum ABMsg
        ABM_NEW = 0
        ABM_REMOVE = 1
        ABM_QUERYPOS = 2
        ABM_SETPOS = 3
        ABM_GETSTATE = 4
        ABM_GETTASKBARPOS = 5
        ABM_ACTIVATE = 6
        ABM_GETAUTOHIDEBAR = 7
        ABM_SETAUTOHIDEBAR = 8
        ABM_WINDOWPOSCHANGED = 9
        ABM_SETSTATE = 10
    End Enum
    Enum ABNotify
        ABN_STATECHANGE = 0
        ABN_POSCHANGED
        ABN_FULLSCREENAPP
        ABN_WINDOWARRANGE
    End Enum
    Enum ABEdge
        ABE_LEFT = 0
        ABE_TOP
        ABE_RIGHT
        ABE_BOTTOM
    End Enum
    Public fBarRegistered As Boolean = False
    <DllImport("SHELL32", CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function SHAppBarMessage(ByVal dwMessage As Integer, ByRef BarrData As APPBARDATA) As Integer
    End Function
    <DllImport("USER32")>
    Public Shared Function GetSystemmetric(ByVal Index As Integer) As Integer
    End Function
    <DllImport("User32.dll", ExactSpelling:=True, CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
    Public Shared Function MoveWindow(ByVal hWnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal cX As Integer, ByVal cY As Integer, ByVal repaint As Boolean) As Boolean
    End Function

    Private uCallBack As Integer

    Public Sub RegisterBar()
        Dim abd As New APPBARDATA
        Dim ret As Integer
        abd.cbSize = Marshal.SizeOf(abd)
        abd.hWnd = Me.Handle
        If fBarRegistered = False Then
            uCallBack = RegisterWindowMessage("AppBarMessage")
            abd.uCallbackMessage = uCallBack

            ret = SHAppBarMessage(ABMsg.ABM_NEW, abd)
            fBarRegistered = True

            ABSetPos()
        Else
            ret = SHAppBarMessage(ABMsg.ABM_REMOVE, abd)

            fBarRegistered = False
        End If
    End Sub

    Private Sub ABSetPos()
        Dim abd As New APPBARDATA
        abd.cbSize = Marshal.SizeOf(abd)
        abd.hWnd = Me.Handle
        abd.uEdge = ABEdge.ABE_BOTTOM

        If abd.uEdge = ABEdge.ABE_LEFT OrElse abd.uEdge = ABEdge.ABE_RIGHT Then
            abd.rc.top = 0
            abd.rc.bottom = SystemInformation.PrimaryMonitorSize.Height
            If abd.uEdge = ABEdge.ABE_LEFT Then
                abd.rc.left = 0
                abd.rc.right = Size.Width
            Else
                abd.rc.right = SystemInformation.PrimaryMonitorSize.Width
                abd.rc.left = abd.rc.right - Size.Width
            End If
        Else
            abd.rc.left = 0
            abd.rc.right = SystemInformation.PrimaryMonitorSize.Width
            If abd.uEdge = ABEdge.ABE_TOP Then
                abd.rc.top = 0
                abd.rc.bottom = Size.Height
            Else
                abd.rc.bottom = SystemInformation.PrimaryMonitorSize.Height
                abd.rc.top = abd.rc.bottom - Size.Height
            End If
        End If

        'Query the system for an approved size and position. 
        SHAppBarMessage(ABMsg.ABM_QUERYPOS, abd)

        'Adjust the rectangle, depending on the edge to which the appbar is anchored. 
        Select Case abd.uEdge
            Case ABEdge.ABE_LEFT
                abd.rc.right = abd.rc.left + Size.Width
            Case ABEdge.ABE_RIGHT
                abd.rc.left = abd.rc.right - Size.Width
            Case ABEdge.ABE_TOP
                abd.rc.top = abd.rc.bottom - Size.Height
            Case ABEdge.ABE_BOTTOM
                abd.rc.bottom = abd.rc.top + Size.Height
        End Select

        'Pass the final bounding rectangle to the system. 
        SHAppBarMessage(ABMsg.ABM_SETPOS, abd)

        'Move and size the appbar so that it conforms to the bounding rectangle passed to the system. 
        MoveWindow(abd.hWnd, abd.rc.left, abd.rc.top, abd.rc.right - abd.rc.left, abd.rc.bottom - abd.rc.top, True)
    End Sub

#End Region

#Region " Notify Tray Icons "
    ' ============================================================
    '  Public events
    ' ============================================================

    Public Event IconAdded As EventHandler(Of TrayIconEventArgs)
    Public Event IconRemoved As EventHandler(Of TrayIconEventArgs)
    Public Event IconLeftClick As EventHandler(Of TrayIconEventArgs)
    Public Event IconRightClick As EventHandler(Of TrayIconEventArgs)
    Public Event BalloonShown As EventHandler(Of BalloonEventArgs)
    Public Event BalloonClosed As EventHandler(Of BalloonEventArgs)
    Public Event BalloonClicked As EventHandler(Of BalloonEventArgs)

    ' ============================================================
    '  Public properties
    ' ============================================================

    <DefaultValue(True)>
    Public Property AutoShowBalloons As Boolean = True

    <DefaultValue(False)>
    Public Property UseSendNotify As Boolean = False

    <DefaultValue(2000)>
    Public Property CleanupInterval As Integer = 2000

    ' ============================================================
    '  P/Invoke
    ' ============================================================

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function PostMessage(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SendNotifyMessage(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As UInteger) As UInteger
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function DestroyWindow(hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function CopyIcon(ByVal hIcon As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function AllowSetForegroundWindow(ByVal dwProcessId As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindow(ByVal hWnd As IntPtr, ByVal uCmd As Integer) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SendNotifyMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", EntryPoint:="GetClassLongW", CharSet:=CharSet.Unicode)>
    Private Shared Function GetClassLong(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As UInteger
    End Function

    <DllImport("user32.dll", EntryPoint:="GetClassLongPtrW", CharSet:=CharSet.Unicode)>
    Private Shared Function GetClassLongPtr64(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
    End Function

    Private Function GetClassLongPtr(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
        If IntPtr.Size = 8 Then
            Return GetClassLongPtr64(hWnd, nIndex)
        Else
            Return New IntPtr(GetClassLong(hWnd, nIndex))
        End If
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function SendMessageTimeout(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr, ByVal flags As Integer, ByVal timeout As Integer, ByRef lpdwResult As IntPtr) As IntPtr
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Private Shared Function OpenProcessToken(ProcessHandle As IntPtr, DesiredAccess As UInteger, ByRef TokenHandle As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Private Shared Function GetTokenInformation(TokenHandle As IntPtr, TokenInformationClass As Integer, TokenInformation As IntPtr, TokenInformationLength As UInteger, ByRef ReturnLength As UInteger) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function CloseHandle(hObject As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function OpenProcess(dwDesiredAccess As UInteger, bInheritHandle As Boolean, dwProcessId As UInteger) As IntPtr
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Private Shared Function IsValidSid(pSid As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Private Shared Function GetSidSubAuthority(pSid As IntPtr, nSubAuthority As UInteger) As IntPtr
    End Function

    ' ============================================================
    '  Constants
    ' ============================================================

    ' Mouse messages
    Private Const WM_MOUSEMOVE As Integer = &H200
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_LBUTTONUP As Integer = &H202
    Private Const WM_LBUTTONDBLCLK As Integer = &H203
    Private Const WM_RBUTTONDOWN As Integer = &H204
    Private Const WM_RBUTTONUP As Integer = &H205
    Private Const WM_RBUTTONDBLCLK As Integer = &H206
    Private Const WM_MBUTTONDOWN As Integer = &H207
    Private Const WM_MBUTTONUP As Integer = &H208
    Private Const WM_MBUTTONDBLCLK As Integer = &H209
    Private Const WM_MOUSEWHEEL As Integer = &H20A
    Private Const WM_CONTEXTMENU As Integer = &H7B

    Private Const MK_LBUTTON As UInteger = &H1
    Private Const MK_RBUTTON As UInteger = &H2
    Private Const MK_MBUTTON As UInteger = &H10

    Private Const NIM_ADD As Integer = 0
    Private Const NIM_MODIFY As Integer = 1
    Private Const NIM_DELETE As Integer = 2

    ' NOTIFYICONDATA uFlags
    Private Const NIF_MESSAGE As Integer = &H1
    Private Const NIF_ICON As Integer = &H2
    Private Const NIF_TIP As Integer = &H4
    Private Const NIF_STATE As Integer = &H8
    Private Const NIF_INFO As Integer = &H10
    Private Const NIF_GUID As Integer = &H20
    Private Const NIF_REALTIME As Integer = &H40
    Private Const NIF_SHOWTIP As Integer = &H80

    ' dwInfoFlags
    Private Const NIIF_NONE As Integer = &H0
    Private Const NIIF_INFO As Integer = &H1
    Private Const NIIF_WARNING As Integer = &H2
    Private Const NIIF_ERROR As Integer = &H3
    Private Const NIIF_USER As Integer = &H4
    Private Const NIIF_NOSOUND As Integer = &H10
    Private Const NIIF_LARGE_ICON As Integer = &H20
    Private Const NIIF_RESPECT_QUIET_TIME As Integer = &H80
    Private Const NIIF_ICON_MASK As Integer = &HF

    ' Tray callback notifications
    Private Const NIN_SELECT As Integer = &H401
    Private Const NIN_BALLOONSHOW As Integer = &H402
    Private Const NIN_BALLOONHIDE As Integer = &H403
    Private Const NIN_BALLOONTIMEOUT As Integer = &H404
    Private Const NIN_BALLOONUSERCLICK As Integer = &H405
    Private Const NIN_POPUPOPEN As Integer = &H406
    Private Const NIN_POPUPCLOSE As Integer = &H407

    Private Const WM_COPYDATA As Integer = &H4A
    Private Const WS_POPUP As Integer = &H80000000
    Private HWND_BROADCAST As IntPtr = New IntPtr(&HFFFF)

    Private Const GCLP_HICONSM As Integer = -34
    Private Const GW_OWNER As Integer = 4
    Private Const TRAY_COPYDATA_SIGNATURE As Integer = &H34753423
    Private Const TokenIntegrityLevel As Integer = 25

    ' ============================================================
    '  Structures
    ' ============================================================

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Pack:=1)>
    Public Structure NOTIFYICONDATA32
        Public cbSize As Integer
        Public hWnd As Integer
        Public uID As Integer
        Public uFlags As Integer
        Public uCallbackMessage As Integer
        Public hIcon As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)>
        Public szTip As String
        Public dwState As Integer
        Public dwStateMask As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)>
        Public szInfo As String
        Public uVersion As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=64)>
        Public szInfoTitle As String
        Public dwInfoFlags As Integer
        Public guidItem As Guid
        Public hBalloonIcon As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure NOTIFYICONDATA64
        Public cbSize As Integer
        Public hWnd As IntPtr
        Public uID As Integer
        Public uFlags As Integer
        Public uCallbackMessage As Integer
        Public hIcon As IntPtr
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)>
        Public szTip As String
        Public dwState As Integer
        Public dwStateMask As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)>
        Public szInfo As String
        Public uVersion As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=64)>
        Public szInfoTitle As String
        Public dwInfoFlags As Integer
        Public guidItem As Guid
        Public hBalloonIcon As IntPtr
    End Structure

    Public Class TrayData
        Public Property HWnd As IntPtr
        Public Property UID As Integer
        Public Property UFlags As Integer
        Public Property CallbackMessage As Integer
        Public Property HIcon As IntPtr
        Public Property SzTip As String
        Public Property SenderIs64Bit As Boolean
        Public Property SzInfo As String
        Public Property SzInfoTitle As String
        Public Property DwInfoFlags As Integer
        Public Property HBalloonIcon As IntPtr
    End Class

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Private Structure WNDCLASSEX
        Public cbSize As Integer
        Public style As Integer
        Public lpfnWndProc As WndProcDelegate
        Public cbClsExtra As Integer
        Public cbWndExtra As Integer
        Public hInstance As IntPtr
        Public hIcon As IntPtr
        Public hCursor As IntPtr
        Public hbrBackground As IntPtr
        Public lpszMenuName As String
        Public lpszClassName As String
        Public hIconSm As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure COPYDATASTRUCT
        Public dwData As IntPtr
        Public cbData As Integer
        Public lpData As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure TOKEN_MANDATORY_LABEL
        Public Sid As IntPtr
        Public Attributes As UInteger
    End Structure

    Private Enum PopupBlacklist
        SndVol
    End Enum

    Private Delegate Function WndProcDelegate(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    Private Shared ReadOnly WM_TASKBARCREATED As Integer = RegisterWindowMessage("TaskbarCreated")

    Private ReadOnly trayIcons As New Dictionary(Of String, TrayIconInfo)()
    Private ReadOnly trayButtons As New Dictionary(Of String, ToolStripButton)()
    Private currentBalloonOwner As TrayIconInfo
    Private _balloonClickedByUser As Boolean   ' distinguishes user-click close from timeout close
    Private _balloonSequence As Integer = 0
    Private ReadOnly _uiLock As New Object()

    Private _myWndProcDelegate As WndProcDelegate
    Private _fakeTrayHandle As IntPtr
    Private _hInstance As IntPtr
    Private _myIntegrityLevel As UInteger

    Private currentPopup As BalloonPopup
    Private _balloonCloseTimer As System.Windows.Forms.Timer

    ' Konstanta pro velikost ikony
    Private Const TRAY_ICON_SIZE As Integer = 24

    Public Sub InitTrayIcons()
        _hInstance = Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly().GetModules()(0))
        _myIntegrityLevel = GetProcessIntegrityLevel()

        CreateTrayWindow()
        PostMessage(HWND_BROADCAST, CUInt(WM_TASKBARCREATED), IntPtr.Zero, IntPtr.Zero)
    End Sub

    Private Sub TrayCleanup()
        SyncLock _uiLock
            Dim dead As New List(Of String)()
            For Each kvp As KeyValuePair(Of String, TrayIconInfo) In trayIcons
                If Not IsWindow(kvp.Value.HWnd) Then dead.Add(kvp.Key)
            Next
            For Each k As String In dead
                Dim info As TrayIconInfo = trayIcons(k)
                trayIcons.Remove(k)
                Dim btn As ToolStripButton = Nothing
                If trayButtons.TryGetValue(k, btn) Then
                    ToolStrip1.Items.Remove(btn)
                    btn.Dispose()
                    trayButtons.Remove(k)
                End If
                info.Icon?.Dispose()
                RaiseEvent IconRemoved(Me, New TrayIconEventArgs(info))
            Next
        End SyncLock
    End Sub

    ' ============================================================
    '  Tray message parsing
    ' ============================================================

    Private Sub ProcessTrayMessage(lParam As IntPtr)
        Try
            Dim cds As COPYDATASTRUCT = Marshal.PtrToStructure(Of COPYDATASTRUCT)(lParam)
            If cds.lpData = IntPtr.Zero OrElse cds.cbData < 12 Then Return

            Dim dwSignature As Integer = Marshal.ReadInt32(cds.lpData, 0)
            If dwSignature <> TRAY_COPYDATA_SIGNATURE Then Return

            Dim dwCommand As Integer = Marshal.ReadInt32(cds.lpData, 4)
            Dim dataPtr As IntPtr = IntPtr.Add(cds.lpData, 8)

            Dim cbSize As Integer = Marshal.ReadInt32(dataPtr, 0)
            Dim data As TrayData = Nothing

            If cbSize = 956 OrElse (cbSize > 0 AndAlso cbSize < 960) Then
                Dim nid As NOTIFYICONDATA32 = Marshal.PtrToStructure(Of NOTIFYICONDATA32)(dataPtr)
                data = New TrayData() With {
                    .HWnd = New IntPtr(nid.hWnd),
                    .UID = nid.uID,
                    .UFlags = nid.uFlags,
                    .CallbackMessage = nid.uCallbackMessage,
                    .HIcon = New IntPtr(nid.hIcon),
                    .SzTip = nid.szTip,
                    .SenderIs64Bit = False,
                    .SzInfo = nid.szInfo,
                    .SzInfoTitle = nid.szInfoTitle,
                    .DwInfoFlags = nid.dwInfoFlags,
                    .HBalloonIcon = New IntPtr(nid.hBalloonIcon)
                }
            ElseIf cbSize = 976 OrElse cbSize >= 960 Then
                Dim nid As NOTIFYICONDATA64 = Marshal.PtrToStructure(Of NOTIFYICONDATA64)(dataPtr)
                data = New TrayData() With {
                    .HWnd = nid.hWnd,
                    .UID = nid.uID,
                    .UFlags = nid.uFlags,
                    .CallbackMessage = nid.uCallbackMessage,
                    .HIcon = nid.hIcon,
                    .SzTip = nid.szTip,
                    .SenderIs64Bit = True,
                    .SzInfo = nid.szInfo,
                    .SzInfoTitle = nid.szInfoTitle,
                    .DwInfoFlags = nid.dwInfoFlags,
                    .HBalloonIcon = nid.hBalloonIcon
                }
            Else
                Debug.WriteLine("Unknown NOTIFYICONDATA cbSize=" & cbSize)
                Return
            End If

            Select Case dwCommand
                Case NIM_DELETE
                    RemoveIcon(data.HWnd, CUInt(data.UID), fromApp:=True)
                Case NIM_ADD, NIM_MODIFY
                    UpdateIcon(data, dwCommand)
            End Select
        Catch ex As Exception
            Debug.WriteLine("Tray parse error: " & ex.Message)
        End Try
    End Sub

    ' ============================================================
    '  Icon add / modify / remove
    ' ============================================================

    Private Sub UpdateIcon(nid As TrayData, command As Integer)
        SyncLock _uiLock
            Dim key As String = MakeKey(nid.HWnd, CUInt(nid.UID))

            Dim ownerPid As UInteger = 0
            GetWindowThreadProcessId(nid.HWnd, ownerPid)
            Dim procName As String = "<unknown>"
            Try
                If ownerPid <> 0 Then
                    Using p As Process = Process.GetProcessById(CInt(ownerPid))
                        procName = p.ProcessName
                    End Using
                End If
            Catch
            End Try

            ' Respect uFlags: only update fields whose NIF_* bit is set.
            Dim isAdd As Boolean = (command = NIM_ADD) OrElse Not trayIcons.ContainsKey(key)
            Dim updateMessage As Boolean = isAdd OrElse (nid.UFlags And NIF_MESSAGE) <> 0
            Dim updateIcon As Boolean = isAdd OrElse (nid.UFlags And NIF_ICON) <> 0
            Dim updateTip As Boolean = isAdd OrElse (nid.UFlags And NIF_TIP) <> 0
            Dim updateInfo As Boolean = isAdd OrElse (nid.UFlags And NIF_INFO) <> 0

            Dim iconBmp As Bitmap = Nothing
            If updateIcon Then
                If nid.HIcon <> IntPtr.Zero Then iconBmp = BitmapFromHIcon(nid.HIcon)
                If iconBmp Is Nothing AndAlso nid.HWnd <> IntPtr.Zero Then iconBmp = GetIconFromWindow(nid.HWnd)
                If iconBmp Is Nothing AndAlso ownerPid <> 0 Then iconBmp = GetIconFromProcess(ownerPid)
            End If

            Dim info As TrayIconInfo
            Dim wasNew As Boolean = False
            If trayIcons.ContainsKey(key) Then
                info = trayIcons(key)
                If updateMessage Then info.CallbackMessage = CUInt(nid.CallbackMessage)
                If updateTip Then info.Tooltip = If(String.IsNullOrEmpty(nid.SzTip), procName, nid.SzTip)
                info.ProcessId = ownerPid
                info.ProcessName = procName
                info.SenderIs64Bit = nid.SenderIs64Bit
                If updateIcon AndAlso iconBmp IsNot Nothing Then
                    info.Icon?.Dispose()
                    info.Icon = iconBmp
                End If
            Else
                wasNew = True
                info = New TrayIconInfo() With {
                    .HWnd = nid.HWnd,
                    .UID = CUInt(nid.UID),
                    .CallbackMessage = CUInt(nid.CallbackMessage),
                    .Tooltip = If(String.IsNullOrEmpty(nid.SzTip), procName, nid.SzTip),
                    .ProcessId = ownerPid,
                    .ProcessName = procName,
                    .Icon = iconBmp,
                    .SenderIs64Bit = nid.SenderIs64Bit
                }
                trayIcons.Add(key, info)
            End If

            ' Create / refresh the ToolStrip button.
            CreateOrUpdateButton(key, info, updateIcon OrElse updateTip OrElse wasNew)

            ' Balloon tooltip detection - only when NIF_INFO is set.
            If updateInfo Then
                Dim balloonText As String = If(nid.SzInfo, "")
                Dim balloonTitle As String = If(nid.SzInfoTitle, "")
                Dim isHide As Boolean = String.IsNullOrEmpty(balloonText)

                _balloonSequence += 1
                info.LastBalloonTitle = balloonTitle
                info.LastBalloonText = balloonText
                info.LastBalloonFlags = nid.DwInfoFlags
                info.LastBalloonTime = DateTime.Now

                If Not isHide Then
                    ShowBalloonForIcon(info)
                Else
                    ' If this icon was the one showing a balloon, dismiss ours too.
                    If currentBalloonOwner Is info Then
                        currentBalloonOwner = Nothing
                    End If
                End If
            End If

            If wasNew Then RaiseEvent IconAdded(Me, New TrayIconEventArgs(info))
        End SyncLock
    End Sub

    Private Sub RemoveIcon(hWnd As IntPtr, uID As UInteger, fromApp As Boolean)
        SyncLock _uiLock
            Dim key As String = MakeKey(hWnd, uID)
            If Not trayIcons.ContainsKey(key) Then Return
            Dim info As TrayIconInfo = trayIcons(key)
            trayIcons.Remove(key)
            Dim btn As ToolStripButton = Nothing
            If trayButtons.TryGetValue(key, btn) Then
                ToolStrip1.Items.Remove(btn)
                btn.Dispose()
                trayButtons.Remove(key)
            End If
            info.Icon?.Dispose()
            RaiseEvent IconRemoved(Me, New TrayIconEventArgs(info))
        End SyncLock
    End Sub


    ' ============================================================
    '  ToolStrip button management
    ' ============================================================

    Private Sub CreateOrUpdateButton(key As String, info As TrayIconInfo, refreshVisual As Boolean)
        Dim btn As ToolStripButton = Nothing
        If trayButtons.TryGetValue(key, btn) Then
            ' Existing button - refresh
            If refreshVisual Then
                If info.Icon IsNot Nothing Then
                    ' *** IMMEDIATE: 16x16 info.Icon ***
                    btn.BackgroundImage = New Bitmap(info.Icon, New Size(TRAY_ICON_SIZE, TRAY_ICON_SIZE))
                    btn.BackgroundImageLayout = ImageLayout.Zoom
                    btn.DisplayStyle = ToolStripItemDisplayStyle.Image
                    btn.Text = ""
                Else
                    btn.BackgroundImage = Nothing
                    btn.DisplayStyle = ToolStripItemDisplayStyle.Text
                    btn.Text = "?"
                End If
            End If
            btn.ToolTipText = BuildButtonTooltip(info)
        Else
            ' New button
            btn = New ToolStripButton()
            btn.AutoSize = False
            btn.BackgroundImageLayout = ImageLayout.Zoom

            Select Case currentDockPos
                Case DockPosition.Top, DockPosition.Bottom
                    btn.Size = New Size(TRAY_ICON_SIZE, ToolStrip1.Height)
                Case DockPosition.Left, DockPosition.Right
                    btn.Size = New Size(ToolStrip1.Width, TRAY_ICON_SIZE)
            End Select

            If info.Icon IsNot Nothing Then
                ' *** IMMEDIATE: 16x16 info.Icon ***
                btn.BackgroundImage = New Bitmap(info.Icon, New Size(TRAY_ICON_SIZE, TRAY_ICON_SIZE))
                btn.DisplayStyle = ToolStripItemDisplayStyle.Image
            Else
                btn.Text = "?"
                btn.DisplayStyle = ToolStripItemDisplayStyle.Text
            End If

            btn.ToolTipText = BuildButtonTooltip(info)
            btn.Tag = info

            AddHandler btn.MouseDown, AddressOf OnButtonMouseDown
            AddHandler btn.MouseUp, AddressOf OnButtonMouseUp
            AddHandler btn.MouseMove, AddressOf OnButtonMouseMove
            AddHandler btn.DoubleClick, AddressOf OnButtonDoubleClick
            AddHandler btn.Click, AddressOf OnButtonClick

            trayButtons.Add(key, btn)
            ToolStrip1.Items.Add(btn)
        End If
    End Sub

    Private Function BuildButtonTooltip(info As TrayIconInfo) As String
        Dim s As String = info.Tooltip
        Return s
    End Function

    ' ----- Mouse event handlers - the heart of the "real ToolStrip" behavior -----

    Private Sub OnButtonMouseDown(sender As Object, e As MouseEventArgs)
        Dim info As TrayIconInfo = TryCast(DirectCast(sender, ToolStripButton).Tag, TrayIconInfo)
        If info Is Nothing Then Return
        Select Case e.Button
            Case MouseButtons.Left : SendMouse(info, WM_LBUTTONDOWN, MK_LBUTTON)
            Case MouseButtons.Right : SendMouse(info, WM_RBUTTONDOWN, MK_RBUTTON)
            Case MouseButtons.Middle : SendMouse(info, WM_MBUTTONDOWN, MK_MBUTTON)
        End Select
    End Sub

    Private Sub OnButtonMouseUp(sender As Object, e As MouseEventArgs)
        Dim btn As ToolStripButton = DirectCast(sender, ToolStripButton)
        Dim info As TrayIconInfo = TryCast(btn.Tag, TrayIconInfo)
        If info Is Nothing Then Return
        Select Case e.Button
            Case MouseButtons.Left
                SendMouse(info, WM_LBUTTONUP, 0)
                ' NIN_SELECT is sent by Windows AFTER WM_LBUTTONUP for tray clicks.
                SendTrayCallback(info, NIN_SELECT)
                RaiseEvent IconLeftClick(Me, New TrayIconEventArgs(info))
            Case MouseButtons.Right
                SendMouse(info, WM_RBUTTONUP, 0)
                ' Windows sends NIN_POPUPOPEN after right-click, then WM_CONTEXTMENU.
                SendTrayCallback(info, NIN_POPUPOPEN)
                ' Screen coords of the cursor at click time -> WM_CONTEXTMENU's lParam.
                Dim pt As Point = Cursor.Position
                Dim wParam As IntPtr = info.HWnd
                Dim lParam As IntPtr = New IntPtr(MAKELONG(pt.X, pt.Y))
                DoSend(info.HWnd, CUInt(WM_CONTEXTMENU), wParam, lParam)
                RaiseEvent IconRightClick(Me, New TrayIconEventArgs(info))
            Case MouseButtons.Middle
                SendMouse(info, WM_MBUTTONUP, 0)
        End Select
    End Sub

    Private Sub OnButtonMouseMove(sender As Object, e As MouseEventArgs)
        Dim info As TrayIconInfo = TryCast(DirectCast(sender, ToolStripButton).Tag, TrayIconInfo)
        If info Is Nothing Then Return
        ' Forward as WM_MOUSEMOVE so apps that track hover (e.g. to show a
        ' custom tooltip) get the signal. We pass cursor screen coords.
        Dim pt As Point = Cursor.Position
        Dim wParam As IntPtr = New IntPtr(0)
        Dim lParam As IntPtr = New IntPtr(MAKELONG(pt.X, pt.Y))
        DoSend(info.HWnd, CUInt(WM_MOUSEMOVE), wParam, lParam)
    End Sub

    Private Sub OnButtonClick(sender As Object, e As EventArgs)
        'Anything goes automatically
    End Sub

    Private Sub OnButtonDoubleClick(sender As Object, e As EventArgs)
        Dim info As TrayIconInfo = TryCast(DirectCast(sender, ToolStripButton).Tag, TrayIconInfo)
        If info Is Nothing Then Return
        ' Full double-click sequence: DOWN -> UP -> DBLCLK -> UP -> NIN_SELECT
        SendMouse(info, WM_LBUTTONDOWN, MK_LBUTTON)
        SendMouse(info, WM_LBUTTONUP, 0)
        SendMouse(info, WM_LBUTTONDBLCLK, MK_LBUTTON)
        SendMouse(info, WM_LBUTTONUP, 0)
        SendTrayCallback(info, NIN_SELECT)
    End Sub

    ' ============================================================
    '  Balloon display via NotifyIcon proxy
    ' ============================================================

    Private Sub ShowBalloonForIcon(info As TrayIconInfo)
        If Not AutoShowBalloons Then Return
        If info Is Nothing Then Return

        If currentPopup IsNot Nothing Then
            CloseBalloon(sendTimeout:=True)
        End If

        ' Check if Disabled
        Dim EnabledValue As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Windows\Explorer", "EnableLegacyBalloonNotifications", True)
        If Not Enabled Then Return

        If currentPopup IsNot Nothing AndAlso currentPopup.Title = "<unknown>" Then
            Return
        End If

        ' Check Blacklist
        For Each str As String In [Enum].GetNames(GetType(PopupBlacklist))
            Dim bpr As Process = GetProcessFromWindowHandle(info.HWnd)
            If bpr IsNot Nothing AndAlso bpr.ProcessName = str Then
                Return
            End If
        Next

        currentBalloonOwner = info
        _balloonClickedByUser = False

        ' ===================================================================
        '  ICON LOGIC:
        ' ===================================================================
        Dim iconType As Integer = info.LastBalloonFlags And NIIF_ICON_MASK
        Dim useSystemIcon As Boolean = (iconType = NIIF_INFO) OrElse
                                    (iconType = NIIF_WARNING) OrElse
                                    (iconType = NIIF_ERROR)

        Dim popup As New BalloonPopup() With {
        .Title = If(String.IsNullOrEmpty(info.LastBalloonTitle),
                    info.ProcessName,
                    info.LastBalloonTitle),
        .Body = If(info.LastBalloonText, ""),
        .ToolTipIcon = If(useSystemIcon,
                          BalloonFlagsToToolTipIcon(info.LastBalloonFlags),
                          ToolTipIcon.None),
        .Duration = 10000,
        .IconImage = If(useSystemIcon, Nothing, info.Icon)
    }

        ' ===================================================================
        '  Location
        ' ===================================================================
        popup.PopupPosition = ComputeBalloonPosition()

        AddHandler popup.BalloonClicked, Sub(s, e)
                                             _balloonClickedByUser = True
                                             SendTrayCallback(currentBalloonOwner, NIN_BALLOONUSERCLICK)
                                             RaiseEvent BalloonClicked(Me, New BalloonEventArgs(currentBalloonOwner))
                                         End Sub

        AddHandler popup.BalloonClosed, Sub(s, e)
                                            Dim reason As String = DirectCast(s, BalloonPopup).CloseReasonValue
                                            If reason = "timeout" AndAlso Not _balloonClickedByUser Then
                                                SendTrayCallback(currentBalloonOwner, NIN_BALLOONTIMEOUT)
                                            End If
                                            RaiseEvent BalloonClosed(Me, New BalloonEventArgs(currentBalloonOwner))
                                            currentBalloonOwner = Nothing
                                            currentPopup = Nothing
                                            _balloonClickedByUser = False
                                        End Sub

        currentPopup = popup

        SendTrayCallback(info, NIN_BALLOONSHOW)
        RaiseEvent BalloonShown(Me, New BalloonEventArgs(currentBalloonOwner))

        popup.Show()
    End Sub

    Private Function ComputeBalloonPosition() As Point
        Const POPUP_W As Integer = 280
        Const POPUP_H As Integer = 80
        Const MARGIN As Integer = 10

        Dim workArea As Rectangle = Screen.FromControl(ToolStrip1).WorkingArea
        Dim tsH As Integer = ToolStrip1.Height
        Dim tsW As Integer = ToolStrip1.Width

        Select Case currentDockPos
            Case DockPosition.Bottom
                Return New Point(SystemInformation.WorkingArea.Width - POPUP_W, SystemInformation.WorkingArea.Height - POPUP_H)
                'New Point(workArea.Right - POPUP_W - MARGIN, workArea.Bottom - POPUP_H - tsH - MARGIN)

            Case DockPosition.Right
                Return New Point(SystemInformation.WorkingArea.Width - POPUP_W, SystemInformation.WorkingArea.Height - POPUP_H)
                'New Point(workArea.Right - POPUP_W - tsW - MARGIN, workArea.Top + MARGIN)

            Case DockPosition.Left
                Return New Point(Me.Width, SystemInformation.WorkingArea.Height - POPUP_H)
                'New Point(workArea.Left + tsW + MARGIN, workArea.Top + MARGIN)

            Case DockPosition.Top
                Return New Point(SystemInformation.WorkingArea.Width - POPUP_W, Me.Height)
                'New Point(workArea.Right - POPUP_W - MARGIN, workArea.Top + tsH + MARGIN)

            Case Else
                Return New Point(workArea.Right - POPUP_W - MARGIN, workArea.Bottom - POPUP_H - MARGIN)
        End Select
    End Function

    Private Sub CloseBalloon(sendTimeout As Boolean)
        If currentPopup Is Nothing Then Return

        If sendTimeout AndAlso currentBalloonOwner IsNot Nothing AndAlso Not _balloonClickedByUser Then
            SendTrayCallback(currentBalloonOwner, NIN_BALLOONTIMEOUT)
        End If

        RemoveHandler currentPopup.BalloonClicked, Nothing   ' (lambda handler se odpojí s popup.Dispose)
        currentPopup.Close()
        currentPopup = Nothing
        currentBalloonOwner = Nothing
        _balloonClickedByUser = False
    End Sub

    Private Sub OnBalloonProxyShown(sender As Object, e As EventArgs)
        If currentBalloonOwner Is Nothing Then Return
        SendTrayCallback(currentBalloonOwner, NIN_BALLOONSHOW)
        RaiseEvent BalloonShown(Me, New BalloonEventArgs(currentBalloonOwner))
    End Sub

    Private Sub OnBalloonProxyClicked(sender As Object, e As EventArgs)
        If currentBalloonOwner Is Nothing Then Return
        _balloonClickedByUser = True
        SendTrayCallback(currentBalloonOwner, NIN_BALLOONUSERCLICK)
        RaiseEvent BalloonClicked(Me, New BalloonEventArgs(currentBalloonOwner))
    End Sub

    Private Sub OnBalloonProxyClosed(sender As Object, e As EventArgs)
        If currentBalloonOwner Is Nothing Then Return
        ' .NET BalloonTipClosed fires for both timeout and user-click. If the
        ' user clicked, OnBalloonProxyClicked already sent NIN_BALLOONUSERCLICK
        ' (which the app will treat as the close reason). We only send
        ' NIN_BALLOONTIMEOUT when the user did NOT click.
        If Not _balloonClickedByUser Then
            SendTrayCallback(currentBalloonOwner, NIN_BALLOONTIMEOUT)
        End If
        RaiseEvent BalloonClosed(Me, New BalloonEventArgs(currentBalloonOwner))
        ' Hide the phantom tray icon until the next balloon.
        currentBalloonOwner = Nothing
        _balloonClickedByUser = False
    End Sub

    ' ============================================================
    '  Send helpers
    ' ============================================================

    Private Function DoSend(hWnd As IntPtr, msg As UInteger, wParam As IntPtr, lParam As IntPtr) As Boolean
        If UseSendNotify Then
            Return SendNotifyMessage(hWnd, msg, wParam, lParam)
        Else
            Return PostMessage(hWnd, msg, wParam, lParam)
        End If
    End Function

    ' Send a raw WM_* mouse message in Callback mode (wParam=uID, lParam=msg)
    ' or Direct mode (wParam=mk, lParam=coords). For TrayShell we always use
    ' Callback mode because that's the only mode the source app's tray
    ' handler expects messages in.
    Private Sub SendMouse(info As TrayIconInfo, mouseMsg As Integer, mkKeys As UInteger)
        If info Is Nothing Then Return
        If CheckUIPIBlocked(info) Then Return
        If info.CallbackMessage = 0 Then
            Return
        End If
        Dim wParam As IntPtr = New IntPtr(info.UID)
        Dim lParam As IntPtr = New IntPtr(mouseMsg)
        Dim ok As Boolean = DoSend(info.HWnd, info.CallbackMessage, wParam, lParam)
    End Sub

    Private Sub SendTrayCallback(info As TrayIconInfo, code As Integer)
        If info Is Nothing Then Return
        If CheckUIPIBlocked(info) Then Return
        If info.CallbackMessage = 0 Then
            Return
        End If
        Dim wParam As IntPtr = New IntPtr(info.UID)
        Dim lParam As IntPtr = New IntPtr(code)
        Dim ok As Boolean = DoSend(info.HWnd, info.CallbackMessage, wParam, lParam)
    End Sub

    Private Function CheckUIPIBlocked(info As TrayIconInfo) As Boolean
        Dim targetIL As UInteger = GetProcessIntegrityLevel(info.ProcessId)
        If targetIL > _myIntegrityLevel Then
            Return True
        End If
        Return False
    End Function

    ' ============================================================
    '  Shell_TrayWnd hijack
    ' ============================================================

    Public Sub CreateTrayWindow()
        _myWndProcDelegate = AddressOf FakeTrayWndProc

        Dim wc As New WNDCLASSEX()
        wc.cbSize = Marshal.SizeOf(wc)
        wc.lpfnWndProc = _myWndProcDelegate
        wc.lpszClassName = "Shell_TrayWnd"
        wc.hInstance = Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly().GetModules()(0))

        If RegisterClassEx(wc) <> 0 Or Marshal.GetLastWin32Error() = 1410 Then
            _fakeTrayHandle = CreateWindowEx(0, "Shell_TrayWnd", "", WS_POPUP, 0, 0, 0, 0, IntPtr.Zero, IntPtr.Zero, wc.hInstance, IntPtr.Zero)

            If _fakeTrayHandle <> IntPtr.Zero Then
                Debug.WriteLine("SUCCESS: Shell_TrayWnd Active. Handle: " & _fakeTrayHandle.ToString())
            End If
        End If
    End Sub

    Private Function FakeTrayWndProc(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
        If msg = WM_COPYDATA Then
            ProcessTrayMessage(lParam)
            Return New IntPtr(1)
        End If
        Return DefWindowProc(hWnd, msg, wParam, lParam)
    End Function


    ' ============================================================
    '  Integrity level (UIPI detection)
    ' ============================================================

    Private Function GetProcessIntegrityLevel(Optional pid As Integer = -1) As UInteger
        Try
            Dim hProc As IntPtr
            If pid = -1 Then
                hProc = Process.GetCurrentProcess().Handle
            Else
                Const PROCESS_QUERY_LIMITED_INFORMATION As UInteger = &H1000
                hProc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, CUInt(pid))
                If hProc = IntPtr.Zero Then Return 0
            End If

            Try
                Dim hToken As IntPtr = IntPtr.Zero
                If Not OpenProcessToken(hProc, &H8, hToken) Then Return 0
                Try
                    Dim needed As UInteger = 0
                    GetTokenInformation(hToken, TokenIntegrityLevel, IntPtr.Zero, 0, needed)
                    If needed = 0 Then Return 0
                    Dim buf As IntPtr = Marshal.AllocHGlobal(CInt(needed))
                    Try
                        If Not GetTokenInformation(hToken, TokenIntegrityLevel, buf, needed, needed) Then Return 0
                        Dim label As TOKEN_MANDATORY_LABEL = Marshal.PtrToStructure(Of TOKEN_MANDATORY_LABEL)(buf)
                        Dim pSid As IntPtr = label.Sid
                        If pSid = IntPtr.Zero Then Return 0
                        If Not IsValidSid(pSid) Then Return 0
                        Dim subAuthCount As Byte = Marshal.ReadByte(pSid, 1)
                        If subAuthCount = 0 Then Return 0
                        Dim pSubAuth As IntPtr = GetSidSubAuthority(pSid, subAuthCount - 1)
                        If pSubAuth = IntPtr.Zero Then Return 0
                        Return CUInt(Marshal.ReadInt32(pSubAuth))
                    Finally
                        Marshal.FreeHGlobal(buf)
                    End Try
                Finally
                    CloseHandle(hToken)
                End Try
            Finally
                If pid <> -1 Then CloseHandle(hProc)
            End Try
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Private Shared Function IntegrityLevelName(il As UInteger) As String
        Select Case il
            Case 0 : Return "Unknown"
            Case 4096 : Return "Low"
            Case 8192 : Return "Medium"
            Case 8448 : Return "Medium+"
            Case 12288 : Return "High"
            Case 16384 : Return "System"
            Case Else : Return "IL=" & il
        End Select
    End Function

    ' ============================================================
    '  Icon helpers
    ' ============================================================

    Private Function BitmapFromHIcon(hIcon As IntPtr) As Bitmap
        If hIcon = IntPtr.Zero Then Return Nothing
        Try
            Dim hCopy As IntPtr = CopyIcon(hIcon)
            If hCopy = IntPtr.Zero Then hCopy = hIcon
            Using ico As Icon = Icon.FromHandle(hCopy)
                If ico Is Nothing Then Return Nothing

                Return CType(ico.ToBitmap(), Bitmap)
            End Using
        Catch
            Return Nothing
        End Try
    End Function

    Private Function GetIconFromWindow(hWnd As IntPtr) As Bitmap
        Try
            Dim result As IntPtr = IntPtr.Zero
            SendMessageTimeout(hWnd, WM_GETICON, New IntPtr(ICON_SMALL), IntPtr.Zero, 2, 100, result)
            Dim hIcon As IntPtr = result
            If hIcon = IntPtr.Zero Then hIcon = GetClassLongPtr(hWnd, GCLP_HICONSM)
            If hIcon <> IntPtr.Zero Then Return BitmapFromHIcon(hIcon)
        Catch
        End Try
        Return Nothing
    End Function

    Private Function GetIconFromProcess(pid As UInteger) As Bitmap
        Try
            Using proc As Process = Process.GetProcessById(CInt(pid))
                Dim path As String = proc.MainModule.FileName
                Using ico As Icon = Icon.ExtractAssociatedIcon(path)
                    Return CType(ico.ToBitmap(), Bitmap)
                End Using
            End Using
        Catch
            Return Nothing
        End Try
    End Function


    ' ============================================================
    '  Misc helpers
    ' ============================================================

    Private Shared Function MakeKey(hWnd As IntPtr, uID As UInteger) As String
        Return "tray_" & hWnd.ToInt64() & "_" & uID
    End Function

    Private Shared Function MAKELONG(lo As Integer, hi As Integer) As Integer
        Return (hi << 16) Or (lo And &HFFFF)
    End Function

    Private Shared Function UFlagsToString(flags As Integer) As String
        Dim parts As New List(Of String)()
        If (flags And NIF_MESSAGE) <> 0 Then parts.Add("MESSAGE")
        If (flags And NIF_ICON) <> 0 Then parts.Add("ICON")
        If (flags And NIF_TIP) <> 0 Then parts.Add("TIP")
        If (flags And NIF_STATE) <> 0 Then parts.Add("STATE")
        If (flags And NIF_INFO) <> 0 Then parts.Add("INFO")
        If (flags And NIF_GUID) <> 0 Then parts.Add("GUID")
        If (flags And NIF_REALTIME) <> 0 Then parts.Add("REALTIME")
        If (flags And NIF_SHOWTIP) <> 0 Then parts.Add("SHOWTIP")
        If parts.Count = 0 Then Return "none"
        Return String.Join("+", parts)
    End Function

    Private Shared Function BalloonFlagsToString(flags As Integer) As String
        Dim parts As New List(Of String)()
        Select Case flags And NIIF_ICON_MASK
            Case NIIF_NONE : parts.Add("NONE")
            Case NIIF_INFO : parts.Add("INFO")
            Case NIIF_WARNING : parts.Add("WARNING")
            Case NIIF_ERROR : parts.Add("ERROR")
            Case NIIF_USER : parts.Add("USER")
        End Select
        If (flags And NIIF_NOSOUND) <> 0 Then parts.Add("NOSOUND")
        If (flags And NIIF_LARGE_ICON) <> 0 Then parts.Add("LARGE_ICON")
        If (flags And NIIF_RESPECT_QUIET_TIME) <> 0 Then parts.Add("QUIET_TIME")
        Return If(parts.Count = 0, "none", String.Join("+", parts))
    End Function

    Private Shared Function BalloonFlagsToToolTipIcon(flags As Integer) As ToolTipIcon
        Select Case flags And NIIF_ICON_MASK
            Case NIIF_INFO : Return ToolTipIcon.Info
            Case NIIF_WARNING : Return ToolTipIcon.Warning
            Case NIIF_ERROR : Return ToolTipIcon.Error
            Case Else : Return ToolTipIcon.None
        End Select
    End Function
#End Region

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function RegisterClassEx(ByRef lpwcx As WNDCLASSEX) As Short
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function CreateWindowEx(ByVal dwExStyle As Integer, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hWndParent As IntPtr, ByVal hMenu As IntPtr, ByVal hInstance As IntPtr, ByVal lpParam As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function DefWindowProc(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function RegisterWindowMessage(ByVal lpString As String) As Integer
    End Function

    Private Declare Function ShowWindow Lib "user32.dll" (
    ByVal hWnd As IntPtr,
    ByVal nCmdShow As SHOW_WINDOW
    ) As Boolean

    <Flags()>
    Private Enum SHOW_WINDOW As Integer
        SW_HIDE = 0
        SW_SHOWNORMAL = 1
        SW_NORMAL = 1
        SW_SHOWMINIMIZED = 2
        SW_SHOWMAXIMIZED = 3
        SW_MAXIMIZE = 3
        SW_SHOWNOACTIVATE = 4
        SW_SHOW = 5
        SW_MINIMIZE = 6
        SW_SHOWMINNOACTIVE = 7
        SW_SHOWNA = 8
        SW_RESTORE = 9
        SW_SHOWDEFAULT = 10
        SW_FORCEMINIMIZE = 11
        SW_MAX = 11
    End Enum

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

    Enum ShowWindowCommands As Integer
        Hide = 0
        Normal = 1
        ShowMinimized = 2
        Maximize = 3
        ShowMaximized = 3
        ShowNoActivate = 4
        Show = 5
        Minimize = 6
        ShowMinNoActive = 7
        ShowNA = 8
        Restore = 9
        ShowDefault = 10
        ForceMinimize = 11
    End Enum

    Public Structure POINTAPI
        Public X As Integer
        Public Y As Integer
        Public Sub New(ByVal X As Integer, ByVal Y As Integer)
            Me.X = X
            Me.Y = Y
        End Sub
    End Structure

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetWindowPos(
    ByVal hWnd As IntPtr,
    ByVal hWndInsertAfter As IntPtr,
    ByVal X As Integer,
    ByVal Y As Integer,
    ByVal cx As Integer,
    ByVal cy As Integer,
    ByVal uFlags As UInteger) As Boolean
    End Function

    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOZORDER As UInteger = &H4
    Private Const SWP_SHOWWINDOW As UInteger = &H40

    Private Sub MoveAndResizeWindow(ByVal hWnd As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer)
        SetWindowPos(hWnd, IntPtr.Zero, X, Y, Width, Height, SWP_NOZORDER Or SWP_SHOWWINDOW)
    End Sub

    <DllImport("shell32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function ExtractIconEx(
    ByVal lpszFile As String,
    ByVal nIconIndex As Integer,
    ByVal phiconLarge() As IntPtr,
    ByVal phiconSmall() As IntPtr,
    ByVal nIcons As UInteger
) As UInteger
    End Function

    Private Const SHGFI_ICON As UInteger = &H100
    Private Const SHGFI_SMALLICON As UInteger = &H1
    Private Const SHGFI_LARGEICON As UInteger = &H0
    Private Const SHGFI_USEFILEATTRIBUTES As UInteger = &H10
    Private Const SHGFI_DISPLAYNAME As UInteger = &H200
    Private Const SHGFI_PIDL As UInteger = &H8

    Private Const FILE_ATTRIBUTE_NORMAL As UInteger = &H80
    Private Const FILE_ATTRIBUTE_DIRECTORY As UInteger = &H10
    Public Function GetFileIcon(filePath As String, ByVal isLargeIcon As Boolean) As Icon
        filePath = EncodeFilePath(filePath)

        Dim shInfo As New SHFILEINFO()
        Dim flags As UInteger = SHGFI_ICON

        If isLargeIcon Then
            flags = flags Or SHGFI_LARGEICON
        Else
            flags = flags Or SHGFI_SMALLICON
        End If

        Dim attributes As UInteger = FILE_ATTRIBUTE_NORMAL

        If String.IsNullOrEmpty(filePath) OrElse System.IO.Directory.Exists(filePath) Then
            If Not String.IsNullOrEmpty(filePath) AndAlso System.IO.Directory.Exists(filePath) Then
                attributes = FILE_ATTRIBUTE_DIRECTORY
            ElseIf String.IsNullOrEmpty(filePath) Then
                flags = flags Or SHGFI_USEFILEATTRIBUTES
                attributes = FILE_ATTRIBUTE_DIRECTORY
            End If
        End If

        Dim result As IntPtr

        Try
            result = Desktop.SHGetFileInfo(
            filePath,
            attributes,
            shInfo,
            CUInt(Marshal.SizeOf(shInfo)),
            flags)
        Catch ex As Exception
            result = IntPtr.Zero
        End Try

        If Not result.Equals(IntPtr.Zero) AndAlso Not shInfo.hIcon.Equals(IntPtr.Zero) Then
            Try
                Dim iconResult As Icon = Icon.FromHandle(shInfo.hIcon)
                Dim finalIcon As Icon = CType(iconResult.Clone(), Icon)

                DestroyIcon(shInfo.hIcon)

                Return finalIcon
            Catch ex As Exception
                Return Nothing
            End Try
        End If

        Return Nothing
    End Function

    Public Function GetFolderIcon(folderPath As String, ByVal isLargeIcon As Boolean) As Icon
        folderPath = EncodeFilePath(folderPath)

        If Not System.IO.Directory.Exists(folderPath) Then
            Return Nothing
        End If

        Dim shInfo As New SHFILEINFO()

        Dim flags As UInteger = SHGFI_ICON

        If isLargeIcon Then
            flags = flags Or SHGFI_LARGEICON
        Else
            flags = flags Or SHGFI_SMALLICON
        End If

        Dim result As IntPtr = Desktop.SHGetFileInfo(
            folderPath,
            FILE_ATTRIBUTE_DIRECTORY,
            shInfo,
            CUInt(Marshal.SizeOf(shInfo)),
            flags)

        If Not result.Equals(IntPtr.Zero) AndAlso Not shInfo.hIcon.Equals(IntPtr.Zero) Then
            Try
                Dim iconResult As Icon = Icon.FromHandle(shInfo.hIcon)
                Dim finalIcon As Icon = CType(iconResult.Clone(), Icon)

                DestroyIcon(shInfo.hIcon)

                Return finalIcon
            Catch ex As Exception
                Return Nothing
            End Try
        End If

        Return Nothing
    End Function

    Public Shared Function GetIcon(filePath As String, ByVal iconIndex As Integer, ByVal isSmallIcon As Boolean) As Icon
        filePath = EncodeFilePath(filePath)

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
    Private Const DWMWA_BORDER_COLOR As Integer = 34
    Private Const DWMWA_WINDOW_CORNER_PREFERENCE As Integer = 33
    Private Const DWMWCP_DONOTROUND As Integer = 1

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

    Public Shared Function GetInactiveAccentColor() As Color
        Try
            Dim regValue = Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\DWM", "AccentColorInactive", Nothing)

            If regValue IsNot Nothing Then
                Dim colorInt As Integer = CInt(regValue)

                Dim r As Byte = CByte(colorInt And &HFF)
                Dim g As Byte = CByte((colorInt >> 8) And &HFF)
                Dim b As Byte = CByte((colorInt >> 16) And &HFF)

                Return Color.FromArgb(255, r, g, b)
            Else
                Return Color.FromArgb(170, 170, 170)
            End If
        Catch
            Return Color.Gray
        End Try
    End Function

    Delegate Sub WinEventDelegate(hWinEventHook As IntPtr, eventType As UInteger, hwnd As IntPtr, idObject As Integer, idChild As Integer, dwEventThread As UInteger, dwmsEventTime As UInteger)

    <DllImport("user32.dll")>
    Private Shared Function SetWinEventHook(eventMin As UInteger, eventMax As UInteger, hmodWinEventProc As IntPtr, lpfnWinEventProc As WinEventDelegate, idProcess As UInteger, idThread As UInteger, dwFlags As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function UnhookWinEvent(hWinEventHook As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindowText(hWnd As IntPtr, lpString As StringBuilder, nMaxCount As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function GetClassName(hWnd As IntPtr, lpClassName As StringBuilder, nMaxCount As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function IsWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function DrawAnimatedRects(hwnd As IntPtr, idAni As Integer, ByRef lprcFrom As RECT, ByRef lprcTo As RECT) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetProgmanWindow(hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetShellWindow(hWnd As IntPtr) As Boolean
    End Function

    Private Const IDANI_CAPTION As Integer = &H3

    Private hHook As IntPtr
    Private procDelegate As WinEventDelegate = AddressOf WinEventProc

    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)

        hHook = SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_OBJECT_DESTROY, IntPtr.Zero, procDelegate, 0, 0, WINEVENT_OUTOFCONTEXT)
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If hHook <> IntPtr.Zero Then
            UnhookWinEvent(hHook)
        End If
        MyBase.OnFormClosing(e)
    End Sub
    Private Const EVENT_SYSTEM_MINIMIZESTART As UInteger = &H16
    Private Const EVENT_SYSTEM_MINIMIZEEND As UInteger = &H17

    Public Shared currentDockPos As DockPosition = DockPosition.Bottom
    Public Shared isSmallIcons As Boolean = False
    Public Shared defSize As Integer = 37
    Public Shared defSmallSize As Integer = 24

    Private SavedWindowPositions As New Dictionary(Of IntPtr, Point)
    Public Sub AnimateMinimizeToButton(hwnd As IntPtr)
        Dim btn = ProcessStrip.Items.Cast(Of ToolStripItem).FirstOrDefault(Function(x) x.Tag.Equals(hwnd))
        If btn Is Nothing Then Return

        Dim rcFrom As New RECT()
        GetWindowRect(hwnd, rcFrom)

        Dim btnPoint As Point = ProcessStrip.PointToScreen(btn.Bounds.Location)
        Dim rcTo As New RECT With {
        .left = btnPoint.X,
        .top = btnPoint.Y,
        .right = btnPoint.X + btn.Width,
        .bottom = btnPoint.Y + btn.Height
    }

        DrawAnimatedRects(hwnd, IDANI_CAPTION, rcFrom, rcTo)

        SetWindowPos(hwnd, IntPtr.Zero, 50, SystemInformation.PrimaryMonitorSize.Height + 5000, 0, 0, SWP_NOSIZE Or SWP_NOZORDER)
    End Sub
    Private Const EVENT_OBJECT_STATECHANGE As UInteger = &H800A
    Private Sub WinEventProc(hWinEventHook As IntPtr, eventType As UInteger, hwnd As IntPtr, idObject As Integer, idChild As Integer, dwEventThread As UInteger, dwmsEventTime As UInteger)
        If Not Me.IsHandleCreated OrElse Me.IsDisposed Then Return

        If idObject <> OBJID_WINDOW Then Return

        If IsHungAppWindow(hwnd) Then Return

        Select Case eventType
            Case EVENT_SYSTEM_FOREGROUND
                Me.Invoke(Sub() CleanupDeadWindows())

                If BlockingProcesses IsNot Nothing Then CheckIfProcessBlocked(hwnd)

                Me.Invoke(Sub()
                              Dim newActiveHWnd As IntPtr = hwnd

                              Dim rootHWnd As IntPtr = GetAncestor(newActiveHWnd, GA_ROOTOWNER)
                              If rootHWnd <> IntPtr.Zero Then
                                  newActiveHWnd = rootHWnd
                              End If

                              Dim myFormHandle As IntPtr = System.Windows.Forms.Application.OpenForms(0).Handle

                              If newActiveHWnd = myFormHandle OrElse
                              newActiveHWnd = WAT.Handle OrElse
                              newActiveHWnd = Startmenu.Handle OrElse
                              newActiveHWnd = ProcessStrip.Handle OrElse
                              newActiveHWnd = AeroPeekScreen.Handle OrElse
                              newActiveHWnd = ScreenDrawer.Handle Then
                                  Return
                              End If

                              ActiveWindowHandle = newActiveHWnd

                              System.Windows.Forms.Application.OpenForms(0).Invoke(Sub()
                                                                                       If LastActiveWindowHandle <> IntPtr.Zero AndAlso LastActiveWindowHandle <> newActiveHWnd Then
                                                                                           UpdateTaskbarSelection(newActiveHWnd)
                                                                                       End If

                                                                                       If IsTaskbarWindow(newActiveHWnd) Then
                                                                                           UpdateTaskbarSelection(hwnd)
                                                                                       End If
                                                                                   End Sub)

                              TrayCleanup()
                          End Sub)

            Case EVENT_SYSTEM_ALERT, EVENT_OBJECT_STATECHANGE
                Me.Invoke(Sub() CleanupDeadWindows())

                If idObject = OBJID_WINDOW Then
                    Me.Invoke(Sub()
                                  TrayCleanup()

                                  CheckIfShouldFlash(hwnd)
                              End Sub)
                End If

            Case EVENT_OBJECT_CREATE
                Me.Invoke(Sub() CleanupDeadWindows())

                If BlockingProcesses IsNot Nothing Then CheckIfProcessBlocked(hwnd)

                Me.Invoke(Sub()
                              Dim targetHWnd As IntPtr = hwnd

                              Dim myFormHandle As IntPtr = System.Windows.Forms.Application.OpenForms(0).Handle
                              Dim watHandle As IntPtr = WAT.Handle
                              Dim startMenuHandle As IntPtr = Startmenu.Handle
                              Dim processStripHandle As IntPtr = ProcessStrip.Handle
                              Dim aeroPeekHandle As IntPtr = AeroPeekScreen.Handle
                              Dim screenDrawerHandle As IntPtr = ScreenDrawer.Handle

                              Task.Run(Sub()
                                           Dim attempts As Integer = 0
                                           Dim maxAttempts As Integer = 100
                                           Dim isReady As Boolean = False

                                           If targetHWnd = myFormHandle OrElse
                                  targetHWnd = watHandle OrElse
                                  targetHWnd = startMenuHandle OrElse
                                  targetHWnd = processStripHandle OrElse
                                  targetHWnd = aeroPeekHandle OrElse
                                  targetHWnd = screenDrawerHandle Then
                                               Return
                                           End If

                                           While attempts < maxAttempts
                                               System.Threading.Thread.Sleep(15)

                                               Dim titleLength As Integer = GetWindowTextLength(targetHWnd)
                                               Dim visible As Boolean = IsWindowVisible(targetHWnd)

                                               If titleLength > 0 AndAlso visible Then
                                                   isReady = True
                                               End If

                                               attempts += 1

                                               If isReady Then
                                                   Exit While
                                               End If
                                           End While

                                           If isReady Then
                                               System.Windows.Forms.Application.OpenForms(0).Invoke(Sub()
                                                                                                        If IsTaskbarWindow(targetHWnd) AndAlso Not IsProcessHidden(targetHWnd) Then
                                                                                                            Dim exists As Boolean = ProcessStrip.Items.Cast(Of ToolStripItem)().Any(Function(x) x.Tag IsNot Nothing AndAlso CType(x.Tag, IntPtr) = targetHWnd)
                                                                                                            If Not exists Then
                                                                                                                AddToolStripButton(hwnd, GetWindowTitleSafe(hwnd))
                                                                                                            End If
                                                                                                        End If
                                                                                                    End Sub)
                                           End If
                                       End Sub)
                              TrayCleanup()
                          End Sub)

            Case EVENT_OBJECT_DESTROY
                Me.Invoke(Sub() CleanupDeadWindows())

                Me.Invoke(Sub()
                              If idObject = 0 AndAlso idChild = 0 Then
                                  Dim exists As Boolean = ProcessStrip.Items.Cast(Of ToolStripItem)().Any(Function(x) x.Tag IsNot Nothing AndAlso CType(x.Tag, IntPtr) = hwnd)
                                  If exists Then
                                      RemoveToolStripButton(hwnd)
                                  End If
                              End If

                              TrayCleanup()
                          End Sub)

            Case EVENT_OBJECT_NAMECHANGE
                Me.Invoke(Sub() CleanupDeadWindows())

                If BlockingProcesses IsNot Nothing Then CheckIfProcessBlocked(hwnd)

                Me.Invoke(Sub()
                              TrayCleanup()

                              Dim targetHWnd As IntPtr = hwnd
                              Dim index As Integer = ProcessStrip.Items.Cast(Of ToolStripItem)().ToList().
                              FindIndex(Function(x) x.Tag IsNot Nothing AndAlso CType(x.Tag, IntPtr) = targetHWnd)

                              If index <> -1 Then
                                  UpdateWindowInfo(hwnd)
                              Else
                                  If IsTaskbarWindow(targetHWnd) Then
                                      Dim exists As Boolean = ProcessStrip.Items.Cast(Of ToolStripItem)().Any(Function(x) x.Tag IsNot Nothing AndAlso CType(x.Tag, IntPtr) = targetHWnd)
                                      If Not exists Then
                                          AddToolStripButton(hwnd, GetWindowTitleSafe(hwnd))
                                      End If
                                  End If
                              End If
                          End Sub)

            Case EVENT_OBJECT_LOCATIONCHANGE
                If BlockingProcesses IsNot Nothing Then CheckIfProcessBlocked(hwnd)

                If OnlyShowWindowsOnCurrentMonitor Then
                    Me.Invoke(Sub() RefreshTaskbarForMonitor())
                End If
            Case EVENT_SYSTEM_MINIMIZESTART
                Me.BeginInvoke(Sub()
                                   Dim rect As New RECT()
                                   GetWindowRect(hwnd, rect)

                                   If Not SavedWindowPositions.ContainsKey(hwnd) Then
                                       SavedWindowPositions.Add(hwnd, New Point(rect.left, rect.top))
                                   End If

                                   AnimateMinimizeToButton(hwnd)
                               End Sub)

            Case EVENT_SYSTEM_MINIMIZEEND
                Me.BeginInvoke(Sub()
                                   If SavedWindowPositions.ContainsKey(hwnd) Then
                                       Dim originalPos = SavedWindowPositions(hwnd)

                                       SavedWindowPositions.Remove(hwnd)
                                   End If
                               End Sub)
        End Select
    End Sub
    Private Shared Property hiddenProcesses As New List(Of String)()
    Private Shared Property blockedProcesses As New List(Of String)()

    Public Shared Sub RefreshBlockedProcesses()
        blockedProcesses.Clear()

        Try
            Using key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\BlockedProcesses\.")
                If key IsNot Nothing Then
                    For Each valueName As String In key.GetValueNames()
                        If Not String.IsNullOrEmpty(valueName) Then
                            Dim procName As String = valueName.Trim().ToLower()

                            If procName.EndsWith(".exe") Then
                                procName = procName.Substring(0, procName.Length - 4)
                            End If

                            If Not String.IsNullOrEmpty(procName) AndAlso Not blockedProcesses.Contains(procName) Then
                                blockedProcesses.Add(procName)
                            End If
                        End If
                    Next
                End If
            End Using
        Catch
        End Try
    End Sub

    Public Sub CheckIfProcessBlocked(ByVal hWnd As IntPtr)
        If hWnd = IntPtr.Zero Then Return

        If IsProcessBlocked(hWnd) Then
            Dim pr As Process = GetProcessFromWindowHandle(hWnd)
            If pr Is Nothing Then Return

            If [Enum].GetNames(GetType(CriticalProcesses)).Contains(pr.ProcessName, StringComparer.OrdinalIgnoreCase) Then Return

            Dim killMethod As Object = Microsoft.Win32.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "KillMethod", "1")
            If killMethod IsNot Nothing AndAlso killMethod.ToString() = "0" Then
                pr.Kill()
            Else
                pr.CloseMainWindow()
            End If
        End If
    End Sub

    Public Shared Sub RefreshHiddenProcesses()
        hiddenProcesses.Clear()

        Try
            Using key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\HiddenProcesses")
                If key IsNot Nothing Then
                    For Each valueName As String In key.GetValueNames()
                        If Not String.IsNullOrEmpty(valueName) Then
                            Dim procName As String = valueName.Trim().ToLower()

                            If procName.EndsWith(".exe") Then
                                procName = procName.Substring(0, procName.Length - 4)
                            End If

                            If Not String.IsNullOrEmpty(procName) AndAlso Not hiddenProcesses.Contains(procName) Then
                                hiddenProcesses.Add(procName)
                            End If
                        End If
                    Next
                End If
            End Using
        Catch
        End Try
    End Sub

    Private Shared Function IsProcessBlocked(ByVal hWnd As IntPtr) As Boolean
        If blockedProcesses.Count = 0 Then Return False
        If hWnd = IntPtr.Zero Then Return False

        Dim rootHWnd As IntPtr = GetAncestor(hWnd, GA_ROOTOWNER)
        Dim checkHWnd As IntPtr = If(rootHWnd <> IntPtr.Zero, rootHWnd, hWnd)

        Dim processId As Integer = 0
        GetWindowThreadProcessId(checkHWnd, processId)

        If processId <> 0 Then
            Try
                Using proc As Process = Process.GetProcessById(processId)
                    Dim procName As String = proc.ProcessName.ToLower()
                    Return blockedProcesses.Contains(procName)
                End Using
            Catch
                Return False
            End Try
        End If
        Return False
    End Function

    Private Shared Function IsProcessHidden(ByVal hWnd As IntPtr) As Boolean
        If hiddenProcesses.Count = 0 Then Return False
        If hWnd = IntPtr.Zero Then Return False

        Dim rootHWnd As IntPtr = GetAncestor(hWnd, GA_ROOTOWNER)
        Dim checkHWnd As IntPtr = If(rootHWnd <> IntPtr.Zero, rootHWnd, hWnd)

        Dim processId As Integer = 0
        GetWindowThreadProcessId(checkHWnd, processId)

        If processId <> 0 Then
            Try
                Using proc As Process = Process.GetProcessById(processId)
                    Dim procName As String = proc.ProcessName.ToLower()
                    Return hiddenProcesses.Contains(procName)
                End Using
            Catch
                Return False
            End Try
        End If
        Return False
    End Function

    Public UseExplorerFP As Boolean = True
    Public UseExplorerFM As Boolean = True
    Public CustomFMPath As String = String.Empty

    Public Shared AlocateAppbarWorkingArea As Boolean = True

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Check if Registry Settings are present...
        If My.Computer.Registry.CurrentUser.OpenSubKey("Shell") IsNot Nothing Then

            Dim currentpath As String = Application.StartupPath

            Dim pathsuccessful As String = String.Empty

            Dim result = MessageBox.Show("Krr's Shell didn't found its Registry key where all the settings will be. Without it the shell cannot continue, to avoid the Shell being corrupted." & Environment.NewLine & Environment.NewLine &
                            "If you downloaded the Shell from GitHub, please check if you have another file called ""Import_me.reg"" and execute it first." & Environment.NewLine &
                            "Do you want to locate the file from the Shell directly?", "No Settings found.", MessageBoxButtons.YesNo, MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1)

            Select Case result
                Case DialogResult.Yes
                    For Each f As String In Directory.GetFiles(currentpath)

                        ' Searching for Import_me.reg...
                        If Path.GetFileName(f).ToLower = "import_me.reg" Then
                            pathsuccessful = f
                            Exit For
                        End If

                    Next

                    If File.Exists(pathsuccessful) AndAlso Path.GetFileName(pathsuccessful).ToLower = "import_me.reg" Then
                        Dim result1 = MessageBox.Show("SUCCESS! The file ""Import_me.reg"" has been found, do you want to import it right now?", "Import_me.reg found", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                        Select Case result1
                            Case DialogResult.Yes
                                Try
                                    Shell(pathsuccessful, AppWinStyle.NormalFocus, True)

                                    MessageBox.Show("Now please check your Registry Editor and look to those key path: ""HKEY_CURRENT_USER\Shell"" and if the key is there, now close this message box, start the program again and enjoy!", "Program Restart required", MessageBoxButtons.OK, MessageBoxIcon.Information)

                                    End
                                Catch ex As Exception
                                    MessageBox.Show(ex.Message)

                                    End
                                End Try
                            Case DialogResult.No
                                End
                        End Select
                    Else
                        MessageBox.Show("File not found! Please check the file ""Import_me.reg"" somewhere else to continue. This Shell will be now ended.", "Failed to locate Import_me.reg", MessageBoxButtons.OK, MessageBoxIcon.Error)

                        End
                    End If
                Case DialogResult.No
                    MessageBox.Show("This Shell will be now ended. As the Registry keys are the main part the Shell needs to communicate with.", "Canceled.", MessageBoxButtons.OK, MessageBoxIcon.Error)

                    End
            End Select
        End If


        ' Setting up Keyboard hooks
        If GlobalKeyboardHook.SetHook() Then
            Debug.WriteLine("Keyboard Hooks set up!")
        Else
            MessageBox.Show("Keyboard Hooks failed to set up!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        AddHandler GlobalKeyboardHook.HotkeyPressed, AddressOf GlobalKeyboardHook_HotkeyPressed


        ' Startup Apps
        Dim currentDefShell As String = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe")
        If currentDefShell.ToLower.Contains("krrshell") Then

            Dim startupFolderPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Startup)

            Dim runningProcesses As New HashSet(Of String)(
    Process.GetProcesses().Select(Function(p) p.ProcessName),
    StringComparer.OrdinalIgnoreCase
)

            For Each i As String In Directory.GetFiles(startupFolderPath)
                Dim fi As New FileInfo(i)

                If fi.Name.ToLower = "desktop.ini" Then Continue For
                If String.IsNullOrWhiteSpace(fi.FullName) Then Continue For

                ' Checks if the process is already running.
                If runningProcesses.Contains(Path.GetFileNameWithoutExtension(fi.Name)) Then Continue For

                If Not fi.Extension.ToLower = ".lnk" Then
                    Process.Start(New ProcessStartInfo(i) With {.UseShellExecute = True})
                Else
                    Try
                        Dim pathName As String = GetShortcutTarget(fi.FullName)

                        Process.Start(New ProcessStartInfo() With {
                                      .UseShellExecute = True,
                                      .FileName = If(pathName, fi.FullName)})
                    Catch : End Try
                End If
            Next
        End If


        ' Load non-form related settings
        LoadOtherSettings()


        ' Desktop
        If SystemInformation.MonitorCount > 1 Then
            Dim screens As Screen() = Screen.AllScreens
            For Each screen As Screen In screens
                Dim newDesktop As New Desktop With {
                    .Location = screen.Bounds.Location,
                    .Size = screen.Bounds.Size
                }
                newDesktop.Show()
                newDesktop.SendToBack()
            Next
        Else
            Desktop.Location = New Point(0, 0)
            Desktop.Size = SystemInformation.PrimaryMonitorSize
            Desktop.Show()
            Desktop.SendToBack()
        End If

        ' Screen Filter
        If SystemInformation.MonitorCount > 1 Then
            Dim screens As Screen() = Screen.AllScreens
            For Each screen As Screen In screens
                Dim newFilter As New ScreenDrawer With {
                    .Location = screen.Bounds.Location,
                    .Size = screen.Bounds.Size
                }
                newFilter.Show()
                newFilter.BringToFront()
            Next
        Else
            ScreenDrawer.Location = New Point(0, 0)
            ScreenDrawer.Size = SystemInformation.PrimaryMonitorSize
            ScreenDrawer.Show()
            ScreenDrawer.BringToFront()
        End If


        ' Apply Theme and sizing
        Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layout", 2)
            Case 0 ' DWM
                Try
                    Dim side As Side = side
                    side.Left = -1
                    side.Right = -1
                    side.Top = -1
                    side.Bottom = -1
                    Dim result As Integer = DwmExtendFrameIntoClientArea(Me.Handle, side)
                Catch ex As Exception
                    Dim side As Side = side
                    side.Left = 0
                    side.Right = 0
                    side.Top = 0
                    side.Bottom = 0
                    Dim result As Integer = DwmExtendFrameIntoClientArea(Me.Handle, side)
                End Try

                AppbarProperties.rb_ABCdwm.Checked = True

                Me.Opacity = 1

            Case 1 ' Transparent
                Dim side As Side = side
                side.Left = 0
                side.Right = 0
                side.Top = 0
                side.Bottom = 0
                Dim result As Integer = DwmExtendFrameIntoClientArea(Me.Handle, side)

                AppbarProperties.rb_ABCdefault.Checked = True

                Try
                    Me.Opacity = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "TransparencyLevel", 90) / 100
                Catch ex As Exception
                    Me.Opacity = 0.9
                End Try

            Case 2 ' Opaque
                Dim side As Side = side
                side.Left = 0
                side.Right = 0
                side.Top = 0
                side.Bottom = 0
                Dim result As Integer = DwmExtendFrameIntoClientArea(Me.Handle, side)

                AppbarProperties.rb_ABCopaque.Checked = True

                Me.Opacity = 1
        End Select

        SetWindowTheme(Me.Handle, "Taskbar", Nothing)
        WAT.CreateControl()
        AltTab.CreateControl()

        ' Position and form options
        defSize = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\Alignment", "DefSize", 37)
        defSmallSize = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\Alignment", "DefSmallSize", 24)

        currentDockPos = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\Alignment", "Dock", DockPosition.Bottom)
        isSmallIcons = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\Alignment", "SmallIcons", False)

        If Not Me.WindowState = FormWindowState.Normal Then
            Me.WindowState = FormWindowState.Normal
        End If

        Me.TopMost = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "OnTop", True)

        LockAppbarToolStripMenuItem.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Locked", False)
        AutoHideToolStripMenuItem.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "AutoHide", 0)


        ' Working area
        AlocateAppbarWorkingArea = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "AppbarAlign", True)

        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "Enabled", False) = 1 Then
            Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "Type", 0)

                Case 0 ' Explorer WorkingArea
                    fBarRegistered = False
                    RegisterBar()

                Case 1 ' User32.dll WorkingArea
                    UpdateWorkingArea()

            End Select
        End If

        Me.Location = New Point(Me.Location.X - 1, Me.Location.Y - 1)
        Me.Size = New Size(Me.Width - 1, Me.Height - 1)


        ' Color and BackImage
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "UseSystemColor", 1) = 1 Then
            Dim accentColor As Color = GetAccentColor()

            Me.BackColor = accentColor
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", ColorTranslator.ToHtml(accentColor))
        Else
            Me.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", "#B77358"))
        End If

        Dim backImagePath As String = EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImage", ""))
        Dim isBackImageEnabled As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowBackImage", False)

        If isBackImageEnabled Then
            Select Case File.Exists(backImagePath)
                Case True
                    Try
                        Me.BackgroundImage = Image.FromFile(backImagePath)
                        AppbarProperties.pb_CIPreview.Image = Image.FromFile(backImagePath)
                        AppbarProperties.cmb_CIPathToImage.Text = backImagePath
                    Catch
                        Me.BackgroundImage = My.Resources.AppBarMain
                        AppbarProperties.pb_CIPreview.Image = My.Resources.AppBarMain
                        AppbarProperties.cmb_CIPathToImage.Text = ""
                    End Try

                Case False
                    Me.BackgroundImage = My.Resources.AppBarMain
                    AppbarProperties.pb_CIPreview.Image = My.Resources.AppBarMain
                    AppbarProperties.cmb_CIPathToImage.Text = ""
            End Select

            ' Set backimage layout
            Me.BackgroundImageLayout = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImageLayout", "1")
        End If

        Splitter1.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "SplitterColor", "#000000"))
        Splitter2.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "SplitterColor", "#000000"))
        Splitter3.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "SplitterColor", "#000000"))


        ' Toolbars
        btn_ShowDesktop.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowDesktopButton", True)

        If LockAppbarToolStripMenuItem.Checked = True Then
            Splitter1.Visible = False
            Splitter2.Visible = False
            Splitter3.Visible = False
        Else
            Splitter1.Visible = True
            Splitter2.Visible = True
            Splitter3.Visible = True
        End If

        Select Case currentDockPos
            Case DockPosition.Top, DockPosition.Bottom
                LoadBandSizes(False)
            Case DockPosition.Left, DockPosition.Right
                LoadBandSizes(True)
        End Select

        ' Start menu and button
        Start.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "StartButton", True)
        StartButtonToolStripMenuItem.Checked = Start.Visible

        Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0")
            Case 0 ' Default
                AppbarProperties.rb_SBDefault.Checked = True
                Start.BackgroundImage = My.Resources.StartRight

            Case 1 ' Custom Image
                AppbarProperties.rb_SBCustomImage.Checked = True

                Try
                    Start.BackgroundImage = Image.FromFile(EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Normal", "")))
                Catch ex As Exception
                    Start.BackgroundImage = My.Resources.StartRight
                End Try

            Case 2 ' ORB Logic
                Dim orbPath As String = EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "CustomORB", ""))

                If orbPath IsNot Nothing AndAlso File.Exists(orbPath) Then
                    Try
                        LoadAndSplitOrb(orbPath)

                        Start.BackgroundImage = OrbNormal
                    Catch ex As Exception
                        Start.BackgroundImage = My.Resources.StartRight
                    End Try
                End If
        End Select

        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Layout", "0") = 0 Then
            Start.BackgroundImageLayout = ImageLayout.Stretch
        Else
            Start.BackgroundImageLayout = ImageLayout.Center
        End If

        Startmenu.CreateControl()

        ' AppStrip
        Panel1.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "PinnedBar", True)
        PinnedBarToolStripMenuItem.Checked = Panel1.Visible

        Dim PADdefault As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Internet Explorer\Quick Launch"
        Dim PADtoSet As String = EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "PinnedAppsLocation", PADdefault).ToString.Trim)
        Dim PADitExist As Boolean = False

        While Not PADitExist
            Select Case Directory.Exists(PADtoSet)
                Case True
                    PinnedAppsDir = PADtoSet
                    LoadPinnedApps()

                    PADitExist = True
                Case False
                    If String.IsNullOrWhiteSpace(PADtoSet) Then
                        ' When the string is empty, set up a default directory
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "PinnedAppsLocation", PinnedAppsDir, Microsoft.Win32.RegistryValueKind.String)

                        PinnedAppsDir = PADdefault
                        LoadPinnedApps()

                        PADitExist = True
                    Else
                        ' If it isn't empty, ask the user
                        Select Case MsgBox($"Failed to register path for pinned apps: ""{PADtoSet}""" & Environment.NewLine & "Do you want to change it to something else? If you select ""No"", the AppStrip will stay empty.", MsgBoxStyle.YesNo, "Failed to register AppStrip Path")
                            Case MsgBoxResult.Yes
                                Dim FBD As New FolderBrowserDialog With {
                                .Description = "Select a default location where your Pinned apps will be:",
                                .RootFolder = Environment.SpecialFolder.Desktop,
                                .ShowNewFolderButton = True,
                                .SelectedPath = If(PinnedAppsDir, PADdefault)
                                }

                                If FBD.ShowDialog = DialogResult.OK Then
                                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "PinnedAppsLocation", FBD.SelectedPath, Microsoft.Win32.RegistryValueKind.String)
                                    PADtoSet = FBD.SelectedPath

                                    PinnedAppsDir = PADdefault
                                    LoadPinnedApps()
                                End If

                                PADitExist = True

                            Case Else
                                PADitExist = True
                        End Select
                    End If
            End Select
        End While


        ' ProcessStrip
        procDelegate = New WinEventDelegate(AddressOf WinEventProc)
        hHook = SetWinEventHook(EVENT_SYSTEM_ALERT, EVENT_OBJECT_NAMECHANGE, IntPtr.Zero, procDelegate, 0, 0, WINEVENT_OUTOFCONTEXT)

        RefreshHiddenProcesses()
        RefreshBlockedProcesses()
        LoadApps()


        ' Tray Icons
        Panel4.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "MenuBar", True)
        MenuBarToolStripMenuItem.Checked = Panel4.Visible

        InitTrayIcons()

        Dim oskPath As String = EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "Keyboard", Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\osk.exe"))
        If ShowKeyboardButton = True AndAlso File.Exists(oskPath) Then
            ToolStripButton2.Visible = True
        Else
            ToolStripButton2.Visible = False
        End If

        UpdatePowerUI()
        UpdateNetworkUI()


        ' Language List

        Try : btn_LanguageBar.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "LanguageList", True) : Catch
            btn_LanguageBar.Visible = True
        End Try

        UpdateLanguageDisplay()


        ' Clock and Alarms
        isAutoRefreshing = False
        StartClockUpdate(1000)

        Dim timerInterval As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Network", "AutoRefreshInterval", 5000)

        If timerInterval < 1000 Then timerInterval = 5000
        StartNetworkTimer(timerInterval / 1000)

        Dim rvA = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\", "DisableAlarms", 0)
        If rvA = 0 Then
            DisableAlarmOption.Checked = False
            AlarmController.Enabled = True
        Else
            DisableAlarmOption.Checked = True
            AlarmController.Enabled = False
        End If


        ' Initialize Shell to System
        shellHookMessageId = RegisterWindowMessage("SHELLHOOK")
        SetShellWindow(Me.Handle)
        SetProgmanWindow(Desktop.Handle)


        ' Final code
        Me.BringToFront()
    End Sub

    Private shellHookMessageId As Integer

    Dim layer1 As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer1", "50")
    Dim layer2 As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer2", "107")
    Dim layer3 As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer3", "200")

    Public Sub LoadOtherSettings()
        ' Setting Custom File manager (if detected)
        Try ' FP = File Picker
            Dim avoidefp As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "AvoidExplorerFileDialog", False)
            avoidefp = Not UseExplorerFM
        Catch ex As Exception
            UseExplorerFP = True
        End Try

        Try ' FM = File Manager
            Dim avoidefm As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "UseCustomFileManager", False)
            avoidefm = Not UseExplorerFM
        Catch ex As Exception
            UseExplorerFM = True
        End Try

        Try
            Dim cfm As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "CustomFileManager", String.Empty)
            Dim encodedCFM As String = EncodeFilePath(cfm)

            If Not String.IsNullOrWhiteSpace(encodedCFM) AndAlso File.Exists(encodedCFM) Then
                If New FileInfo(EncodeFilePath(encodedCFM)).Extension.ToLower = ".exe" Then
                    CustomFMPath = encodedCFM
                End If
            End If
        Catch ex As Exception
            CustomFMPath = String.Empty
        End Try


        ' Blocking processes
        ToggleProcessMonitoring(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "Enabled", False))


        ' Custom paths
        Dim defsndvol As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\SndVol.exe"
        Dim soundcntpath As String = EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "SoundControl", defsndvol))

        If String.IsNullOrWhiteSpace(soundcntpath) OrElse Not File.Exists(soundcntpath) Then
            If File.Exists(defsndvol) Then soundcontrolPath = soundcntpath
        Else
            soundcontrolPath = soundcntpath
        End If
    End Sub

    Public Sub LoadBandSizes(ByVal top As Boolean)
        If top = True Then
            Start.Height = layer1
            Panel1.Height = layer2
            Panel4.Height = layer3
            btn_ShowDesktop.Height = 12
        Else
            Start.Width = layer1
            Panel1.Width = layer2
            Panel4.Width = layer3
            btn_ShowDesktop.Width = 12
        End If
    End Sub

    Public SC_isSecondsEnabled As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSecondsInSystemClock", False)
    Public SC_ShowTime As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\DateTime", "ShowTime", True)
    Public SC_ShowDay As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\DateTime", "ShowDay", True)
    Public SC_ShowDate As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\DateTime", "ShowDate", True)

    Public Async Sub StartClockUpdate(ByVal intervalTicks As Integer)
        If isAutoRefreshing Then Return
        isAutoRefreshing = True

        If intervalTicks <= 0 Then intervalTicks = 300

        While isAutoRefreshing
            Dim currentTime As String = String.Empty
            Dim currentDay As String = String.Empty
            Dim currentDate As String = String.Empty

            ' Options
            If SC_ShowTime = True Then : If SC_isSecondsEnabled Then
                    currentTime = DateTime.Now.ToString("T")
                Else
                    currentTime = DateTime.Now.ToString("t")
                End If
            End If

            If SC_ShowDay = True Then currentDay = DateTime.Now.ToString("dddd")
            If SC_ShowDate = True Then currentDate = DateTime.Now.ToString("d")

            ' Showing time
            Dim updateUI As Action = Sub()
                                         Dim displayText As String = currentTime

                                         If SC_ShowDay AndAlso Not String.IsNullOrEmpty(currentDay) Then
                                             displayText += Environment.NewLine & currentDay
                                         End If
                                         If SC_ShowDate AndAlso Not String.IsNullOrEmpty(currentDate) Then
                                             displayText += Environment.NewLine & currentDate
                                         End If

                                         TimeLabel.Text = displayText
                                     End Sub

            If Me.InvokeRequired Then
                Me.Invoke(updateUI)
            Else
                updateUI()
            End If

            If Me.IsDisposed OrElse Not Me.IsHandleCreated Then Exit While

            ' Time sync
            Dim msToNextTick As Integer

            If intervalTicks >= 60000 Then
                msToNextTick = 60000 - Now.Millisecond - (Now.Second * 1000)
            Else
                Dim totalMilliseconds As Long = Now.Ticks \ TimeSpan.TicksPerMillisecond
                Dim remainder As Integer = CInt(totalMilliseconds Mod intervalTicks)

                msToNextTick = intervalTicks - remainder
            End If

            If msToNextTick <= 0 Then msToNextTick = intervalTicks

            Await Task.Delay(msToNextTick)
        End While
    End Sub
    Public ShowKeyboardButton As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowKeyboard", False)
    Private Sub GlobalKeyboardHook_HotkeyPressed(sender As Object, e As EventArgs)
        If Startmenu.Visible = True Then
            Startmenu.FST = False
            StartMenuShowAndHide()
            UpdateStartButtonImage(Not IsMouseOnStart)
        Else
            Startmenu.FST = True
            StartMenuShowAndHide()
            UpdateStartButtonImage()
        End If
    End Sub

    Private Sub ToolStrip1_ItemAdded(sender As Object, e As ToolStripItemEventArgs) Handles AppStrip.ItemAdded
        e.Item.DisplayStyle = ToolStripItemDisplayStyle.Image
    End Sub

    Private Sub ProcessTimer_Tick(sender As Object, e As EventArgs) Handles ProcessTimer.Tick
        ProcessTimer.Stop()

        'Await Task.Run(Sub() CheckForProcessUpdates())

        TrayCleanup()

        ProcessTimer.Start()
    End Sub

    <DllImport("user32.dll")>
    Private Shared Function IsIconic(hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetSystemMenu(ByVal hWnd As IntPtr, ByVal bRevert As Boolean) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function TrackPopupMenuEx(ByVal hMenu As IntPtr, ByVal fuFlags As UInteger, ByVal x As Integer, ByVal y As Integer, ByVal hWnd As IntPtr, ByVal lptpm As IntPtr) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function PostMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean
    End Function

    Private Const TPM_RETURNCMD As UInteger = &H100
    Private Const WM_SYSCOMMAND As UInteger = &H112
    Private Const WM_POPUPSYSTEMMENU As UInteger = &H313

    Private LastActiveWindowHandle As IntPtr = IntPtr.Zero
    Private Sub ProcessStripItemClick(sender As Object, e As ToolStripItemClickedEventArgs) Handles ProcessStrip.ItemClicked
        WAT.Hide()

        Dim item = TryCast(e.ClickedItem, ToolStripButton)

        If item.Tag IsNot Nothing AndAlso TypeOf item.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(item.Tag, IntPtr)

            If My.Computer.Keyboard.CtrlKeyDown Then
                Dim startInfo As New ProcessStartInfo With {
                    .FileName = GetProcessFromWindowHandle(windowHandle).MainModule.FileName,
                    .UseShellExecute = True
                }

                If My.Computer.Keyboard.ShiftKeyDown Then
                    startInfo.Verb = "runas"
                End If

                Process.Start(startInfo)
                Exit Sub
            End If

            StopAlert(windowHandle)

            LastActiveWindowHandle = windowHandle

            Dim p As Process = GetProcessFromWindowHandle(windowHandle)

            Try
                Dim currentState As String = GetWindowState(windowHandle)

                If item.Checked = False Then
                    Select Case currentState
                        Case "Normal"
                            SetForegroundWindow(windowHandle)
                            ActiveWindowHandle = windowHandle
                            ShowWindow(windowHandle, SHOW_WINDOW.SW_NORMAL)

                        Case "Maximized"
                            SetForegroundWindow(windowHandle)
                            ActiveWindowHandle = windowHandle
                            ShowWindow(windowHandle, SHOW_WINDOW.SW_MAXIMIZE)

                        Case "Minimized"
                            SetForegroundWindow(windowHandle)
                            ActiveWindowHandle = windowHandle
                            ShowWindow(windowHandle, SHOW_WINDOW.SW_RESTORE)

                        Case Else
                            SetForegroundWindow(windowHandle)
                            ActiveWindowHandle = windowHandle
                            ShowWindow(windowHandle, SHOW_WINDOW.SW_RESTORE)
                    End Select

                    For Each item_ As ToolStripItem In Me.ProcessStrip.Items

                        If TypeOf item_ Is ToolStripButton Then
                            CType(item_, ToolStripButton).Checked = False

                        ElseIf TypeOf item_ Is ToolStripDropDownButton Then
                            Dim dropDown As ToolStripDropDownButton = CType(item_, ToolStripDropDownButton)

                            For Each subItem As ToolStripItem In dropDown.DropDownItems
                                If TypeOf subItem Is ToolStripMenuItem Then
                                    CType(subItem, ToolStripMenuItem).Checked = False
                                End If
                            Next
                        End If
                    Next

                    item.Checked = True
                Else
                    If LastActiveWindowHandle <> IntPtr.Zero Then ActiveWindowHandle = LastActiveWindowHandle
                    ShowWindow(windowHandle, SHOW_WINDOW.SW_MINIMIZE)

                    item.Checked = False
                End If

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Cannot focus a process")
            End Try

            item.Select()
        End If
    End Sub

    Public mainPeek As New AeroPeek
    Private Sub ProcessStrip_MouseMove(sender As Object, e As MouseEventArgs) Handles ProcessStrip.MouseMove
        If cm_Main.Visible OrElse cm_ProcessItem.Visible Then
            WAT.Hide()
            Exit Sub
        End If

        Dim Item As ToolStripButton = ProcessStripSelectedItem

        If ProcessStrip.GetItemAt(e.Location) Is Nothing Then
            ProcessStripSelectedItem = Nothing
        End If

        If Item IsNot Nothing Then
            If TypeOf Item.Tag Is IntPtr Then
                Dim windowHandle As IntPtr = CType(Item.Tag, IntPtr)

                Try
                    If WAT.Visible AndAlso Object.Equals(WAT.Tag, windowHandle) Then
                        CloseTimer.Stop()
                        Exit Sub
                    End If

                    If IsHungAppWindow(windowHandle) Then Exit Sub

                    cm_ProcessItem.Tag = windowHandle
                    WAT.Tag = windowHandle

                    WAT.Label1.Text = Item.Text
                    WAT.Button4.Image = Nothing
                    WAT.Button4.BackgroundImage = Nothing

                    Dim targetHeight As Integer = CInt(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband", "MinThumbSizePx", 178))
                    Dim magicWidth As Integer = 178

                    Dim finalWidth As Integer = magicWidth
                    Dim finalHeight As Integer = targetHeight

                    Try
                        Dim wRect As WAT.Rect = WAT.GetRealWindowRect(windowHandle)
                        Dim wWidth As Integer = wRect.Right - wRect.Left
                        Dim wHeight As Integer = wRect.Bottom - wRect.Top

                        If wHeight > 0 AndAlso wWidth > 0 Then
                            If wHeight > targetHeight Then
                                Dim aspectRatio As Double = CDbl(wWidth) / CDbl(wHeight)
                                finalWidth = CInt(targetHeight * aspectRatio)
                                finalHeight = targetHeight
                            Else
                                finalWidth = magicWidth
                                finalHeight = wHeight
                            End If
                        End If
                    Catch
                        finalWidth = magicWidth
                        finalHeight = targetHeight
                    End Try

                    finalWidth = Math.Max(finalWidth, WAT.MinimumSize.Width)
                    finalHeight = Math.Max(finalHeight, WAT.MinimumSize.Height)

                    Dim itemScreenBounds As Rectangle = Item.Bounds
                    If Item.Owner IsNot Nothing Then
                        itemScreenBounds = Item.Owner.RectangleToScreen(Item.Bounds)
                    End If

                    Dim targetX As Integer
                    Dim targetY As Integer

                    Select Case currentDockPos
                        Case DockPosition.Bottom
                            targetX = CInt(itemScreenBounds.Left + (itemScreenBounds.Width / 2) - (finalWidth / 2))
                            targetY = Me.Top - finalHeight

                        Case DockPosition.Top
                            targetX = CInt(itemScreenBounds.Left + (itemScreenBounds.Width / 2) - (finalWidth / 2))
                            targetY = Me.Bottom

                        Case DockPosition.Left
                            targetX = Me.Right
                            targetY = CInt(itemScreenBounds.Top + (itemScreenBounds.Height / 2) - (finalHeight / 2))

                        Case DockPosition.Right
                            targetX = Me.Left - finalWidth
                            targetY = CInt(itemScreenBounds.Top + (itemScreenBounds.Height / 2) - (finalHeight / 2))
                    End Select

                    CloseTimer.Stop()
                    Dim targetRect As New Rectangle(targetX, targetY, finalWidth, finalHeight)
                    WAT.MoveTo(targetRect)

                    WAT.WATPeek.ShowControlPeek(windowHandle, WAT, WAT.Button4, True)

                Catch ex As Exception
                    WAT.Label1.Text = Item.Text
                    WAT.Button4.BackgroundImage = Nothing
                    WAT.Button4.Image = My.Resources.ProgramMedium

                    Dim fallbackX As Integer = Control.MousePosition.X - (WAT.Width / 2)
                    Dim fallbackY As Integer = SystemInformation.WorkingArea.Height - WAT.Height

                    WAT.MoveTo(New Rectangle(fallbackX, fallbackY, 168, 62))
                End Try
            End If
        End If
    End Sub

    Private Sub ProcessStrip_MouseEnter(sender As Object, e As EventArgs) Handles ProcessStrip.MouseEnter
        If Startmenu.Visible Then
            Startmenu.Focus()
            Exit Sub
        Else
            ProcessStrip.Focus()
            ProcessStrip.Select()

            Hidden = False

            Startmenu.FST = False
            AutoHideOnOff()
        End If
    End Sub

    Public nextItem As ToolStripItem
    Private Sub ProcessStrip_MouseWheel(sender As Object, e As MouseEventArgs) Handles ProcessStrip.MouseWheel
        Dim item As ToolStripItem = TryCast(ProcessStripSelectedItem, ToolStripItem)

        If item IsNot Nothing AndAlso item.Tag IsNot Nothing Then
            Dim handle As IntPtr = CType(item.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(handle)

            If My.Computer.Keyboard.CtrlKeyDown Then
                ' Change App volume
                Dim newVolumePercent As Integer = ChangeAppVolume(pr, e.Delta)

                If newVolumePercent >= 0 Then
                    Dim mousePos As Point = ProcessStrip.PointToClient(Cursor.Position)

                    mousePos.Offset(15, 15)

                    MainToolTip.Show($"Volume: {newVolumePercent} %", ProcessStrip, mousePos, 1000)
                End If
            Else
                ' Cycle thru taskbar
                Select Case currentDockPos
                    Case DockPosition.Top, DockPosition.Bottom
                        If e.Delta > 0 Then
                            nextItem = ProcessStrip.GetNextItem(nextItem, ArrowDirection.Right)
                        Else
                            nextItem = ProcessStrip.GetNextItem(nextItem, ArrowDirection.Left)
                        End If

                    Case DockPosition.Left, DockPosition.Right
                        If e.Delta > 0 Then
                            nextItem = ProcessStrip.GetNextItem(nextItem, ArrowDirection.Left)
                        Else
                            nextItem = ProcessStrip.GetNextItem(nextItem, ArrowDirection.Right)
                        End If
                End Select

                If nextItem IsNot Nothing Then
                    nextItem.Select()
                    If Not My.Computer.Keyboard.ShiftKeyDown Then nextItem.PerformClick()
                End If
            End If
        End If
    End Sub

    Public Function ChangeAppVolume(targetPid As Process, delta As Integer) As Integer
        If targetPid Is Nothing Then Return -1

        Dim enumerator As New MMDeviceEnumerator()
        Dim device = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia)
        Dim sessions = device.AudioSessionManager.Sessions

        Dim targetPidId As Integer = targetPid.Id
        Dim lastVolumePercent As Integer = -1

        For i As Integer = 0 To sessions.Count - 1
            Using session = sessions(i)
                Dim sessionPid As UInteger = 0

                Try
                    sessionPid = session.GetProcessID
                Catch ex As Exception
                    Continue For
                End Try

                If sessionPid = targetPidId OrElse IsChildProcess(targetPidId, CInt(sessionPid)) Then
                    Dim currentVol As Single = session.SimpleAudioVolume.Volume
                    Dim stepSize As Single = 0.05F ' 5 %

                    If delta > 0 Then
                        currentVol = Math.Min(1.0F, currentVol + stepSize)
                    Else
                        currentVol = Math.Max(0.0F, currentVol - stepSize)
                    End If

                    session.SimpleAudioVolume.Volume = currentVol
                    lastVolumePercent = CInt(Math.Round(currentVol * 100))
                End If
            End Using
        Next

        Return lastVolumePercent
    End Function

    Private Function IsChildProcess(parentPid As Integer, childPid As Integer) As Boolean
        Try
            Using child = Process.GetProcessById(childPid)
                Using parent = Process.GetProcessById(parentPid)
                    Return child.ProcessName.Equals(parent.ProcessName, StringComparison.OrdinalIgnoreCase)
                End Using
            End Using
        Catch
            Return False
        End Try
    End Function

    Private Sub UpdateTaskbarSelection(activeHwnd As IntPtr)
        If activeHwnd = Me.Handle Then Exit Sub
        If activeHwnd = Desktop.Handle Then Exit Sub
        'If activeHwnd = WAT.Handle Then Exit Sub
        If Startmenu.Visible = True Then
            If activeHwnd = Startmenu.Handle Then Exit Sub
        End If

        For Each item As ToolStripItem In ProcessStrip.Items
            If TypeOf item Is ToolStripButton Then
                Dim btn = DirectCast(item, ToolStripButton)

                If btn.Tag IsNot Nothing AndAlso DirectCast(btn.Tag, IntPtr) = activeHwnd Then
                    btn.Checked = True
                    nextItem = TryCast(btn, ToolStripItem)
                Else
                    btn.Checked = False
                End If
            End If
        Next
    End Sub

    Private Sub DefaultToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DefaultToolStripMenuItem.Click
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_SHOWDEFAULT)
        End If
    End Sub

    Private Sub MaximalizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MaximalizeToolStripMenuItem.Click
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_MAXIMIZE)
        End If

        'Dim App As Process = Process.GetProcessById(TPCM.Tag)
        'If App.Id > 0 Then
        '   AppActivate(App.Id)
        '   ShowWindow(App.MainWindowHandle, SHOW_WINDOW.SW_MAXIMIZE)
        'End If
    End Sub

    Private Sub NormalizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NormalizeToolStripMenuItem.Click
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_NORMAL)
        End If
    End Sub

    Private Sub MinimalizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MinimalizeToolStripMenuItem.Click
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_MINIMIZE)
        End If
    End Sub

    Private Sub ForceMinimalizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ForceMinimalizeToolStripMenuItem.Click
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_FORCEMINIMIZE)
        End If
    End Sub

    Private Sub HIDEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HIDEToolStripMenuItem.Click
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            If MsgBox("Are you sure do you want to hide this window process? [" & GetProcessFromWindowHandle(windowHandle).ProcessName & "]", MsgBoxStyle.YesNo, "Confirm Box") = MsgBoxResult.Yes Then
                ShowWindow(windowHandle, SHOW_WINDOW.SW_HIDE)
            End If
        End If
    End Sub

    Private Sub SwitchToToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SwitchToToolStripMenuItem.Click
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            SetForegroundWindow(windowHandle)
            'ShowWindow(windowHandle, SHOW_WINDOW.SW_SHOW)
            ActiveWindowHandle = windowHandle
        End If
    End Sub

    Private Sub SwitchButNotActiveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SwitchButNotActiveToolStripMenuItem.Click
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_SHOWNA)
        End If
    End Sub

    Private Const HWND_TOPMOST = -1 '-- Bring to top and stay there
    Private Const HWND_NOTOPMOST = -2 '-- Put the window into a normal position

    Const WM_CLOSE As Long = &H10
    Const WM_DESTROY As Long = &H2

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            SendMessage(windowHandle, WM_CLOSE, IntPtr.Zero, IntPtr.Zero)

            WAT.Hide()
        End If
    End Sub

    Private Sub StartSameProcessToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartSameProcessToolStripMenuItem.Click
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Try
                Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
                Process.Start(GetProcessFromWindowHandle(windowHandle).MainModule.FileName)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub TPCM_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cm_ProcessItem.Opening
        If cm_Main.Visible Then cm_Main.Close()

        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

            Try
                Dim fvi As FileVersionInfo = FileVersionInfo.GetVersionInfo(pr.MainModule.FileName)
                StartSameProcessToolStripMenuItem.Text = fvi.FileDescription
            Catch
                StartSameProcessToolStripMenuItem.Text = "Start Same program"
            End Try

            Try
                Dim ico As Icon = Icon.ExtractAssociatedIcon(pr.MainModule.FileName)
                StartSameProcessToolStripMenuItem.Image = ico.ToBitmap
            Catch
                StartSameProcessToolStripMenuItem.Image = My.Resources.ProgramMedium
            End Try

            For Each item As ToolStripMenuItem In SetPriorityLevelToolStripMenuItem.DropDownItems
                item.Checked = False
            Next

            Try
                Select Case pr.PriorityClass
                    Case ProcessPriorityClass.Idle
                        PRprcmb6.Checked = True
                        SetPriorityLevelToolStripMenuItem.Enabled = True

                    Case ProcessPriorityClass.BelowNormal
                        PRprcmb5.Checked = True
                        SetPriorityLevelToolStripMenuItem.Enabled = True

                    Case ProcessPriorityClass.Normal
                        PRprcmb4.Checked = True
                        SetPriorityLevelToolStripMenuItem.Enabled = True

                    Case ProcessPriorityClass.AboveNormal
                        PRprcmb3.Checked = True
                        SetPriorityLevelToolStripMenuItem.Enabled = True

                    Case ProcessPriorityClass.High
                        PRprcmb2.Checked = True
                        SetPriorityLevelToolStripMenuItem.Enabled = True

                    Case ProcessPriorityClass.RealTime
                        PRprcmb1.Checked = True
                        SetPriorityLevelToolStripMenuItem.Enabled = True

                    Case Else
                        SetPriorityLevelToolStripMenuItem.Enabled = False
                End Select
            Catch
                SetPriorityLevelToolStripMenuItem.Enabled = False
            End Try

            ' Working directory currently unavalible.
            OpenProgramsWorkingDirectoryToolStripMenuItem.Enabled = False

            Dim endTaskEnabled As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings", "TaskbarEndTask", False)
            EndTaskToolStripMenuItem.Visible = endTaskEnabled
        End If
    End Sub

    Private Sub PRprcmb1_Click(sender As Object, e As EventArgs) Handles PRprcmb1.Click
        On Error Resume Next
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            GetProcessFromWindowHandle(windowHandle).PriorityClass = ProcessPriorityClass.RealTime
        End If
    End Sub

    Private Sub PRprcmb2_Click(sender As Object, e As EventArgs) Handles PRprcmb2.Click
        On Error Resume Next
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            GetProcessFromWindowHandle(windowHandle).PriorityClass = ProcessPriorityClass.High
        End If
    End Sub

    Private Sub PRprcmb3_Click(sender As Object, e As EventArgs) Handles PRprcmb3.Click
        On Error Resume Next
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            GetProcessFromWindowHandle(windowHandle).PriorityClass = ProcessPriorityClass.AboveNormal
        End If
    End Sub

    Private Sub PRprcmb4_Click(sender As Object, e As EventArgs) Handles PRprcmb4.Click
        On Error Resume Next
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            GetProcessFromWindowHandle(windowHandle).PriorityClass = ProcessPriorityClass.Normal
        End If
    End Sub

    Private Sub PRprcmb5_Click(sender As Object, e As EventArgs) Handles PRprcmb5.Click
        On Error Resume Next
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            GetProcessFromWindowHandle(windowHandle).PriorityClass = ProcessPriorityClass.BelowNormal
        End If
    End Sub

    Private Sub PRprcmb6_Click(sender As Object, e As EventArgs) Handles PRprcmb6.Click
        On Error Resume Next
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            GetProcessFromWindowHandle(windowHandle).PriorityClass = ProcessPriorityClass.Idle
        End If
    End Sub

    Public Sub StartMenuShowAndHide()
        Select Case currentDockPos
            Case DockPosition.Bottom
                Startmenu.StartDefaultLocation = New Point(0 - SystemInformation.BorderSize.Width, SystemInformation.PrimaryMonitorSize.Height - Me.Height - Startmenu.Height + SystemInformation.BorderSize.Height)
            Case DockPosition.Left
                Startmenu.StartDefaultLocation = New Point(Me.Width - SystemInformation.BorderSize.Width, 0 - SystemInformation.BorderSize.Height)
            Case DockPosition.Top
                Startmenu.StartDefaultLocation = New Point(0 - SystemInformation.BorderSize.Width, Me.Height)
            Case DockPosition.Right
                Startmenu.StartDefaultLocation = New Point(SystemInformation.PrimaryMonitorSize.Width - Me.Width - Startmenu.Width + SystemInformation.BorderSize.Width, 0 - SystemInformation.BorderSize.Height)
        End Select

        Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0")
            Case 0 'Normal Start menu look
                If Startmenu.Visible = True Then
                    Startmenu.Hide()
                Else
                    Startmenu.Show()
                End If

                UpdateStartButtonImage()

            Case 1 'Custom Images start menu look
                If Startmenu.Visible = True Then
                    Startmenu.Hide()

                    Try
                        Start.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Normal", ""))
                    Catch ex As Exception
                        Start.BackgroundImage = My.Resources.StartRight
                    End Try
                Else
                    Startmenu.Show()

                    Try
                        Start.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Pressed", ""))
                    Catch ex As Exception
                        Start.BackgroundImage = My.Resources.StartLeft
                    End Try
                End If

            Case 2 'ORB Start menu look
                If Startmenu.Visible = True Then
                    Startmenu.Hide()

                    Try
                        Start.BackgroundImage = OrbNormal
                    Catch ex As Exception
                        Start.BackgroundImage = My.Resources.StartRight
                    End Try
                Else
                    Startmenu.Show()

                    Try
                        Start.BackgroundImage = OrbPressed
                    Catch ex As Exception
                        Start.BackgroundImage = My.Resources.StartLeft
                    End Try
                End If

            Case Else
                If Startmenu.Visible = True Then
                    Startmenu.Hide()
                Else
                    Startmenu.Show()
                End If

                UpdateStartButtonImage()
        End Select
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Start.Click
        If Startmenu.Visible = True Then
            Startmenu.FST = False
            StartMenuShowAndHide()
            UpdateStartButtonImage()
        Else
            Startmenu.FST = True
            StartMenuShowAndHide()
            UpdateStartButtonImage(True)
        End If
    End Sub
    Public UseLegacyVolumeSlider As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Volume", "UseLegacy", False)
    Public CanClose As Boolean = False
    Private Sub AppBar_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Dim success As Boolean = False
        Dim workingAreaReset As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "ResetAfterShutdown", True)
        Dim workspace As New WorkspaceManager

        If e.CloseReason = CloseReason.TaskManagerClosing OrElse e.CloseReason = CloseReason.WindowsShutDown OrElse CanClose = True Then success = True

        If success Then
            e.Cancel = False
            If GlobalKeyboardHook.Unhook() Then
                Debug.WriteLine("Hook ended.")
            Else
                Debug.WriteLine("Hook ending failed.")
            End If

            If workingAreaReset Then workspace.UpdateWorkingArea(0, 0, 0, 0)
            If hHook <> IntPtr.Zero Then UnhookWinEvent(hHook)

            SetShellWindow(IntPtr.Zero)
            SetProgmanWindow(IntPtr.Zero)
        Else
            e.Cancel = True
            SA.ShowDialog()
        End If
    End Sub

    Private Sub LockAppbarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LockAppbarToolStripMenuItem.Click
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Locked", LockAppbarToolStripMenuItem.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        If LockAppbarToolStripMenuItem.Checked = True Then
            Splitter1.Visible = False
            Splitter2.Visible = False
            Splitter3.Visible = False
        Else
            Splitter1.Visible = True
            Splitter2.Visible = True
            Splitter3.Visible = True
        End If
    End Sub

    Private Sub TaskManagerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TaskManagerToolStripMenuItem.Click
        On Error Resume Next
        Process.Start(New ProcessStartInfo(EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "Taskmgr", Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\taskmgr.exe"))) With {.UseShellExecute = True})
    End Sub

    <DllImport("user32.dll")>
    Private Shared Function IsHungAppWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    Private IsDragging As Boolean = False
    Private DraggedItem As ToolStripItem = Nothing
    Private OriginalIndex As Integer = -1

    Public Sub LoadApps()
        If Me.InvokeRequired Then
            Me.Invoke(Sub() LoadApps())
            Exit Sub
        End If

        For Each hwnd In AlertTimers.Keys.ToList()
            StopAlert(hwnd)
        Next
        AlertTimers.Clear()

        For i As Integer = ProcessStrip.Items.Count - 1 To 0 Step -1
            Dim item = ProcessStrip.Items(i)
            If item.Image IsNot Nothing Then item.Image.Dispose()
            item.Dispose()
        Next
        ProcessStrip.Items.Clear()

        LoadBlacklistFromRegistry()
        LoadBlocklistFromRegistry()

        EnumWindows(AddressOf EnumWindowsCallback, IntPtr.Zero)
    End Sub

    Private Sub StartButtonToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartButtonToolStripMenuItem.Click
        Start.Visible = StartButtonToolStripMenuItem.Checked
        If LockAppbarToolStripMenuItem.Checked = False Then Splitter1.Visible = StartButtonToolStripMenuItem.Checked

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "StartButton", StartButtonToolStripMenuItem.Checked, RegistryValueKind.DWord)
    End Sub

    Private Sub PinnedBarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PinnedBarToolStripMenuItem.Click
        Panel1.Visible = PinnedBarToolStripMenuItem.Checked
        If LockAppbarToolStripMenuItem.Checked = False Then Splitter2.Visible = PinnedBarToolStripMenuItem.Checked

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "PinnedBar", PinnedBarToolStripMenuItem.Checked, RegistryValueKind.DWord)
    End Sub

    Private Sub MenuBarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MenuBarToolStripMenuItem.Click
        Panel4.Visible = MenuBarToolStripMenuItem.Checked
        If LockAppbarToolStripMenuItem.Checked = False Then Splitter3.Visible = MenuBarToolStripMenuItem.Checked

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "MenuBar", MenuBarToolStripMenuItem.Checked, RegistryValueKind.DWord)
    End Sub

    Public Sub AutoHideOnOff()
        If AutoHideToolStripMenuItem.Checked = False Then
            Hidden = False
            Exit Sub
        End If

        fBarRegistered = AutoHideToolStripMenuItem.Checked
        RegisterBar()
        If AutoHideToolStripMenuItem.Checked = True Then
            Hidden = True
            'Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - 2)
        Else
            Hidden = False
            'Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - Me.Height)
        End If
    End Sub

    Private Sub AutoHideToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutoHideToolStripMenuItem.Click
        AutoHideOnOff()
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "AutoHide", AutoHideToolStripMenuItem.Checked, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub
    Public Property Hidden As Boolean = False
    Private Sub AppBar_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
        'If AutoHideToolStripMenuItem.Checked = True Then
        'Else
        Select Case currentDockPos
            Case DockPosition.Bottom
                If Not Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - If(isSmallIcons, defSmallSize, defSize)) Then
                    Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - If(isSmallIcons, defSmallSize, defSize))
                End If

            Case DockPosition.Top, DockPosition.Left
                If Not Me.Location = New Point(0, 0) Then
                    Me.Location = New Point(0, 0)
                End If

            Case DockPosition.Bottom, DockPosition.Top
                If isSmallIcons Then
                    If Not Me.Height = defSmallSize Then
                        Me.Height = defSmallSize
                    End If
                Else
                    If Not Me.Height = defSize Then
                        Me.Height = defSize
                    End If
                End If

            Case DockPosition.Right
                If Not Me.Location = New Point(SystemInformation.PrimaryMonitorSize.Width - If(isSmallIcons, defSmallSize, defSize), 0) Then
                    Me.Location = New Point(SystemInformation.PrimaryMonitorSize.Width - If(isSmallIcons, defSmallSize, defSize), 0)
                End If

        End Select
        'End If
    End Sub

    Private Sub Internal_MouseEnter(sender As Object, e As EventArgs) Handles ToolStrip1.MouseEnter, Start.MouseEnter, AppStrip.MouseEnter, TimeLabel.MouseEnter, btn_LanguageBar.MouseEnter
        If WAT.Visible AndAlso Not WAT.HideFadeTimer.Enabled Then WAT.Hide()
    End Sub

    Private Sub ToolStrip1_MouseEnter(sender As Object, e As EventArgs) Handles ToolStrip1.MouseEnter ', ProcessStrip.MouseEnter, AppStrip.MouseEnter, TimeLabel.MouseEnter, btn_LanguageBar.MouseEnter
        If Startmenu.Visible Then
            Startmenu.Focus()
            Exit Sub
        Else
            ToolStrip1.Focus()
            ToolStrip1.Select()

            Hidden = False

            Startmenu.FST = False
            AutoHideOnOff()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btn_ShowDesktop.Click, ShowDesktopToolStripMenuItem.Click, ShowDesktopToolStripMenuItem1.Click
        Desktop.ShowDesktop()
    End Sub

    Private Sub PropertiesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PropertiesToolStripMenuItem.Click
        On Error Resume Next
        AppbarProperties.Show(Me)
    End Sub

    Public WithEvents BlockingProcesses As System.Management.ManagementEventWatcher

    Public Sub ToggleProcessMonitoring(ByVal enable As Boolean)
        If enable Then
            If BlockingProcesses Is Nothing Then
                Try
                    Dim blockedNames As New List(Of String)
                    Using registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\BlockedProcesses\.\")
                        If registryKey IsNot Nothing Then
                            For Each N In registryKey.GetValueNames()
                                Dim exeName As String = If(N.EndsWith(".exe", StringComparison.OrdinalIgnoreCase), Name, Name & ".exe")
                                blockedNames.Add($"TargetInstance.Name='{exeName}'")
                            Next
                        End If
                    End Using

                    If blockedNames.Count > 0 Then
                        Dim wqlCondition As String = String.Join(" OR ", blockedNames)
                        Dim queryString As String = $"SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process' AND ({wqlCondition})"

                        Dim query As New System.Management.EventQuery(queryString)
                        BlockingProcesses = New System.Management.ManagementEventWatcher(query)
                        AddHandler BlockingProcesses.EventArrived, AddressOf processWatcher_EventArrived
                        BlockingProcesses.Start()
                    Else
                        BlockingProcesses = Nothing
                    End If
                Catch
                    BlockingProcesses = Nothing
                End Try
            End If
        Else
            If BlockingProcesses IsNot Nothing Then
                Try
                    BlockingProcesses.Stop()
                    RemoveHandler BlockingProcesses.EventArrived, AddressOf processWatcher_EventArrived
                    BlockingProcesses.Dispose()
                Catch
                Finally
                    BlockingProcesses = Nothing
                End Try
            End If
        End If
    End Sub

    Private Sub processWatcher_EventArrived(sender As Object, e As EventArrivedEventArgs)
        Dim targetInstance As ManagementBaseObject = CType(e.NewEvent("TargetInstance"), ManagementBaseObject)
        Dim processName As String = Path.GetFileNameWithoutExtension(targetInstance("Name").ToString())
        Dim processId As Integer = Convert.ToInt32(targetInstance("ProcessId"))

        CheckAndTerminateProcess(processName, processId)
    End Sub

    Private Sub CheckAndTerminateProcess(processName As String, processId As Integer)
        If [Enum].GetNames(GetType(CriticalProcesses)).Contains(processName, StringComparer.OrdinalIgnoreCase) Then Return

        Using registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\BlockedProcesses\.\")
            If registryKey IsNot Nothing AndAlso registryKey.GetValueNames().Contains(processName) Then
                Try
                    Dim pr As Process = Process.GetProcessById(processId)
                    Dim killMethod As Object = Microsoft.Win32.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "KillMethod", "0")

                    If killMethod IsNot Nothing AndAlso killMethod.ToString() = "0" Then
                        pr.Kill()
                    Else
                        pr.CloseMainWindow()
                    End If
                Catch
                End Try
            End If
        End Using
    End Sub

    Private Sub CMMAIN_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cm_Main.Opening
        Dim item As ToolStripButton = ProcessStrip.GetItemAt(ProcessStrip.PointToClient(MousePosition))
        If item IsNot Nothing AndAlso TypeOf item.Tag Is IntPtr Then
            e.Cancel = True

            If My.Computer.Keyboard.CtrlKeyDown Then
                Dim windowHandle As IntPtr = item.Tag

                WAT.Hide()

                cm_ProcessItem.Tag = windowHandle

                Try
                    Dim pt As Point = Cursor.Position
                    Dim lParam As New IntPtr((pt.Y << 16) Or (pt.X And &HFFFF))

                    SetForegroundWindow(windowHandle)

                    PostMessage(windowHandle, WM_POPUPSYSTEMMENU, IntPtr.Zero, lParam)
                Catch ex As Exception
                    WindowStateToolStripMenuItem.DropDown.Show(Cursor.Position)
                End Try
            Else
                cm_ProcessItem.Tag = item.Tag
                cm_ProcessItem.Show(MousePosition)
            End If
        End If

        cm_Main.RightToLeft = RightToLeft.No

        SmallIconsToolStripMenuItem.Checked = isSmallIcons

        If My.Computer.Keyboard.CtrlKeyDown AndAlso My.Computer.Keyboard.ShiftKeyDown Then
            ToolStripSeparator5.Visible = True
            RestartShellToolStripMenuItem.Visible = True
            EndShellToolStripMenuItem.Visible = True
        Else
            ToolStripSeparator5.Visible = False
            RestartShellToolStripMenuItem.Visible = False
            EndShellToolStripMenuItem.Visible = False
        End If
    End Sub

    Private Sub TimeAndDateSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TimeAndDateSettingsToolStripMenuItem.Click
        On Error Resume Next
        Process.Start(New ProcessStartInfo(EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "TimeDateProperties", Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\timedate.cpl"))) With {.UseShellExecute = True})
    End Sub

    Private Sub DisableAlarmOption_Click(sender As Object, e As EventArgs) Handles DisableAlarmOption.Click
        If DisableAlarmOption.Checked = True Then
            AlarmController.Enabled = False
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\", "DisableAlarms", "1", Microsoft.Win32.RegistryValueKind.DWord)
        Else
            AlarmController.Enabled = True
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\", "DisableAlarms", "0", Microsoft.Win32.RegistryValueKind.DWord)
        End If
    End Sub

    Private Sub ClearToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearToolStripMenuItem.Click
        If MessageBox.Show("Do you really want to clear the Clipboard data?", "Microsoft Windows", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            Clipboard.Clear()
        End If
    End Sub

    Private Sub DaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DaToolStripMenuItem.Click
        ClipboardViewer.Show()
    End Sub

    Private Sub ActionCenterToolStripMenuItem_Click(sender As Object, e As EventArgs)
        On Error Resume Next
        Process.Start(CType(sender, ToolStripMenuItem).Tag)
    End Sub

    Private Sub ActionCM_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cm_Action.Opening
        cm_Action.Items.Clear()

        Try
            Dim basePath As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\Windows\WinX\Group")

            For i As Integer = 1 To 10
                Dim currentGroupPath As String = basePath & i

                If IO.Directory.Exists(currentGroupPath) Then
                    Dim files As String() = IO.Directory.GetFiles(currentGroupPath)
                    Dim addedAny As Boolean = False

                    For Each filePath As String In files
                        Dim fileInfo As New IO.FileInfo(filePath)

                        If fileInfo.Name.ToLower() <> "desktop.ini" Then
                            Dim item As ToolStripMenuItem = CreateWinXMenuItem(fileInfo)
                            cm_Action.Items.Add(item)
                            addedAny = True
                        End If
                    Next

                    If addedAny Then
                        cm_Action.Items.Add(New ToolStripSeparator())
                    End If
                Else
                    Exit For
                End If
            Next

            cm_Action.Items.Add(ClipboardToolStripMenuItem)
            cm_Action.Items.Add(PerformanceToolStripMenuItem)
            cm_Action.Items.Add(New ToolStripSeparator)
            cm_Action.Items.Add(LogonToolStripMenuItem)
            cm_Action.Items.Add(PowerToolStripMenuItem)

            ' Loads all Power options
            RefreshPowerMenus()

        Catch ex As Exception
        End Try
    End Sub

    Private Function CreateWinXMenuItem(fileInfo As IO.FileInfo) As ToolStripMenuItem
        Dim cleanName As String = IO.Path.GetFileNameWithoutExtension(fileInfo.FullName).Trim()

        If Char.IsNumber(cleanName(0)) Then
            Dim firstSpaceIndex As Integer = cleanName.IndexOf(" ")
            If firstSpaceIndex <> -1 AndAlso cleanName.Length > firstSpaceIndex + 3 Then
                cleanName = cleanName.Substring(firstSpaceIndex + 3).TrimStart()
            End If
        End If

        Dim item As New ToolStripMenuItem(cleanName)
        item.Tag = fileInfo.FullName

        Try
            Dim ico As Icon = GetFileIcon(fileInfo.FullName, False)
            If ico IsNot Nothing Then
                item.Image = ico.ToBitmap()
            End If
        Catch
        End Try

        AddHandler item.Click, AddressOf ActionCenterToolStripMenuItem_Click
        Return item
    End Function

    Public primaryForeColor As Color = Color.Black
    Private Sub AppBar_BackColorChanged(sender As Object, e As EventArgs) Handles Me.BackColorChanged
        ' Save current BackColor
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", ColorTranslator.ToHtml(Me.BackColor), Microsoft.Win32.RegistryValueKind.String)

        ' Calculate text contrast
        Dim luminance As Double = (0.299 * Me.BackColor.R + 0.587 * Me.BackColor.G + 0.114 * Me.BackColor.B) / 255

        Dim contrastColor As Color = If(luminance > 0.5, Color.Black, Color.White)
        primaryForeColor = contrastColor

        For Each ctrl As Control In Me.Controls
            ctrl.ForeColor = contrastColor
        Next

        TimeLabel.ForeColor = primaryForeColor
        btn_LanguageBar.ForeColor = primaryForeColor
    End Sub

    Private Sub KillToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EndTaskToolStripMenuItem.Click
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

            Try
                pr.Kill()

                WAT.Hide()
            Catch ex As Exception
                MsgBox("Unable to kill the process. " & ex.Message, MsgBoxStyle.Critical, "Error")
            End Try
        End If
    End Sub

    Private Sub DestroyProgramToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DestroyWindowToolStripMenuItem.Click
        If MsgBox("!!!WARNING!!!" & Environment.NewLine & """Destroy a Window"" means, that everything that the selected Window has, including data on your memory will be ""destroyed"" which I don't recommend. Well this warning is to prevent it happening one time. Do you really want to continue?", MsgBoxStyle.YesNo, "Warning") = MsgBoxResult.Yes Then
            If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
                Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
                SendMessage(windowHandle, WM_DESTROY, IntPtr.Zero, IntPtr.Zero)
            End If
        End If
    End Sub

    Private Sub OpenProgramsLocationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenProgramsLocationToolStripMenuItem.Click
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

            If IO.File.Exists(pr.MainModule.FileName) = True Then
                Dim fi As New IO.FileInfo(pr.MainModule.FileName)

                OpenDestination(fi.DirectoryName)
            End If
        End If
    End Sub

    Public Sub OpenDestination(ByVal afterText)
        Try
            If UseExplorerFM = True Then
                Process.Start("explorer.exe", afterText)
            Else
                If Not String.IsNullOrWhiteSpace(CustomFMPath) Then
                    Process.Start(CustomFMPath, """" & afterText & """")
                Else
                    MsgBox("Failed to execute a custom File Manager. Path to it is empty", MsgBoxStyle.Critical)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub MoreOptionsDropDownLol(sender As Object, e As EventArgs) Handles MoreOptionsToolStripMenuItem.DropDownOpening
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

            If IO.Directory.Exists(pr.StartInfo.WorkingDirectory) = False Then
                OpenProgramsWorkingDirectoryToolStripMenuItem.Enabled = False
                OpenProgramsWorkingDirectoryToolStripMenuItem.ToolTipText = "This process hasn't associated any custom Working directory."
            Else
                OpenProgramsWorkingDirectoryToolStripMenuItem.Enabled = True
                OpenProgramsWorkingDirectoryToolStripMenuItem.ToolTipText = ""
            End If

            If IsWindowAlwaysOnTop(windowHandle) Then
                AlwaysOnTopToolStripMenuItem.Checked = True
            Else
                AlwaysOnTopToolStripMenuItem.Checked = False
            End If

            FullscreenWin32ToolStripMenuItem.Checked = WindowFullscreenHelper.IsAppFullscreen(windowHandle)

            BlockProcessToolStripMenuItem.Enabled = BlockingProcesses IsNot Nothing
        End If
    End Sub

    Private Sub OpenProgramsWorkingDirectoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenProgramsWorkingDirectoryToolStripMenuItem.Click
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

            If IO.Directory.Exists(pr.StartInfo.WorkingDirectory) = True Then
                Dim fi As New IO.DirectoryInfo(pr.StartInfo.WorkingDirectory)

                OpenDestination(fi.FullName)
            End If
        End If
    End Sub

    Private Sub AlwaysOnTopToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AlwaysOnTopToolStripMenuItem.Click
        On Error Resume Next

        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim hWnd As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(hWnd)

            If IsWindowAlwaysOnTop(hWnd) Then
                SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
            Else
                SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
            End If
        End If
    End Sub

    Private Sub HideProcessFromAppbarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HideProcessFromAppbarToolStripMenuItem.Click
        On Error Resume Next
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\HiddenProcesses\", pr.ProcessName, "")

            'LoadApps()
            MainToolTip.Show("The """ & Process.GetProcessById(cm_ProcessItem.Tag).ProcessName & """ has been successfully set to hidden from Appbar!", Me, Me.Width - ToolStrip1.Width / 2, Me.Location.Y + ToolStrip1.Height + 150, 5000)
        End If
    End Sub

    Private Sub BlockProcessToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BlockProcessToolStripMenuItem.Click
        If MessageBox.Show("Are you sure do you want to block [" & Process.GetProcessById(cm_ProcessItem.Tag).ProcessName & "] process?" & Environment.NewLine & "This means, when this shell will be running, then that process you'll set as ""Blocked"" will no longer be working. To allow the process again: Right click on the Appbar, select properties and remove the app from Blocklist.", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
                Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
                Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses\.\", pr.ProcessName, "")

                LoadApps()
                MainToolTip.Show("The """ & Process.GetProcessById(cm_ProcessItem.Tag).ProcessName & """ process, has been successfully blocked!", Me, Me.Width - ToolStrip1.Width / 2, Me.Location.Y + ToolStrip1.Height + 150, 5000)
            End If
        End If
    End Sub

    Private Sub VolumeControlToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VolumeControlToolStripMenuItem.Click
        VolumeControl.Show(Me)
    End Sub
    Public soundcontrolPath = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\SndVol.exe"
    Private Async Sub SoundControlToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SoundControlToolStripMenuItem.Click
        If Not File.Exists(soundcontrolPath) Then Exit Sub
        If Not New FileInfo(soundcontrolPath).Extension.ToLower = ".exe" Then Exit Sub

        Try : Dim p As Process = Process.Start(soundcontrolPath)

            Dim hwnd As IntPtr = IntPtr.Zero
            Dim attempts As Integer = 0

            While hwnd = IntPtr.Zero AndAlso attempts < 20
                Await Task.Delay(50)
                p.Refresh()
                hwnd = p.MainWindowHandle
                attempts += 1
            End While

            If hwnd <> IntPtr.Zero Then
                'Dim rect As New RECT()
                'GetWindowRect(hwnd, rect)

                'Dim width As Integer = rect.right - rect.left
                'Dim height As Integer = rect.bottom - rect.top

                SetWindowPos(hwnd, HWND_NOTOPMOST, SystemInformation.WorkingArea.Width - 440, SystemInformation.WorkingArea.Height - 340, 450, 400, SWP_NOZORDER Or SWP_SHOWWINDOW)
            End If
        Catch
        End Try
    End Sub

    Private Sub AboutShellToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles tsmi_AboutShell.Click
        If AboutDialog.Visible = True Then
            AboutDialog.Focus()
            AboutDialog.BringToFront()
        Else
            AboutDialog.Show()
        End If
    End Sub

    Private Sub ClipboardViewerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClipboardViewerToolStripMenuItem.Click
        If ClipboardViewer.Visible = True Then
            ClipboardViewer.BringToFront()
            ClipboardViewer.Focus()
        Else
            ClipboardViewer.Show()
        End If
    End Sub
    Public PinnedAppsDir As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Internet Explorer\Quick Launch"
    Public Sub LoadPinnedApps()
        AppStrip.Items.Clear()

        'i.Substring(i.LastIndexOf("\") + 1)
        Try
            For Each i As String In IO.Directory.GetFiles(PinnedAppsDir)

                Dim FI As New FileInfo(i)

                If Not String.Equals(FI.Name.ToLower, "desktop.ini") Then

                    Dim ico As Icon = GetFileIcon(FI.FullName, False)

                    Dim Item As New ToolStripButton
                    With Item

                        Item.DisplayStyle = ToolStripItemDisplayStyle.Image
                        Item.Text = Path.GetFileNameWithoutExtension(FI.FullName)
                        Item.Tag = FI.FullName

                        If ico IsNot Nothing Then
                            Item.Image = ico.ToBitmap
                        Else
                            Item.Image = My.Resources.ProgramSmall
                        End If

                        AddHandler Item.MouseUp, AddressOf PinnedAppMouseUp
                    End With

                    AppStrip.Items.Add(Item)
                End If
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PinToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PinToolStripMenuItem.Click
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)
            Dim fi As New FileInfo(pr.MainModule.FileName)

            Try
                Dim wshShell As Object = CreateObject("WScript.Shell")

                Dim shortcut As Object = wshShell.CreateShortcut(PinnedAppsDir & "\" & fi.Name & ".lnk")

                shortcut.TargetPath = pr.MainModule.FileName

                shortcut.Save()

                System.Runtime.InteropServices.Marshal.ReleaseComObject(shortcut)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wshShell)

                shortcut = Nothing
                wshShell = Nothing

                LoadPinnedApps()

            Catch uex As UnauthorizedAccessException
                MsgBox($"Failed to pin ""{pr.ProcessName}""! Because it has no access to the Taskbar's directory where the pinned apps are stored. You can have this shell running as Administrator and this error will be not shown.", MsgBoxStyle.Critical, "Unauthorized Access")
                LoadPinnedApps()

            Catch ex As Exception
                MsgBox(ex.Message)
                LoadPinnedApps()

            End Try
        End If

    End Sub

    Dim currentPinnedItem As ToolStripButton = Nothing
    Private Sub PinnedAppMouseUp(ByVal sender As Object, e As MouseEventArgs)
        Dim item As ToolStripButton = CType(sender, ToolStripButton)

        If File.Exists(item.Tag) Then
            Select Case e.Button
                Case MouseButtons.Left
                    Dim startInfo As New ProcessStartInfo()

                    startInfo.FileName = item.Tag
                    startInfo.UseShellExecute = True

                    If My.Computer.Keyboard.CtrlKeyDown AndAlso My.Computer.Keyboard.ShiftKeyDown Then
                        startInfo.Verb = "runas"
                    End If

                    Process.Start(startInfo)

                Case MouseButtons.Right
                    cm_AppStrip.Tag = item.Tag

                    currentPinnedItem = item
                    item.Checked = True
                    cm_AppStrip.Show(MousePosition)

                    'Desktop.ShowShellContextMenu(item.Tag, MousePosition)
            End Select
        End If
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        On Error Resume Next
        Process.Start(cm_AppStrip.Tag)
    End Sub

    Private Sub OpenFileLocationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenFileLocationToolStripMenuItem.Click
        On Error Resume Next
        Dim fi As New FileInfo(cm_AppStrip.Tag)

        If fi.Extension.ToLower = ".lnk" Then

            Dim targetPath As String = GetShortcutTarget(fi.FullName)

            If File.Exists(targetPath) Then
                Dim linkfi As New FileInfo(targetPath)
                OpenDestination(linkfi.DirectoryName)
            Else
                OpenDestination(fi.DirectoryName)
            End If
        Else
            OpenDestination(fi.DirectoryName)
        End If
    End Sub

    Public Function GetShortcutTarget(shortcutPath As String) As String
        shortcutPath = EncodeFilePath(shortcutPath)

        Try
            If Not File.Exists(shortcutPath) Then Return String.Empty

            Dim fileBytes As Byte() = File.ReadAllBytes(shortcutPath)

            If fileBytes.Length < &H4C OrElse fileBytes(0) <> &H4C Then Return String.Empty

            Dim flags As UInteger = BitConverter.ToUInt32(fileBytes, &H14)
            If (flags And &H2) = 0 Then Return String.Empty ' HasLinkInfo flag

            Dim offset As Integer = &H4C
            If (flags And &H1) <> 0 Then ' HasLinkTargetIDList flag
                Dim idListSize As Short = BitConverter.ToInt16(fileBytes, offset)
                offset += idListSize + 2
            End If

            Dim localBaseNameOffset As Integer = BitConverter.ToInt32(fileBytes, offset + &H10)
            Dim finalOffset As Integer = offset + localBaseNameOffset

            Dim pathBuilder As New StringBuilder()
            While finalOffset < fileBytes.Length AndAlso fileBytes(finalOffset) <> 0
                pathBuilder.Append(Chr(fileBytes(finalOffset)))
                finalOffset += 1
            End While

            Return pathBuilder.ToString()
        Catch
            Return String.Empty
        End Try
    End Function

    Private Sub RemoveEntirelyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemoveEntirelyToolStripMenuItem.Click
        On Error Resume Next
        If IO.File.Exists(cm_AppStrip.Tag) Then
            Dim fi As New FileInfo(cm_AppStrip.Tag)
            If MessageBox.Show("Are you sure do you want remove """ & fi.Name & """ Item from Pinned Bar?", "Remove Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                My.Computer.FileSystem.DeleteFile(cm_AppStrip.Tag, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)

                LoadPinnedApps()
            End If
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        On Error Resume Next
        OpenDestination(PinnedAppsDir)
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        Dim ofd As New OpenFileDialog With {
            .Title = "Select a program to be Pinned into Pinned Bar...",
            .Multiselect = False,
            .AutoUpgradeEnabled = UseExplorerFP,
            .CheckFileExists = True,
            .SupportMultiDottedExtensions = True,
            .Filter = "Programs (*.exe;*.pif;*.bat;*.cmd)|*.exe;*.pif;*.bat;*.cmd|All files (*.*)|*.*"
        }

        If ofd.ShowDialog = DialogResult.OK Then
            Dim fi As New FileInfo(ofd.FileName)

            Try
                Dim wshShell As Object = CreateObject("WScript.Shell")

                Dim shortcut As Object = wshShell.CreateShortcut(PinnedAppsDir & "\" & fi.Name & ".lnk")

                shortcut.TargetPath = fi.FullName

                shortcut.Save()

                System.Runtime.InteropServices.Marshal.ReleaseComObject(shortcut)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wshShell)

                shortcut = Nothing
                wshShell = Nothing

                LoadPinnedApps()

            Catch uex As UnauthorizedAccessException
                MsgBox($"Failed to pin ""{ofd.FileName}""! Because it has no access to the Taskbar's directory where the pinned apps are stored. You can have this shell running as Administrator and this error will be not shown.", MsgBoxStyle.Critical, "Unauthorized Access")
                LoadPinnedApps()

            Catch ex As Exception
                MsgBox(ex.Message)
                LoadPinnedApps()

            End Try

            'Shell("mklink """ & Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Internet Explorer\Quick Launch\" & fi.Name & """ """ & fi.FullName & """", AppWinStyle.NormalFocus, True)

        End If
    End Sub

    Public OrbNormal As Bitmap = Nothing
    Public OrbHover As Bitmap = Nothing
    Public OrbPressed As Bitmap = Nothing
    Public Sub LoadAndSplitOrb(filePath As String)
        filePath = EncodeFilePath(filePath)

        Try
            Dim sourceImage As New Bitmap(filePath)
            Dim partHeight As Integer = sourceImage.Height \ 3
            Dim partWidth As Integer = sourceImage.Width

            For i As Integer = 0 To 2

                Dim rect As New Rectangle(0, i * partHeight, partWidth, partHeight)
                Dim part As New Bitmap(partWidth, partHeight)

                Using g As Graphics = Graphics.FromImage(part)
                    g.DrawImage(sourceImage, New Rectangle(0, 0, partWidth, partHeight), rect, GraphicsUnit.Pixel)
                End Using

                Select Case i
                    Case 0
                        If OrbNormal IsNot Nothing Then OrbNormal.Dispose()

                        OrbNormal = part
                    Case 1
                        If OrbHover IsNot Nothing Then OrbHover.Dispose()

                        OrbHover = part
                    Case 2
                        If OrbPressed IsNot Nothing Then OrbPressed.Dispose()

                        OrbPressed = part
                End Select
            Next

            sourceImage.Dispose()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub AppBar_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged, Me.Resize
        'If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "SpaceForEmergeTray", "0") = 0 Then
        Select Case currentDockPos
            Case DockPosition.Top, DockPosition.Bottom

                If isSmallIcons Then
                    If Not Me.Height = defSmallSize Then
                        Me.Height = defSmallSize
                    End If
                Else
                    If Not Me.Height = defSize Then
                        Me.Height = defSize
                    End If
                End If

                If Not Me.Width = SystemInformation.PrimaryMonitorSize.Width Then
                    Me.Width = SystemInformation.PrimaryMonitorSize.Width

                    AppStrip.LayoutStyle = ToolStripLayoutStyle.HorizontalStackWithOverflow
                    ProcessStrip.LayoutStyle = ToolStripLayoutStyle.HorizontalStackWithOverflow
                    ToolStrip1.LayoutStyle = ToolStripLayoutStyle.HorizontalStackWithOverflow

                    Start.Dock = DockStyle.Left
                    Splitter1.Dock = DockStyle.Left
                    Panel1.Dock = DockStyle.Left
                    Splitter2.Dock = DockStyle.Left
                    Panel4.Dock = DockStyle.Right
                    Splitter3.Dock = DockStyle.Right

                    AppStrip.Location = New Point(5, 0)
                    AppStrip.Size = New Size(Panel1.Width - 5, Panel1.Height + 2)

                    ToolStrip1.Location = New Point(0, 0)
                    ToolStrip1.Size = New Size(TrayNotifyWnd.Width, TrayNotifyWnd.Height + 2)
                    ToolStrip1.RightToLeft = RightToLeft.Yes

                    btn_ShowDesktop.Dock = DockStyle.Right
                    TimeDatePanel.Dock = DockStyle.Right
                    TimeLabel.Dock = DockStyle.Right
                    btn_LanguageBar.Dock = DockStyle.Right

                    LoadBandSizes(False)
                End If

            Case DockPosition.Left, DockPosition.Right
                If Not Me.Width = If(isSmallIcons, defSmallSize, defSize) Then
                    Me.Width = If(isSmallIcons, defSmallSize, defSize)
                End If

                If Not Me.Height = SystemInformation.PrimaryMonitorSize.Height Then
                    Me.Height = SystemInformation.PrimaryMonitorSize.Height
                    Me.Width = If(isSmallIcons, defSmallSize, defSize)

                    AppStrip.LayoutStyle = ToolStripLayoutStyle.VerticalStackWithOverflow
                    ProcessStrip.LayoutStyle = ToolStripLayoutStyle.VerticalStackWithOverflow
                    ToolStrip1.LayoutStyle = ToolStripLayoutStyle.VerticalStackWithOverflow

                    Start.Dock = DockStyle.Top
                    Splitter1.Dock = DockStyle.Top
                    Panel1.Dock = DockStyle.Top
                    Splitter2.Dock = DockStyle.Top
                    Panel4.Dock = DockStyle.Bottom
                    Splitter3.Dock = DockStyle.Bottom

                    AppStrip.Location = New Point(0, 5)
                    AppStrip.Size = New Size(Panel1.Width, Panel1.Height - 3)

                    btn_ShowDesktop.Dock = DockStyle.Bottom
                    TimeDatePanel.Dock = DockStyle.Bottom
                    TimeLabel.Dock = DockStyle.Bottom
                    btn_LanguageBar.Dock = DockStyle.Bottom

                    ToolStrip1.Location = New Point(0, 0)
                    ToolStrip1.Size = New Size(TrayNotifyWnd.Width, TrayNotifyWnd.Height + 2)
                    ToolStrip1.RightToLeft = RightToLeft.No

                    LoadBandSizes(True)
                End If
        End Select

        If Not Me.WindowState = FormWindowState.Normal Then Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub AppBar_StyleChanged(sender As Object, e As EventArgs) Handles Me.StyleChanged
        Me.WindowState = FormWindowState.Normal
    End Sub

    Public Sub UpdateStartButtonImage(Optional isNormal As Boolean = False)
        Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0")
            Case 0 'Default
                If Startmenu.Visible = False Then
                    Start.BackgroundImage = My.Resources.StartRight
                Else
                    Start.BackgroundImage = My.Resources.StartLeft
                End If

            Case 1 'Custom Image Logic
                Dim normalBtn As String = EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Normal", ""))
                Dim hoverBtn As String = EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Hover", ""))
                Dim pressedBtn As String = EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Pressed", ""))

                If Startmenu.Visible = False Then
                    Try
                        Start.BackgroundImage = Image.FromFile(If(isNormal, normalBtn, hoverBtn))
                    Catch
                        Start.BackgroundImage = My.Resources.StartRight
                    End Try
                Else
                    Try
                        Start.BackgroundImage = Image.FromFile(pressedBtn)
                    Catch
                        Start.BackgroundImage = My.Resources.StartLeft
                    End Try
                End If

            Case 2 'ORB Logic
                If Startmenu.Visible = False Then
                    Try
                        Start.BackgroundImage = If(isNormal, OrbNormal, OrbHover)
                    Catch
                        Start.BackgroundImage = My.Resources.StartRight
                    End Try
                Else
                    Try
                        Start.BackgroundImage = OrbPressed
                    Catch
                        Start.BackgroundImage = My.Resources.StartLeft
                    End Try
                End If
        End Select
    End Sub

    Public IsMouseOnStart As Boolean = False
    Private Sub Button1_MouseHover(sender As Object, e As EventArgs) Handles Start.MouseHover
        IsMouseOnStart = True

        UpdateStartButtonImage()
    End Sub

    Private Sub Button1_MouseLeave(sender As Object, e As EventArgs) Handles Start.MouseLeave
        IsMouseOnStart = False

        If Not Startmenu.Visible Then
            Hidden = True

            Startmenu.FST = False

            UpdateStartButtonImage(True)
        End If
    End Sub

    Private Sub Button1_MouseEnter(sender As Object, e As EventArgs) Handles Start.MouseEnter
        IsMouseOnStart = True

        Hidden = False

        If Startmenu.Visible = False Then
            Start.Focus()
        End If

        Startmenu.FST = True

        UpdateStartButtonImage()
    End Sub

    Private Sub ReloadAppsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReloadAppsToolStripMenuItem.Click
        LoadApps()

        UpdateAppbarAccent()
    End Sub

    Public Sub UpdateAppbarAccent()
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "UseSystemColor", 1) = 1 Then
            Dim accentColor As Color = GetAccentColor()

            AppbarProperties.btn_PickerSC.BackColor = accentColor
            AppbarProperties.pb_CIPreview.BackColor = accentColor
            Me.BackColor = accentColor
        End If
    End Sub

    Private Sub AppStrip_MouseUp(sender As Object, e As MouseEventArgs) Handles AppStrip.MouseUp
        Select Case e.Button
            Case MouseButtons.Right
                cm_AppStripF.Show(MousePosition)
        End Select
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Dim FBD As New FolderBrowserDialog With {
        .Description = "Select a default location where your Pinned apps will be:",
        .RootFolder = Environment.SpecialFolder.Desktop,
        .ShowNewFolderButton = True,
        .SelectedPath = PinnedAppsDir
        }

        Select Case FBD.ShowDialog
            Case DialogResult.OK
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "PinnedAppsLocation", FBD.SelectedPath, Microsoft.Win32.RegistryValueKind.String)
                PinnedAppsDir = FBD.SelectedPath

                LoadPinnedApps()
        End Select
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem.Click
        LoadPinnedApps()
    End Sub

    Private lastMousePos As Point
    Private Sub ToolStrip1_MouseWheel(sender As Object, e As MouseEventArgs) Handles ToolStrip1.MouseWheel
        Dim item = ToolStrip1.GetItemAt(e.Location)

        If item IsNot Nothing AndAlso item.Name = "ToolStripButton1" Then

            Dim curVol As Integer = VolumeControl.GetCurrentVolume * 100
            Dim volAdd As Integer = 2

            If e.Delta > 0 Then

                ' Mouse wheel up
                If curVol + 2 >= 100 Then
                    VolumeControl.SetVolume(1)
                Else
                    VolumeControl.SetVolume((curVol / 100) + (volAdd / 100))
                End If

            Else

                ' Mouse wheel down
                If curVol - 2 <= 0 Then
                    VolumeControl.SetVolume(0)
                Else
                    VolumeControl.SetVolume((curVol / 100) - (volAdd / 100))
                End If

            End If

            ToolStripButton1.Text = $"Volume ({VolumeControl.GetCurrentVolume * 100}%)"
            VolumeControl.UpdateIcons()
        End If
    End Sub

    Private Sub ToolStripButton1_MouseEnter(sender As Object, e As EventArgs) Handles ToolStripButton1.MouseEnter
        ToolStripButton1.Text = $"Volume ({Math.Round(VolumeControl.GetCurrentVolume * 100)}%)"
    End Sub

    Private Sub VCM_Opening(sender As Object, e As CancelEventArgs) Handles cm_Volume.Opening
        MuteToolStripMenuItem.Checked = VolumeControl.IsSysMuted
    End Sub

    Private Sub MuteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MuteToolStripMenuItem.Click
        VolumeControl.ToggleMute(MuteToolStripMenuItem.Checked)
        VolumeControl.UpdateIcons()
    End Sub

    Private Sub VCM_Closing(sender As Object, e As ToolStripDropDownClosingEventArgs) Handles cm_Volume.Closing
        ToolStripButton1.Checked = False
    End Sub

    Private Sub MCM_Closing(sender As Object, e As ToolStripDropDownClosingEventArgs) Handles cm_Menu.Closing
        ToolStripButton3.Checked = False
    End Sub

    Private Sub TimeLabel_Click(sender As Object, e As EventArgs) Handles TimeLabel.Click, ShowTimeAndDateToolStripMenuItem.Click
        If DateAndTime.Visible = True Then
            DateAndTime.Focus()
        Else
            DateAndTime.Show()
            DateAndTime.TabControl1.SelectTab(3)
        End If
    End Sub

    Private Sub AlarmController_Tick(sender As Object, e As EventArgs) Handles AlarmController.Tick
        ' Future Alarms!
        AlarmController.Enabled = False
    End Sub

    Private Sub PIcm_Opening(sender As Object, e As CancelEventArgs) Handles cm_AppStrip.Opening
        If currentPinnedItem IsNot Nothing Then
            currentPinnedItem.Checked = True
        End If
    End Sub

    Private Sub PIcm_Closing(sender As Object, e As ToolStripDropDownClosingEventArgs) Handles cm_AppStrip.Closing
        If currentPinnedItem IsNot Nothing Then
            currentPinnedItem.Checked = False

            currentPinnedItem = Nothing
        End If
    End Sub

    Public WithEvents CloseTimer As New System.Windows.Forms.Timer With {.Interval = 200}

    Private Sub CloseTimer_Tick(sender As Object, e As EventArgs) Handles CloseTimer.Tick
        CloseTimer.Stop()
        WAT.Hide()
    End Sub

    Private Sub VolumeSliderWindowsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VolumeSliderWindowsToolStripMenuItem.Click
        On Error Resume Next
        Process.Start(soundcontrolPath, "-f")
    End Sub

    Private Sub LLCM_Opening(sender As Object, e As CancelEventArgs) Handles cm_Language.Opening
        cm_Language.Items.Clear()

        For Each lang As InputLanguage In InputLanguage.InstalledInputLanguages
            Dim displayName As String = lang.Culture.EnglishName

            For Each item As ToolStripMenuItem In cm_Language.Items
                If item.Text = displayName Then
                    item.Text = $"{item.Text} ({DirectCast(item.Tag, InputLanguage).LayoutName})"
                    displayName = $"{lang.Culture.EnglishName} ({lang.LayoutName})"
                End If
            Next

            Dim menuItem As New ToolStripMenuItem(displayName) With {
                .Tag = lang
            }

            If lang.Equals(InputLanguage.CurrentInputLanguage) Then
                menuItem.Checked = True
            End If

            AddHandler menuItem.Click, AddressOf LanguageItem_Click
            cm_Language.Items.Add(menuItem)
        Next
    End Sub

    Public Sub UpdateLanguageDisplay()
        Dim currentCulture As CultureInfo = InputLanguage.CurrentInputLanguage.Culture
        btn_LanguageBar.Text = currentCulture.ThreeLetterISOLanguageName.ToUpper()
    End Sub

    Private Sub LanguageItem_Click(sender As Object, e As EventArgs)
        Dim clickedItem As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        Dim targetLang As InputLanguage = DirectCast(clickedItem.Tag, InputLanguage)

        InputLanguage.CurrentInputLanguage = targetLang
        UpdateLanguageDisplay()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles btn_LanguageBar.Click
        cm_Language.Show(MousePosition)
    End Sub

    <DllImport("user32.dll")>
    Private Shared Function GetKeyboardLayout(idThread As Integer) As IntPtr
    End Function

    Private Const WM_INPUTLANGCHANGEREQUEST As UInteger = &H50

    Private Function GetGlobalInputLanguageHandle() As IntPtr
        Dim foregroundWindow As IntPtr = GetForegroundWindow()
        Dim threadId As Integer = GetWindowThreadProcessId(foregroundWindow, 0)
        Return GetKeyboardLayout(threadId)
    End Function

    Private Sub Button3_MouseWheel(sender As Object, e As MouseEventArgs) Handles btn_LanguageBar.MouseWheel
        Dim languages As InputLanguageCollection = InputLanguage.InstalledInputLanguages
        Dim currentHandle As IntPtr = GetGlobalInputLanguageHandle()
        Dim currentIndex As Integer = -1

        For i As Integer = 0 To languages.Count - 1
            If languages(i).Handle = currentHandle Then
                currentIndex = i
                Exit For
            End If
        Next

        If currentIndex = -1 Then Exit Sub

        Dim newIndex As Integer = If(e.Delta > 0, currentIndex - 1, currentIndex + 1)
        If newIndex < 0 Then newIndex = languages.Count - 1
        If newIndex >= languages.Count Then newIndex = 0

        Dim targetHandle As IntPtr = languages(newIndex).Handle

        InputLanguage.CurrentInputLanguage = languages(newIndex)
        UpdateLanguageDisplay()

        PostMessage(HWND_BROADCAST, WM_INPUTLANGCHANGEREQUEST, IntPtr.Zero, targetHandle)
    End Sub

    Private Sub Button3_MouseEnter(sender As Object, e As EventArgs) Handles btn_LanguageBar.MouseEnter
        btn_LanguageBar.Focus()
    End Sub

    Public WithEvents TimerPeek As New System.Windows.Forms.Timer() With {.Interval = 250}
    Public IsDesktopPeekEnabled As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "TaskbarSd", True)
    Public IsAeroPeekEnabled As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\DWM", "EnableAeroPeek", True)

    Private Sub TimerPeek_Tick(sender As Object, e As EventArgs) Handles TimerPeek.Tick
        TimerPeek.Stop()

        If Not IsAeroPeekEnabled Then Exit Sub

        If Desktop.IsHandleCreated Then
            Dim targetHWnd As IntPtr = Desktop.Handle
            Dim targetPos = mainPeek.GetVisibleWindowRect(targetHWnd)

            AeroPeekScreen.Show()
            Me.BringToFront()

            mainPeek.ShowPeek(targetHWnd, AeroPeekScreen.Handle, targetPos)
        End If
    End Sub

    Private Sub Button2_MouseEnter(sender As Object, e As EventArgs) Handles btn_ShowDesktop.MouseEnter
        If Not IsAeroPeekEnabled Then Exit Sub

        If IsDesktopPeekEnabled Then
            If Not Desktop.IsShowDesktopMode Then TimerPeek.Start()
        End If
    End Sub

    Private Sub Button2_MouseLeave(sender As Object, e As EventArgs) Handles btn_ShowDesktop.MouseLeave
        If Not IsAeroPeekEnabled Then Exit Sub
        If Not IsDesktopPeekEnabled Then Exit Sub

        TimerPeek.Stop()
        mainPeek.HidePeek()

        AeroPeekScreen.Hide()
    End Sub

    Private Sub ShowDesktopCM_Opening(sender As Object, e As CancelEventArgs) Handles cm_ShowDesktop.Opening
        PeekDesktopToolStripMenuItem.Checked = IsDesktopPeekEnabled
    End Sub

    Private Sub PeekDesktopToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PeekDesktopToolStripMenuItem.Click
        If IsDesktopPeekEnabled = True Then IsDesktopPeekEnabled = False Else IsDesktopPeekEnabled = True

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "TaskbarSd", IsDesktopPeekEnabled, RegistryValueKind.DWord)
    End Sub

    Private Const SC_MONITORPOWER As Integer = &HF170

    ' lParam:
    ' 1 = Low power
    ' 2 = Power off
    ' -1 = Power on 

    Public Shared Sub TurnOffMonitor(handle As IntPtr)
        SendMessage(handle, WM_SYSCOMMAND, New IntPtr(SC_MONITORPOWER), New IntPtr(2))
    End Sub

    Private Sub RefreshPowerMenus()
        LogonToolStripMenuItem.DropDownItems.Clear()
        Dim sessionActions As String() = {"Lock", "Switch User", "Log Off"}

        For Each act In sessionActions
            Dim item As New ToolStripMenuItem(act)
            AddHandler item.Click, AddressOf SessionButtonPressed
            LogonToolStripMenuItem.DropDownItems.Add(item)
        Next

        PowerToolStripMenuItem.DropDownItems.Clear()
        Dim powerActions = PowerManager.GetAvailablePowerActions

        For Each act In powerActions
            Dim item As New ToolStripMenuItem(act)
            AddHandler item.Click, AddressOf PowerButtonPressed
            PowerToolStripMenuItem.DropDownItems.Add(item)
        Next
    End Sub

    Private Sub PowerButtonPressed(sender As Object, e As EventArgs)
        Dim item = DirectCast(sender, ToolStripMenuItem)

        Dim isBrutalMode As Boolean = False
        Dim isForceMode As Boolean = False

        If (ModifierKeys And Keys.Control) = Keys.Control Then
            isBrutalMode = True ' EWX_FORCE
        ElseIf SA.chkForceMode.Checked Then
            isForceMode = True ' EWX_FORCEIFHUNG
        End If

        PowerManager.ExecuteAction(item.Text, isForceMode, 0, PowerManager.IsFastStartupEnabled, SA.chkDisableWakeEvent.Checked, isBrutalMode)
    End Sub

    Private Sub SessionButtonPressed(sender As Object, e As EventArgs)
        Dim item = DirectCast(sender, ToolStripMenuItem)
        PowerManager.ExecuteSessionAction(item.Text)
    End Sub

    Public Shared Sub UpdateWorkingArea()
        Dim workspace As New WorkspaceManager()

        Try
            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "CurrentPreset", "Default") Is Nothing Then Exit Try

            Dim workdata As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea\Presets\", My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "CurrentPreset", "Default"), "Default")
            Dim parts As String() = workdata.Split(" "c)

            If parts.Length >= 4 Then
                Dim workTop As String = parts(0)
                Dim workBottom As String = parts(1)
                Dim workLeft As String = parts(2)
                Dim workRight As String = parts(3)

                If AlocateAppbarWorkingArea Then
                    Select Case currentDockPos
                        Case DockPosition.Top
                            workTop += If(isSmallIcons, defSmallSize, defSize)
                        Case DockPosition.Bottom
                            workBottom += If(isSmallIcons, defSmallSize, defSize)
                        Case DockPosition.Left
                            workLeft += If(isSmallIcons, defSmallSize, defSize)
                        Case DockPosition.Right
                            workRight += If(isSmallIcons, defSmallSize, defSize)
                    End Select
                End If

                workspace.UpdateWorkingArea(workTop, workBottom, workLeft, workRight)
            End If

        Catch ex As Exception
            workspace.UpdateWorkingArea(AppbarProperties.defWorkTop, AppbarProperties.defWorkBottom, AppbarProperties.defWorkLeft, AppbarProperties.defWorkRight)
        End Try
    End Sub

    Private Sub SmallIconsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SmallIconsToolStripMenuItem.Click
        isSmallIcons = Not SmallIconsToolStripMenuItem.Checked
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\Alignment", "SmallIcons", isSmallIcons, Microsoft.Win32.RegistryValueKind.DWord)

        Me.Location = New Point(Me.Location.X - 1, Me.Location.Y - 1)
        Me.Height -= 1

        LoadApps()
        UpdateWorkingArea()
    End Sub

    Private Sub AligmentToolStripMenuItem_DropDownOpening(sender As Object, e As EventArgs) Handles AligmentToolStripMenuItem.DropDownOpening
        For Each item As ToolStripMenuItem In AligmentToolStripMenuItem.DropDownItems
            item.Checked = False
        Next

        Select Case currentDockPos
            Case DockPosition.Top
                TopToolStripMenuItem.Checked = True
            Case DockPosition.Left
                LeftToolStripMenuItem.Checked = True
            Case DockPosition.Right
                RightToolStripMenuItem.Checked = True
            Case DockPosition.Bottom
                BottomToolStripMenuItem.Checked = True
        End Select
    End Sub

    Private Sub TopToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TopToolStripMenuItem.Click
        currentDockPos = DockPosition.Top
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\Alignment", "Dock", currentDockPos, Microsoft.Win32.RegistryValueKind.DWord)

        Me.Location = New Point(Me.Location.X - 1, Me.Location.Y - 1)
        Me.Height -= 1

        LoadApps()
        UpdateWorkingArea()
    End Sub

    Private Sub LeftToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LeftToolStripMenuItem.Click
        currentDockPos = DockPosition.Left
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\Alignment", "Dock", currentDockPos, Microsoft.Win32.RegistryValueKind.DWord)

        Me.Location = New Point(Me.Location.X - 1, Me.Location.Y - 1)
        Me.Height -= 1

        LoadApps()
        UpdateWorkingArea()
    End Sub

    Private Sub RightToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RightToolStripMenuItem.Click
        currentDockPos = DockPosition.Right
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\Alignment", "Dock", currentDockPos, Microsoft.Win32.RegistryValueKind.DWord)

        Me.Location = New Point(Me.Location.X - 1, Me.Location.Y - 1)
        Me.Height -= 1

        LoadApps()
        UpdateWorkingArea()
    End Sub

    Private Sub BottomToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BottomToolStripMenuItem.Click
        currentDockPos = DockPosition.Bottom
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\Alignment", "Dock", currentDockPos, Microsoft.Win32.RegistryValueKind.DWord)

        Me.Location = New Point(Me.Location.X - 1, Me.Location.Y - 1)
        Me.Height -= 1

        LoadApps()
        UpdateWorkingArea()
    End Sub

    Private Sub ProcessStrip_ItemAdded(sender As Object, e As ToolStripItemEventArgs) Handles ProcessStrip.ItemAdded
        e.Item.ForeColor = primaryForeColor
    End Sub

    Private Sub Start_MouseEnter(sender As Object, e As EventArgs) Handles Start.MouseEnter
        If Not Startmenu.Visible Then
            Start.Focus()
            Start.Select()
        End If
    End Sub

    Private Sub ToolStrip1_SizeChanged(sender As Object, e As EventArgs) Handles ToolStrip1.SizeChanged
        Dim isVertical As Boolean = False

        Select Case currentDockPos
            Case DockPosition.Top, DockPosition.Bottom
                isVertical = False

            Case DockPosition.Left, DockPosition.Right
                isVertical = True

        End Select

        For Each btn As ToolStripButton In ToolStrip1.Items
            If isVertical Then
                btn.Size = New Size(ToolStrip1.Width, 24)
            Else
                btn.Size = New Size(24, ToolStrip1.Height)
            End If
        Next
    End Sub

    Private Sub ShowDesktopToolStripMenuItem1_MouseEnter(sender As Object, e As EventArgs) Handles ShowDesktopToolStripMenuItem1.MouseEnter
        btn_ShowDesktop.Focus()
        btn_ShowDesktop.Select()
    End Sub

    Private Sub RefreshToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem1.Click
        TrayCleanup()
        PostMessage(HWND_BROADCAST, CUInt(WM_TASKBARCREATED), IntPtr.Zero, IntPtr.Zero)
    End Sub

    Private Sub ToolStrip1_MouseUp(sender As Object, e As MouseEventArgs) Handles ToolStrip1.MouseUp
        Dim clickedItem As ToolStripItem = ToolStrip1.GetItemAt(e.Location)

        If clickedItem Is Nothing Then
            cm_TrayIcons.Show(MousePosition)
        End If
    End Sub

    Private Async Sub ToolStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked
        If TypeOf e.ClickedItem.Tag Is TrayIconInfo Then Exit Sub

        Select Case e.ClickedItem.Tag
            Case "Volume"
                If UseLegacyVolumeSlider Then
                    Try
                        Dim p As Process = Process.Start(soundcontrolPath, "-f")

                        Dim hwnd As IntPtr = IntPtr.Zero
                        Dim attempts As Integer = 0

                        While hwnd = IntPtr.Zero AndAlso attempts < 20
                            Await Task.Delay(50)
                            p.Refresh()
                            hwnd = p.MainWindowHandle
                            attempts += 1
                        End While

                        If hwnd <> IntPtr.Zero Then
                            Dim cursorPoint = Cursor.Position

                            Select Case currentDockPos
                                Case DockPosition.Top
                                    SetWindowPos(hwnd, IntPtr.Zero, cursorPoint.X - 45, cursorPoint.Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_SHOWWINDOW)
                                Case DockPosition.Left
                                    SetWindowPos(hwnd, IntPtr.Zero, 0, cursorPoint.Y - 200, cursorPoint.Y - 200, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_SHOWWINDOW)
                                Case DockPosition.Right
                                    SetWindowPos(hwnd, IntPtr.Zero, cursorPoint.X - 45, cursorPoint.Y - 200, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_SHOWWINDOW)
                                Case DockPosition.Bottom
                                    SetWindowPos(hwnd, IntPtr.Zero, cursorPoint.X - 45, cursorPoint.Y - 290, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_SHOWWINDOW)
                            End Select
                        End If
                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.Critical)
                    End Try
                Else ' Classic Shell's volume slider
                    Try
                        If VolumeControl.Visible = True Then
                            VolumeControl.Focus()
                        Else
                            ToolStripButton1.Checked = True
                            VolumeControl.Show(Me)
                        End If
                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.Critical)
                    End Try
                End If

            Case "Action"
                ToolStripButton3.Checked = True
                cm_Menu.Show(MousePosition)

            Case "Keyboard"
                Try
                    Dim oskPath As String = EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "Keyboard", Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\osk.exe"))
                    Process.Start(oskPath)

                    If ActiveWindowHandle <> IntPtr.Zero Then SetForegroundWindow(ActiveWindowHandle)
                Catch
                End Try

            Case "Battery"
                If PowerControl.Visible = True Then
                    PowerControl.Focus()
                Else
                    ToolStripButton4.Checked = True

                    PowerControl.Show(Me)
                End If

            Case "Network"
                If NetworkControl.Visible = True Then
                    NetworkControl.Focus()
                Else
                    ToolStripButton5.Checked = True

                    NetworkControl.Show(Me)
                End If

        End Select
    End Sub

    Private Sub ToolStrip1_MouseDown(sender As Object, e As MouseEventArgs) Handles ToolStrip1.MouseDown
        Dim clickedItem As ToolStripItem = ToolStrip1.GetItemAt(e.Location)

        If clickedItem Is Nothing Then Exit Sub

        Select Case clickedItem.Tag
            Case "Volume"
            Case "Action"
            Case "Keyboard"
            Case "Battery"
            Case "Network"
                '...
        End Select
    End Sub

    Private Sub ToolStripButton1_MouseUp(sender As Object, e As MouseEventArgs) Handles ToolStripButton1.MouseUp
        If e.Button = MouseButtons.Right Then
            ToolStripButton1.Checked = True
            cm_Volume.Show(MousePosition)
        End If
    End Sub

    Private Sub ToolStripButton2_MouseUp(sender As Object, e As MouseEventArgs) Handles ToolStripButton2.MouseUp
        If e.Button = MouseButtons.Right Then
            Try
                Dim oskPath As String = EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "Keyboard", Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\osk.exe"))
                Process.Start(oskPath)

                If LastActiveWindowHandle <> IntPtr.Zero Then SetForegroundWindow(LastActiveWindowHandle)
            Catch
            End Try
        End If
    End Sub

    Private Sub ToolStripButton3_MouseUp(sender As Object, e As MouseEventArgs) Handles ToolStripButton3.MouseUp
        If e.Button = MouseButtons.Right Then
            ToolStripButton3.Checked = True
            cm_Menu.Show(MousePosition)
        End If
    End Sub

    Private Sub ToolStripButton4_MouseUp(sender As Object, e As MouseEventArgs) Handles ToolStripButton4.MouseUp
        If e.Button = MouseButtons.Right Then
            If PowerControl.Visible = True Then
                PowerControl.Focus()
            Else
                ToolStripButton4.Checked = True

                PowerControl.Show(Me)
            End If
        End If
    End Sub

    Private Sub ToolStripButton5_MouseUp(sender As Object, e As MouseEventArgs) Handles ToolStripButton5.MouseUp
        If e.Button = MouseButtons.Right Then
            If NetworkControl.Visible = True Then
                NetworkControl.Focus()
            Else
                ToolStripButton5.Checked = True

                NetworkControl.Show(Me)
            End If
        End If
    End Sub

    Private Sub cm_TrayIcons_Opening(sender As Object, e As CancelEventArgs) Handles cm_TrayIcons.Opening
        Dim clickedItem As ToolStripItem = ToolStrip1.GetItemAt(ToolStrip1.PointToClient(MousePosition))

        If clickedItem Is Nothing OrElse TypeOf clickedItem.Tag Is TrayIconInfo Then
            e.Cancel = False
        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub FullscreenWin32ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FullscreenWin32ToolStripMenuItem.Click
        If cm_ProcessItem.Tag IsNot Nothing AndAlso TypeOf cm_ProcessItem.Tag Is IntPtr Then
            Dim hWnd As IntPtr = CType(cm_ProcessItem.Tag, IntPtr)

            WindowFullscreenHelper.ToggleFullscreen(hWnd)
        End If
    End Sub

    Private Sub Splitters_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles Splitter1.SplitterMoved, Splitter2.SplitterMoved, Splitter3.SplitterMoved
        Dim sp As Splitter = CType(sender, Splitter)

        ' Prepared for multiple toolbars
        'Dim layerNumber As String = sp.Name.Replace("Splitter", "")

        Select Case currentDockPos
            Case DockPosition.Top, DockPosition.Bottom
                If sp Is Splitter1 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer1", Start.Width, Microsoft.Win32.RegistryValueKind.String)
                ElseIf sp Is Splitter2 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer2", Panel1.Width, Microsoft.Win32.RegistryValueKind.String)
                ElseIf sp Is Splitter3 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer3", Panel4.Width, Microsoft.Win32.RegistryValueKind.String)
                End If

            Case DockPosition.Left, DockPosition.Right
                If sp Is Splitter1 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer1", Start.Height, Microsoft.Win32.RegistryValueKind.String)
                ElseIf sp Is Splitter2 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer2", Panel1.Height, Microsoft.Win32.RegistryValueKind.String)
                ElseIf sp Is Splitter3 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer3", Panel4.Height, Microsoft.Win32.RegistryValueKind.String)
                End If

        End Select
    End Sub

    Private Sub Splitters_Paint(sender As Object, e As PaintEventArgs) Handles Splitter1.Paint, Splitter2.Paint, Splitter3.Paint
        Dim g As Graphics = e.Graphics
        Dim sp As Splitter = CType(sender, Splitter)

        g.SmoothingMode = SmoothingMode.HighQuality
        g.PixelOffsetMode = PixelOffsetMode.HighQuality

        If Me.BackgroundImage IsNot Nothing Then
            Dim splitterLocationInForm As Point = Me.PointToClient(sp.PointToScreen(New Point(0, 0)))
            Using bgBrush As New TextureBrush(Me.BackgroundImage)
                bgBrush.TranslateTransform(-splitterLocationInForm.X, -splitterLocationInForm.Y)
                g.FillRectangle(bgBrush, sp.ClientRectangle)
            End Using
        End If

        Dim isHorizontal As Boolean = (sp.Width > sp.Height)

        Dim dotSize As Integer = 3
        Dim dotGap As Integer = 2
        Dim gripLength As Integer = 50

        Dim patternWidth As Integer = If(isHorizontal, dotSize + dotGap, sp.ClientRectangle.Width)
        Dim patternHeight As Integer = If(isHorizontal, sp.ClientRectangle.Height, dotSize + dotGap)

        Dim spOffsetX As Integer
        Dim spOffsetY As Integer

        If isHorizontal Then
            spOffsetX = 1
            spOffsetY = (sp.ClientRectangle.Height - dotSize) \ 2
        Else
            spOffsetX = (sp.ClientRectangle.Width - dotSize) \ 2
            spOffsetY = 1
        End If

        Dim mainColor As Color = sp.BackColor
        Dim shadowColor As Color = Color.FromArgb(110, Color.Black)

        Using gripBitmap As New Bitmap(patternWidth, patternHeight)
            Using gripGraphics As Graphics = Graphics.FromImage(gripBitmap)
                gripGraphics.SmoothingMode = SmoothingMode.None
                gripGraphics.Clear(Color.Transparent)

                Using mainBrush As New SolidBrush(mainColor), shadowBrush As New SolidBrush(shadowColor)
                    ' 1. Shadow
                    gripGraphics.FillRectangle(shadowBrush, spOffsetX + 1, spOffsetY + 1, dotSize, dotSize)

                    ' 2. Main square
                    gripGraphics.FillRectangle(mainBrush, spOffsetX, spOffsetY, dotSize, dotSize)
                End Using
            End Using

            Using patternBrush As New TextureBrush(gripBitmap)
                patternBrush.WrapMode = WrapMode.Tile

                Dim gripRect As Rectangle
                If isHorizontal Then
                    ' For Width (X)
                    Dim gripLeft As Integer = (sp.ClientRectangle.Width - gripLength) \ 2
                    gripRect = New Rectangle(gripLeft, 0, gripLength, sp.ClientRectangle.Height)
                Else
                    ' For Height (Y)
                    Dim gripTop As Integer = (sp.ClientRectangle.Height - gripLength) \ 2
                    gripRect = New Rectangle(0, gripTop, sp.ClientRectangle.Width, gripLength)
                End If

                g.FillRectangle(patternBrush, gripRect)
            End Using
        End Using
    End Sub

    Private Sub EndShellToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EndShellToolStripMenuItem.Click
        Dim workingAreaReset As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "ResetAfterShutdown", True)
        Dim workspace As New WorkspaceManager

        If workingAreaReset Then workspace.UpdateWorkingArea(0, 0, 0, 0)

        CanClose = True
        Desktop.CanClose = True

        Application.Exit()
        End
    End Sub

    Private Sub RestartShellToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestartShellToolStripMenuItem.Click
        CanClose = True
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

    Public Shared Function EncodeFilePath(ByVal path As String) As String
        If String.IsNullOrEmpty(path) Then Return path

        If path.ToLower.Contains("%cd%") Then
            Dim baseDir As String = Application.StartupPath
            path = System.Text.RegularExpressions.Regex.Replace(path, "%cd%", baseDir, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        End If

        path = Environment.ExpandEnvironmentVariables(path)

        Dim brackets As Char() = {"["c, "]"c, "{"c, "}"c, "("c, ")"c, """"c, "'"c}
        For Each c As Char In brackets
            path = path.Replace(c, String.Empty)
        Next

        Return path
    End Function
End Class

Public Class WorkspaceManager

    <StructLayout(LayoutKind.Sequential)>
    Public Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SystemParametersInfo(
        ByVal uiAction As UInteger,
        ByVal uiParam As UInteger,
        ByVal pvParam As IntPtr,
        ByVal fWinIni As UInteger) As Boolean
    End Function

    Private Const SPI_SETWORKAREA As UInteger = 47
    Private Const SPIF_UPDATEINIFILE As UInteger = &H1
    Private Const SPIF_SENDCHANGE As UInteger = &H2

    Private Const WM_SETTINGCHANGE As Integer = &H1A
    Private Const SMTO_ABORTIFHUNG As Integer = &H2

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function SendMessageTimeout(
    ByVal hWnd As IntPtr,
    ByVal Msg As Integer,
    ByVal wParam As IntPtr,
    ByVal lParam As IntPtr,
    ByVal fuFlags As Integer,
    ByVal uTimeout As Integer,
    <Out> ByRef lpdwResult As IntPtr
) As IntPtr
    End Function

    Public Sub UpdateWorkingArea(ByVal topOffset As Integer, ByVal bottomOffset As Integer, ByVal leftOffset As Integer, ByVal rightOffset As Integer)
        Dim screenBounds As Rectangle = Screen.PrimaryScreen.Bounds

        If topOffset < 0 OrElse bottomOffset < 0 OrElse leftOffset < 0 OrElse rightOffset < 0 Then
            MessageBox.Show("Offsets cannot be negative values.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim minWidth As Integer = 100
        Dim minHeight As Integer = 100

        If (leftOffset + rightOffset) > (screenBounds.Width - minWidth) Then
            MessageBox.Show("Left and Right offsets are too large. They would overlap or leave too little space.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If (topOffset + bottomOffset) > (screenBounds.Height - minHeight) Then
            MessageBox.Show("Top and Bottom offsets are too large. They would overlap or leave too little space.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim newWorkArea As New RECT With {
            .Left = leftOffset,
            .Top = topOffset,
            .Right = screenBounds.Width - rightOffset,
            .Bottom = screenBounds.Height - bottomOffset
        }

        ' Alokujeme strukturu RECT přímo v nespravované paměti, aby k ní mělo user32.dll čistý přístup
        Dim ptr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(newWorkArea))
        Dim success As Boolean = False

        Try
            Marshal.StructureToPtr(newWorkArea, ptr, False)

            success = SystemParametersInfo(
                SPI_SETWORKAREA,
                0,
                ptr,
                SPIF_UPDATEINIFILE)
        Finally
            Marshal.FreeHGlobal(ptr)
        End Try

        If Not success Then
            Dim errorCode As Integer = Marshal.GetLastWin32Error()
            MessageBox.Show($"Failed to set Working Area. Error code: {errorCode}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class

Module GlobalKeyboardHook
    'Window Manipulation
    <DllImport("user32.dll")>
    Private Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, ByRef lpdwProcessId As UInteger) As UInteger
    End Function

    <DllImport("user32.dll")>
    Private Function ShowWindow(ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Function SetWindowPos(ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    Private Const SW_RESTORE As Integer = 9
    Private Const SW_MINIMIZE As Integer = 6
    Private Const SW_MAXIMIZE As Integer = 3
    Private Const SW_SHOWNOACTIVATE As Integer = 4

    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOZORDER As UInteger = &H4
    Private Const SWP_SHOWWINDOW As UInteger = &H40

    Private Const WM_CLOSE As UInteger = &H10

    Private Sub MaximizeWindow(ByVal hWnd As IntPtr)
        ShowWindow(hWnd, SW_MAXIMIZE)
    End Sub

    Private Sub MinimizeWindow(ByVal hWnd As IntPtr)
        ShowWindow(hWnd, SW_MINIMIZE)
    End Sub

    Private Sub RestoreWindow(ByVal hWnd As IntPtr)
        ShowWindow(hWnd, SW_RESTORE)
    End Sub

    Private Sub MoveAndResizeWindow(ByVal hWnd As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer)
        SetWindowPos(hWnd, IntPtr.Zero, X, Y, Width, Height, SWP_NOZORDER Or SWP_SHOWWINDOW)
    End Sub
    Private Sub CloseWindow(ByVal hWnd As IntPtr)
        SendMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero)
    End Sub

    <StructLayout(LayoutKind.Sequential)>
    Private Structure POINTAPI
        Public x As Integer
        Public y As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure WINDOWPLACEMENT
        Public length As UInteger
        Public flags As UInteger
        Public showCmd As UInteger
        Public ptMinPosition As POINTAPI
        Public ptMaxPosition As POINTAPI
        Public rcNormalPosition As RECT
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function GetWindowPlacement(ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function

    Private Const SW_SHOWNORMAL As UInteger = 1
    Private Const SW_SHOWMAXIMIZED As UInteger = 3
    Private Const SW_SHOWMINIMIZED As UInteger = 2

    ' Function for detecting WindowState
    Public Function GetWindowState(ByVal hWnd As IntPtr, Optional toInt As Boolean = False) As String
        Dim wp As WINDOWPLACEMENT
        wp.length = CType(Marshal.SizeOf(wp), UInteger)

        If toInt = False Then
            If GetWindowPlacement(hWnd, wp) Then
                Select Case wp.showCmd
                    Case SW_SHOWNORMAL
                        Return "Normal"
                    Case SW_SHOWMAXIMIZED
                        Return "Maximized"
                    Case SW_SHOWMINIMIZED
                        Return "Minimized"
                    Case Else
                        Return "Unknown"
                End Select
            Else
                Return "Error by getting Information."
            End If
        Else
            If GetWindowPlacement(hWnd, wp) Then
                Return wp.showCmd
            Else
                Return "Error by getting Information."
            End If
        End If
    End Function

    Private Const GWL_EXSTYLE As Integer = -20

    ' Extended Window Styles
    Private Const WS_EX_TOPMOST As UInteger = &H8

    Private ReadOnly HWND_TOPMOST As New IntPtr(-1)
    Private ReadOnly HWND_NOTOPMOST As New IntPtr(-2)

    Private Const SWP_NOACTIVATE As UInteger = &H10

    ' Declaration GetWindowLongPtr for 64-bit apps
    <DllImport("user32.dll", EntryPoint:="GetWindowLongPtrA", SetLastError:=True)>
    Private Function GetWindowLongPtr(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
    End Function

    ' Declaration GetWindowLong for 32-bit apps
    <DllImport("user32.dll", EntryPoint:="GetWindowLongA", SetLastError:=True)>
    Private Function GetWindowLong(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
    End Function


    ' Always On Top detection
    Public Function IsWindowAlwaysOnTop(ByVal hWnd As IntPtr) As Boolean
        Dim exStyle As UInteger = 0

        If IntPtr.Size = 8 Then ' 64-bit process
            exStyle = CType(GetWindowLongPtr(hWnd, GWL_EXSTYLE).ToInt64, UInteger)
        Else ' 32-bit process
            exStyle = CType(GetWindowLong(hWnd, GWL_EXSTYLE), UInteger)
        End If

        Return (exStyle And WS_EX_TOPMOST) = WS_EX_TOPMOST
    End Function

    ' Always On Top function
    Public Sub ToggleAlwaysOnTop(ByVal hWnd As IntPtr)
        If IsWindowAlwaysOnTop(hWnd) Then
            SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
        Else
            SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
        End If
    End Sub


    'Key Events
    Public WinKeyPressed As Boolean = False
    Public CtrlKeyPressed As Boolean = False
    Public ShiftKeyPressed As Boolean = False
    Public AltKeyPressed As Boolean = False

    Public UsedMultipleKey As Boolean = False

    Public Enum Hotkey
        AltG
        CtrlS
        WinLeft
        WinRight
    End Enum

    ' All keys
    Public Const VK_WIN_LEFT As Integer = 91
    Public Const VK_WIN_RIGHT As Integer = 92

    Public Const VK_CTRL_LEFT As Integer = 162
    Public Const VK_CTRL_RIGHT As Integer = 163

    Public Const VK_SHIFT_LEFT As Integer = 160
    Public Const VK_SHIFT_RIGHT As Integer = 161

    Public Const VK_ALT_LEFT As Integer = 164
    Public Const VK_ALT_RIGHT As Integer = 165 'Or Alt GR

    Public Const WH_KEYBOARD_LL As Integer = 13

    <StructLayout(LayoutKind.Sequential)>
    Public Structure KBDLLHOOKSTRUCT
        Public vkCode As Integer
        Public scanCode As Integer
        Public flags As Integer
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure

    Public Delegate Function HookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr

    <DllImport("user32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Function SetWindowsHookEx(ByVal idHook As Integer, ByVal lpfn As HookProc, ByVal hMod As IntPtr, ByVal dwThreadId As Integer) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Function UnhookWindowsHookEx(ByVal hhk As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Function CallNextHookEx(ByVal hhk As IntPtr, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Function GetModuleHandle(ByVal lpModuleName As String) As IntPtr
    End Function

    Private hookHandle As IntPtr = IntPtr.Zero
    Private hookProcInstance As HookProc

    Public Event HotkeyPressed As EventHandler

    Private Const WM_SYSKEYDOWN As Integer = &H104
    Private Const WM_SYSKEYUP As Integer = &H105

    Private Const WM_SYSCOMMAND As Integer = &H112
    Private Const SC_SIZE As Integer = &HF000
    Private Const GWL_STYLE As Integer = -16
    Private Const WS_THICKFRAME As Integer = &H40000
    Private Const WS_MAXIMIZEBOX As Integer = &H10000

    Private ReadOnly windowRestoreBounds As New Dictionary(Of IntPtr, Rectangle)

    Private Function CanSnapWindow(ByVal hWnd As IntPtr) As Boolean
        Dim style As Integer = GetWindowLong(hWnd, GWL_STYLE)
        If (style And WS_THICKFRAME) = 0 Then Return False

        Dim title As String = GetWindowTitle(hWnd)
        If String.IsNullOrEmpty(title) Then Return False

        If hWnd = AppBar.Handle Then Return False

        Return True
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function GetWindowRect(ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Function GetWindowText(ByVal hWnd As IntPtr, ByVal lpString As StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function GetWindowTextLength(ByVal hWnd As IntPtr) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function MoveWindow(ByVal hWnd As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
    End Function

    Private Function GetWindowPosition(ByVal hWnd As IntPtr) As Rectangle
        Dim r As New RECT
        If GetWindowRect(hWnd, r) Then
            Return New Rectangle(r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top)
        End If
        Return Rectangle.Empty
    End Function

    Private Function GetWindowTitle(ByVal hWnd As IntPtr) As String
        Dim length As Integer = GetWindowTextLength(hWnd)
        If length = 0 Then Return String.Empty

        Dim sb As New StringBuilder(length + 1)
        GetWindowText(hWnd, sb, sb.Capacity)
        Return sb.ToString()
    End Function

    Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As IntPtr) As Boolean

    Public Sub SmartSnap(ByVal hWnd As IntPtr, ByVal mode As SnapMode)
        If Not CanSnapWindow(hWnd) Then Exit Sub

        Dim screenArea As Rectangle = Screen.FromHandle(hWnd).WorkingArea
        Dim halfWidth As Integer = screenArea.Width \ 2
        Dim halfHeight As Integer = screenArea.Height \ 2

        If Not windowRestoreBounds.ContainsKey(hWnd) AndAlso GetWindowState(hWnd, True) = 1 Then
            windowRestoreBounds(hWnd) = GetWindowPosition(hWnd)
        End If

        Dim targetRect As Rectangle

        Select Case mode
            Case SnapMode.LeftHalf
                targetRect = New Rectangle(screenArea.X, screenArea.Y, halfWidth, screenArea.Height)
            Case SnapMode.RightHalf
                targetRect = New Rectangle(screenArea.X + halfWidth, screenArea.Y, halfWidth, screenArea.Height)
            Case SnapMode.BottomQuarter
                targetRect = New Rectangle(screenArea.X, screenArea.Y + halfHeight, screenArea.Width, halfHeight)
            Case SnapMode.TopQuarter
                targetRect = New Rectangle(screenArea.X, screenArea.Y, screenArea.Width, halfHeight)
            Case SnapMode.TopLeftQuarter
                targetRect = New Rectangle(screenArea.X, screenArea.Y, halfWidth, halfHeight)
            Case SnapMode.TopRightQuarter
                targetRect = New Rectangle(screenArea.X + halfWidth, screenArea.Y, halfWidth, halfHeight)
            Case SnapMode.BottomLeftQuarter
                targetRect = New Rectangle(screenArea.X, screenArea.Y + halfHeight, halfWidth, halfHeight)
            Case SnapMode.BottomRightQuarter
                targetRect = New Rectangle(screenArea.X + halfWidth, screenArea.Y + halfHeight, halfWidth, halfHeight)

            Case SnapMode.Restore
                If windowRestoreBounds.ContainsKey(hWnd) Then
                    Dim r = windowRestoreBounds(hWnd)
                    ShowWindow(hWnd, 9) ' SW_RESTORE
                    SetWindowPos(hWnd, IntPtr.Zero, r.X, r.Y, r.Width, r.Height, &H40)
                    windowRestoreBounds.Remove(hWnd)
                    Exit Sub
                End If
        End Select

        If Not targetRect.IsEmpty Then
            ShowWindow(hWnd, 9)
            SetWindowPos(hWnd, IntPtr.Zero, targetRect.X, targetRect.Y, targetRect.Width, targetRect.Height, &H40)
        End If
    End Sub

    Public Sub MoveWindowToNextMonitor(ByVal hWnd As IntPtr)
        If Screen.AllScreens.Length <= 1 Then
            Exit Sub
        End If

        Dim windowRect As New RECT()
        If Not GetWindowRect(hWnd, windowRect) Then
            Exit Sub
        End If

        Dim wWidth As Integer = windowRect.Right - windowRect.Left
        Dim wHeight As Integer = windowRect.Bottom - windowRect.Top

        Dim currentScreen As Screen = Screen.FromHandle(hWnd)
        Dim allScreens As Screen() = Screen.AllScreens
        Dim currentIndex As Integer = Array.IndexOf(allScreens, currentScreen)

        Dim nextIndex As Integer = (currentIndex + 1) Mod allScreens.Length
        Dim nextScreen As Screen = allScreens(nextIndex)

        Dim relativeX As Double = (windowRect.Left - currentScreen.Bounds.X) / currentScreen.Bounds.Width
        Dim relativeY As Double = (windowRect.Top - currentScreen.Bounds.Y) / currentScreen.Bounds.Height

        Dim newX As Integer = nextScreen.Bounds.X + CInt(relativeX * nextScreen.Bounds.Width)
        Dim newY As Integer = nextScreen.Bounds.Y + CInt(relativeY * nextScreen.Bounds.Height)

        MoveWindow(hWnd, newX, newY, wWidth, wHeight, True)
    End Sub

    Public Sub MoveWindowToPreviousMonitor(ByVal hWnd As IntPtr)
        If Screen.AllScreens.Length <= 1 Then
            Exit Sub
        End If

        Dim windowRect As New RECT()
        If Not GetWindowRect(hWnd, windowRect) Then
            Exit Sub
        End If

        Dim wWidth As Integer = windowRect.Right - windowRect.Left
        Dim wHeight As Integer = windowRect.Bottom - windowRect.Top

        Dim currentScreen As Screen = Screen.FromHandle(hWnd)
        Dim allScreens As Screen() = Screen.AllScreens
        Dim currentIndex As Integer = Array.IndexOf(allScreens, currentScreen)

        Dim prevIndex As Integer = currentIndex - 1
        If prevIndex < 0 Then
            prevIndex = allScreens.Length - 1
        End If
        Dim prevScreen As Screen = allScreens(prevIndex)

        Dim relativeX As Double = (windowRect.Left - currentScreen.Bounds.X) / currentScreen.Bounds.Width
        Dim relativeY As Double = (windowRect.Top - currentScreen.Bounds.Y) / currentScreen.Bounds.Height

        Dim newX As Integer = prevScreen.Bounds.X + CInt(relativeX * prevScreen.Bounds.Width)
        Dim newY As Integer = prevScreen.Bounds.Y + CInt(relativeY * prevScreen.Bounds.Height)

        MoveWindow(hWnd, newX, newY, wWidth, wHeight, True)
    End Sub

    <DllImport("user32.dll")>
    Private Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Function SetLayeredWindowAttributes(hWnd As IntPtr, crKey As UInteger, bAlpha As Byte, dwFlags As UInteger) As Boolean
    End Function

    Private Const WS_EX_LAYERED As Integer = &H80000
    Private Const LWA_ALPHA As Integer = &H2

    Public Enum SnapMode
        LeftHalf
        RightHalf
        TopLeftQuarter
        TopRightQuarter
        TopQuarter
        BottomQuarter
        BottomLeftQuarter
        BottomRightQuarter
        Restore
        Minimize
    End Enum

    Private Function SetWindowOpacity(hWnd As IntPtr, opacityPercent As Integer) As Boolean
        If hWnd = AppBar.Handle Then Return Nothing
        If hWnd = Startmenu.Handle Then Return Nothing
        If hWnd = WAT.Handle Then Return Nothing

        If opacityPercent < 10 Then opacityPercent = 10
        If opacityPercent > 100 Then opacityPercent = 100

        Dim opacityByte As Byte = CByte((opacityPercent / 100) * 255)

        Dim currentRes As Integer = GetWindowLong(hWnd, GWL_EXSTYLE)
        SetWindowLong(hWnd, GWL_EXSTYLE, currentRes Or WS_EX_LAYERED)
        Return SetLayeredWindowAttributes(hWnd, 0, opacityByte, LWA_ALPHA)
    End Function

    Private Sub ExecutePinnedBar(ByVal index As Integer, admin As Boolean)
        Dim liveBar = Application.OpenForms.OfType(Of AppBar)().FirstOrDefault()

        If liveBar IsNot Nothing Then
            If liveBar.AppStrip.Items.Count > index Then
                Dim item = AppBar.AppStrip.Items(index)

                If item IsNot Nothing Then
                    Dim executePath As String = DirectCast(item, ToolStripItem).Tag

                    If File.Exists(executePath) Then
                        Dim psi As New ProcessStartInfo(executePath) With {.UseShellExecute = True}
                        If admin Then psi.Verb = "runas"

                        Process.Start(psi)
                    End If
                End If
            End If
        End If
    End Sub

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function PostMessage(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As IntPtr) As Boolean
    End Function

    Private Const WM_INPUTLANGCHANGEREQUEST As UInteger = &H50

    Private Function KeyboardHookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        If nCode < 0 Then Return CallNextHookEx(hookHandle, nCode, wParam, lParam)

        Dim hookStruct As KBDLLHOOKSTRUCT = Marshal.PtrToStructure(lParam, GetType(KBDLLHOOKSTRUCT))
        Dim msg As Integer = wParam.ToInt32()
        Dim vkCode As Integer = hookStruct.vkCode

        Dim liveBar = Application.OpenForms.OfType(Of AppBar)().FirstOrDefault()

        Select Case msg
            Case WM_KEYDOWN, WM_SYSKEYDOWN
                ' Update modifier states
                Select Case vkCode
                    Case VK_WIN_LEFT, VK_WIN_RIGHT : WinKeyPressed = True
                    Case VK_CTRL_LEFT, VK_CTRL_RIGHT : CtrlKeyPressed = True
                    Case VK_SHIFT_LEFT, VK_SHIFT_RIGHT : ShiftKeyPressed = True
                    Case VK_ALT_LEFT, VK_ALT_RIGHT : AltKeyPressed = True

                    Case Keys.Tab
                        Dim AltTabEnabled As Object = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "Enabled", False)
                        Dim isCustomAltTabEnabled As Boolean = Convert.ToBoolean(AltTabEnabled)

                        If (AltKeyPressed AndAlso isCustomAltTabEnabled) OrElse WinKeyPressed Then
                            UsedMultipleKey = True

                            If Not AltTab.Visible Then
                                Dim hWndALTTAB As IntPtr = GetForegroundWindow()
                                Dim processId As UInteger = 0
                                GetWindowThreadProcessId(hWndALTTAB, processId)

                                AltTab.ActiveProcessID = processId
                                AltTab.OpenWithDelay()
                            Else
                                AltTab.MoveNextItem()
                            End If

                            Return New IntPtr(1)
                        End If

                    Case 37 ' LEFT
                        If liveBar.Focused AndAlso liveBar.ProcessStrip.Focused AndAlso liveBar.ProcessStripSelectedItem IsNot Nothing Then
                            UsedMultipleKey = True

                            If ShiftKeyPressed AndAlso AltKeyPressed Then
                                'liveBar.MoveProcessStripItem(liveBar.ProcessStripSelectedItem, 0)
                                'Return New IntPtr(1)
                            End If
                        End If

                    Case 39 ' RIGHT
                        If liveBar.Focused AndAlso liveBar.ProcessStrip.Focused AndAlso liveBar.ProcessStripSelectedItem IsNot Nothing Then
                            UsedMultipleKey = True

                            If ShiftKeyPressed AndAlso AltKeyPressed Then
                                'liveBar.MoveProcessStripItem(liveBar.ProcessStripSelectedItem, 1)
                                'Return New IntPtr(1)
                            End If
                        End If

                    Case Keys.F4
                        If CtrlKeyPressed AndAlso AltKeyPressed Then
                            UsedMultipleKey = True

                            Dim winToKill As IntPtr = GetForegroundWindow()
                            Dim prID As Integer
                            GetWindowThreadProcessId(winToKill, prID)

                            Dim pr As Process = Process.GetProcessById(prID)

                            If pr IsNot Nothing Then
                                If CanSnapWindow(pr.MainWindowHandle) Then
                                    pr.Kill()
                                End If
                            End If
                        End If

                    Case 82 ' R
                        If WinKeyPressed Then
                            UsedMultipleKey = True

                            Dim openRunDialog = Async Sub()
                                                    Await Task.Delay(20)

                                                    If Not RunDialog.Visible Then
                                                        RunDialog.Show()
                                                        Await Task.Delay(30)
                                                        SetForegroundWindow(RunDialog.Handle)
                                                    Else
                                                        RunDialog.Focus()
                                                        Await Task.Delay(30)
                                                        SetForegroundWindow(RunDialog.Handle)
                                                    End If
                                                End Sub

                            liveBar.Invoke(openRunDialog)

                            Return New IntPtr(1)
                        End If

                    Case 188 ' comma (,)
                        If WinKeyPressed Then
                            UsedMultipleKey = True
                            If liveBar.IsDesktopPeekEnabled Then
                                If AeroPeekScreen.Visible Then Return New IntPtr(1)

                                If Not Desktop.IsShowDesktopMode Then liveBar.TimerPeek.Start()

                                Return New IntPtr(1)
                            End If
                        End If

                    Case Keys.Escape
                        If Not ShiftKeyPressed AndAlso Not AltKeyPressed AndAlso CtrlKeyPressed Then
                            RaiseEvent HotkeyPressed(Nothing, EventArgs.Empty)
                            Return New IntPtr(1)
                        End If

                    Case Keys.VolumeUp
                        Dim curVol As Integer = VolumeControl.GetCurrentVolume * 100
                        Dim volAdd As Integer = 2

                        ' up
                        If curVol + 2 >= 100 Then
                            VolumeControl.SetVolume(1)
                        Else
                            VolumeControl.SetVolume((curVol / 100) + (volAdd / 100))
                        End If

                        ScreenDrawer.OnVolumeChanged()

                    Case Keys.VolumeDown
                        Dim curVol As Integer = VolumeControl.GetCurrentVolume * 100
                        Dim volAdd As Integer = 2

                        ' down
                        If curVol - 2 <= 0 Then
                            VolumeControl.SetVolume(0)
                        Else
                            VolumeControl.SetVolume((curVol / 100) - (volAdd / 100))
                        End If

                        ScreenDrawer.OnVolumeChanged()

                    Case Keys.VolumeMute
                        Dim curMuteState As Boolean = VolumeControl.IsSysMuted

                        If curMuteState = True Then
                            VolumeControl.ToggleMute(False)
                        Else
                            VolumeControl.ToggleMute(True)
                        End If

                        ScreenDrawer.OnVolumeChanged()

                End Select

                ' Logic for Combinations
                If WinKeyPressed AndAlso vkCode <> VK_WIN_LEFT AndAlso vkCode <> VK_WIN_RIGHT Then
                    If HandleWindowsShortcuts(vkCode) Then
                        Return New IntPtr(1)
                    End If
                End If


            Case WM_KEYUP, WM_SYSKEYUP
                Select Case vkCode
                    Case 82 ' R
                        If WinKeyPressed Then
                            Return New IntPtr(1)
                        End If

                    Case 188 ' comma (,)
                        UsedMultipleKey = True
                        If WinKeyPressed Then
                            If liveBar.IsDesktopPeekEnabled Then
                                liveBar.TimerPeek.Stop()
                                liveBar.mainPeek.HidePeek()

                                AeroPeekScreen.Hide()
                                Return New IntPtr(1)
                            End If
                        End If

                    Case VK_WIN_LEFT, VK_WIN_RIGHT
                        If AppBar.cm_Language.Visible Then
                            For Each item As ToolStripMenuItem In AppBar.cm_Language.Items
                                If item.Checked Then
                                    Dim targetLang As InputLanguage = DirectCast(item.Tag, InputLanguage)
                                    InputLanguage.CurrentInputLanguage = targetLang

                                    PostMessage(GetForegroundWindow(), WM_INPUTLANGCHANGEREQUEST, IntPtr.Zero, targetLang.Handle)

                                    AppBar.UpdateLanguageDisplay()
                                    Exit For
                                End If
                            Next
                            AppBar.cm_Language.Close()
                        End If

                        If Not UsedMultipleKey Then
                            RaiseEvent HotkeyPressed(Nothing, EventArgs.Empty)
                        End If

                        liveBar.TimerPeek.Stop()
                        liveBar.mainPeek.HidePeek()

                        AeroPeekScreen.Hide()

                        WinKeyPressed = False

                        If AltTab.Visible Then
                            AltTab.ActivateSelectedWindow()
                        End If

                    Case VK_CTRL_LEFT, VK_CTRL_RIGHT : CtrlKeyPressed = False
                    Case VK_ALT_LEFT, VK_ALT_RIGHT
                        AltKeyPressed = False

                        If AltTab.Visible Then
                            AltTab.ActivateSelectedWindow()
                        End If

                    Case VK_SHIFT_LEFT, VK_SHIFT_RIGHT
                        ShiftKeyPressed = False

                        If liveBar.nextItem IsNot Nothing AndAlso liveBar.ProcessStrip.Bounds.Contains(liveBar.ProcessStrip.PointToClient(MousePosition)) Then
                            'liveBar.nextItem.PerformClick()
                        End If

                    Case Keys.Pause
                        If Not WinKeyPressed Then
                            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "PauseKeyToCaptureWindow", 0) = 1 Then
                                Dim hwnd As IntPtr = GetForegroundWindow()

                                Using capturedWindow As Bitmap = liveBar.RenderWindow(hwnd, False)
                                    If capturedWindow IsNot Nothing Then
                                        Dim dataObj As New DataObject()

                                        dataObj.SetData(DataFormats.Bitmap, True, capturedWindow)

                                        Clipboard.SetDataObject(dataObj, True)
                                    Else
                                        MsgBox("Nothing captured from HWND: " & hwnd.ToString())
                                    End If
                                End Using
                            End If
                        End If
                End Select
        End Select

        Return CallNextHookEx(hookHandle, nCode, wParam, lParam)
    End Function

    Public Function SetHook() As Boolean
        hookProcInstance = New HookProc(AddressOf KeyboardHookProc)
        Using process As Process = Process.GetCurrentProcess()
            Using modul As ProcessModule = process.MainModule
                Dim hModule As IntPtr = GetModuleHandle(modul.ModuleName)
                hookHandle = SetWindowsHookEx(WH_KEYBOARD_LL, hookProcInstance, hModule, 0)
            End Using
        End Using
        Return (hookHandle <> IntPtr.Zero)
    End Function

    Public CurrentHotkeyPreset As String = ""
    Private Function HandleWindowsShortcuts(ByVal vkCode As Integer) As Boolean
        Dim hWnd As IntPtr = GetForegroundWindow()
        Dim handled As Boolean = True
        UsedMultipleKey = True

        Dim state As Integer = GetWindowState(hWnd, True)

        Dim isCustomHotkey As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "UseHotkeysForPinned", 0)
        Dim isTextStylizer As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "TextStylizer", False)

        Dim startHotkeyPath As String = "HKEY_CURRENT_USER\Software\Shell\CustomPaths\Hotkeys"
        CurrentHotkeyPreset = My.Computer.Registry.GetValue(startHotkeyPath, "CurrentPreset", String.Empty)

        Dim hotkeyStart As String = startHotkeyPath & "\" & CurrentHotkeyPreset

        Dim liveBar = Application.OpenForms.OfType(Of AppBar)().FirstOrDefault()

        Select Case vkCode
            Case 32 ' Space
                If WinKeyPressed Then
                    UsedMultipleKey = True

                    If Not AppBar.cm_Language.Visible Then
                        AppBar.cm_Language.Show(New Point(SystemInformation.WorkingArea.Width - AppBar.cm_Language.Width, SystemInformation.WorkingArea.Height - AppBar.cm_Language.Height))
                    End If

                    If AppBar.cm_Language.Items.Count > 0 Then
                        Dim currentIdx As Integer = -1

                        For i As Integer = 0 To AppBar.cm_Language.Items.Count - 1
                            If DirectCast(AppBar.cm_Language.Items(i), ToolStripMenuItem).Checked Then
                                currentIdx = i
                                DirectCast(AppBar.cm_Language.Items(i), ToolStripMenuItem).Checked = False
                                Exit For
                            End If
                        Next

                        Dim nextIdx As Integer = (currentIdx + 1) Mod AppBar.cm_Language.Items.Count
                        Dim nextItem = DirectCast(AppBar.cm_Language.Items(nextIdx), ToolStripMenuItem)

                        nextItem.Checked = True
                        nextItem.Select()
                    End If

                    Return True
                End If

            Case 37 ' LEFT
                If CanSnapWindow(hWnd) = True Then
                    If CtrlKeyPressed AndAlso WinKeyPressed Then
                        SmartSnap(hWnd, SnapMode.BottomLeftQuarter)
                    ElseIf ShiftKeyPressed AndAlso WinKeyPressed Then
                        MoveWindowToNextMonitor(hWnd)
                    ElseIf WinKeyPressed Then
                        SmartSnap(hWnd, SnapMode.LeftHalf)
                    End If
                End If

            Case 39 ' RIGHT
                If CanSnapWindow(hWnd) = True Then
                    If CtrlKeyPressed AndAlso WinKeyPressed Then
                        SmartSnap(hWnd, SnapMode.TopRightQuarter)
                    ElseIf ShiftKeyPressed AndAlso WinKeyPressed Then
                        MoveWindowToPreviousMonitor(hWnd)
                    ElseIf WinKeyPressed Then
                        SmartSnap(hWnd, SnapMode.RightHalf)
                    End If
                End If

            Case 38 ' UP
                If CanSnapWindow(hWnd) = True Then
                    If CtrlKeyPressed AndAlso WinKeyPressed Then
                        SmartSnap(hWnd, SnapMode.TopLeftQuarter)
                    ElseIf AltKeyPressed AndAlso WinKeyPressed Then
                        SmartSnap(hWnd, SnapMode.TopQuarter)
                    ElseIf WinKeyPressed Then
                        If state = 2 Then
                            If ShiftKeyPressed Then
                                ShowWindow(hWnd, SW_MAXIMIZE)
                            Else
                                ShowWindow(hWnd, SW_RESTORE)
                            End If

                        ElseIf state = 1 Then
                            ShowWindow(hWnd, SW_MAXIMIZE)
                        End If
                    End If
                End If

            Case 40 ' DOWN
                If CanSnapWindow(hWnd) = True Then
                    If CtrlKeyPressed AndAlso WinKeyPressed Then
                        SmartSnap(hWnd, SnapMode.BottomRightQuarter)
                    ElseIf AltKeyPressed AndAlso WinKeyPressed Then
                        SmartSnap(hWnd, SnapMode.BottomQuarter)
                    ElseIf WinKeyPressed Then
                        If state = 3 Then
                            If ShiftKeyPressed Then
                                ShowWindow(hWnd, SW_RESTORE)
                            Else
                                ShowWindow(hWnd, SW_MINIMIZE)
                            End If
                        Else
                            ShowWindow(hWnd, SW_MINIMIZE)
                        End If
                    End If
                End If

            Case 65, Keys.N ' A / N
                If Not DateAndTime.Visible Then
                    DateAndTime.Show()
                    SetForegroundWindow(DateAndTime.Handle)
                    DateAndTime.Focus()
                Else
                    SetForegroundWindow(DateAndTime.Handle)
                    DateAndTime.Focus()
                End If

            Case 66 ' B
                If CtrlKeyPressed AndAlso ShiftKeyPressed AndAlso WinKeyPressed Then ' WIN + CTRL + Shift + B - Refreshes the Size of the Shell
                    liveBar.Location = New Point(liveBar.Location.X - 1, liveBar.Location.Y - 1)
                    liveBar.Size = New Size(liveBar.Width - 1, liveBar.Height - 1)

                    Desktop.Location = New Point(Desktop.Location.X - 1, Desktop.Location.Y - 1)
                    Desktop.Size = New Size(Desktop.Width - 1, Desktop.Height - 1)
                End If

                If CtrlKeyPressed AndAlso WinKeyPressed Then
                    If isTextStylizer Then
                        TextStylizer.ProcessWin32SelectedText(1, liveBar.ActiveWindowHandle)
                    End If
                End If

                If Not CtrlKeyPressed AndAlso WinKeyPressed Then
                    SetForegroundWindow(liveBar.Handle)
                    liveBar.ToolStrip1.Focus()
                    liveBar.ToolStrip1.Select()
                End If

            Case 67 ' C
                Dim prtoexecute As String = AppBar.EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "Assistant", ""))
                If File.Exists(prtoexecute) Then
                    Process.Start(New ProcessStartInfo(prtoexecute) With {.UseShellExecute = True})
                End If

            Case 68 ' D
                Desktop.ShowDesktop()
                Desktop.BringToFront()

            Case 69 ' E
                If AppBar.UseExplorerFM = True Then
                    Process.Start("explorer.exe", "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}")
                Else
                    If Not String.IsNullOrWhiteSpace(AppBar.CustomFMPath) AndAlso File.Exists(AppBar.CustomFMPath) Then
                        Process.Start(AppBar.CustomFMPath)
                    End If
                End If

            Case 73 ' I
                If CtrlKeyPressed AndAlso WinKeyPressed Then
                    If isTextStylizer Then
                        TextStylizer.ProcessWin32SelectedText(2, liveBar.ActiveWindowHandle)
                    End If
                Else
                    If Not AppbarProperties.Visible Then
                        AppbarProperties.Show(AppBar)
                        SetForegroundWindow(AppbarProperties.Handle)
                        AppbarProperties.Focus()
                    Else
                        SetForegroundWindow(AppbarProperties.Handle)
                        AppbarProperties.Focus()
                    End If
                End If

            Case 76 ' L
                Process.Start("RunDll32.exe", "user32.dll,LockWorkStation")

            Case 77 ' M

                liveBar?.Invoke(Sub()
                                    If CtrlKeyPressed AndAlso WinKeyPressed Then
                                        For Each btn In liveBar.ProcessStrip.Items.OfType(Of ToolStripButton)()
                                            If btn.Tag IsNot Nothing AndAlso TypeOf btn.Tag Is IntPtr Then
                                                Dim windowHandle As IntPtr = CType(btn.Tag, IntPtr)
                                                RestoreWindow(windowHandle)
                                            End If
                                        Next
                                    ElseIf WinKeyPressed Then
                                        For Each btn In liveBar.ProcessStrip.Items.OfType(Of ToolStripButton)()
                                            If btn.Tag IsNot Nothing AndAlso TypeOf btn.Tag Is IntPtr Then
                                                Dim windowHandle As IntPtr = CType(btn.Tag, IntPtr)
                                                MinimizeWindow(windowHandle)
                                            End If
                                        Next
                                    End If
                                End Sub)

            Case 82 ' R
                'Rest in Peace

            Case 83 ' S
                If CtrlKeyPressed AndAlso WinKeyPressed Then
                    If isTextStylizer Then
                        TextStylizer.ProcessWin32SelectedText(4, liveBar.ActiveWindowHandle)
                    End If
                Else
                    Dim prtoexecute As String = AppBar.EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "SearchApp", ""))
                    If File.Exists(prtoexecute) Then
                        Process.Start(New ProcessStartInfo(prtoexecute) With {.UseShellExecute = True})
                    End If
                End If

            Case 84 ' T
                SetForegroundWindow(AppBar.Handle)

                AppBar.ProcessStrip.Focus()
                AppBar.ProcessStrip.Select()

            Case 85 ' U
                If CtrlKeyPressed AndAlso WinKeyPressed Then
                    If isTextStylizer Then
                        TextStylizer.ProcessWin32SelectedText(3, liveBar.ActiveWindowHandle)
                    End If
                End If

            Case 86 ' V
                If Not ClipboardViewer.Visible Then
                    ClipboardViewer.Show()
                    SetForegroundWindow(ClipboardViewer.Handle)
                    ClipboardViewer.Focus()
                Else
                    SetForegroundWindow(ClipboardViewer.Handle)
                    ClipboardViewer.Focus()
                End If

            Case 88 ' X
                If Not AppBar.cm_Action.Visible Then
                    AppBar.cm_Action.Show(AppBar, AppBar.Location)
                    AppBar.cm_Action.BringToFront()
                    AppBar.cm_Action.Focus()
                Else
                    AppBar.cm_Action.BringToFront()
                    AppBar.cm_Action.Focus()
                End If

            Case 89 ' Y
                If hWnd <> IntPtr.Zero Then
                    ToggleAlwaysOnTop(hWnd)
                Else
                    MessageBox.Show("No active window has been found.", "Info")
                End If

#Region " Custom WIN hotkey "
            ' CUSTOM SAVES OF CUSTOM PROGRAMS YOU MADE IN APPBAR PROPERTIES
            Case Keys.D1, Keys.NumPad1 ' 1

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 10)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(0, CtrlKeyPressed)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "1", "")
                        Dim psi As New ProcessStartInfo(executePath) With {.UseShellExecute = True}
                        If CtrlKeyPressed Then psi.Verb = "runas"

                        Process.Start(psi)
                    End If
                End If

            Case Keys.D2, Keys.NumPad2 ' 2

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 20)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(1, CtrlKeyPressed)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "2", "")
                        Dim psi As New ProcessStartInfo(executePath) With {.UseShellExecute = True}
                        If CtrlKeyPressed Then psi.Verb = "runas"

                        Process.Start(psi)
                    End If
                End If

            Case Keys.D3, Keys.NumPad3 ' 3

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 30)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(2, CtrlKeyPressed)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "3", "")
                        Dim psi As New ProcessStartInfo(executePath) With {.UseShellExecute = True}
                        If CtrlKeyPressed Then psi.Verb = "runas"

                        Process.Start(psi)
                    End If
                End If

            Case Keys.D4, Keys.NumPad4 ' 4

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 40)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(3, CtrlKeyPressed)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "4", "")
                        Dim psi As New ProcessStartInfo(executePath) With {.UseShellExecute = True}
                        If CtrlKeyPressed Then psi.Verb = "runas"

                        Process.Start(psi)
                    End If
                End If

            Case Keys.D5, Keys.NumPad5 ' 5

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 50)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(4, CtrlKeyPressed)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "5", "")
                        Dim psi As New ProcessStartInfo(executePath) With {.UseShellExecute = True}
                        If CtrlKeyPressed Then psi.Verb = "runas"

                        Process.Start(psi)
                    End If
                End If

            Case Keys.D6, Keys.NumPad6 ' 6

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 60)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(5, CtrlKeyPressed)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "6", "")
                        Dim psi As New ProcessStartInfo(executePath) With {.UseShellExecute = True}
                        If CtrlKeyPressed Then psi.Verb = "runas"

                        Process.Start(psi)
                    End If
                End If

            Case Keys.D7, Keys.NumPad7 ' 7

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 70)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(6, CtrlKeyPressed)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "7", "")
                        Dim psi As New ProcessStartInfo(executePath) With {.UseShellExecute = True}
                        If CtrlKeyPressed Then psi.Verb = "runas"

                        Process.Start(psi)
                    End If
                End If

            Case Keys.D8, Keys.NumPad8 ' 8

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 80)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(7, CtrlKeyPressed)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "8", "")
                        Dim psi As New ProcessStartInfo(executePath) With {.UseShellExecute = True}
                        If CtrlKeyPressed Then psi.Verb = "runas"

                        Process.Start(psi)
                    End If
                End If

            Case Keys.D9, Keys.NumPad9 ' 9

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 90)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(8, CtrlKeyPressed)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "9", "")
                        Dim psi As New ProcessStartInfo(executePath) With {.UseShellExecute = True}
                        If CtrlKeyPressed Then psi.Verb = "runas"

                        Process.Start(psi)
                    End If
                End If

            Case Keys.D0, Keys.NumPad0 ' 0

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 100)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(9, CtrlKeyPressed)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "0", "")
                        Dim psi As New ProcessStartInfo(executePath) With {.UseShellExecute = True}
                        If CtrlKeyPressed Then psi.Verb = "runas"

                        Process.Start(psi)
                    End If
                End If
#End Region

            Case Keys.Enter
                If CtrlKeyPressed Then
                    Dim narratorPath As String = AppBar.EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "Narrator", Environment.GetFolderPath(Environment.SpecialFolder.Windows)) & "\System32\Narrator.exe")
                    Try
                        If File.Exists(narratorPath) Then
                            Process.Start(New ProcessStartInfo(narratorPath) With {.UseShellExecute = True})
                        End If
                    Catch
                    End Try
                End If

                If ShiftKeyPressed Then
                    If WindowFullscreenHelper.IsAppFullscreen(hWnd) Then
                        WindowFullscreenHelper.ToggleFullscreen(hWnd)
                    Else
                        If CanSnapWindow(hWnd) Then
                            WindowFullscreenHelper.ToggleFullscreen(hWnd)
                        End If
                    End If

                End If

            Case Keys.Home

                liveBar?.Invoke(Sub()
                                    Dim currentForegroundWindow As IntPtr = GetForegroundWindow()

                                    If CtrlKeyPressed AndAlso WinKeyPressed Then
                                        For Each btn In liveBar.ProcessStrip.Items.OfType(Of ToolStripButton)()
                                            If btn.Tag IsNot Nothing AndAlso TypeOf btn.Tag Is IntPtr Then
                                                Dim windowHandle As IntPtr = CType(btn.Tag, IntPtr)

                                                If windowHandle = currentForegroundWindow Then Continue For

                                                RestoreWindow(windowHandle)
                                            End If
                                        Next
                                    ElseIf WinKeyPressed Then
                                        For Each btn In liveBar.ProcessStrip.Items.OfType(Of ToolStripButton)()
                                            If btn.Tag IsNot Nothing AndAlso TypeOf btn.Tag Is IntPtr Then
                                                Dim windowHandle As IntPtr = CType(btn.Tag, IntPtr)

                                                If windowHandle = currentForegroundWindow Then Continue For

                                                MinimizeWindow(windowHandle)
                                            End If
                                        Next
                                    End If

                                    If currentForegroundWindow <> IntPtr.Zero Then
                                        SetForegroundWindow(currentForegroundWindow)
                                    End If
                                End Sub)

            Case Keys.Pause
                If AboutDialog.Visible Then
                    AboutDialog.Focus()
                Else
                    AboutDialog.Show()
                End If

            Case Else
                handled = False
        End Select

        Return handled
    End Function

    Public Function Unhook() As Boolean
        Dim result As Boolean = UnhookWindowsHookEx(hookHandle)
        hookHandle = IntPtr.Zero
        Return result
    End Function

    Private Const WM_KEYDOWN As Integer = &H100
    Private Const WM_KEYUP As Integer = &H101

    <DllImport("user32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Function GetKeyState(ByVal nVirtKey As Integer) As Short
    End Function
End Module

Public Class WindowFullscreenHelper

    Private Const GWL_STYLE As Integer = -16
    Private Const WS_CAPTION As UInteger = &HCC0000
    Private Const WS_THICKFRAME As UInteger = &H40000
    Private Const WS_DLGFRAME As UInteger = &H400000
    Private Const WS_BORDER As UInteger = &H800000

    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOZORDER As UInteger = &H4
    Private Const SWP_FRAMECHANGED As UInteger = &H20
    Private Const SWP_SHOWWINDOW As UInteger = &H40
    Private Const SW_SHOWNORMAL As Integer = 1

    Private Shared ReadOnly HWND_TOPMOST As IntPtr = New IntPtr(-1)
    Private Shared ReadOnly HWND_NOTOPMOST As IntPtr = New IntPtr(-2)

    Private Const ORIGINAL_STYLE_PROP As String = "KrrShell_OriginalStyle"
    Private Const ORIGINAL_PLACEMENT_PROP As String = "KrrShell_OriginalPlacement"
    Private Const IS_FULLSCREEN_PROP As String = "KrrShell_IsFullscreen"

    <StructLayout(LayoutKind.Sequential)>
    Private Structure POINT
        Public X As Integer
        Public Y As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure WINDOWPLACEMENT
        Public length As Integer
        Public flags As Integer
        Public showCmd As Integer
        Public ptMinPosition As POINT
        Public ptMaxPosition As POINT
        Public rcNormalPosition As RECT
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Private Structure MONITORINFO
        Public cbSize As Integer
        Public rcMonitor As RECT
        Public rcWork As RECT
        Public dwFlags As Integer
    End Structure

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetWindowPos(hWnd As IntPtr, hWndInsertAfter As IntPtr, X As Integer, Y As Integer, cx As Integer, cy As Integer, uFlags As UInteger) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowPlacement(hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetWindowPlacement(hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function MonitorFromWindow(hWnd As IntPtr, dwFlags As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function GetMonitorInfo(hMonitor As IntPtr, ByRef lpmi As MONITORINFO) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Private Shared Function SetProp(hWnd As IntPtr, lpString As String, hData As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Private Shared Function GetProp(hWnd As IntPtr, lpString As String) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Private Shared Function RemoveProp(hWnd As IntPtr, lpString As String) As IntPtr
    End Function

    Public Shared Function IsAppFullscreen(hWnd As IntPtr) As Boolean
        If hWnd = IntPtr.Zero Then Return False
        Return GetProp(hWnd, IS_FULLSCREEN_PROP) <> IntPtr.Zero
    End Function

    Public Shared Sub ToggleFullscreen(hWnd As IntPtr)
        If hWnd = IntPtr.Zero Then Return

        If Not IsAppFullscreen(hWnd) Then
            Dim originalStyle As Integer = GetWindowLong(hWnd, GWL_STYLE)
            SetProp(hWnd, ORIGINAL_STYLE_PROP, New IntPtr(originalStyle))

            Dim wp As New WINDOWPLACEMENT()
            wp.length = Marshal.SizeOf(wp)
            If GetWindowPlacement(hWnd, wp) Then
                Dim ptr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(wp))
                Marshal.StructureToPtr(wp, ptr, False)
                SetProp(hWnd, ORIGINAL_PLACEMENT_PROP, ptr)
            End If

            Dim hMonitor As IntPtr = MonitorFromWindow(hWnd, 1)
            Dim mi As New MONITORINFO()
            mi.cbSize = Marshal.SizeOf(mi)
            GetMonitorInfo(hMonitor, mi)

            SetProp(hWnd, IS_FULLSCREEN_PROP, New IntPtr(1))

            Dim newStyle As Integer = originalStyle And Not CInt(WS_THICKFRAME Or WS_DLGFRAME Or WS_BORDER)
            SetWindowLong(hWnd, GWL_STYLE, newStyle)

            ShowWindow(hWnd, SW_SHOWNORMAL)

            Dim width As Integer = mi.rcMonitor.Right - mi.rcMonitor.Left
            Dim height As Integer = mi.rcMonitor.Bottom - mi.rcMonitor.Top

            SetWindowPos(hWnd, HWND_TOPMOST, mi.rcMonitor.Left, mi.rcMonitor.Top, width, height, SWP_FRAMECHANGED Or SWP_SHOWWINDOW)
        Else
            Dim savedStylePtr As IntPtr = GetProp(hWnd, ORIGINAL_STYLE_PROP)
            Dim savedPlacementPtr As IntPtr = GetProp(hWnd, ORIGINAL_PLACEMENT_PROP)

            If savedStylePtr <> IntPtr.Zero Then
                Dim originalStyle As Integer = savedStylePtr.ToInt32()
                SetWindowLong(hWnd, GWL_STYLE, originalStyle)
                RemoveProp(hWnd, ORIGINAL_STYLE_PROP)
            End If

            SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED)

            If savedPlacementPtr <> IntPtr.Zero Then
                Dim wp As WINDOWPLACEMENT = Marshal.PtrToStructure(Of WINDOWPLACEMENT)(savedPlacementPtr)
                SetWindowPlacement(hWnd, wp)

                Marshal.FreeHGlobal(savedPlacementPtr)
                RemoveProp(hWnd, ORIGINAL_PLACEMENT_PROP)
            End If

            SetWindowPos(hWnd, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
            RemoveProp(hWnd, IS_FULLSCREEN_PROP)
        End If
    End Sub
End Class

' TRAY ICONS WORKING
Public Class TrayIconInfo
    Public Property HWnd As IntPtr
    Public Property UID As UInteger
    Public Property CallbackMessage As UInteger
    Public Property Tooltip As String
    Public Property ProcessId As UInteger
    Public Property ProcessName As String
    Public Property Icon As Bitmap
    Public Property SenderIs64Bit As Boolean
    Public Property LastBalloonTitle As String
    Public Property LastBalloonText As String
    Public Property LastBalloonFlags As Integer
    Public Property LastBalloonTime As DateTime
End Class

Public Class TrayIconEventArgs
    Inherits EventArgs
    Public Sub New(icon As TrayIconInfo)
        Me.Icon = icon
    End Sub
    Public ReadOnly Property Icon As TrayIconInfo
End Class

Public Class BalloonEventArgs
    Inherits EventArgs
    Public Sub New(icon As TrayIconInfo)
        Me.Icon = icon
    End Sub
    Public ReadOnly Property Icon As TrayIconInfo
    Public Property Title As String
    Public Property Body As String
    Public Property Flags As Integer
End Class