Imports System.ComponentModel
Imports System.Configuration.Assemblies
Imports System.Drawing.Drawing2D
Imports System.Globalization
Imports System.IO
Imports System.Net.Security
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
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

        If m.Msg = WM_MOUSEACTIVATE Then
            Dim clickedControl As Control = Me.GetChildAtPoint(Me.PointToClient(Cursor.Position))

            If clickedControl IsNot Nothing AndAlso clickedControl Is Me.ProcessStrip Then
                m.Result = CType(MA_NOACTIVATE, IntPtr)
            End If
        End If

        MyBase.WndProc(m)
    End Sub

    <Runtime.InteropServices.DllImport("dwmapi.dll")> Public Shared Function DwmExtendFrameIntoClientArea(ByVal hWnd As IntPtr, ByRef pMarinset As Side) As Integer
    End Function

    <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)> Public Structure Side
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

    Public Function RenderWindow(hWnd As IntPtr, clientAreaOnly As Boolean) As Bitmap
        Dim windowRect As New RECT()
        DwmGetWindowAttribute(hWnd, 9, windowRect, Marshal.SizeOf(Of RECT)()) ' 9 = DWMWA_EXTENDED_FRAME_BOUNDS

        If windowRect.Width <= 0 OrElse windowRect.Height <= 0 Then
            Return GeneratePlaceholder(hWnd)
        End If

        Dim bmp As New Bitmap(windowRect.Width, windowRect.Height)

        Using g = Graphics.FromImage(bmp)
            Dim hdcDest = g.GetHdc()

            Dim success As Boolean = PrintWindow(hWnd, hdcDest, &H2)

            If Not success OrElse IsResultBlack(bmp) Then
                PrintWindow(hWnd, hdcDest, If(clientAreaOnly, &H1, 0))
            End If

            g.ReleaseHdc(hdcDest)
        End Using

        If IsResultBlack(bmp) Then
            Return GeneratePlaceholder(hWnd)
        End If

        Return bmp
    End Function

    Private Shared Function IsResultBlack(bmp As Bitmap) As Boolean
        Dim c = bmp.GetPixel(bmp.Width / 2, bmp.Height / 2)
        Return c.R = 0 AndAlso c.G = 0 AndAlso c.B = 0 AndAlso c.A = 255
    End Function

    Private Shared Function GeneratePlaceholder(hWnd As IntPtr) As Bitmap
        Dim w As Integer = 200
        Dim h As Integer = 120

        Dim appIcon As Icon = Icon.ExtractAssociatedIcon(GetProcessPath(hWnd))
        Dim bmp As New Bitmap(w, h)

        Using g = Graphics.FromImage(bmp)
            g.Clear(GetAccentColor)

            If appIcon IsNot Nothing Then
                g.DrawIcon(appIcon, New Rectangle((w \ 2) - 16, (h \ 2) - 16, 32, 32))
            End If

            TextRenderer.DrawText(g, "Minimized", SystemFonts.CaptionFont, New Point(5, h - 20), Color.White)
        End Using

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

    Public Delegate Function EnumWindowsProc(ByVal hwnd As IntPtr, ByVal lParam As IntPtr) As Boolean

    Public Class TaskbarWindowInfo
        Public Property Handle As IntPtr
        Public Property Title As String
        Public Property PID As Integer
    End Class

    Private ReadOnly BlacklistedProcesses As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

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
                                   btn.ForeColor = Control.DefaultForeColor
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

    Private Function GetWindowTitleSafe(hwnd As IntPtr) As String
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

    Private Sub AddToolStripButton(hwnd As IntPtr, title As String)
        If ProcessStrip.Items.Cast(Of ToolStripItem).Any(Function(x) x.Tag.Equals(hwnd)) Then Return

        Dim btn As New ToolStripButton(title) With {
            .Tag = hwnd,
            .DisplayStyle = ToolStripItemDisplayStyle.Image
        }

        Dim windowImage As Image = GetWindowIcon(hwnd, IconSize.Large)?.ToBitmap()
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
                btn.Size = New Size(48, Me.Height)

            Case 1
                btn.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText
                btn.Size = New Size(150, Me.Height)

                Dim maxWidth As Integer = 20
                If btn.Text.Length > maxWidth Then
                    btn.Text = btn.Text.Substring(0, maxWidth - 3) & "..."
                End If

                btn.TextImageRelation = TextImageRelation.ImageBeforeText
                btn.TextAlign = ContentAlignment.MiddleLeft
                btn.ImageAlign = ContentAlignment.MiddleLeft

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

        AddHandler btn.MouseUp, AddressOf ProcessStripItemClick
        AddHandler btn.MouseEnter, AddressOf ProcessStripItemMouseEnter
        AddHandler btn.MouseLeave, AddressOf ProcessStripItemMouseLeave
        AddHandler btn.MouseMove, Sub(sender As Object, e As MouseEventArgs)
                                      ProcessStrip.Focus()
                                  End Sub

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

#Region " Notify Tray Icons " ' UNDER DESTRUCTION
    Private Const WM_COPYDATA As Integer = &H4A
    Private Const WS_POPUP As Integer = &H80000000
    Private Shared ReadOnly WM_TASKBARCREATED As Integer = RegisterWindowMessage("TaskbarCreated")

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto, Pack:=8)>
    Public Structure NOTIFYICONDATA_X64
        Public cbSize As Integer
        Public padding1 As Integer ' Těchto 4 bajtů opraví posun mezi cbSize a hWnd
        Public hWnd As IntPtr
        Public uID As Integer
        Public uFlags As Integer
        Public uCallbackMessage As Integer
        Public padding2 As Integer ' Dalších 4 bajtů pro zarovnání před hIcon
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

    ' Delegate for Window Procedure
    Private Delegate Function WndProcDelegate(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr

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

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure COPYDATASTRUCT
        Public dwData As IntPtr
        Public cbData As Integer
        Public lpData As IntPtr
    End Structure

    Private _myWndProcDelegate As WndProcDelegate
    Private _fakeTrayHandle As IntPtr

    Private Sub CreateTrayWindow()
        _myWndProcDelegate = AddressOf FakeTrayWndProc

        Dim wc As New WNDCLASSEX()
        wc.cbSize = Marshal.SizeOf(wc)
        wc.lpfnWndProc = _myWndProcDelegate
        wc.lpszClassName = "Shell_TrayWnd"
        wc.hInstance = Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly().GetModules()(0))

        If RegisterClassEx(wc) <> 0 Or Marshal.GetLastWin32Error() = 1410 Then ' 1410 = Class already exists
            _fakeTrayHandle = CreateWindowEx(0, "Shell_TrayWnd", "", WS_POPUP, 0, 0, 0, 0, IntPtr.Zero, IntPtr.Zero, wc.hInstance, IntPtr.Zero)

            If _fakeTrayHandle <> IntPtr.Zero Then
                Debug.WriteLine("SUCCESS: Real Shell_TrayWnd created! Handle: " & _fakeTrayHandle.ToString())

                Dim HWND_BROADCAST As New IntPtr(&HFFFF)
                PostMessage(HWND_BROADCAST, WM_TASKBARCREATED, IntPtr.Zero, IntPtr.Zero)
                'SendMessage(HWND_BROADCAST, WM_TASKBARCREATED, IntPtr.Zero, IntPtr.Zero)
            Else
                Debug.WriteLine("CreateWindowEx failed. Error: " & Marshal.GetLastWin32Error())
            End If
        Else
            Debug.WriteLine("RegisterClassEx failed. Error: " & Marshal.GetLastWin32Error())
        End If
    End Sub

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function CopyIcon(ByVal hIcon As IntPtr) As IntPtr
    End Function

    Private Sub ProcessTrayMessage(lParam As IntPtr)
        Try
            Dim cds As COPYDATASTRUCT = Marshal.PtrToStructure(Of COPYDATASTRUCT)(lParam)
            If cds.cbData < 100 Then Return

            Dim sig As Integer = Marshal.ReadInt32(cds.lpData, 0)

            If sig = &H34753423 Then
                Dim uCallback As Integer = Marshal.ReadInt32(cds.lpData, 8)
                Dim mainHWnd As Integer = Marshal.ReadInt32(cds.lpData, 12)
                Dim alt1 As Integer = Marshal.ReadInt32(cds.lpData, 24)
                Dim alt2 As Integer = Marshal.ReadInt32(cds.lpData, 28)
                Dim uID As Integer = Marshal.ReadInt32(cds.lpData, 16)

                Dim realHWnd As New IntPtr(mainHWnd)
                Dim szTip As String = Marshal.PtrToStringUni(IntPtr.Add(cds.lpData, 32), 128).Split(ControlChars.NullChar)(0)

                'Debug.WriteLine($">>> FINAL TRY: hWnd={realHWnd.ToString("X")}, Text={szTip}")

                Me.BeginInvoke(Sub()
                                   UpdateUIWithInfo(New IntPtr(mainHWnd), uID, szTip, Nothing, New IntPtr(alt1), New IntPtr(alt2))
                               End Sub)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Const GCLP_HICONSM As Integer = -34

    <DllImport("user32.dll", EntryPoint:="GetClassLongPtrW")>
    Private Shared Function GetClassLongPtr(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
    End Function

    Private Function GetIconFromWindow(ByVal hWnd As IntPtr) As Bitmap
        Try
            Dim hIcon As IntPtr = SendMessage(hWnd, WM_GETICON, New IntPtr(ICON_SMALL), IntPtr.Zero)

            If hIcon = IntPtr.Zero Then
                hIcon = GetClassLongPtr(hWnd, GCLP_HICONSM)
            End If

            If hIcon <> IntPtr.Zero Then
                Using ico As Icon = Icon.FromHandle(hIcon)
                    Return ico.ToBitmap()
                End Using
            End If
        Catch

        End Try
        Return Nothing
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function SendMessageTimeout(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr, ByVal flags As Integer, ByVal timeout As Integer, ByRef lpdwResult As IntPtr) As IntPtr
    End Function

    Private Function GetIconAggressive(ByVal hWnd As IntPtr) As Bitmap
        Try
            Dim hIcon As IntPtr = IntPtr.Zero
            Dim result As IntPtr

            SendMessageTimeout(hWnd, WM_GETICON, New IntPtr(ICON_SMALL), IntPtr.Zero, 2, 100, result)
            hIcon = result

            If hIcon = IntPtr.Zero Then hIcon = GetClassLongPtr(hWnd, GCLP_HICONSM)

            If hIcon <> IntPtr.Zero Then
                Return Icon.FromHandle(hIcon).ToBitmap()
            Else
                Dim pid As Integer
                GetWindowThreadProcessId(hWnd, pid)
                If pid <> 0 Then
                    Using proc = Process.GetProcessById(pid)
                        Dim path = proc.MainModule.FileName
                        Return Icon.ExtractAssociatedIcon(path).ToBitmap()
                    End Using
                End If
            End If
        Catch
        End Try
        Return Nothing
    End Function

    Public Class TrayItemData
        Public hWnd As IntPtr
        Public hWndAlt1 As IntPtr ' Offset 24
        Public hWndAlt2 As IntPtr ' Offset 28
        Public uID As Integer
        Public uCallback As Integer
    End Class

    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_LBUTTONUP As Integer = &H202
    Private Const WM_RBUTTONDOWN As Integer = &H204
    Private Const WM_RBUTTONUP As Integer = &H205
    Private Const WM_MOUSEMOVE As Integer = &H200

    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function AllowSetForegroundWindow(ByVal dwProcessId As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindow(ByVal hWnd As IntPtr, ByVal uCmd As Integer) As IntPtr
    End Function
    Private Const GW_OWNER As Integer = 4

    Private Delegate Function EnumThreadDelegate(ByVal hWnd As IntPtr, ByVal lParam As IntPtr) As Boolean
    <DllImport("user32.dll")>
    Private Shared Function EnumThreadWindows(ByVal dwThreadId As Integer, ByVal lpfn As EnumThreadDelegate, ByVal lParam As IntPtr) As Boolean
    End Function

    Private Function GetProcessWindows(ByVal p As Process) As List(Of IntPtr)
        Dim handless As New List(Of IntPtr)
        For Each thread As ProcessThread In p.Threads
            EnumThreadWindows(thread.Id, Function(hWnd, lParam)
                                             handless.Add(hWnd)
                                             Return True
                                         End Function, IntPtr.Zero)
        Next
        Return handless
    End Function

    Private Function PackCoords(ByVal x As Integer, ByVal y As Integer) As IntPtr
        Return New IntPtr((y << 16) Or (x And &HFFFF))
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SendNotifyMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean
    End Function

    Private Sub UpdateUIWithInfo(ByVal rawHWnd As IntPtr, ByVal id As Integer, ByVal text As String, ByVal bmp As Bitmap, ByVal alt1 As IntPtr, ByVal alt2 As IntPtr)
        Dim cleanHWnd As New IntPtr(rawHWnd.ToInt64() And &HFFFFFFFF)
        Dim key As String = "tray_" & cleanHWnd.ToString() & "_" & id.ToString()

        If ToolStrip1.Items.ContainsKey(key) Then Return

        Dim item As New ToolStripButton With {
        .Name = key,
        .AutoSize = False,
        .Size = New Drawing.Size(24, 39),
        .BackgroundImageLayout = ImageLayout.Zoom,
        .DisplayStyle = ToolStripItemDisplayStyle.Image,
        .ImageScaling = ToolStripItemImageScaling.None,
        .ToolTipText = text,
        .Margin = New Padding(1)
    }

        Dim trayData As New TrayItemData With {
        .hWnd = rawHWnd,
        .uID = id,
        .uCallback = &H3BC
    }
        item.Tag = trayData

        If bmp IsNot Nothing Then
            item.BackgroundImage = bmp
        Else
            Task.Run(Sub()
                         Dim bmpLoaded = GetIconAggressive(rawHWnd)
                         If bmpLoaded IsNot Nothing Then
                             Me.BeginInvoke(Sub() item.BackgroundImage = bmpLoaded)
                         End If
                     End Sub)
        End If

        AddHandler item.MouseEnter, Sub(sender As Object, e As EventArgs)
                                        Dim btn = DirectCast(sender, ToolStripButton)
                                        ToolStrip1.Focus()
                                    End Sub
        AddHandler item.MouseUp, Sub(sender As Object, e As MouseEventArgs)
                                     Dim btn = DirectCast(sender, ToolStripButton)
                                     Dim data = DirectCast(btn.Tag, TrayItemData)

                                     Dim pid As Integer = 0
                                     GetWindowThreadProcessId(data.hWnd, pid)

                                     Dim proc As Process = Nothing
                                     Try
                                         proc = Process.GetProcessById(pid)
                                     Catch
                                         Return
                                     End Try

                                     Dim targetHWnd As IntPtr = If(proc.MainWindowHandle <> IntPtr.Zero, proc.MainWindowHandle, data.hWnd)

                                     ' --- LEFT CLICK ---
                                     If e.Button = MouseButtons.Left Then
                                         If IsWindow(targetHWnd) AndAlso IsWindowVisible(targetHWnd) Then
                                             ShowWindow(targetHWnd, SHOW_WINDOW.SW_RESTORE)
                                             SetForegroundWindow(targetHWnd)
                                         Else
                                             Try
                                                 Dim exePath As String = proc.MainModule.FileName
                                                 Process.Start(exePath)
                                             Catch ex As Exception
                                                 Debug.WriteLine("Failed to start process: " & ex.Message)
                                             End Try
                                         End If

                                         ' --- RIGHT CLICK ---
                                     ElseIf e.Button = MouseButtons.Right Then
                                         Dim ctx As New ContextMenuStrip()

                                         ctx.Items.Add(New ToolStripMenuItem("Show && Focus", Nothing, Sub()
                                                                                                           ShowWindow(targetHWnd, SHOW_WINDOW.SW_RESTORE)
                                                                                                           SetForegroundWindow(targetHWnd)
                                                                                                       End Sub))

                                         ctx.Items.Add(New ToolStripMenuItem("Hide Window", Nothing, Sub()
                                                                                                         ShowWindow(targetHWnd, SHOW_WINDOW.SW_HIDE)
                                                                                                     End Sub))

                                         ctx.Items.Add(New ToolStripSeparator())

                                         Dim infoItem As New ToolStripMenuItem($"Process: {proc.ProcessName} (PID: {pid})") With {
                                             .Enabled = False
                                         }
                                         ctx.Items.Add(infoItem)

                                         Try
                                             Dim exePath As String = proc.MainModule.FileName
                                             ctx.Items.Add(New ToolStripMenuItem("Launch New Instance", Nothing, Sub()
                                                                                                                     Try
                                                                                                                         Process.Start(exePath)
                                                                                                                     Catch ex As Exception
                                                                                                                         Debug.WriteLine("Error while start process: " & ex.Message)
                                                                                                                     End Try
                                                                                                                 End Sub))
                                         Catch : End Try

                                         ctx.Items.Add(New ToolStripSeparator())

                                         Dim killBtn As New ToolStripMenuItem("Kill Process") With {
                                             .ForeColor = Color.Red
                                         }
                                         AddHandler killBtn.Click, Sub()
                                                                       If MessageBox.Show($"Are you sure do you want this process ""{proc.ProcessName}""?", "Confirm box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
                                                                           Try : proc.Kill() : Catch : End Try
                                                                       End If
                                                                   End Sub
                                         ctx.Items.Add(killBtn)

                                         ctx.Show(Cursor.Position)
                                     End If
                                 End Sub

        ToolStrip1.Items.Add(item)
    End Sub

    Private Sub TrayCleanup()

        Try
            For i As Integer = ToolStrip1.Items.Count - 1 To 0 Step -1
                Dim item = ToolStrip1.Items(i)

                If TypeOf item.Tag Is TrayItemData Then
                    Dim data = DirectCast(item.Tag, TrayItemData)
                    Dim pid As Integer = 0

                    Try
                        GetWindowThreadProcessId(data.hWnd, pid)

                        Dim procExists As Boolean = False
                        If pid <> 0 Then
                            Try
                                Using p = Process.GetProcessById(pid)
                                    If Not p.HasExited Then procExists = True
                                End Using
                            Catch

                            End Try
                        End If

                        If Not procExists Then
                            Debug.WriteLine($"Removing dead Tray icon: {item.ToolTipText}")
                            ToolStrip1.Items.RemoveAt(i)
                        End If

                    Catch ex As Exception
                        If Me.InvokeRequired = True OrElse Me.IsHandleCreated Then
                            Try : ToolStrip1.Items.RemoveAt(i) : Catch ex2 As Exception
                            End Try
                        End If
                    End Try
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine("Error in Tray item Data: " & ex.Message)
        End Try

    End Sub

    ' This is the message loop for the fake Shell_TrayWnd
    Private Function FakeTrayWndProc(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        If msg = WM_COPYDATA Then
            ProcessTrayMessage(lParam)
        End If
        Return DefWindowProc(hWnd, msg, wParam, lParam)
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
    Public Function GetFileIcon(ByVal filePath As String, ByVal isLargeIcon As Boolean) As Icon
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

    Public Function GetFolderIcon(ByVal folderPath As String, ByVal isLargeIcon As Boolean) As Icon
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

    <DllImport("user32.dll")>
    Private Shared Function IsWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function DrawAnimatedRects(hwnd As IntPtr, idAni As Integer, ByRef lprcFrom As RECT, ByRef lprcTo As RECT) As Boolean
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

        Me.Invoke(Sub() CleanupDeadWindows())

        If idObject <> OBJID_WINDOW Then Return

        Select Case eventType
            Case EVENT_SYSTEM_FOREGROUND
                Me.Invoke(Sub() TrayCleanup())

                ActiveWindowHandle = hwnd

                Task.Delay(100).ContinueWith(Sub()
                                                 Me.Invoke(Sub() UpdateTaskbarSelection(hwnd))
                                             End Sub)

            Case EVENT_SYSTEM_ALERT, EVENT_OBJECT_STATECHANGE
                If idObject = OBJID_WINDOW Then
                    Me.Invoke(Sub()
                                  TrayCleanup()

                                  CheckIfShouldFlash(hwnd)
                              End Sub)
                End If

            Case EVENT_OBJECT_CREATE
                Me.Invoke(Sub() TrayCleanup())

                Task.Delay(100).ContinueWith(Sub()
                                                 Me.Invoke(Sub()
                                                               If IsTaskbarWindow(hwnd) Then
                                                                   AddToolStripButton(hwnd, GetWindowTitleSafe(hwnd))
                                                               End If
                                                           End Sub)
                                             End Sub)


            Case EVENT_OBJECT_DESTROY
                RemoveToolStripButton(hwnd)

                Me.Invoke(Sub() TrayCleanup())

            Case EVENT_OBJECT_NAMECHANGE
                Me.Invoke(Sub()
                              TrayCleanup()

                              UpdateWindowInfo(hwnd)
                          End Sub)

            Case EVENT_OBJECT_LOCATIONCHANGE
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

    Public UseExplorerFP As Boolean = True
    Public UseExplorerFM As Boolean = True
    Public CustomFMPath As String = String.Empty

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

        ' Other
        Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layout", 2)
            Case 0
                Try
                    Dim side As side = side
                    side.Left = -1
                    side.Right = -1
                    side.Top = -1
                    side.Bottom = -1
                    Dim result As Integer = DwmExtendFrameIntoClientArea(Me.Handle, side)
                Catch ex As Exception
                    Dim side As side = side
                    side.Left = 0
                    side.Right = 0
                    side.Top = 0
                    side.Bottom = 0
                    Dim result As Integer = DwmExtendFrameIntoClientArea(Me.Handle, side)
                End Try

                AppbarProperties.RadioButton6.Checked = True

                Me.Opacity = 1
            Case 1
                AppbarProperties.RadioButton1.Checked = True

                Try
                    Me.Opacity = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "TransparencyLevel", 90) / 100
                Catch ex As Exception
                    Me.Opacity = 0.9
                End Try
            Case 2
                Dim side As side = side
                side.Left = 0
                side.Right = 0
                side.Top = 0
                side.Bottom = 0
                Dim result As Integer = DwmExtendFrameIntoClientArea(Me.Handle, side)

                AppbarProperties.RadioButton7.Checked = True

                Me.Opacity = 1
            Case Else
                Me.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", "#B77358"))
        End Select

        Try
            Dim avoidefp As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "AvoidExplorerFileDialog", False)
            If avoidefp = True Then UseExplorerFP = False Else UseExplorerFP = True
        Catch ex As Exception
            UseExplorerFP = True
        End Try

        Try
            Dim avoidefm As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "UseCustomFileManager", False)
            If avoidefm = True Then UseExplorerFM = False Else UseExplorerFM = True
        Catch ex As Exception
            UseExplorerFM = True
        End Try

        Try
            Dim cfm As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "CustomFileManager", String.Empty)
            If Not String.IsNullOrWhiteSpace(cfm) AndAlso File.Exists(cfm) Then
                If New FileInfo(cfm).Extension.ToLower = ".exe" Then
                    CustomFMPath = cfm
                End If
            End If
        Catch ex As Exception
            CustomFMPath = String.Empty
        End Try

        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "UseSystemColor", 1) = 1 Then

            Dim accentColor As Color = GetAccentColor()

            Me.BackColor = accentColor
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", ColorTranslator.ToHtml(accentColor))

        End If

        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "Enabled", False) = 1 Then
            Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\WorkingArea", "Type", 0)

                Case 0 ' Explorer WorkingArea

                    fBarRegistered = False
                    RegisterBar()

                Case 1 ' User32.dll WorkingArea

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

                            workspace.UpdateWorkingArea(workTop, workBottom, workLeft, workRight)
                        End If

                    Catch ex As Exception
                        workspace.UpdateWorkingArea(AppbarProperties.defWorkTop, AppbarProperties.defWorkBottom, AppbarProperties.defWorkLeft, AppbarProperties.defWorkRight)
                    End Try

            End Select
        End If

        Me.BringToFront()
        Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - Me.Height)

        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowBackImage", False) = 1 Then
            If IO.File.Exists(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImage", "")) Then
                Try
                    Me.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImage", ""))
                    AppbarProperties.PictureBox1.Image = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImage", ""))
                    AppbarProperties.ComboBox2.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImage", "")
                Catch ex As Exception
                    Me.BackgroundImage = My.Resources.AppBarMain
                    AppbarProperties.PictureBox1.Image = My.Resources.AppBarMain
                    AppbarProperties.ComboBox2.Text = ""
                End Try
            Else
                Me.BackgroundImage = My.Resources.AppBarMain
                AppbarProperties.PictureBox1.Image = My.Resources.AppBarMain
                AppbarProperties.ComboBox2.Text = ""
            End If
            Me.BackgroundImageLayout = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImageLayout", "1")
        End If

        Splitter1.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "SplitterColor", "#000000"))
        Splitter2.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "SplitterColor", "#000000"))
        Splitter3.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "SplitterColor", "#000000"))

        BlockingProcesses.Enabled = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "Enabled", 0)

        LockAppbarToolStripMenuItem.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Locked", 0)
        If LockAppbarToolStripMenuItem.Checked = True Then
            Splitter1.Visible = False
            Splitter2.Visible = False
            Splitter3.Visible = False
        Else
            Splitter1.Visible = True
            Splitter2.Visible = True
            Splitter3.Visible = True
        End If

        ' Start button

        Start.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "StartButton", True)
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 0 Then

            AppbarProperties.RadioButton8.Checked = True
            Start.BackgroundImage = My.Resources.StartRight

        ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 1 Then
            AppbarProperties.RadioButton9.Checked = True

            Try
                Start.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Normal", ""))
            Catch ex As Exception
                Start.BackgroundImage = My.Resources.StartRight
            End Try

        ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 2 Then
            'ORB Code Here
            Dim orbPath As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "CustomORB", "")

            If orbPath IsNot Nothing AndAlso File.Exists(orbPath) Then
                Try
                    LoadAndSplitOrb(orbPath)

                    Start.BackgroundImage = OrbNormal
                Catch ex As Exception
                    Start.BackgroundImage = My.Resources.StartRight
                End Try
            End If
        End If

        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Layout", "0") = 0 Then
            Start.BackgroundImageLayout = ImageLayout.Stretch
        Else
            Start.BackgroundImageLayout = ImageLayout.Center
        End If

        If Not Me.Width = SystemInformation.PrimaryMonitorSize.Width Then
            Me.Width = SystemInformation.PrimaryMonitorSize.Width
        End If

        If Not Me.WindowState = FormWindowState.Normal Then
            Me.WindowState = FormWindowState.Normal
        End If

        StartButtonToolStripMenuItem.Checked = Start.Visible

        Try
            Start.Width = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer1", "50")
        Catch ex As Exception
            Start.Width = 50
        End Try

        If GlobalKeyboardHook.SetHook() Then
            Debug.WriteLine("Keyboard Hooks set up!")
        Else
            MessageBox.Show("Keyboard Hooks failed to set up!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        AddHandler GlobalKeyboardHook.HotkeyPressed, AddressOf GlobalKeyboardHook_HotkeyPressed

        ' AppStrip

        Panel1.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "PinnedBar", True)
        PinnedBarToolStripMenuItem.Checked = Panel1.Visible
        Try
            Panel1.Width = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer2", "107")
        Catch ex As Exception
            Panel1.Width = 107
        End Try

        Try : Button3.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "LanguageList", 1) : Catch ex As Exception
            Button3.Visible = True
        End Try

        Dim PADtoSet As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "PinnedAppsLocation", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Internet Explorer\Quick Launch").ToString.Trim

CheckAgainIfFileExist:
        If Directory.Exists(PADtoSet) Then
            PinnedAppsDir = PADtoSet

            LoadPinnedApps()
        Else
            If PADtoSet Is Nothing Then
                Select Case MsgBox($"Failed to register path for pinned apps: ""{PADtoSet}""" & Environment.NewLine & "Do you want to change it to something else? If you select ""No"", the AppStrip will stay empty.", MsgBoxStyle.YesNo, "Failed to register Path")
                    Case MsgBoxResult.Yes
                        Dim FBD As New FolderBrowserDialog With {
                        .Description = "Select a default location where your Pinned apps will be:",
                        .RootFolder = Environment.SpecialFolder.Desktop,
                        .ShowNewFolderButton = True,
                        .SelectedPath = PinnedAppsDir
                        }

                        Select Case FBD.ShowDialog
                            Case DialogResult.OK
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "PinnedAppsLocation", FBD.SelectedPath, Microsoft.Win32.RegistryValueKind.String)
                                PADtoSet = FBD.SelectedPath

                                GoTo CheckAgainIfFileExist
                        End Select
                End Select
            Else
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "PinnedAppsLocation", PinnedAppsDir, Microsoft.Win32.RegistryValueKind.String)

                LoadPinnedApps()
            End If
        End If

        BlockingProcesses.Enabled = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "Enabled", "0")

        Try
            BlockingProcesses.Interval = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "Ticks", "1")
        Catch ex As Exception
            BlockingProcesses.Interval = 1
        End Try

        ' ProcessStrip

        MenuBarToolStripMenuItem.Checked = Panel4.Visible
        Button2.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowDesktopButton", True)
        Try
            Panel4.Width = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer3", "200")
        Catch ex As Exception
            Panel4.Width = 200
        End Try

        procDelegate = New WinEventDelegate(AddressOf WinEventProc)
        hHook = SetWinEventHook(EVENT_SYSTEM_ALERT, EVENT_OBJECT_NAMECHANGE, IntPtr.Zero, procDelegate, 0, 0, WINEVENT_OUTOFCONTEXT)

        'EnumWindows(AddressOf EnumWindowCallBack, IntPtr.Zero)
        LoadApps()

        Me.TopMost = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "OnTop", True)
        LockAppbarToolStripMenuItem.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Locked", False)
        AutoHideToolStripMenuItem.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "AutoHide", 0)

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

        ' Screen Filter (NEW)

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

        ' Alarms

        StartClockUpdate()

        Dim rvA = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\", "DisableAlarms", 0)
        If rvA = 0 Then
            DisableAlarmOption.Checked = False
            AlarmController.Enabled = True
        Else
            DisableAlarmOption.Checked = True
            AlarmController.Enabled = False
        End If

        ' Tray Icons

        CreateTrayWindow()

        Dim oskPath As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "Keyboard", Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\osk.exe")
        If ShowKeyboardButton = True AndAlso File.Exists(oskPath) Then
            ToolStripButton2.Visible = True
        Else
            ToolStripButton2.Visible = False
        End If

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

                ' checks if the process is already running.
                If runningProcesses.Contains(Path.GetFileNameWithoutExtension(fi.Name)) Then Continue For

                If Not fi.Extension.ToLower = ".lnk" Then
                    Process.Start(New ProcessStartInfo(i) With {.UseShellExecute = True})
                Else
                    Try
                        Process.Start(New ProcessStartInfo(GetShortcutTarget(fi.FullName)) With {.UseShellExecute = True})
                    Catch ex As Exception

                    End Try
                End If : Next : End If

        ' Language List

        UpdateLanguageDisplay()

        ' Custom paths

        Dim defsndvol As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\SndVol.exe"
        Dim soundcntpath As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "SoundControl", defsndvol)

        If String.IsNullOrWhiteSpace(soundcntpath) OrElse Not File.Exists(soundcntpath) Then
            If File.Exists(defsndvol) Then soundcontrolPath = soundcntpath
        Else
            soundcontrolPath = soundcntpath
        End If
    End Sub

    Public SC_isSecondsEnabled As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSecondsInSystemClock", False)
    Public SC_ShowTime As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\DateTime", "ShowTime", True)
    Public SC_ShowDay As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\DateTime", "ShowDay", True)
    Public SC_ShowDate As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\DateTime", "ShowDate", True)

    Private Sub StartClockUpdate()
        Task.Run(Sub()
                     While True
                         Dim currentTime As String = String.Empty
                         Dim currentDay As String = String.Empty
                         Dim currentDate As String = String.Empty

                         ' Options
                         If SC_ShowTime = True Then : If SC_isSecondsEnabled Then
                                 currentTime = DateTime.Now.ToString("HH:mm:ss")
                             Else
                                 currentTime = DateTime.Now.ToString("HH:mm")
                             End If
                         End If

                         If SC_ShowDay = True Then currentDay = DateTime.Now.DayOfWeek.ToString
                         If SC_ShowDate = True Then currentDate = DateTime.Now.ToString("dd. MM. yyyy")

                         ' Showing time
                         If Me.InvokeRequired Then
                             Me.Invoke(Sub()
                                           TimeLabel.Text = currentTime
                                           If SC_ShowDay = True Then TimeLabel.Text += Environment.NewLine & currentDay
                                           If SC_ShowDate = True Then TimeLabel.Text += Environment.NewLine & currentDate
                                       End Sub)
                         Else
                             TimeLabel.Text = currentTime
                             If SC_ShowDay = True Then TimeLabel.Text += Environment.NewLine & currentDay
                             If SC_ShowDate = True Then TimeLabel.Text += Environment.NewLine & currentDate
                         End If

                         Thread.Sleep(1000)
                     End While
                 End Sub)
    End Sub
    Public ShowKeyboardButton As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowKeyboard", False)
    Private Sub GlobalKeyboardHook_HotkeyPressed(sender As Object, e As EventArgs)
        Me.BringToFront()
        Startmenu.FST = True
        StartMenuShowAndHide()
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
    Private Const WM_CONTEXTMENU As UInteger = &H7B
    Private Const WM_POPUPSYSTEMMENU As UInteger = &H313

    Private LastActiveWindowHandle As IntPtr = IntPtr.Zero
    Private Sub ProcessStripItemClick(sender As Object, e As MouseEventArgs)

        Dim item = CType(sender, ToolStripButton)

        If item.Tag IsNot Nothing AndAlso TypeOf item.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(item.Tag, IntPtr)

            StopAlert(windowHandle)

            LastActiveWindowHandle = windowHandle

            Dim p As Process = GetProcessFromWindowHandle(windowHandle)

            If e.Button = Windows.Forms.MouseButtons.Left Then
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

            ElseIf e.Button = Windows.Forms.MouseButtons.Right Then

                WAT.Close()

                TPCM.Tag = windowHandle

                Try
                    StartSameProcessToolStripMenuItem.ToolTipText = p.StartInfo.FileName
                Catch ex As Exception : End Try

                TPCM.Show(Control.MousePosition)

            ElseIf e.Button = MouseButtons.Middle Then

                WAT.Close()

                TPCM.Tag = windowHandle

                Try
                    Try
                        StartSameProcessToolStripMenuItem.ToolTipText = p.StartInfo.FileName
                    Catch ex2 As Exception : End Try

                    Dim pt As Point = Cursor.Position
                    Dim lParam As New IntPtr((pt.Y << 16) Or (pt.X And &HFFFF))

                    SetForegroundWindow(windowHandle)

                    PostMessage(windowHandle, WM_POPUPSYSTEMMENU, IntPtr.Zero, lParam)
                Catch ex As Exception
                    WindowStateToolStripMenuItem.DropDown.Show(Cursor.Position)
                End Try
            End If
        End If
    End Sub

    Private Sub ProcessStripItemMouseLeave(sender As Object, e As EventArgs)
        CloseTimer.Start()
    End Sub

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
                Else
                    btn.Checked = False
                End If
            End If
        Next
    End Sub

    Private Sub DefaultToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DefaultToolStripMenuItem.Click
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_SHOWDEFAULT)
        End If
    End Sub

    Private Sub MaximalizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MaximalizeToolStripMenuItem.Click
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_MAXIMIZE)
        End If

        'Dim App As Process = Process.GetProcessById(TPCM.Tag)
        'If App.Id > 0 Then
        '   AppActivate(App.Id)
        '   ShowWindow(App.MainWindowHandle, SHOW_WINDOW.SW_MAXIMIZE)
        'End If
    End Sub

    Private Sub NormalizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NormalizeToolStripMenuItem.Click
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_NORMAL)
        End If
    End Sub

    Private Sub MinimalizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MinimalizeToolStripMenuItem.Click
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_MINIMIZE)
        End If
    End Sub

    Private Sub ForceMinimalizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ForceMinimalizeToolStripMenuItem.Click
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_FORCEMINIMIZE)
        End If
    End Sub

    Private Sub HIDEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HIDEToolStripMenuItem.Click
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            If MsgBox("Are you sure do you want to hide this window process? [" & GetProcessFromWindowHandle(windowHandle).ProcessName & "]", MsgBoxStyle.YesNo, "Confirm Box") = MsgBoxResult.Yes Then
                ShowWindow(windowHandle, SHOW_WINDOW.SW_HIDE)
            End If
        End If
    End Sub

    Private Sub SwitchToToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SwitchToToolStripMenuItem.Click
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            SetForegroundWindow(windowHandle)
            'ShowWindow(windowHandle, SHOW_WINDOW.SW_SHOW)
            ActiveWindowHandle = windowHandle
        End If
    End Sub

    Private Sub SwitchButNotActiveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SwitchButNotActiveToolStripMenuItem.Click
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_SHOWNA)
        End If
    End Sub

    Private Const HWND_TOPMOST = -1 '-- Bring to top and stay there
    Private Const HWND_NOTOPMOST = -2 '-- Put the window into a normal position

    Const WM_CLOSE As Long = &H10
    Const WM_DESTROY As Long = &H2

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            SendMessage(windowHandle, WM_CLOSE, 0, 0)
        End If
    End Sub

    Private Sub StartSameProcessToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartSameProcessToolStripMenuItem.Click
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Try
                Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
                Process.Start(GetProcessFromWindowHandle(windowHandle).MainModule.FileName)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub TPCM_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TPCM.Opening

        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

            Try
                Dim fvi As FileVersionInfo = FileVersionInfo.GetVersionInfo(pr.MainModule.FileName)
                StartSameProcessToolStripMenuItem.Text = fvi.FileDescription
            Catch ex As Exception
                StartSameProcessToolStripMenuItem.Text = "Start Same program"
            End Try

            Try
                Dim ico As Icon = Icon.ExtractAssociatedIcon(pr.MainModule.FileName)
                StartSameProcessToolStripMenuItem.Image = ico.ToBitmap
            Catch ex As Exception
                StartSameProcessToolStripMenuItem.Image = My.Resources.ProgramMedium
            End Try

            Try
                Select Case pr.PriorityClass
                    Case ProcessPriorityClass.Idle
                        PRprcmb6.Checked = True
                        PRprcmb5.Checked = False
                        PRprcmb4.Checked = False
                        PRprcmb3.Checked = False
                        PRprcmb2.Checked = False
                        PRprcmb1.Checked = False
                        SetPriorityLevelToolStripMenuItem.Enabled = True

                    Case ProcessPriorityClass.BelowNormal
                        PRprcmb6.Checked = False
                        PRprcmb5.Checked = True
                        PRprcmb4.Checked = False
                        PRprcmb3.Checked = False
                        PRprcmb2.Checked = False
                        PRprcmb1.Checked = False
                        SetPriorityLevelToolStripMenuItem.Enabled = True

                    Case ProcessPriorityClass.Normal
                        PRprcmb6.Checked = False
                        PRprcmb5.Checked = False
                        PRprcmb4.Checked = True
                        PRprcmb3.Checked = False
                        PRprcmb2.Checked = False
                        PRprcmb1.Checked = False
                        SetPriorityLevelToolStripMenuItem.Enabled = True

                    Case ProcessPriorityClass.AboveNormal
                        PRprcmb6.Checked = False
                        PRprcmb5.Checked = False
                        PRprcmb4.Checked = False
                        PRprcmb3.Checked = True
                        PRprcmb2.Checked = False
                        PRprcmb1.Checked = False
                        SetPriorityLevelToolStripMenuItem.Enabled = True

                    Case ProcessPriorityClass.High
                        PRprcmb6.Checked = False
                        PRprcmb5.Checked = False
                        PRprcmb4.Checked = False
                        PRprcmb3.Checked = False
                        PRprcmb2.Checked = True
                        PRprcmb1.Checked = False
                        SetPriorityLevelToolStripMenuItem.Enabled = True

                    Case ProcessPriorityClass.RealTime
                        PRprcmb6.Checked = False
                        PRprcmb5.Checked = False
                        PRprcmb4.Checked = False
                        PRprcmb3.Checked = False
                        PRprcmb2.Checked = False
                        PRprcmb1.Checked = True
                        SetPriorityLevelToolStripMenuItem.Enabled = True

                    Case Else
                        SetPriorityLevelToolStripMenuItem.Enabled = False
                End Select
            Catch ex As Exception
                SetPriorityLevelToolStripMenuItem.Enabled = False
            End Try

            Try
                If pr.StartInfo.WorkingDirectory IsNot Nothing Then
                    OpenProgramsWorkingDirectoryToolStripMenuItem.Enabled = True
                Else
                    OpenProgramsWorkingDirectoryToolStripMenuItem.Enabled = False
                End If
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub PRprcmb1_Click(sender As Object, e As EventArgs) Handles PRprcmb1.Click
        On Error Resume Next
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            GetProcessFromWindowHandle(windowHandle).PriorityClass = ProcessPriorityClass.RealTime
            PRprcmb6.Checked = False
            PRprcmb5.Checked = False
            PRprcmb4.Checked = False
            PRprcmb3.Checked = False
            PRprcmb2.Checked = False
            PRprcmb1.Checked = True
        End If
    End Sub

    Private Sub PRprcmb2_Click(sender As Object, e As EventArgs) Handles PRprcmb2.Click
        On Error Resume Next
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            GetProcessFromWindowHandle(windowHandle).PriorityClass = ProcessPriorityClass.High
            PRprcmb6.Checked = False
            PRprcmb5.Checked = False
            PRprcmb4.Checked = False
            PRprcmb3.Checked = False
            PRprcmb2.Checked = True
            PRprcmb1.Checked = False
        End If
    End Sub

    Private Sub PRprcmb3_Click(sender As Object, e As EventArgs) Handles PRprcmb3.Click
        On Error Resume Next
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            GetProcessFromWindowHandle(windowHandle).PriorityClass = ProcessPriorityClass.AboveNormal
            PRprcmb6.Checked = False
            PRprcmb5.Checked = False
            PRprcmb4.Checked = False
            PRprcmb3.Checked = True
            PRprcmb2.Checked = False
            PRprcmb1.Checked = False
        End If
    End Sub

    Private Sub PRprcmb4_Click(sender As Object, e As EventArgs) Handles PRprcmb4.Click
        On Error Resume Next
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            GetProcessFromWindowHandle(windowHandle).PriorityClass = ProcessPriorityClass.Normal
            PRprcmb6.Checked = False
            PRprcmb5.Checked = False
            PRprcmb4.Checked = True
            PRprcmb3.Checked = False
            PRprcmb2.Checked = False
            PRprcmb1.Checked = False
        End If
    End Sub

    Private Sub PRprcmb5_Click(sender As Object, e As EventArgs) Handles PRprcmb5.Click
        On Error Resume Next
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            GetProcessFromWindowHandle(windowHandle).PriorityClass = ProcessPriorityClass.BelowNormal
            PRprcmb6.Checked = False
            PRprcmb5.Checked = True
            PRprcmb4.Checked = False
            PRprcmb3.Checked = False
            PRprcmb2.Checked = False
            PRprcmb1.Checked = False
        End If
    End Sub

    Private Sub PRprcmb6_Click(sender As Object, e As EventArgs) Handles PRprcmb6.Click
        On Error Resume Next
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            GetProcessFromWindowHandle(windowHandle).PriorityClass = ProcessPriorityClass.Idle
            PRprcmb6.Checked = True
            PRprcmb5.Checked = False
            PRprcmb4.Checked = False
            PRprcmb3.Checked = False
            PRprcmb2.Checked = False
            PRprcmb1.Checked = False
        End If
    End Sub

    Public Sub StartMenuShowAndHide()
        Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0")
            ' Normal Start menu look
            Case 0
                If Startmenu.Visible = True Then
                    Startmenu.Hide()
                    Start.BackgroundImage = My.Resources.StartRight
                Else
                    Startmenu.Show()
                    Start.BackgroundImage = My.Resources.StartLeft
                End If

            ' Custom Images start menu look
            Case 1
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

            ' ORB Start menu look
            Case 2
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
                    Start.BackgroundImage = My.Resources.StartRight
                Else
                    Startmenu.Show()
                    Start.BackgroundImage = My.Resources.StartLeft
                End If
        End Select
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Start.Click
        Me.BringToFront()
        Startmenu.FST = True
        StartMenuShowAndHide()
    End Sub
    Public UseLegacyVolumeSlider As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Volume", "UseLegacy", False)
    Private Async Sub ToolStripButton1_MouseUp(sender As Object, e As MouseEventArgs) Handles ToolStripButton1.MouseUp
        If e.Button = MouseButtons.Left Then
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
                        SetWindowPos(hwnd, IntPtr.Zero, cursorPoint.X - 45, cursorPoint.Y - 290, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_SHOWWINDOW)
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

        ElseIf e.Button = MouseButtons.Right Then
            ToolStripButton1.Checked = True

            VCM.Show(MousePosition)
        End If
    End Sub
    Public CanClose As Boolean = False
    Private Sub AppBar_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.TaskManagerClosing OrElse e.CloseReason = CloseReason.WindowsShutDown Then
            e.Cancel = False
            If GlobalKeyboardHook.Unhook() Then
                Debug.WriteLine("Hook ended.")
            Else
                Debug.WriteLine("Hook ending failed.")
            End If
        Else
            If CanClose = False Then
                e.Cancel = True
                SA.ShowDialog()
            Else
                If GlobalKeyboardHook.Unhook() Then
                    Debug.WriteLine("Hook ended.")
                Else
                    Debug.WriteLine("Hook ending failed.")
                End If
            End If
        End If

        If hHook <> IntPtr.Zero Then UnhookWinEvent(hHook)
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
        Process.Start(New ProcessStartInfo(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "Taskmgr", Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\taskmgr.exe")) With {.UseShellExecute = True})
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

        EnumWindows(AddressOf EnumWindowsCallback, IntPtr.Zero)
    End Sub

    Private Async Sub ProcessStripItemMouseEnter(sender As Object, e As EventArgs)
        If Startmenu.Visible Then
            Startmenu.Focus()
            Exit Sub
        End If

        Dim item As ToolStripButton = TryCast(sender, ToolStripButton)
        If item Is Nothing OrElse item.Tag Is Nothing Then Exit Sub

        Try
            Dim windowHandle As IntPtr = CType(item.Tag, IntPtr)

            If WAT.Visible AndAlso WAT.Tag.Equals(windowHandle) Then
                CloseTimer.Stop()
                Exit Sub
            End If

            TPCM.Tag = windowHandle
            WAT.Tag = windowHandle

            If IsHungAppWindow(windowHandle) Then Exit Sub

            Dim renderedImage As Image = Await Task.Run(Function()
                                                            Return RenderWindow(windowHandle, False)
                                                        End Function)
            WAT.Label1.Text = item.Text
            WAT.Button4.Image = Nothing
            WAT.Button4.BackgroundImage = If(renderedImage, item.Image)

            Dim targetHeight As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband", "MinThumbSizePx", 178)
            Dim magicWidth As Integer = 178

            Dim finalWidth As Integer
            Dim finalHeight As Integer

            If WAT.Button4.BackgroundImage IsNot Nothing Then
                Dim img As Image = WAT.Button4.BackgroundImage

                If img.Height > targetHeight Then
                    Dim ratio As Double = targetHeight / img.Height
                    finalHeight = targetHeight
                    finalWidth = CInt(img.Width * ratio) - (WAT.Panel3.Width * 4) - (WAT.Panel2.Width * 4)
                Else
                    finalHeight = img.Height
                    finalWidth = img.Width - WAT.Panel3.Width - WAT.Panel2.Width
                End If
            Else
                finalHeight = targetHeight
                finalWidth = magicWidth
            End If

            finalWidth = Math.Max(finalWidth, WAT.MinimumSize.Width)
            finalHeight = Math.Max(finalHeight, WAT.MinimumSize.Height)

            Dim itemScreenBounds As Rectangle = item.Bounds
            If item.Owner IsNot Nothing Then
                itemScreenBounds = item.Owner.RectangleToScreen(item.Bounds)
            End If

            Dim targetX As Integer = CInt(itemScreenBounds.Left + (itemScreenBounds.Width / 2) - (finalWidth / 2))
            Dim targetY As Integer = SystemInformation.WorkingArea.Height - finalHeight - 5

            Dim targetRect As New Rectangle(targetX, targetY, finalWidth, finalHeight)
            CloseTimer.Stop()
            WAT.MoveTo(targetRect)

        Catch ex As Exception
            WAT.Label1.Text = item.Text
            WAT.Button4.BackgroundImage = Nothing
            WAT.Button4.Image = My.Resources.ProgramMedium

            Dim fallbackX As Integer = Control.MousePosition.X - (WAT.Width / 2)
            Dim fallbackY As Integer = SystemInformation.WorkingArea.Height - WAT.Height
            WAT.MoveTo(New Rectangle(fallbackX, fallbackY, 150, 100))
        End Try
    End Sub

    Private Sub StartButtonToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartButtonToolStripMenuItem.Click
        Start.Visible = StartButtonToolStripMenuItem.Checked
        If LockAppbarToolStripMenuItem.Checked = False Then Splitter1.Visible = StartButtonToolStripMenuItem.Checked
    End Sub

    Private Sub PinnedBarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PinnedBarToolStripMenuItem.Click
        Panel1.Visible = PinnedBarToolStripMenuItem.Checked
        If LockAppbarToolStripMenuItem.Checked = False Then Splitter2.Visible = PinnedBarToolStripMenuItem.Checked
    End Sub

    Private Sub MenuBarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MenuBarToolStripMenuItem.Click
        Panel4.Visible = MenuBarToolStripMenuItem.Checked
        If LockAppbarToolStripMenuItem.Checked = False Then Splitter3.Visible = MenuBarToolStripMenuItem.Checked
    End Sub

    Public Sub AutoHideOnOff()
        fBarRegistered = AutoHideToolStripMenuItem.Checked
        RegisterBar()
        If AutoHideToolStripMenuItem.Checked = True Then
            Hidden = True
            Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - 2)
        Else
            Hidden = False
            Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - Me.Height)
        End If
    End Sub

    Private Sub AutoHideToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutoHideToolStripMenuItem.Click
        AutoHideOnOff()
    End Sub
    Public Hidden As Boolean = False
    Private Sub AppBar_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
        'If AutoHideToolStripMenuItem.Checked = True Then
        'Else
        If Not Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - Me.Height) Then
            Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - Me.Height)
        End If
        'End If
    End Sub

    Private Sub ProcessStrip_MouseLeave1(sender As Object, e As EventArgs) Handles ProcessStrip.MouseLeave, AppStrip.MouseLeave, TimeLabel.MouseLeave
        Hidden = True
        'Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - 2)
        AutoHideOnOff()
    End Sub

    Private Sub ProcessStrip_MouseLeave(sender As Object, e As EventArgs) Handles ProcessStrip.MouseLeave

    End Sub

    Private Sub ToolStrip1_MouseEnter(sender As Object, e As EventArgs) Handles ToolStrip1.MouseEnter, ProcessStrip.MouseEnter, AppStrip.MouseEnter, TimeLabel.MouseEnter
        Hidden = False

        Startmenu.FST = False
        AutoHideOnOff()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, ShowDesktopToolStripMenuItem.Click, ShowDesktopToolStripMenuItem1.Click
        Desktop.ShowDesktop()
    End Sub

    Private Sub PropertiesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PropertiesToolStripMenuItem.Click
        On Error Resume Next
        AppbarProperties.Show(Me)
    End Sub

    Private Sub BlockingProcesses_Tick(sender As Object, e As EventArgs) Handles BlockingProcesses.Tick
        On Error Resume Next
        For Each i As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\BlockedProcesses\.\").GetValueNames
            For Each pr As Process In Process.GetProcesses
                If pr.ProcessName = i Then
                    If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "KillMethod", "0") = 0 Then
                        pr.Kill()
                    Else
                        pr.CloseMainWindow()
                    End If
                End If
            Next
        Next
    End Sub

    Private Sub CMMAIN_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles CMMAIN.Opening
        CMMAIN.RightToLeft = RightToLeft.No
    End Sub

    Private Sub TimeAndDateSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TimeAndDateSettingsToolStripMenuItem.Click
        On Error Resume Next
        Process.Start(New ProcessStartInfo(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "TimeDateProperties", Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\timedate.cpl")) With {.UseShellExecute = True})
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
        If MessageBox.Show("Do you really want to clear the Clipboard data?", "Microsoft Windows", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
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

    Private Sub ActionCM_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ActionCM.Opening
        ActionCM.Items.Clear()

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
                            ActionCM.Items.Add(item)
                            addedAny = True
                        End If
                    Next

                    If addedAny Then
                        ActionCM.Items.Add(New ToolStripSeparator())
                    End If
                Else
                    Exit For
                End If
            Next

            ActionCM.Items.Add(ClipboardToolStripMenuItem)
            ActionCM.Items.Add(PerformanceToolStripMenuItem)
            ActionCM.Items.Add(New ToolStripSeparator)
            ActionCM.Items.Add(LogonToolStripMenuItem)
            ActionCM.Items.Add(PowerToolStripMenuItem)

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

    Private Sub ToolStripButton3_MouseUp(sender As Object, e As MouseEventArgs) Handles ToolStripButton3.MouseUp
        'Shell("RunDll32.exe shell32.dll,Control_RunDLL ncpa.cpl")
        ToolStripButton3.Checked = True
        MCM.Show(MousePosition)
    End Sub

    Private Sub AppBar_BackColorChanged(sender As Object, e As EventArgs) Handles Me.BackColorChanged
        Dim luminance As Double = (0.299 * Me.BackColor.R + 0.587 * Me.BackColor.G + 0.114 * Me.BackColor.B) / 255

        Dim contrastColor As Color = If(luminance > 0.5, Color.Black, Color.White)

        For Each ctrl As Control In Me.Controls
            ctrl.ForeColor = contrastColor
        Next
    End Sub

    Private Sub KillToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KillToolStripMenuItem.Click
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

            If MsgBox("Are you sure do you want to kill this process? [" & pr.ProcessName & "]", MsgBoxStyle.YesNo, "Confirm Box") = MsgBoxResult.Yes Then
                Try
                    pr.Kill()

                    'LoadApps()
                Catch ex As Exception
                    MsgBox("Unable to kill the process. " & ex.Message, MsgBoxStyle.Critical, "Error")
                End Try

            End If
        End If
    End Sub

    Private Sub DestroyProgramToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DestroyWindowToolStripMenuItem.Click
        If MsgBox("!!!WARNING!!!" & Environment.NewLine & """Destroy a Window"" means, that everything that the selected Window has, including data on your memory will be ""destroyed"" which I don't recommend. Well this warning is to prevent it happening one time. Do you really want to continue?", MsgBoxStyle.YesNo, "Warning") = MsgBoxResult.Yes Then
            If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
                Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
                SendMessage(windowHandle, WM_DESTROY, 0, 0)
            End If
        End If
    End Sub

    Private Sub OpenProgramsLocationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenProgramsLocationToolStripMenuItem.Click
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
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
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

            If IO.Directory.Exists(pr.StartInfo.WorkingDirectory) = False Then
                OpenProgramsWorkingDirectoryToolStripMenuItem.Enabled = False
                OpenProgramsWorkingDirectoryToolStripMenuItem.ToolTipText = "This process hasn't associated any custom Working directory."
            Else
                OpenProgramsWorkingDirectoryToolStripMenuItem.Enabled = True
                OpenProgramsWorkingDirectoryToolStripMenuItem.ToolTipText = ""
            End If

            BlockProcessToolStripMenuItem.Enabled = BlockingProcesses.Enabled
        End If
    End Sub

    Private Sub OpenProgramsWorkingDirectoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenProgramsWorkingDirectoryToolStripMenuItem.Click
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

            If IO.Directory.Exists(pr.StartInfo.WorkingDirectory) = True Then
                Dim fi As New IO.DirectoryInfo(pr.StartInfo.WorkingDirectory)

                OpenDestination(fi.FullName)
            End If
        End If
    End Sub

    Private Sub AlwaysOnTopToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AlwaysOnTopToolStripMenuItem.Click
        On Error Resume Next
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)
            SetWindowPos(windowHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        End If
    End Sub

    Private Sub AlwaysOnTopToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AlwaysOnTopToolStripMenuItem1.Click
        On Error Resume Next
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)
            SetWindowPos(windowHandle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        End If
    End Sub

    Private Sub HideProcessFromAppbarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HideProcessFromAppbarToolStripMenuItem.Click
        On Error Resume Next
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\HiddenProcesses\", pr.ProcessName, "")

            'LoadApps()
            MainToolTip.Show("The """ & Process.GetProcessById(TPCM.Tag).ProcessName & """ has been successfully set to hidden from Appbar!", Me, Me.Width - ToolStrip1.Width / 2, Me.Location.Y + ToolStrip1.Height + 150, 5000)
        End If
    End Sub

    Private Sub BlockProcessToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BlockProcessToolStripMenuItem.Click
        If MessageBox.Show("Are you sure do you want to block [" & Process.GetProcessById(TPCM.Tag).ProcessName & "] process?" & Environment.NewLine & "This means, when this shell will be running, then that process you'll set as ""Blocked"" will no longer be working. To allow the process again: Right click on the Appbar, select properties and remove the app from Blocklist.", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
                Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
                Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses\.\", pr.ProcessName, "")

                LoadApps()
                MainToolTip.Show("The """ & Process.GetProcessById(TPCM.Tag).ProcessName & """ process, has been successfully blocked!", Me, Me.Width - ToolStrip1.Width / 2, Me.Location.Y + ToolStrip1.Height + 150, 5000)
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
        Catch ex As Exception

        End Try
    End Sub

    Private Sub AboutShellToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutShellToolStripMenuItem.Click
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
        If TPCM.Tag IsNot Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
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
                    Process.Start(item.Tag)

                Case MouseButtons.Right
                    PIcm.Tag = item.Tag

                    currentPinnedItem = item
                    item.Checked = True
                    PIcm.Show(MousePosition)

                    'Desktop.ShowShellContextMenu(item.Tag, MousePosition)
            End Select
        End If
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        On Error Resume Next
        Process.Start(PIcm.Tag)
    End Sub

    Private Sub OpenFileLocationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenFileLocationToolStripMenuItem.Click
        On Error Resume Next
        Dim fi As New FileInfo(PIcm.Tag)

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
        If IO.File.Exists(PIcm.Tag) Then
            Dim fi As New FileInfo(PIcm.Tag)
            If MessageBox.Show("Are you sure do you want remove """ & fi.Name & """ Item from Pinned Bar?", "Remove Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                My.Computer.FileSystem.DeleteFile(PIcm.Tag, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)

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
        If Not Me.Width = SystemInformation.PrimaryMonitorSize.Width Then
            Me.Width = SystemInformation.PrimaryMonitorSize.Width
        End If

        If Not Me.WindowState = FormWindowState.Normal Then Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub AppBar_StyleChanged(sender As Object, e As EventArgs) Handles Me.StyleChanged
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub Button1_MouseHover(sender As Object, e As EventArgs) Handles Start.MouseHover
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 1 Then
            If Startmenu.Visible = False Then
                Try
                    Start.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Hover", ""))
                Catch ex As Exception
                    Start.BackgroundImage = My.Resources.StartRight
                End Try
            Else
                Try
                    Start.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Pressed", ""))
                Catch ex As Exception
                    Start.BackgroundImage = My.Resources.StartLeft
                End Try
            End If
        ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 2 Then
            If Startmenu.Visible = False Then
                Try
                    Start.BackgroundImage = OrbHover
                Catch ex As Exception
                    Start.BackgroundImage = My.Resources.StartRight
                End Try
            Else
                Try
                    Start.BackgroundImage = OrbPressed
                Catch ex As Exception
                    Start.BackgroundImage = My.Resources.StartLeft
                End Try
            End If
        End If
    End Sub

    Private Sub Button1_MouseLeave(sender As Object, e As EventArgs) Handles Start.MouseLeave
        Hidden = True

        Startmenu.FST = False

        'Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - 2)
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 1 Then
            If Startmenu.Visible = False Then
                Try
                    Start.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Normal", ""))
                Catch ex As Exception
                    Start.BackgroundImage = My.Resources.StartRight
                End Try
            Else
                Try
                    Start.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Pressed", ""))
                Catch ex As Exception
                    Start.BackgroundImage = My.Resources.StartLeft
                End Try
            End If
        ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 2 Then
            If Startmenu.Visible = False Then
                Try
                    Start.BackgroundImage = OrbNormal
                Catch ex As Exception
                    Start.BackgroundImage = My.Resources.StartRight
                End Try
            Else
                Try
                    Start.BackgroundImage = OrbPressed
                Catch ex As Exception
                    Start.BackgroundImage = My.Resources.StartLeft
                End Try
            End If
        End If
    End Sub

    Private Sub Button1_MouseEnter(sender As Object, e As EventArgs) Handles Start.MouseEnter
        Hidden = False

        If Startmenu.Visible = False Then
            Start.Focus()
        End If

        Startmenu.FST = True

        'Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - 2)
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 1 Then
            If Startmenu.Visible = False Then
                Try
                    Start.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Hover", ""))
                Catch ex As Exception
                    Start.BackgroundImage = My.Resources.StartRight
                End Try
            Else
                Try
                    Start.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Pressed", ""))
                Catch ex As Exception
                    Start.BackgroundImage = My.Resources.StartLeft
                End Try
            End If
        ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 2 Then
            If Startmenu.Visible = False Then
                Try
                    Start.BackgroundImage = OrbHover
                Catch ex As Exception
                    Start.BackgroundImage = My.Resources.StartRight
                End Try
            Else
                Try
                    Start.BackgroundImage = OrbPressed
                Catch ex As Exception
                    Start.BackgroundImage = My.Resources.StartLeft
                End Try
            End If
        End If
    End Sub

    Private Sub ReloadAppsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReloadAppsToolStripMenuItem.Click
        LoadApps()

        UpdateAppbarAccent()
    End Sub

    Public Sub UpdateAppbarAccent()
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "UseSystemColor", 1) = 1 Then
            Dim accentColor As Color = GetAccentColor()

            AppbarProperties.Label12.BackColor = accentColor
            AppbarProperties.PictureBox1.BackColor = accentColor
            Me.BackColor = accentColor
        End If
    End Sub

    Private Sub AppStrip_MouseUp(sender As Object, e As MouseEventArgs) Handles AppStrip.MouseUp
        Select Case e.Button
            Case MouseButtons.Right
                Pcm.Show(MousePosition)
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

    Private Sub ProcessStrip_KeyUp(sender As Object, e As KeyEventArgs) Handles ProcessStrip.KeyUp
        If e.KeyCode = Keys.Enter Then
            ' Keyboard settings here!
        End If
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

    Private Sub VCM_Opening(sender As Object, e As CancelEventArgs) Handles VCM.Opening
        MuteToolStripMenuItem.Checked = VolumeControl.IsSysMuted
    End Sub

    Private Sub MuteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MuteToolStripMenuItem.Click
        VolumeControl.ToggleMute(MuteToolStripMenuItem.Checked)
        VolumeControl.UpdateIcons()
    End Sub

    Private Sub VCM_Closing(sender As Object, e As ToolStripDropDownClosingEventArgs) Handles VCM.Closing
        ToolStripButton1.Checked = False
    End Sub

    Private Sub MCM_Closing(sender As Object, e As ToolStripDropDownClosingEventArgs) Handles MCM.Closing
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

    Private Sub PIcm_Opening(sender As Object, e As CancelEventArgs) Handles PIcm.Opening
        If currentPinnedItem IsNot Nothing Then
            currentPinnedItem.Checked = True
        End If
    End Sub

    Private Sub PIcm_Closing(sender As Object, e As ToolStripDropDownClosingEventArgs) Handles PIcm.Closing
        If currentPinnedItem IsNot Nothing Then
            currentPinnedItem.Checked = False

            currentPinnedItem = Nothing
        End If
    End Sub

    Public WithEvents CloseTimer As New System.Windows.Forms.Timer With {.Interval = 200}

    Private Sub CloseTimer_Tick(sender As Object, e As EventArgs) Handles CloseTimer.Tick
        CloseTimer.Stop()
        WAT.Close()
    End Sub

    Private Sub VolumeSliderWindowsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VolumeSliderWindowsToolStripMenuItem.Click
        On Error Resume Next
        Process.Start(soundcontrolPath, "-f")
    End Sub

    Private Sub LLCM_Opening(sender As Object, e As CancelEventArgs) Handles LLCM.Opening
        LLCM.Items.Clear()

        For Each lang As InputLanguage In InputLanguage.InstalledInputLanguages
            Dim displayName As String = lang.Culture.EnglishName

            For Each item As ToolStripMenuItem In LLCM.Items
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
            LLCM.Items.Add(menuItem)
        Next
    End Sub

    Public Sub UpdateLanguageDisplay()
        Dim currentCulture As CultureInfo = InputLanguage.CurrentInputLanguage.Culture
        Button3.Text = currentCulture.ThreeLetterISOLanguageName.ToUpper()
    End Sub

    Private Sub LanguageItem_Click(sender As Object, e As EventArgs)
        Dim clickedItem As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        Dim targetLang As InputLanguage = DirectCast(clickedItem.Tag, InputLanguage)

        InputLanguage.CurrentInputLanguage = targetLang
        UpdateLanguageDisplay()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        LLCM.Show(MousePosition)
    End Sub

    <DllImport("user32.dll")>
    Private Shared Function GetKeyboardLayout(idThread As Integer) As IntPtr
    End Function

    Private Const WM_INPUTLANGCHANGEREQUEST As UInteger = &H50
    Private HWND_BROADCAST As New IntPtr(&HFFFF)

    Private Function GetGlobalInputLanguageHandle() As IntPtr
        Dim foregroundWindow As IntPtr = GetForegroundWindow()
        Dim threadId As Integer = GetWindowThreadProcessId(foregroundWindow, 0)
        Return GetKeyboardLayout(threadId)
    End Function

    Private Sub Button3_MouseWheel(sender As Object, e As MouseEventArgs) Handles Button3.MouseWheel
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

    Private Sub Button3_MouseEnter(sender As Object, e As EventArgs) Handles Button3.MouseEnter
        Button3.Focus()
    End Sub

    Private WithEvents TimerPeek As New System.Windows.Forms.Timer() With {.Interval = 250}
    Public isDesktopPeekEnabled As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "TaskbarSd", True)
    Public isAeroPeekEnabled As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\DWM", "EnableAeroPeek", True)

    Private Sub TimerPeek_Tick(sender As Object, e As EventArgs) Handles TimerPeek.Tick
        TimerPeek.Stop()

        If Not isAeroPeekEnabled Then Exit Sub

        If Desktop.IsHandleCreated Then
            Dim targetHWnd As IntPtr = Desktop.Handle
            Dim targetPos = AeroPeek.GetVisibleWindowRect(targetHWnd)

            AeroPeekScreen.Show()
            Me.BringToFront()

            AeroPeek.ShowPeek(targetHWnd, AeroPeekScreen.Handle, targetPos)
        End If
    End Sub

    Private Sub Button2_MouseEnter(sender As Object, e As EventArgs) Handles Button2.MouseEnter
        If Not isAeroPeekEnabled Then Exit Sub

        If isDesktopPeekEnabled Then
            If Not Desktop.IsShowDesktopMode Then TimerPeek.Start()
        End If
    End Sub

    Private Sub Button2_MouseLeave(sender As Object, e As EventArgs) Handles Button2.MouseLeave
        If Not isAeroPeekEnabled Then Exit Sub
        If Not isDesktopPeekEnabled Then Exit Sub

        TimerPeek.Stop()
        AeroPeek.HidePeek()

        AeroPeekScreen.Hide()
    End Sub

    Private Sub ShowDesktopCM_Opening(sender As Object, e As CancelEventArgs) Handles ShowDesktopCM.Opening
        PeekDesktopToolStripMenuItem.Checked = isDesktopPeekEnabled
    End Sub

    Private Sub PeekDesktopToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PeekDesktopToolStripMenuItem.Click
        If isDesktopPeekEnabled = True Then isDesktopPeekEnabled = False Else isDesktopPeekEnabled = True

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "TaskbarSd", isDesktopPeekEnabled, RegistryValueKind.DWord)
    End Sub

    Private Sub ProcessStrip_MouseWheel(sender As Object, e As MouseEventArgs) Handles ProcessStrip.MouseWheel
        Dim item As ToolStripItem = ProcessStrip.GetItemAt(e.Location)

        If item IsNot Nothing AndAlso item.Tag IsNot Nothing Then
            Dim handle As IntPtr = CType(item.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(handle)

            ChangeAppVolume(pr, e.Delta)
        End If
    End Sub

    Public Sub ChangeAppVolume(targetPid As Process, delta As Integer)
        Dim enumerator As New MMDeviceEnumerator()
        Dim device = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia)
        Dim sessions = device.AudioSessionManager.Sessions

        For i As Integer = 0 To sessions.Count - 1
            Dim session = sessions(i)
            If session.GetProcessID = targetPid.Id Then
                Dim currentVol As Single = session.SimpleAudioVolume.Volume
                Dim stepSize As Single = 0.05F '5 %

                If delta > 0 Then
                    currentVol = Math.Min(1.0F, currentVol + stepSize)
                Else
                    currentVol = Math.Max(0.0F, currentVol - stepSize)
                End If

                session.SimpleAudioVolume.Volume = currentVol
                Exit For
            End If
        Next
    End Sub

    Private Const SC_MONITORPOWER As Integer = &HF170

    ' Hodnoty pro lParam:
    ' 1 = Low power (úsporný režim)
    ' 2 = Power off (úplné vypnutí podsvícení)
    ' -1 = Power on (zapnutí)

    Public Shared Sub TurnOffMonitor(handle As IntPtr)
        SendMessage(handle, WM_SYSCOMMAND, New IntPtr(SC_MONITORPOWER), New IntPtr(2))
    End Sub

    Private Sub ToolStripButton2_MouseUp(sender As Object, e As MouseEventArgs) Handles ToolStripButton2.MouseUp
        On Error Resume Next

        Dim oskPath As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "Keyboard", Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\osk.exe")
        Process.Start(oskPath)
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
        ElseIf sa.chkForcemode.Checked Then
            isForceMode = True ' EWX_FORCEIFHUNG
        End If

        PowerManager.ExecuteAction(item.Text, isForceMode, 0, PowerManager.IsFastStartupEnabled, SA.chkDisableWakeEvent.Checked, isBrutalMode)
    End Sub

    Private Sub SessionButtonPressed(sender As Object, e As EventArgs)
        Dim item = DirectCast(sender, ToolStripMenuItem)
        PowerManager.ExecuteSessionAction(item.Text)
    End Sub
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
        ByRef pvParam As RECT,
        ByVal fWinIni As UInteger) As Boolean
    End Function

    Private Const SPI_SETWORKAREA As UInteger = 47
    Private Const SPIF_UPDATEINIFILE As UInteger = &H1
    Private Const SPIF_SENDCHANGE As UInteger = &H2

    Public Sub UpdateWorkingArea(ByVal topOffset As Integer, ByVal bottomOffset As Integer, ByVal leftOffset As Integer, ByVal rightOffset As Integer)
        Dim screenBounds As Rectangle = Screen.PrimaryScreen.Bounds
        Dim newWorkArea As New RECT With {
            .Left = leftOffset,
            .Top = topOffset,
            .Right = screenBounds.Width - rightOffset,
            .Bottom = screenBounds.Height - bottomOffset
        }

        Dim success As Boolean = SystemParametersInfo(
            SPI_SETWORKAREA,
            0,
            newWorkArea,
            SPIF_UPDATEINIFILE Or SPIF_SENDCHANGE)

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

    Public Sub SnapWindow(ByVal hWnd As IntPtr, ByVal mode As SnapMode)

        Dim screenArea As Rectangle = Screen.FromHandle(hWnd).WorkingArea
        Dim newPos As New Rectangle

        Select Case mode
            Case SnapMode.LeftHalf
                newPos = New Rectangle(screenArea.X, screenArea.Y, screenArea.Width \ 2, screenArea.Height)

            Case SnapMode.RightHalf
                newPos = New Rectangle(screenArea.X + (screenArea.Width \ 2), screenArea.Y, screenArea.Width \ 2, screenArea.Height)

            Case SnapMode.TopLeftQuarter
                newPos = New Rectangle(screenArea.X, screenArea.Y, screenArea.Width \ 2, screenArea.Height \ 2)

            Case SnapMode.TopRightQuarter
                newPos = New Rectangle(screenArea.X + (screenArea.Width \ 2), screenArea.Y, screenArea.Width \ 2, screenArea.Height \ 2)

        End Select

        ShowWindow(hWnd, 9) ' SW_RESTORE

        SetWindowPos(hWnd, IntPtr.Zero, newPos.X, newPos.Y, newPos.Width, newPos.Height, &H40)
    End Sub

    Private Sub ExecutePinnedBar(ByVal index As Integer)
        Dim liveBar = Application.OpenForms.OfType(Of AppBar)().FirstOrDefault()

        If liveBar IsNot Nothing Then
            If liveBar.AppStrip.Items.Count > index Then
                Dim item = AppBar.AppStrip.Items(index)

                If item IsNot Nothing Then
                    Dim executePath As String = DirectCast(item, ToolStripItem).Tag
                    If File.Exists(executePath) Then Process.Start(New ProcessStartInfo(executePath) With {.UseShellExecute = True})
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
                If WinKeyPressed Then
                    If vkCode <> VK_WIN_LEFT AndAlso vkCode <> VK_WIN_RIGHT Then
                        If HandleWindowsShortcuts(vkCode) Then
                            Return New IntPtr(1)
                        End If
                    End If
                ElseIf AltKeyPressed AndAlso vkCode = 9 Then ' Alt + Tab
                    If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "Enabled", "0") = 1 Then
                        AltTab.Show()
                        Return New IntPtr(1)
                    End If
                End If

            Case WM_KEYUP, WM_SYSKEYUP
                Select Case vkCode
                    Case VK_WIN_LEFT, VK_WIN_RIGHT
                        If AppBar.LLCM.Visible Then
                            For Each item As ToolStripMenuItem In AppBar.LLCM.Items
                                If item.Checked Then
                                    Dim targetLang As InputLanguage = DirectCast(item.Tag, InputLanguage)
                                    InputLanguage.CurrentInputLanguage = targetLang

                                    PostMessage(GetForegroundWindow(), WM_INPUTLANGCHANGEREQUEST, IntPtr.Zero, targetLang.Handle)

                                    AppBar.UpdateLanguageDisplay()
                                    Exit For
                                End If
                            Next
                            AppBar.LLCM.Close()
                        End If

                        If Not UsedMultipleKey Then
                            RaiseEvent HotkeyPressed(Nothing, EventArgs.Empty)
                        End If
                        WinKeyPressed = False
                        UsedMultipleKey = False

                    Case VK_CTRL_LEFT, VK_CTRL_RIGHT : CtrlKeyPressed = False
                    Case VK_SHIFT_LEFT, VK_SHIFT_RIGHT : ShiftKeyPressed = False
                    Case VK_ALT_LEFT, VK_ALT_RIGHT
                        AltKeyPressed = False
                        AltTab.Close()

                    Case Keys.Pause
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

    Private Function HandleWindowsShortcuts(ByVal vkCode As Integer) As Boolean
        Dim hWnd As IntPtr = GetForegroundWindow()
        Dim handled As Boolean = True
        UsedMultipleKey = True

        Dim state As Integer = GetWindowState(hWnd, True)

        Dim hotkeyStart As String = "HKEY_CURRENT_USER\Software\Shell\CustomPaths\Hotkeys"
        Dim isCustomHotkey As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "UseHotkeysForPinned", 0)

        Dim liveBar = Application.OpenForms.OfType(Of AppBar)().FirstOrDefault()

        Select Case vkCode
            Case 9 ' TAB
                If AltKeyPressed OrElse WinKeyPressed Then
                    UsedMultipleKey = True

                    If Not AltTab.Visible Then
                        Dim hWndALTTAB As IntPtr = GetForegroundWindow()
                        Dim processId As UInteger = 0
                        GetWindowThreadProcessId(hWndALTTAB, processId)

                        AltTab.ActiveProcessID = processId
                        AltTab.Show()
                        AltTab.Focus()

                        AltTab.MenuStrip1.Focus()
                    End If

                    If AltTab.MenuStrip1.Items.Count <> 0 Then

                        Dim currentSelItem As Integer = 0

                        For i As Integer = 0 To AltTab.MenuStrip1.Items.Count - 1
                            If AltTab.MenuStrip1.Items.Item(i).Selected Then
                                currentSelItem = i
                            End If
                        Next

                        If currentSelItem + 1 > AltTab.MenuStrip1.Items.Count - 1 Then
                            AltTab.MenuStrip1.Items.Item(0).Select()
                        Else
                            AltTab.MenuStrip1.Items.Item(currentSelItem + 1).Select()
                        End If
                    End If

                    Return True
                End If

            Case 32 ' Space
                If WinKeyPressed Then
                    UsedMultipleKey = True

                    If Not AppBar.LLCM.Visible Then
                        AppBar.LLCM.Show(New Point(SystemInformation.WorkingArea.Width - AppBar.LLCM.Width, SystemInformation.WorkingArea.Height - AppBar.LLCM.Height))
                    End If

                    If AppBar.LLCM.Items.Count > 0 Then
                        Dim currentIdx As Integer = -1

                        For i As Integer = 0 To AppBar.LLCM.Items.Count - 1
                            If DirectCast(AppBar.LLCM.Items(i), ToolStripMenuItem).Checked Then
                                currentIdx = i
                                DirectCast(AppBar.LLCM.Items(i), ToolStripMenuItem).Checked = False
                                Exit For
                            End If
                        Next

                        Dim nextIdx As Integer = (currentIdx + 1) Mod AppBar.LLCM.Items.Count
                        Dim nextItem = DirectCast(AppBar.LLCM.Items(nextIdx), ToolStripMenuItem)

                        nextItem.Checked = True
                        nextItem.Select()
                    End If

                    Return True
                End If

            Case 37 ' LEFT
                If CanSnapWindow(hWnd) = True Then
                    If CtrlKeyPressed AndAlso WinKeyPressed Then
                        SmartSnap(hWnd, SnapMode.BottomLeftQuarter)
                    ElseIf WinKeyPressed Then
                        SmartSnap(hWnd, SnapMode.LeftHalf)
                    End If
                End If

            Case 39 ' RIGHT
                If CanSnapWindow(hWnd) = True Then
                    If CtrlKeyPressed AndAlso WinKeyPressed Then
                        SmartSnap(hWnd, SnapMode.TopRightQuarter)
                    ElseIf WinKeyPressed Then
                        SmartSnap(hWnd, SnapMode.RightHalf)
                    End If
                End If

            Case 38 ' UP
                If CanSnapWindow(hWnd) = True Then
                    If CtrlKeyPressed AndAlso WinKeyPressed Then
                        SmartSnap(hWnd, SnapMode.TopLeftQuarter)
                    ElseIf WinKeyPressed Then
                        If state = 2 Then
                            ShowWindow(hWnd, 9)
                        ElseIf state = 1 Then
                            ShowWindow(hWnd, 3)
                        End If
                    End If
                End If

            Case 40 ' DOWN
                If CanSnapWindow(hWnd) = True Then
                    If CtrlKeyPressed AndAlso WinKeyPressed Then
                        SmartSnap(hWnd, SnapMode.BottomRightQuarter)
                    ElseIf WinKeyPressed Then
                        If state = 3 Then
                            ShowWindow(hWnd, 9)
                        Else
                            ShowWindow(hWnd, 6)
                        End If
                    End If
                End If

            Case 65 ' A
                If Not DateAndTime.Visible Then
                    DateAndTime.Show()
                    SetForegroundWindow(DateAndTime.Handle)
                    DateAndTime.Focus()
                Else
                    SetForegroundWindow(DateAndTime.Handle)
                    DateAndTime.Focus()
                End If

            Case 66 ' B
                AppBar.ToolStrip1.Focus()
                AppBar.ToolStrip1.Select()

            Case 67 ' C
                Dim prtoexecute As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "Assistant", "")
                If File.Exists(prtoexecute) Then
                    Process.Start(New ProcessStartInfo(prtoexecute) With {.UseShellExecute = True})
                End If

            Case 68 ' D
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
                If Not AppbarProperties.Visible Then
                    AppbarProperties.Show(AppBar)
                    SetForegroundWindow(AppbarProperties.Handle)
                    AppbarProperties.Focus()
                Else
                    SetForegroundWindow(AppbarProperties.Handle)
                    AppbarProperties.Focus()
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
                If Not RunDialog.Visible Then
                    RunDialog.Show()
                    Task.Delay(30).ContinueWith(Sub() SetForegroundWindow(RunDialog.Handle))
                Else
                    RunDialog.Focus()
                    Task.Delay(30).ContinueWith(Sub() SetForegroundWindow(RunDialog.Handle))
                End If

            Case 83 ' S
                Dim prtoexecute As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "SearchApp", "")
                If File.Exists(prtoexecute) Then
                    Process.Start(New ProcessStartInfo(prtoexecute) With {.UseShellExecute = True})
                End If

            Case 84 ' T
                AppBar.ProcessStrip.Focus()
                AppBar.ProcessStrip.Select()

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
                If Not AppBar.ActionCM.Visible Then
                    AppBar.ActionCM.Show(AppBar, AppBar.Location)
                    AppBar.ActionCM.BringToFront()
                    AppBar.ActionCM.Focus()
                Else
                    AppBar.ActionCM.BringToFront()
                    AppBar.ActionCM.Focus()
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
                        ExecutePinnedBar(0)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "1", "")
                        If File.Exists(executePath) Then Process.Start(New ProcessStartInfo(executePath) With {.UseShellExecute = True})
                    End If
                End If

            Case Keys.D2, Keys.NumPad2 ' 2

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 20)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(1)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "2", "")
                        If File.Exists(executePath) Then Process.Start(New ProcessStartInfo(executePath) With {.UseShellExecute = True})
                    End If
                End If

            Case Keys.D3, Keys.NumPad3 ' 3

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 30)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(2)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "3", "")
                        If File.Exists(executePath) Then Process.Start(New ProcessStartInfo(executePath) With {.UseShellExecute = True})
                    End If
                End If

            Case Keys.D4, Keys.NumPad4 ' 4

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 40)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(3)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "4", "")
                        If File.Exists(executePath) Then Process.Start(New ProcessStartInfo(executePath) With {.UseShellExecute = True})
                    End If
                End If

            Case Keys.D5, Keys.NumPad5 ' 5

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 50)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(4)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "5", "")
                        If File.Exists(executePath) Then Process.Start(New ProcessStartInfo(executePath) With {.UseShellExecute = True})
                    End If
                End If

            Case Keys.D6, Keys.NumPad6 ' 6

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 60)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(5)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "6", "")
                        If File.Exists(executePath) Then Process.Start(New ProcessStartInfo(executePath) With {.UseShellExecute = True})
                    End If
                End If

            Case Keys.D7, Keys.NumPad7 ' 7

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 70)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(6)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "7", "")
                        If File.Exists(executePath) Then Process.Start(New ProcessStartInfo(executePath) With {.UseShellExecute = True})
                    End If
                End If

            Case Keys.D8, Keys.NumPad8 ' 8

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 80)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(7)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "8", "")
                        If File.Exists(executePath) Then Process.Start(New ProcessStartInfo(executePath) With {.UseShellExecute = True})
                    End If
                End If

            Case Keys.D9, Keys.NumPad9 ' 9

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 90)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(8)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "9", "")
                        If File.Exists(executePath) Then Process.Start(New ProcessStartInfo(executePath) With {.UseShellExecute = True})
                    End If
                End If

            Case Keys.D0, Keys.NumPad0 ' 0

                If AltKeyPressed Then
                    SetWindowOpacity(hWnd, 100)
                    Return True
                Else
                    If isCustomHotkey Then
                        ExecutePinnedBar(9)
                    Else
                        Dim executePath As String = My.Computer.Registry.GetValue(hotkeyStart, "0", "")
                        If File.Exists(executePath) Then Process.Start(New ProcessStartInfo(executePath) With {.UseShellExecute = True})
                    End If
                End If
#End Region
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