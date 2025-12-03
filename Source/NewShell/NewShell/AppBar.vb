Imports System.ComponentModel
Imports System.Configuration.Assemblies
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Net.Security
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports Microsoft.VisualBasic.Devices

Public Class AppBar

    Private Const WM_MOUSEACTIVATE As Integer = &H21
    Private Const MA_NOACTIVATE As Integer = 3
    Protected Overrides Sub WndProc(ByRef m As Message)
        MyBase.WndProc(m)

        If m.Msg = WM_MOUSEACTIVATE Then
            Dim clickedControl As Control = Me.GetChildAtPoint(Me.PointToClient(Cursor.Position))

            If Not clickedControl Is Nothing AndAlso clickedControl Is Me.ProcessStrip Then
                m.Result = CType(MA_NOACTIVATE, IntPtr)
            End If
        End If
    End Sub
    <Runtime.InteropServices.DllImport("dwmapi.dll")> Public Shared Function DwmExtendFrameIntoClientArea(ByVal hWnd As IntPtr, ByRef pMarinset As side) As Integer
    End Function

    <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)> Public Structure side
        Public Left As Integer
        Public Right As Integer
        Public Top As Integer
        Public Bottom As Integer
    End Structure

    Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

    Private Function GetActiveWindow() As Long
        Dim xHwnd As Long
        xHwnd = GetForegroundWindow()
        GetActiveWindow = xHwnd
    End Function

    Sub GetWindowSize(ByVal hWnd As Long, Width As Long, Height As Long)
        Dim rc As RECT
        GetWindowRect(hWnd, rc)
        Width = rc.right - rc.left
        Height = rc.bottom - rc.top
    End Sub

    'Window Capturing
    <DllImport("user32.dll", SetLastError:=True)>
    Friend Shared Function PrintWindow(hwnd As IntPtr, hDC As IntPtr, nFlags As UInteger) As Boolean
    End Function

    <DllImport("dwmapi.dll", SetLastError:=True)>
    Friend Shared Function DwmGetWindowAttribute(hwnd As IntPtr, dwAttribute As DWMWINDOWATTRIBUTE, ByRef pvAttribute As Rectangle, cbAttribute As Integer) As Integer
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
    Public Shared Function RenderWindow(hWnd As IntPtr, clientAreaOnly As Boolean, Optional tryGetFullContent As Boolean = False) As Bitmap
        Dim printOption = If(clientAreaOnly, PrintWindowOptions.PW_CLIENTONLY, PrintWindowOptions.PW_DEFAULT)
        printOption = If(tryGetFullContent, PrintWindowOptions.PW_RENDERFULLCONTENT, printOption)

        Dim dwmRect = New Rectangle()
        Dim hResult = DwmGetWindowAttribute(hWnd, DWMWINDOWATTRIBUTE.DWMWA_EXTENDED_FRAME_BOUNDS, dwmRect, Marshal.SizeOf(Of Rectangle)())
        If hResult < 0 Then
            Marshal.ThrowExceptionForHR(hResult)
            Return Nothing
        End If

        Dim bmp = New Bitmap(dwmRect.Width - dwmRect.X, dwmRect.Height - dwmRect.Y)
        Using g = Graphics.FromImage(bmp)
            Dim hDC = g.GetHdc()

            Try
                Dim success = PrintWindow(hWnd, hDC, printOption)
                ' success result not fully handled here
                If Not success Then
                    Dim win32Error = Marshal.GetLastWin32Error()
                    Return Nothing
                End If
                Return bmp
            Finally
                g.ReleaseHdc(hDC)
            End Try

        End Using
    End Function

#Region " Process System"
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function GetClassLongPtr(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
    End Function

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

    Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (
    ByVal hwnd As IntPtr,
    ByVal lpString As String,
    ByVal cch As Integer) As Integer

    Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (
    ByVal hwnd As IntPtr) As Integer

    Public Declare Function IsWindowVisible Lib "user32" (
    ByVal hwnd As IntPtr) As Boolean

    Public Declare Function GetAncestor Lib "user32" (
    ByVal hwnd As IntPtr,
    ByVal gaFlags As Integer) As IntPtr

    Public Declare Function GetParent Lib "user32" (
    ByVal hwnd As IntPtr) As IntPtr

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

    Private Function IsTaskbarWindow(ByVal hwnd As IntPtr) As Boolean
        If Not IsWindowVisible(hwnd) Then Return False

        If GetParent(hwnd) <> IntPtr.Zero Then Return False

        Dim rootOwner As IntPtr = GetAncestor(hwnd, GA_ROOTOWNER)
        If rootOwner <> hwnd Then Return False

        Dim exStyle As Long = GetWindowLong(hwnd, GWL_EXSTYLE)


        If (exStyle And WS_EX_TOOLWINDOW) = WS_EX_TOOLWINDOW AndAlso
       (exStyle And WS_EX_APPWINDOW) <> WS_EX_APPWINDOW Then
            Return False
        End If


        If GetWindowTextLength(hwnd) = 0 Then
            If (exStyle And WS_EX_APPWINDOW) <> WS_EX_APPWINDOW Then
                Return False
            End If
        End If

        Return True
    End Function


    Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As IntPtr) As Boolean

    Public Const WM_GETICON As Integer = &H7F
    Public Const ICON_SMALL As Integer = 0
    Public Const ICON_BIG As Integer = 1
    Public Const GCL_HICON As Integer = -14
    Public Const GCL_HICONSM As Integer = -34

    Private Sub DisposeIconCache()
        If Not ToolStripIconCache Is Nothing Then
            For Each img In ToolStripIconCache.Values
                img.Dispose()
            Next
            ToolStripIconCache.Clear()
        End If
    End Sub

    Public Function GetWindowIcon(ByVal hWnd As IntPtr) As Icon
        Dim hIcon As IntPtr = IntPtr.Zero

        hIcon = SendMessage(hWnd, WM_GETICON, CType(ICON_BIG, IntPtr), IntPtr.Zero)
        If hIcon = IntPtr.Zero Then
            hIcon = SendMessage(hWnd, WM_GETICON, CType(ICON_SMALL, IntPtr), IntPtr.Zero)
        End If

        If hIcon = IntPtr.Zero Then
            hIcon = GetClassLongPtr(hWnd, GCL_HICON)
        End If
        If hIcon = IntPtr.Zero Then
            hIcon = GetClassLongPtr(hWnd, GCL_HICONSM)
        End If

        If hIcon <> IntPtr.Zero Then
            Try
                Dim iconCopy As Icon = Icon.FromHandle(hIcon).Clone()
                Return iconCopy
            Catch ex As Exception
                Return Nothing
            End Try
        End If

        Return Nothing
    End Function

    Private WindowList As New List(Of TaskbarWindowInfo)

    Private Function EnumWindowsCallback(ByVal hwnd As IntPtr, ByVal lParam As IntPtr) As Boolean
        If IsTaskbarWindow(hwnd) Then
            Dim titleLength As Integer = GetWindowTextLength(hwnd)
            If titleLength > 0 Then
                Dim windowTitle As String = New String(ChrW(0), titleLength + 1)
                GetWindowText(hwnd, windowTitle, titleLength + 1)

                Dim processId As Integer = 0
                GetWindowThreadProcessId(hwnd, processId)

                WindowList.Add(New TaskbarWindowInfo() With {
                .Handle = hwnd,
                .Title = windowTitle.TrimEnd(ChrW(0)),
                .PID = processId
            })
            End If
        End If

        Return True
    End Function
    Private LastKnownWindowCount As Integer = 0
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

        If Not winIcon Is Nothing Then
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
    Private IsReordering As Boolean = False
    Private LastKnownPids As New List(Of Integer)
    Public ActiveWindowHandle As IntPtr = IntPtr.Zero
    Private Sub CheckForProcessUpdates()
        Dim currentPids As List(Of Integer) = GetTaskbarApplications().Select(Function(w) w.PID).Distinct().ToList()

        Dim currentWindows As List(Of TaskbarWindowInfo) = GetTaskbarApplications()
        Dim currentWindowCount As Integer = currentWindows.Count

        If IsReordering Then
            LastKnownPids = currentPids.ToList()
            IsReordering = False
            Return
        End If

        If currentWindowCount <> LastKnownWindowCount Then
            ToolStripIconCache.Clear()
            LoadApps()

            LastKnownWindowCount = currentWindowCount
        End If
    End Sub

    Private ReadOnly IconCache As New Dictionary(Of String, Integer)

    Private ReadOnly ProcessIconCache As New Dictionary(Of String, Integer)
    Private ReadOnly WindowIconCache As New Dictionary(Of IntPtr, Integer)
#End Region

#Region " AppBar "
    <StructLayout(LayoutKind.Sequential)> Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
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
    <DllImport("User32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function RegisterWindowMessage(ByVal msg As String) As Integer
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
    Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Integer) As Integer

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

    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long,
                                                        ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long,
                                                        ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
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

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "UseSystemColor", 1) = 1 Then

            Dim accentColor As Color = GetAccentColor

            Me.BackColor = accentColor
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", ColorTranslator.ToHtml(accentColor))

        End If

        fBarRegistered = False
        RegisterBar()
        Me.BringToFront()
        Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - Me.Height)

        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowBackImage", False) = 1 Then
            If IO.File.Exists(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImage", "")) Then
                Try
                    Me.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImage", ""))
                    AppbarProperties.PictureBox1.Image = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImage", ""))
                    AppbarProperties.ComboBox2.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackImage", "")
                Catch ex As Exception
                    Me.BackgroundImage = My.Resources.AppBarMainTransparent1
                    AppbarProperties.PictureBox1.Image = My.Resources.AppBarMainTransparent1
                    AppbarProperties.ComboBox2.Text = ""
                End Try
            Else
                Me.BackgroundImage = My.Resources.AppBarMainTransparent1
                AppbarProperties.PictureBox1.Image = My.Resources.AppBarMainTransparent1
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
            Panel6.Visible = False
        Else
            Splitter1.Visible = True
            Splitter2.Visible = True
            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "SpaceForEmergeTray", "0") = 0 Then
                Splitter3.Visible = True
            Else
                Panel6.Visible = True
                Splitter3.Visible = False
            End If
        End If

        ' Start button

        Button1.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "StartButton", True)
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 0 Then
            AppbarProperties.RadioButton8.Checked = True
            Button1.BackgroundImage = My.Resources.StartRight
        ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 1 Then
            AppbarProperties.RadioButton9.Checked = True
            Try
                Button1.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Normal", ""))
            Catch ex As Exception
                Button1.BackgroundImage = My.Resources.StartRight
            End Try
            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Layout", "0") = 0 Then
                Button1.BackgroundImageLayout = ImageLayout.Stretch
            Else
                Button1.BackgroundImageLayout = ImageLayout.Center
            End If
        ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 2 Then
            'ORB Code Here
        End If

        '--------EMERGE-TRAY---------

        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "SpaceForEmergeTray", "0") = 0 Then
            ' If emerge Desktop is not installed:

            If Not Me.Width = SystemInformation.PrimaryMonitorSize.Width Then
                Me.Width = SystemInformation.PrimaryMonitorSize.Width
            End If

            AppbarProperties.CheckBox9.Checked = False
            Panel4.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "MenuBar", True)
            ClockTray.Hide()
            'Controller.Enabled = True
            'ClockTray.Controller.Enabled = False
            Try
                Dim hWnd As Long
                hWnd = FindWindow("EmergeDesktopApplet", vbNullString)
                Dim wp As WINDOWPLACEMENT
                wp.Length = Marshal.SizeOf(wp)
                GetWindowPlacement(FindWindow("EmergeDesktopApplet", ""), wp)

                Dim wp2 As WINDOWPLACEMENT
                wp2.showCmd = ShowWindowCommands.ShowNoActivate
                wp2.ptMinPosition = wp.ptMinPosition
                wp2.ptMaxPosition = New POINTAPI(0, 0)
                wp2.rcNormalPosition = New RECT(Me.Width - Panel4.Width, Me.Location.Y - 1, Me.Width - Button2.Width - DayLabel.Width - ToolStripButton1.Width - ToolStripButton3.Width - 3, Me.Location.Y + Me.Height)

                wp2.flags = wp.flags
                wp2.Length = Marshal.SizeOf(wp2)

                SetWindowPlacement(FindWindow("EmergeDesktopApplet", ""), wp2)
                SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

                For Each pr As Process In Process.GetProcesses
                    If pr.ProcessName = "emergeCore" Then
                        pr.Kill()
                    End If

                    If pr.ProcessName = "emergeWorkspace" Then
                        ShowWindow(pr.MainWindowHandle, SHOW_WINDOW.SW_HIDE)
                    End If
                Next
            Catch ex As Exception
            End Try

        ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "SpaceForEmergeTray", "0") = 1 Then
            ' If emerge desktop is installed:

            Dim ValueW As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "CustomWidth", SystemInformation.PrimaryMonitorSize.Width / 2 + SystemInformation.PrimaryMonitorSize.Width / 4)

            AppbarProperties.CheckBox9.Checked = True
            SizeLocked = False

            If Not Me.Width = ValueW Then
                Me.Width = ValueW
            End If

            Panel4.Visible = False
            ClockTray.Show()

            Try
                Dim hWnd As IntPtr = FindWindow("EmergeDesktopApplet", vbNullString)
                If hWnd <> IntPtr.Zero Then
                    Dim newX As Integer = Me.Width
                    Dim newY As Integer = Me.Location.Y
                    Dim newWidth As Integer = SystemInformation.PrimaryMonitorSize.Width - Me.Width - ClockTray.Width
                    Dim newHeight As Integer = Me.Height

                    MoveAndResizeWindow(hWnd, newX, newY, newWidth, newHeight)
                Else
                    MessageBox.Show("No Window has been found.", "Info")
                End If

                For Each pr As Process In Process.GetProcesses
                    If pr.ProcessName = "emergeCore" AndAlso pr.ProcessName = "Explorer.exe" Then
                        pr.Kill()
                    End If

                    If pr.ProcessName = "emergeWorkspace" Then
                        ShowWindow(pr.MainWindowHandle, SHOW_WINDOW.SW_HIDE)
                    End If
                Next
            Catch ex As Exception

            End Try
            SizeLocked = True
        End If
        If Not Me.WindowState = FormWindowState.Normal Then
            Me.WindowState = FormWindowState.Normal
        End If

        StartButtonToolStripMenuItem.Checked = Button1.Visible
        Try
            Button1.Width = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer1", "50")
        Catch ex As Exception
            Button1.Width = 50
        End Try

        If GlobalKeyboardHook.SetHook() Then
            Debug.WriteLine("Hooks set up!")
        Else
            MessageBox.Show("Hooks failed to set up!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
            Panel4.Width = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer3", "164")
        Catch ex As Exception
            Panel4.Width = 164
        End Try

        LoadApps()

        Me.TopMost = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "OnTop", True)
        LockAppbarToolStripMenuItem.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Locked", False)
        AutoHideToolStripMenuItem.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "AutoHide", 0)

        ' Desktop

        If SystemInformation.MonitorCount > 1 Then
            Dim screens As Screen() = Screen.AllScreens
            For Each screen As Screen In screens
                Dim newDesktop As New Desktop
                newDesktop.Location = screen.Bounds.Location
                newDesktop.Size = screen.Bounds.Size
                newDesktop.Show()
                newDesktop.SendToBack()
            Next
        Else
            Desktop.Location = New Point(0, 0)
            Desktop.Size = SystemInformation.PrimaryMonitorSize
            Desktop.Show()
            Desktop.SendToBack()
        End If

        ' Alarms

        StartClockUpdate()

        Try
0:
            For Each i In My.Computer.Registry.CurrentUser.OpenSubKey("ALARMS\", True).GetSubKeyNames
                Dialog3.AlarmList.Items.Add(i.Substring(i.LastIndexOf("\") + 1))
            Next
        Catch ex As Exception
            My.Computer.Registry.CurrentUser.CreateSubKey("ALARMS", True)
            GoTo 0
        End Try
        ALC.Items.AddRange(Dialog3.AlarmList.Items)
        Dim rvA = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\", "DisableAlarms", "0")
        If rvA = "0" Then
            DisableAlarmOption.Checked = False
            AlarmController.Enabled = True
        Else
            DisableAlarmOption.Checked = True
            AlarmController.Enabled = False
        End If
    End Sub

    Private Sub StartClockUpdate()
        Task.Run(Sub()
                     While True
                         Dim currentTime As String = DateTime.Now.ToString("HH:mm:ss")
                         Dim currentDay As String = DateTime.Now.DayOfWeek.ToString
                         Dim currentDate As String = DateTime.Now.ToString("dd. MM. yyyy")

                         If Me.InvokeRequired Then
                             Me.Invoke(Sub()
                                           TimeLabel.Text = currentTime
                                           DayLabel.Text = currentDay
                                           DateLabel.Text = currentDate
                                       End Sub)
                         Else
                             TimeLabel.Text = currentTime
                             DayLabel.Text = currentDay
                             DateLabel.Text = currentDate
                         End If

                         Thread.Sleep(1000)
                     End While
                 End Sub)
    End Sub

    Private Sub GlobalKeyboardHook_HotkeyPressed(sender As Object, e As EventArgs)
        Me.BringToFront()
        Startmenu.FST = True
        StartMenuShowAndHide()
    End Sub

    Private Sub ToolStrip1_ItemAdded(sender As Object, e As ToolStripItemEventArgs) Handles AppStrip.ItemAdded
        e.Item.DisplayStyle = ToolStripItemDisplayStyle.Image
    End Sub

    Private Sub ProcessTimer_Tick(sender As Object, e As EventArgs) Handles ProcessTimer.Tick
        CheckForProcessUpdates()
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
    Private Sub OldProcessStripItemClick(sender As Object)
        If CType(sender, ToolStripMenuItem).Checked = True Then
            Dim App As Process = Process.GetProcessById(CType(sender, ToolStripMenuItem).Tag)
            'Dim foregroundHandle As IntPtr = GetForegroundWindow() - Detecting a focused process
            'Dim foregroundProcessId As Integer = 0
            'GetWindowThreadProcessId(foregroundHandle, foregroundProcessId)

            If App.Id > 0 Then
                If IsIconic(App.MainWindowHandle) Then
                    ShowWindow(App.MainWindowHandle, SHOW_WINDOW.SW_RESTORE)
                Else
                    AppActivate(App.Id)
                End If
            Else
                Try
                    AppActivate(App.Id)
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK)
                End Try
            End If
        Else
            Dim App As Process = Process.GetProcessById(CType(sender, ToolStripMenuItem).Tag)
            If App.Id > 0 Then
                AppActivate(App.Id)
                ShowWindow(App.MainWindowHandle, SHOW_WINDOW.SW_RESTORE)
            End If
        End If
    End Sub
    Private LastActiveWindowHandle As IntPtr = IntPtr.Zero
    Private Sub ProcessStripItemClick(sender As Object, e As MouseEventArgs)

        Dim item = CType(sender, ToolStripButton)

        SaveToolStripOrder()

        If Not item.Tag Is Nothing AndAlso TypeOf item.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(item.Tag, IntPtr)

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
                Catch ex As Exception

                End Try

                TPCM.Show(Control.MousePosition)

            ElseIf e.Button = MouseButtons.Middle Then

                WAT.Close()

                TPCM.Tag = windowHandle

                Try
                    StartSameProcessToolStripMenuItem.ToolTipText = p.StartInfo.FileName
                Catch ex As Exception

                End Try

                WindowStateToolStripMenuItem.DropDown.Show(Cursor.Position)
            End If
        End If
    End Sub

    Private Sub ProcessStripSubItemClick(sender As Object, e As MouseEventArgs)

        Dim item = CType(sender, ToolStripMenuItem)

        SaveToolStripOrder()

        If Not item.Tag Is Nothing AndAlso TypeOf item.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(item.Tag, IntPtr)

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

                TPCM.Tag = windowHandle

                Try
                    StartSameProcessToolStripMenuItem.ToolTipText = p.StartInfo.FileName
                Catch ex As Exception

                End Try

                TPCM.Show(Control.MousePosition)

            ElseIf e.Button = MouseButtons.Middle Then

                TPCM.Tag = CType(sender, ToolStripMenuItem).Tag

                Try
                    StartSameProcessToolStripMenuItem.ToolTipText = p.StartInfo.FileName
                Catch ex As Exception

                End Try

                WindowStateToolStripMenuItem.DropDown.Show(Cursor.Position)
            End If
        End If
    End Sub

    Private Sub ProcessStripItemClickOld(sender As Object, e As MouseEventArgs)
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Try
                Dim hWnd As IntPtr = GetMainWindowHandleByProcessId(CType(sender, ToolStripMenuItem).Tag)
                If hWnd <> IntPtr.Zero Then
                    Dim currentState As String = GetWindowState(hWnd)
                    Select Case currentState
                        Case "Normal"
                            AppActivate(CType(sender, ToolStripMenuItem).Tag)
                            ShowWindow(hWnd, SHOW_WINDOW.SW_NORMAL)
                        Case "Maximized"
                            AppActivate(CType(sender, ToolStripMenuItem).Tag)
                            ShowWindow(hWnd, SHOW_WINDOW.SW_MAXIMIZE)
                        Case "Minimized"
                            AppActivate(CType(sender, ToolStripMenuItem).Tag)
                            ShowWindow(hWnd, SHOW_WINDOW.SW_RESTORE)
                        Case Else
                            AppActivate(CType(sender, ToolStripMenuItem).Tag)
                            ShowWindow(hWnd, SHOW_WINDOW.SW_RESTORE)
                    End Select
                Else
                    ' Action for non-detecting any app.
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Cannot focus a process")
            End Try
        ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
            TPCM.Tag = CType(sender, ToolStripMenuItem).Tag
            Try
                StartSameProcessToolStripMenuItem.ToolTipText = Process.GetProcessById(CType(sender, ToolStripMenuItem).Tag).StartInfo.FileName
            Catch ex As Exception

            End Try
            TPCM.Show(Control.MousePosition)
        ElseIf e.Button = MouseButtons.Middle Then
            TPCM.Tag = CType(sender, ToolStripMenuItem).Tag
            Try
                StartSameProcessToolStripMenuItem.ToolTipText = Process.GetProcessById(CType(sender, ToolStripMenuItem).Tag).StartInfo.FileName
            Catch ex As Exception

            End Try
            WindowStateToolStripMenuItem.DropDown.Show(Cursor.Position)
        End If
    End Sub
    Public MouseAway As Boolean = False

    Private Sub ProcessStripItemMouseLeave(sender As Object, e As EventArgs)
        MouseAway = True
        'WAT.Close()
    End Sub

    Private Sub ProcessStrip_ItemAdded(sender As Object, e As ToolStripItemEventArgs) Handles ProcessStrip.ItemAdded
        SaveToolStripOrder()

        For Each i As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\HiddenProcesses\").GetValueNames
            If e.Item.ToolTipText = i Then
                e.Item.Visible = False
            End If
        Next
    End Sub

    Private Sub DefaultToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DefaultToolStripMenuItem.Click
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_SHOWDEFAULT)
        End If
    End Sub

    Private Sub MaximalizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MaximalizeToolStripMenuItem.Click
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
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
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_NORMAL)
        End If
    End Sub

    Private Sub MinimalizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MinimalizeToolStripMenuItem.Click
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_MINIMIZE)
        End If
    End Sub

    Private Sub ForceMinimalizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ForceMinimalizeToolStripMenuItem.Click
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_FORCEMINIMIZE)
        End If
    End Sub

    Private Sub HIDEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HIDEToolStripMenuItem.Click
        If MsgBox("Are you sure do you want to hide this window process? [" & Process.GetProcessById(TPCM.Tag).ProcessName & "]", MsgBoxStyle.YesNo, "Confirm Box") = MsgBoxResult.Yes Then
            If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
                Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
                ShowWindow(windowHandle, SHOW_WINDOW.SW_HIDE)
            End If
        End If
    End Sub

    Private Sub SwitchToToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SwitchToToolStripMenuItem.Click
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_SHOW)
        End If
    End Sub

    Private Sub SwitchButNotActiveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SwitchButNotActiveToolStripMenuItem.Click
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            ShowWindow(windowHandle, SHOW_WINDOW.SW_SHOWNOACTIVATE)
        End If
    End Sub

    Private Const HWND_TOPMOST = -1 '-- Bring to top and stay there
    Private Const HWND_NOTOPMOST = -2 '-- Put the window into a normal position

    Const WM_CLOSE As Long = &H10
    Const WM_DESTROY As Long = &H2

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            SendMessage(windowHandle, WM_CLOSE, 0, 0)
        End If
    End Sub

    Private Sub StartSameProcessToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartSameProcessToolStripMenuItem.Click
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            Process.Start(GetProcessFromWindowHandle(windowHandle).MainModule.FileName)
        End If
    End Sub

    Private Sub TPCM_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TPCM.Opening

        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
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
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
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
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
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
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
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
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
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
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
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
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
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
                    Startmenu.Visible = False
                    Button1.BackgroundImage = My.Resources.StartRight
                Else
                    Startmenu.Visible = True
                    Button1.BackgroundImage = My.Resources.StartLeft
                End If

            ' Custom Images start menu look
            Case 1
                If Startmenu.Visible = True Then
                    Startmenu.Visible = False
                    Try
                        Button1.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Normal", ""))
                    Catch ex As Exception
                        Button1.BackgroundImage = My.Resources.StartRight
                    End Try
                Else
                    Startmenu.Visible = True
                    Try
                        Button1.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Pressed", ""))
                    Catch ex As Exception
                        Button1.BackgroundImage = My.Resources.StartLeft
                    End Try
                End If

            ' ORB Start menu look (working in progress...)
            Case 2

            Case Else
                If Startmenu.Visible = True Then
                    Startmenu.Visible = False
                    Button1.BackgroundImage = My.Resources.StartRight
                Else
                    Startmenu.Visible = True
                    Button1.BackgroundImage = My.Resources.StartLeft
                End If
        End Select
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.BringToFront()
        Startmenu.FST = True
        StartMenuShowAndHide()
    End Sub

    Private Sub ToolStripButton1_MouseUp(sender As Object, e As MouseEventArgs) Handles ToolStripButton1.MouseUp
        If e.Button = MouseButtons.Left Then
            Try
                VolumeControl.ShowDialog(Me)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End Try
        ElseIf e.Button = MouseButtons.Right Then
            VCM.Show(MousePosition)
        End If
    End Sub

    Private Sub RunDialogToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RunDialogToolStripMenuItem.Click
        RunDialog.Show()
        RunDialog.Activate()
    End Sub

    Private Sub Panel4_Resize(sender As Object, e As EventArgs) Handles Panel4.Resize
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "SpaceForEmergeTray", "0") = 0 Then
            Dim wp As WINDOWPLACEMENT
            wp.Length = Marshal.SizeOf(wp)
            GetWindowPlacement(FindWindow("EmergeDesktopApplet", ""), wp)

            Dim wp2 As WINDOWPLACEMENT
            wp2.showCmd = ShowWindowCommands.ShowNoActivate
            wp2.ptMinPosition = wp.ptMinPosition
            wp2.ptMaxPosition = New POINTAPI(0, 0)
            wp2.rcNormalPosition = New RECT(Me.Width - Panel4.Width, Me.Location.Y, Me.Width - Button2.Width - DayLabel.Width - ToolStripButton1.Width - ToolStripButton3.Width - 3, Me.Location.Y + Me.Height)

            wp2.flags = wp.flags
            wp2.Length = Marshal.SizeOf(wp2)
            SetWindowPlacement(FindWindow("EmergeDesktopApplet", ""), wp2)
        End If
    End Sub
    Public CanClose As Boolean = False
    Private Sub AppBar_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
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
        'For Each pr As Process In Process.GetProcesses
        ' If pr.ProcessName = "emergeTray.exe" Then
        '        pr.Kill()
        ' End If
        ' Next
    End Sub

    Private Sub LockAppbarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LockAppbarToolStripMenuItem.Click
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Locked", LockAppbarToolStripMenuItem.Checked, Microsoft.Win32.RegistryValueKind.DWord)
        If LockAppbarToolStripMenuItem.Checked = True Then
            Splitter1.Visible = False
            Splitter2.Visible = False
            Splitter3.Visible = False
            Panel6.Visible = False
        Else
            Splitter1.Visible = True
            Splitter2.Visible = True
            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "SpaceForEmergeTray", "0") = 0 Then
                Splitter3.Visible = True
            Else
                Panel6.Visible = True
                Splitter3.Visible = False
            End If
        End If
    End Sub

    Private Sub TaskManagerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TaskManagerToolStripMenuItem.Click
        Process.Start("taskmgr.exe", "/d")
    End Sub

    Public Sub LoadAppsOld()
        ProcessStrip.Items.Clear()
        For Each p As Process In Process.GetProcesses
            If Not p.MainWindowTitle = "" Then
                Try
                    Dim ico As Icon = Icon.ExtractAssociatedIcon(p.MainModule.FileName)
                    Dim item As New ToolStripMenuItem
                    With item
                        item.Text = p.MainWindowTitle
                        item.Image = ico.ToBitmap
                        item.ToolTipText = p.ProcessName.ToString
                        item.DisplayStyle = ToolStripItemDisplayStyle.Image
                        item.CheckOnClick = True
                        item.Size = New Size(55, 40)
                        item.Tag = p.Id
                        If p.Responding = False Then
                            .BackColor = Color.Red
                        End If
                        AddHandler item.MouseUp, AddressOf ProcessStripItemClick
                        'AddHandler item.MouseHover, AddressOf ProcessStripItemMouseHover
                        AddHandler item.MouseLeave, AddressOf ProcessStripItemMouseLeave
                    End With
                    ProcessStrip.Items.Add(item)
                Catch ex As Exception
                    Dim item As New ToolStripMenuItem
                    With item
                        item.Text = p.MainWindowTitle
                        item.Image = My.Resources.Program
                        item.ToolTipText = p.ProcessName.ToString
                        item.DisplayStyle = ToolStripItemDisplayStyle.Image
                        item.CheckOnClick = True
                        item.Size = New Size(55, 40)
                        item.Tag = p.Id
                        If p.Responding = False Then
                            .BackColor = Color.Red
                        End If
                        AddHandler item.MouseUp, AddressOf ProcessStripItemClick
                        'AddHandler item.MouseHover, AddressOf ProcessStripItemMouseHover
                        AddHandler item.MouseLeave, AddressOf ProcessStripItemMouseLeave
                    End With
                    ProcessStrip.Items.Add(item)
                End Try
            End If
        Next
    End Sub

    Private Sub SaveToolStripOrder()
        Dim orderedPids As New List(Of Integer)

        For Each item As ToolStripItem In ProcessStrip.Items

            Dim pidToSave As Integer = -1

            If TypeOf item Is ToolStripButton Then
                If Not item.Tag Is Nothing AndAlso TypeOf item.Tag Is IntPtr Then
                    Dim windowHandle As IntPtr = CType(item.Tag, IntPtr)
                    Dim p As Process = GetProcessFromWindowHandle(windowHandle)
                    If Not p Is Nothing Then pidToSave = p.Id
                End If

            ElseIf TypeOf item Is ToolStripDropDownButton Then

                Dim dropDown As ToolStripDropDownButton = CType(item, ToolStripDropDownButton)

                If dropDown.DropDownItems.Count > 0 Then
                    Dim firstSubItem As ToolStripMenuItem = CType(dropDown.DropDownItems(0), ToolStripMenuItem)
                    If Not firstSubItem.Tag Is Nothing AndAlso TypeOf firstSubItem.Tag Is IntPtr Then
                        Dim windowHandle As IntPtr = CType(firstSubItem.Tag, IntPtr)
                        Dim p As Process = GetProcessFromWindowHandle(windowHandle)
                        If Not p Is Nothing Then pidToSave = p.Id
                    End If
                End If
            End If

            If pidToSave > 0 Then
                orderedPids.Add(pidToSave)
            End If
        Next

        My.Settings.ToolStripOrder = String.Join(",", orderedPids)
        My.Settings.Save()

    End Sub
    Private IsDragging As Boolean = False
    Private DraggedItem As ToolStripItem = Nothing
    Private OriginalIndex As Integer = -1
    Public Sub LoadApps()
        If Not ProcessStrip Is Nothing Then
            ProcessStrip.Items.Clear()
        End If
        DisposeIconCache()

        Dim applications As List(Of TaskbarWindowInfo) = GetTaskbarApplications()
        Dim groupedWindows As List(Of IGrouping(Of Integer, TaskbarWindowInfo)) = applications.GroupBy(Function(w) w.PID).ToList()

        Dim savedOrderString As String = My.Settings.ToolStripOrder
        Dim savedOrderPids As New List(Of Integer)
        If Not String.IsNullOrEmpty(savedOrderString) Then
            For Each pidStr In savedOrderString.Split(","c)
                Dim pid As Integer
                If Integer.TryParse(pidStr, pid) Then
                    savedOrderPids.Add(pid)
                End If
            Next
        End If

        Dim sortedGroups As New List(Of IGrouping(Of Integer, TaskbarWindowInfo))
        Dim usedPids As New HashSet(Of Integer)()

        For Each pid In savedOrderPids
            Dim group = groupedWindows.FirstOrDefault(Function(g) g.Key = pid)
            If Not group Is Nothing AndAlso Not usedPids.Contains(pid) Then
                sortedGroups.Add(group)
                usedPids.Add(pid)
            End If
        Next

        Dim newGroups = groupedWindows.Where(Function(g) Not usedPids.Contains(g.Key)).ToList()

        Dim reliablyOrderedNewGroups = newGroups.OrderBy(Function(g) g.First().Title).ToList()

        sortedGroups.AddRange(reliablyOrderedNewGroups)

        For Each group In sortedGroups
            Try
                Dim p As Process = Process.GetProcessById(group.Key)
                Dim processName As String = p.ProcessName
                Dim processPath As String = p.MainModule.FileName

                Dim windowsCount As Integer = group.Count()

                Dim processImage As Image = Nothing
                Try
                    Dim processIcon As Icon = System.Drawing.Icon.ExtractAssociatedIcon(processPath)
                    If processIcon IsNot Nothing Then
                        processImage = processIcon.ToBitmap()
                    End If
                Catch : End Try

                If windowsCount = 1 Then
                    Dim windowInfo As TaskbarWindowInfo = group.First()

                    Dim windowImage As Image = GetProcessIcon(windowInfo.Handle, processImage)

                    Dim item As New ToolStripButton()
                    item.Text = windowInfo.Title
                    item.Tag = windowInfo.Handle
                    item.Image = windowImage

                    item.ToolTipText = windowInfo.Title

                    item.AutoSize = False
                    item.AutoToolTip = True

                    item.ImageScaling = ToolStripItemImageScaling.None

                    Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\", "CombineMode", "0")
                        Case 0
                            item.DisplayStyle = ToolStripItemDisplayStyle.Image
                            item.Size = New Size(48, Me.Height)

                        Case 1
                            item.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText
                            item.Size = New Size(150, Me.Height)

                            Dim maxWidth As Integer = 20
                            If item.Text.Length > maxWidth Then
                                item.Text = item.Text.Substring(0, maxWidth - 3) & "..."
                            End If

                            item.TextImageRelation = TextImageRelation.ImageBeforeText
                            item.TextAlign = ContentAlignment.MiddleLeft
                            item.ImageAlign = ContentAlignment.MiddleLeft

                            If Me.BackColor.R >= 150 OrElse Me.BackColor.G >= 130 AndAlso Me.BackColor.B >= 130 Then
                                item.ForeColor = Color.Black
                            Else
                                item.ForeColor = Color.White
                            End If

                    End Select

                    If windowInfo.Handle = ActiveWindowHandle Then
                        item.Checked = True
                    End If

                    If p.Responding = False Then
                        item.BackColor = Color.Red
                    End If

                    AddHandler item.MouseUp, AddressOf ProcessStripItemClick
                    AddHandler item.MouseEnter, AddressOf ProcessStripItemMouseEnter
                    AddHandler item.MouseLeave, AddressOf ProcessStripItemMouseLeave
                    AddHandler item.MouseMove, Sub(sender As Object, e As MouseEventArgs)
                                                   ProcessStrip.Focus()
                                               End Sub

                    ProcessStrip.Items.Add(item)

                Else
                    Dim processItem As New ToolStripDropDownButton()
                    processItem.Text = $"{processName}.exe"
                    processItem.Image = processImage

                    processItem.ToolTipText = p.ProcessName.ToString

                    processItem.AutoSize = False

                    Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\", "CombineMode", "0")
                        Case 0
                            processItem.DisplayStyle = ToolStripItemDisplayStyle.Image
                            processItem.Size = New Size(48, Me.Height)

                        Case 1
                            processItem.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText
                            processItem.Size = New Size(150, Me.Height)

                            Dim maxWidth As Integer = 15
                            If processItem.Text.Length > maxWidth Then
                                processItem.Text = processItem.Text.Substring(0, maxWidth - 3) & "..."
                            End If

                            If Me.BackColor.R >= 150 OrElse Me.BackColor.G >= 130 AndAlso Me.BackColor.B >= 130 Then
                                processItem.ForeColor = Color.Black
                            Else
                                processItem.ForeColor = Color.White
                            End If

                    End Select

                    AddHandler processItem.MouseEnter, AddressOf ProcessStripSubItemMouseEnter

                    For Each windowInfo In group
                        Dim item As New ToolStripMenuItem()
                        item.Text = windowInfo.Title
                        item.Tag = windowInfo.Handle

                        If windowInfo.Handle = ActiveWindowHandle Then
                            item.Checked = True
                        End If

                        item.Image = GetProcessIcon(windowInfo.Handle, processImage)

                        AddHandler item.MouseUp, AddressOf ProcessStripSubItemClick

                        If p.Responding = False Then
                            item.BackColor = Color.Yellow
                        End If

                        processItem.DropDownItems.Add(item)
                    Next

                    ProcessStrip.Items.Add(processItem)
                End If

            Catch ex As Exception
                Dim errorItem As New ToolStripLabel()
                errorItem.Text = $"Error/Ended process (PID: {group.Key})"
                ProcessStrip.Items.Add(errorItem)
            End Try
        Next
    End Sub

    Private Sub ProcessStripItemMouseEnter(sender As Object, e As EventArgs)
        If Startmenu.Visible = False Then
            ProcessStrip.Focus()
        End If

        Dim item As ToolStripButton = CType(sender, ToolStripButton)
        MouseAway = False

        Try
            If Not item.Tag Is Nothing AndAlso TypeOf item.Tag Is IntPtr Then
                Dim windowHandle As IntPtr = CType(item.Tag, IntPtr)
                Dim renderedImage As Image = RenderWindow(windowHandle, True)

                TPCM.Tag = windowHandle

                WAT.Label1.Text = item.Text

                If renderedImage IsNot Nothing Then
                    WAT.Button4.BackgroundImage = renderedImage
                Else
                    WAT.Button4.BackgroundImage = item.Image
                End If

                Dim targetHeight As Integer = 140
                If WAT.Button4.BackgroundImage.Height > targetHeight Then
                    Dim aspectRatio As Double = CDbl(WAT.Button4.BackgroundImage.Width) / CDbl(WAT.Button4.BackgroundImage.Height)
                    Dim targetWidth As Integer = CInt(targetHeight * aspectRatio)
                    WAT.Size = New Size(targetWidth + 10, 140)
                Else
                    WAT.Size = New Size(178, WAT.Button4.BackgroundImage.Height)
                End If

                Dim SpawnPoint As Point ' = Me.PointToClient(item.Bounds.Location)

                SpawnPoint.X = MousePosition.X - WAT.Width / 2
                SpawnPoint.Y = SystemInformation.WorkingArea.Height - WAT.Height

                WAT.Location = SpawnPoint
                WAT.Show()
            End If

        Catch ex As Exception
            WAT.Label1.Text = item.Text

            WAT.Button4.BackgroundImage = Nothing
            WAT.Button4.Image = My.Resources.ProgramMedium

            WAT.Location = New Point(Control.MousePosition.X - WAT.Width / 2, SystemInformation.WorkingArea.Height - WAT.Height)
            WAT.Show()
        End Try
    End Sub

    Private Sub ProcessStripSubItemMouseEnter(sender As Object, e As EventArgs)
        If Startmenu.Visible = False Then
            ProcessStrip.Focus()

            Dim item As ToolStripDropDownButton = CType(sender, ToolStripDropDownButton)

            If item.HasDropDownItems = True Then
                item.ShowDropDown()
            End If
        End If
    End Sub

    Private Sub StartButtonToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartButtonToolStripMenuItem.Click
        Button1.Visible = StartButtonToolStripMenuItem.Checked
        Splitter1.Visible = StartButtonToolStripMenuItem.Checked
    End Sub

    Private Sub PinnedBarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PinnedBarToolStripMenuItem.Click
        Panel1.Visible = PinnedBarToolStripMenuItem.Checked
        Splitter2.Visible = PinnedBarToolStripMenuItem.Checked
    End Sub

    Private Sub MenuBarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MenuBarToolStripMenuItem.Click
        Panel4.Visible = MenuBarToolStripMenuItem.Checked
        Splitter3.Visible = MenuBarToolStripMenuItem.Checked
    End Sub

    Private Sub Controller_Tick(sender As Object, e As EventArgs) 'Handles Controller.Tick
        'Time
        TimeLabel.Text = DateTime.Now.ToString("HH:mm:ss")
        DayLabel.Text = DateTime.Now.DayOfWeek.ToString
        DateLabel.Text = DateTime.Now.ToString("dd. MM. yyyy")
    End Sub

    Private Sub AutoHideToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutoHideToolStripMenuItem.Click
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
    Public Hidden As Boolean = False
    Private Sub AppBar_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
        If AutoHideToolStripMenuItem.Checked = True Then
            If Hidden = True Then
                Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - 2)
            Else
                Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - Me.Height)
            End If
        Else
            If Not Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - Me.Height) Then
                Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - Me.Height)
            End If
        End If
    End Sub

    Private Sub ProcessStrip_MouseLeave1(sender As Object, e As EventArgs) Handles ProcessStrip.MouseLeave, AppStrip.MouseLeave, TimeLabel.MouseLeave, DateLabel.MouseLeave, DayLabel.MouseLeave
        Hidden = True
        MouseAway = True
        Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - 2)
    End Sub

    Private Sub ProcessStrip_MouseLeave(sender As Object, e As EventArgs) Handles ProcessStrip.MouseLeave
        MouseAway = True
        WAT.MouseAway = True
    End Sub

    Private Sub ToolStrip1_MouseEnter(sender As Object, e As EventArgs) Handles ToolStrip1.MouseEnter, ProcessStrip.MouseEnter, AppStrip.MouseEnter, TimeLabel.MouseEnter, DateLabel.MouseEnter, DayLabel.MouseEnter
        Hidden = False
        MouseAway = False
        Startmenu.FST = False
        Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - Me.Height)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, ShowDesktopToolStripMenuItem.Click
        Desktop.BringToFront()
    End Sub

    Private Sub PropertiesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PropertiesToolStripMenuItem.Click
        AppbarProperties.ShowDialog(Me)
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

    Public Sub AlarmLoadDefault()
        On Error Resume Next
        REM 1
        Dim RV11 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomB1", Nothing)
        Dim RV11_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "CustomB1", Nothing)
        Dim RV12 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableB2", Nothing)
        Dim RV12_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomB2", Nothing)
        Dim RV12__ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "CustomB2", Nothing)
        Dim RV13 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomPicture", Nothing)
        Dim RV13_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "CustomPicturePath", Nothing)
        Dim RV14 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomTitle", Nothing)
        Dim RV14_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "CustomTitle", Nothing)
        Dim RV15 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "Position", Nothing)
        Dim RV15X = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "X", Nothing)
        Dim RV15Y = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "Y", Nothing)
        Dim RV16 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableRinging", Nothing)
        Dim RV17 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingMode", Nothing)
        Dim RV18 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingALIST", Nothing)
        Dim RV19 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingPath", Nothing)
        Dim RV1a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingSYS", Nothing)
        Dim RV1b = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingRepeat", Nothing)
        Dim RV1c = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingTimesRepeat", Nothing)
        Dim RV1d = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingVol", Nothing)

        APD.NumericUpDown1.Maximum = SystemInformation.PrimaryMonitorSize.Width
        APD.NumericUpDown2.Maximum = SystemInformation.PrimaryMonitorSize.Height
        If RV11 = "1" Then
            APD.remDUD4 = 1
            APD.DomainUpDown4.SelectedIndex = 1
            If Not RV11_ = String.Empty Then
                APD.remTXT3 = RV11_
                APD.TextBox3.Text = RV11_
            End If
        Else
            APD.remDUD4 = 0
            APD.DomainUpDown4.SelectedIndex = 0
            APD.remTXT3 = ""
            APD.TextBox3.Text = ""
        End If
        If RV12 = "1" Then
            APD.remCHB3 = True
            APD.CheckBox2.Checked = True
            If RV12_ = "1" Then
                APD.remDUD5 = 1
                APD.DomainUpDown5.SelectedIndex = 1
                If Not RV12__ = String.Empty Then
                    APD.remTXT4 = RV12__
                    APD.TextBox2.Text = RV12__
                End If
            Else
                APD.remDUD5 = 0
                APD.DomainUpDown5.SelectedIndex = 0
                APD.remTXT4 = ""
                APD.TextBox2.Text = ""
            End If
        Else
            APD.remCHB3 = False
            APD.CheckBox2.Checked = False
        End If
        If RV13 = "2" Then
            APD.remDUD3 = 2
            APD.DomainUpDown2.SelectedIndex = 2
            If Not RV13_ = String.Empty Then
                APD.remTXT2 = RV13_
                APD.TextBox4.Text = RV13_
            End If
        ElseIf RV13 = "1" Then
            APD.remDUD3 = 1
            APD.DomainUpDown2.SelectedIndex = 1
            APD.remTXT2 = ""
            APD.TextBox4.Text = ""
        Else
            APD.remDUD3 = 0
            APD.DomainUpDown2.SelectedIndex = 0
            APD.remTXT2 = ""
            APD.TextBox4.Text = ""
        End If
        If RV14 = "1" Then
            APD.remDUD2 = 1
            APD.DomainUpDown3.SelectedIndex = 1
            If Not RV14_ = String.Empty Then
                APD.remTXT1 = RV14_
                APD.TextBox1.Text = RV14_
            End If
        Else
            APD.remDUD2 = 0
            APD.DomainUpDown3.SelectedIndex = 0
            APD.remTXT1 = ""
            APD.TextBox1.Text = ""
        End If
        If RV15 >= 9 Then
            APD.remCB1 = 9
            APD.ComboBox1.SelectedIndex = 9
            APD.remNUPX = RV15X
            APD.remNUPY = RV15Y
            APD.NumericUpDown1.Value = RV15X
            APD.NumericUpDown2.Value = RV15Y
        ElseIf RV15 < 9 Then
            APD.remCB1 = RV15
            APD.ComboBox1.SelectedIndex = RV15
            APD.remNUPX = RV15X
            APD.remNUPY = RV15Y
            APD.NumericUpDown1.Value = RV15X
            APD.NumericUpDown2.Value = RV15Y
        End If
        If RV16 = "1" Then
            APD.remCHB1 = True
            APD.CheckBox1.Checked = True
        Else
            APD.remCHB1 = True
            APD.CheckBox1.Checked = True
        End If
        If RV1b = "1" Then
            APD.remCHB2 = True
            APD.CheckBox16.Checked = True
        Else
            APD.remCHB2 = False
            APD.CheckBox16.Checked = False
        End If
        If RV1c >= 1 Then
            APD.remNUP2 = RV1c
            APD.ThmN.Value = RV1c
        End If
        If RV1d >= 0 AndAlso RV1d <= 100 Then
            APD.remNUP1 = RV1d
            APD.NumericUpDown12.Value = RV1d
        End If
        If RV17 = "0" Then
            APD.remDUD1 = 0
            APD.DomainUpDown1.SelectedIndex = 0
            APD.remCB2 = RV1a
            APD.ComboBox2.SelectedIndex = RV1a
        ElseIf RV17 = "1" Then
            Dim i As String
            Dim CTSPath As String = "C:\Windows\Media\Alarms"
            If My.Computer.FileSystem.DirectoryExists(CTSPath) Then
                For Each i In My.Computer.FileSystem.GetFiles(CTSPath)
                    APD.ComboBox3.Items.Add(i.Substring(i.LastIndexOf("\") + 1))
                Next
            End If
            APD.remDUD1 = 1
            APD.DomainUpDown1.SelectedIndex = 1
            APD.remCB3 = RV18
            APD.ComboBox3.SelectedIndex = RV18
        Else
            APD.remDUD1 = 2
            APD.DomainUpDown1.SelectedIndex = 2
            APD.remCB4 = RV19
            APD.ComboBox4.Text = RV19
        End If
        REM 2
        Dim RV21 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingMode", Nothing)
        Dim RV22 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingALIST", Nothing)
        Dim RV23 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingPath", Nothing)
        Dim RV24 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingSYS", Nothing)
        Dim RV25 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingRepeat", Nothing)
        Dim RV26 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingTimesRepeat", Nothing)
        Dim RV27 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingVol", Nothing)

        If RV25 = "1" Then
            APD.remCHB11 = True
            APD.CheckBox3.Checked = True
        Else
            APD.remCHB11 = False
            APD.CheckBox3.Checked = False
        End If
        If RV26 >= 1 Then
            APD.remNUP12 = RV26
            APD.NumericUpDown3.Value = RV26
        End If
        If RV27 >= 0 AndAlso RV27 <= 100 Then
            APD.remNUP11 = RV27
            APD.NumericUpDown4.Value = RV27
        End If
        If RV21 = "0" Then
            APD.remDUD1 = 0
            APD.DomainUpDown1.SelectedIndex = 0
            APD.remCB11 = RV24
            APD.ComboBox6.SelectedIndex = RV24
        ElseIf RV21 = "1" Then
            Dim i As String
            Dim CTSPath As String = "C:\Windows\Media\Alarms"
            If My.Computer.FileSystem.DirectoryExists(CTSPath) Then
                For Each i In My.Computer.FileSystem.GetFiles(CTSPath)
                    APD.ComboBox7.Items.Add(i.Substring(i.LastIndexOf("\") + 1))
                Next
            End If
            APD.remDUD1 = 1
            APD.DomainUpDown1.SelectedIndex = 1
            APD.remCB12 = RV22
            APD.ComboBox7.SelectedIndex = RV22
        Else
            APD.remDUD1 = 2
            APD.DomainUpDown1.SelectedIndex = 2
            APD.remCB13 = RV23
            APD.ComboBox5.Text = RV23
        End If
        REM 3
        Dim RV31 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\", "MODE", Nothing)
        Dim RV32 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\", "Path", Nothing)

        If RV31 = "0" Then
            APD.remRB21 = False
            APD.RadioButton3.Checked = False
            APD.RadioButton2.Checked = True
        Else
            APD.remRB21 = True
            APD.RadioButton3.Checked = True
            APD.RadioButton2.Checked = False
        End If
        If IO.File.Exists(RV32) = True Then
            APD.remCB21 = RV32
            APD.ComboBox8.Text = RV32
        End If
        REM 3.1
        Dim RV311 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "AppWinStyle", Nothing)
        Dim RV312 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "Arguments", Nothing)
        Dim RV313 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "CreateNoWindow", Nothing)
        Dim RV314 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "EnableCustomWorkingDir", Nothing)
        Dim RV315 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "CustomWorkDir", Nothing)
        Dim RV316 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "EnableDomain", Nothing)
        Dim RV316_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "Domain", Nothing)
        Dim RV317 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RunAs", Nothing)
        Dim RV318 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "Username", Nothing)
        Dim RV319 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "LoadUserProfile", Nothing)
        Dim RV31a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "UseShellExecute", Nothing)
        Dim RV31b = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "ErrorDialog", Nothing)
        Dim RV31c = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "ErrorDialogParentHandle", Nothing)
        Dim RV31d = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RSError", Nothing)
        Dim RV31e = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RSInput", Nothing)
        Dim RV31f = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RSOutput", Nothing)

        If RV311 < 4 Then
            APD.remCB211 = RV311
            APD.ComboBox10.SelectedIndex = RV311
        End If
        APD.remTXT211 = RV312
        APD.TextBox5.Text = RV312
        If RV313 = "1" Then
            APD.remCHB211 = True
            APD.CheckBox7.Checked = True
        Else
            APD.remCHB211 = False
            APD.CheckBox7.Checked = False
        End If
        If RV31a = "1" Then
            APD.remCHB212 = True
            APD.CheckBox8.Checked = True
        Else
            APD.remCHB212 = False
            APD.CheckBox8.Checked = False
        End If
        If RV319 = "1" Then
            APD.remCHB214 = True
            APD.CheckBox9.Checked = True
        Else
            APD.remCHB214 = False
            APD.CheckBox9.Checked = False
        End If
        If RV31b = "1" Then
            APD.remCHB217 = True
            APD.CheckBox12.Checked = True
        Else
            APD.remCHB217 = False
            APD.CheckBox12.Checked = False
        End If
        If RV31c = "1" Then
            APD.remCHB218 = True
            APD.CheckBox13.Checked = True
        Else
            APD.remCHB218 = False
            APD.CheckBox13.Checked = False
        End If
        If RV31d = "1" Then
            APD.remCHB219 = True
            APD.CheckBox14.Checked = True
        Else
            APD.remCHB219 = False
            APD.CheckBox14.Checked = False
        End If
        If RV31e = "1" Then
            APD.remCHB21a = True
            APD.CheckBox15.Checked = True
        Else
            APD.remCHB21a = False
            APD.CheckBox15.Checked = False
        End If
        If RV31f = "1" Then
            APD.remCHB21b = True
            APD.CheckBox17.Checked = True
        Else
            APD.remCHB21b = False
            APD.CheckBox17.Checked = False
        End If
        APD.remTXT214 = RV315
        APD.TextBox7.Text = RV315
        If RV314 = "1" Then
            APD.remCHB215 = True
            APD.CheckBox10.Checked = True
        Else
            APD.remCHB215 = False
            APD.CheckBox10.Checked = False
        End If
        APD.remTXT215 = RV316_
        APD.TextBox8.Text = RV316_
        If RV316 = "1" Then
            APD.remCHB216 = True
            APD.CheckBox11.Checked = True
        Else
            APD.remCHB216 = False
            APD.CheckBox11.Checked = False
        End If
        APD.remTXT212 = RV318
        APD.TextBox6.Text = RV318
        If RV317 = "1" Then
            APD.remCHB213 = True
            APD.CheckBox4.Checked = True
        Else
            APD.remCHB213 = False
            APD.CheckBox4.Checked = False
        End If
        REM 3.2
        Dim RV321 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\2\", "AppWinStyle", Nothing)
        Dim RV322 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\2\", "Wait", Nothing)
        Dim RV323 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\2\", "Timeout", Nothing)

        If RV321 < 6 Then
            APD.remCB221 = RV321
            APD.ComboBox9.SelectedIndex = RV321
        End If
        If RV322 = "1" Then
            APD.remCHB221 = True
            APD.CheckBox6.Checked = True
        Else
            APD.remCHB221 = False
            APD.CheckBox6.Checked = False
        End If
        If RV323 >= -1 Then
            APD.remNUP221 = RV323
            APD.NumericUpDown5.Value = RV323
        Else
            APD.remNUP221 = -1
            APD.NumericUpDown5.Value = -1
        End If
    End Sub

    Public Sub NWSet()
        'Read
        Dim readValue1 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "EnableCustomTitle", Nothing)
        Dim readValue1_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "CustomTitle", Nothing)
        Dim readValue2 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "EnableCustomIcon", Nothing)
        Dim readValue2_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "CustomIcon", Nothing)
        If IO.File.Exists(readValue2_) = True Then
            Dim FI As New IO.FileInfo(readValue2_)
            If Not FI.Extension = ".ico" Then
                readValue2_ = "C:\Windows\NotificationWindow\NotifyIcon.ico"
            End If
        Else
            readValue2_ = "C:\Windows\NotificationWindow\NotifyIcon.ico"
        End If
        Dim readValue3 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "EnableCustomBackColor", Nothing)
        Dim readValue3_R = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "CustomBR", Nothing)
        Dim readValue3_G = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "CustomBG", Nothing)
        Dim readValue3_B = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "CustomBB", Nothing)
        Dim readValue3_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "CAffectControls", Nothing)
        Dim readValue4 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "BorderStyle", Nothing)
        Dim readValue5 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "OnTop", Nothing)
        Dim readValue6 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "CloseButton", Nothing)
        Dim readValue7 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "FocusOnActivation", Nothing)
        Dim readValue8 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "SaveBTNChecked", Nothing)
        Dim readValue9 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "EnableCustomTextColor", Nothing)
        Dim readValue9_R = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "CustomTR", Nothing)
        Dim readValue9_G = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "CustomTG", Nothing)
        Dim readValue9_B = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\", "CustomTB", Nothing)
        'Set it up
        If readValue1 = 0 Then
            NotifyWindowD.ComboBox1.SelectedIndex = 0
        Else
            NotifyWindowD.ComboBox1.SelectedIndex = 1
            NotifyWindowD.TextBox1.Text = readValue1_
        End If
        If readValue2 = 0 Then
            NotifyWindowD.ComboBox2.SelectedIndex = 0
        Else
            NotifyWindowD.ComboBox2.SelectedIndex = 1
            NotifyWindowD.ComboBox3.Text = readValue2_
        End If
        If readValue3 = 0 Then
            NotifyWindowD.CheckBox1.Checked = False
        ElseIf readValue3 = 1 Then
            NotifyWindowD.CheckBox1.Checked = True
            NotifyWindowD.RadioButton1.Checked = True
            NotifyWindowD.CL.BackColor = Color.FromArgb(255, readValue3_R, readValue3_G, readValue3_B)
            If readValue3_ = 1 Then
                NotifyWindowD.CheckBox2.Checked = True
            Else
                NotifyWindowD.CheckBox2.Checked = False
            End If
            'NotifyWindowD.CL.BackColor = RGB(readValue3_R, readValue3_G, readValue3_B)
        ElseIf readValue3 = 2 Then
            NotifyWindowD.CheckBox1.Checked = True
            NotifyWindowD.RadioButton2.Checked = True
            NotifyWindowD.CL.BackColor = Color.LightBlue
            If readValue3_ = 1 Then
                NotifyWindowD.CheckBox2.Checked = True
            Else
                NotifyWindowD.CheckBox2.Checked = False
            End If
        End If
        If readValue4 = 0 Then
            NotifyWindowD.ComboBox4.SelectedIndex = 0
        Else
            NotifyWindowD.ComboBox4.SelectedIndex = 1
        End If
        If readValue5 = 0 Then
            NotifyWindowD.CheckBox3.Checked = False
        Else
            NotifyWindowD.CheckBox3.Checked = True
        End If
        If readValue6 = 0 Then
            NotifyWindowD.CheckBox4.Checked = False
        Else
            NotifyWindowD.CheckBox4.Checked = True
        End If
        If readValue7 = 0 Then
            NotifyWindowD.CheckBox5.Checked = False
        Else
            NotifyWindowD.CheckBox5.Checked = True
        End If
        If readValue8 = 0 Then
            NotifyWindowD.CheckBox6.Checked = False
        Else
            NotifyWindowD.CheckBox6.Checked = True
        End If
        If readValue9 = 0 Then
            NotifyWindowD.CheckBox7.Checked = False
        ElseIf readValue3 = 1 Then
            NotifyWindowD.CheckBox7.Checked = True
            NotifyWindowD.CL2.BackColor = Color.FromArgb(255, readValue9_R, readValue9_G, readValue9_B)
        End If
    End Sub

    Private Sub AlarmController_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AlarmController.Tick
        On Error Resume Next
0:
        If Not Dialog3.AlarmList.Items.Count = 0 Then
            If ALC.SelectedIndex = ALC.Items.Count - 1 Then
                If My.Computer.Registry.CurrentUser.OpenSubKey("ALARMS\" & ALC.SelectedItem & "\Repeat\", False) Is Nothing Then
                    Dim RV = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem, "Date", Nothing)
                    Dim RVh = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem, "Hour", Nothing)
                    Dim RVm = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem, "Minute", Nothing)
                    Dim RVs = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem, "Second", Nothing)
                    If RV = DateLabel.Text Then
                        If RVh = System.DateTime.Now.Hour Then
                            If RVm = System.DateTime.Now.Minute Then
                                If RVs = System.DateTime.Now.Second Then
                                    System.Threading.Thread.Sleep(1050)
                                    Dialog3.AlarmList.SelectedIndex = ALC.SelectedIndex
                                    ALACT()
                                    Dialog3.ACTION()
                                End If
                            End If
                        End If
                    End If
                Else
                    Dim RVh = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem, "Hour", Nothing)
                    Dim RVm = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem, "Minute", Nothing)
                    Dim RVs = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem, "Second", Nothing)
                    If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Repeat\", "Monday", Nothing) IsNot Nothing Then
                        If DayLabel.Text = "Monday" Then
                            If RVh = System.DateTime.Now.Hour Then
                                If RVm = System.DateTime.Now.Minute Then
                                    If RVs = System.DateTime.Now.Second Then
                                        System.Threading.Thread.Sleep(1005)
                                        Dialog3.AlarmList.SelectedIndex = ALC.SelectedIndex
                                        ALACT()
                                        Dialog3.ACTION()
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Repeat\", "Tuesday", Nothing) IsNot Nothing Then
                        If DayLabel.Text = "Tuesday" Then
                            If RVh = System.DateTime.Now.Hour Then
                                If RVm = System.DateTime.Now.Minute Then
                                    If RVs = System.DateTime.Now.Second Then
                                        System.Threading.Thread.Sleep(1005)
                                        Dialog3.AlarmList.SelectedIndex = ALC.SelectedIndex
                                        ALACT()
                                        Dialog3.ACTION()
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Repeat\", "Wednesday", Nothing) IsNot Nothing Then
                        If DayLabel.Text = "Wednesday" Then
                            If RVh = System.DateTime.Now.Hour Then
                                If RVm = System.DateTime.Now.Minute Then
                                    If RVs = System.DateTime.Now.Second Then
                                        System.Threading.Thread.Sleep(1005)
                                        Dialog3.AlarmList.SelectedIndex = ALC.SelectedIndex
                                        ALACT()
                                        Dialog3.ACTION()
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Repeat\", "Thursday", Nothing) IsNot Nothing Then
                        If DayLabel.Text = "Thursday" Then
                            If RVh = System.DateTime.Now.Hour Then
                                If RVm = System.DateTime.Now.Minute Then
                                    If RVs = System.DateTime.Now.Second Then
                                        System.Threading.Thread.Sleep(1005)
                                        Dialog3.AlarmList.SelectedIndex = ALC.SelectedIndex
                                        ALACT()
                                        Dialog3.ACTION()
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Repeat\", "Friday", Nothing) IsNot Nothing Then
                        If DayLabel.Text = "Friday" Then
                            If RVh = System.DateTime.Now.Hour Then
                                If RVm = System.DateTime.Now.Minute Then
                                    If RVs = System.DateTime.Now.Second Then
                                        System.Threading.Thread.Sleep(1005)
                                        Dialog3.AlarmList.SelectedIndex = ALC.SelectedIndex
                                        ALACT()
                                        Dialog3.ACTION()
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Repeat\", "Saturday", Nothing) IsNot Nothing Then
                        If DayLabel.Text = "Saturday" Then
                            If RVh = System.DateTime.Now.Hour Then
                                If RVm = System.DateTime.Now.Minute Then
                                    If RVs = System.DateTime.Now.Second Then
                                        System.Threading.Thread.Sleep(1005)
                                        Dialog3.AlarmList.SelectedIndex = ALC.SelectedIndex
                                        ALACT()
                                        Dialog3.ACTION()
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Repeat\", "Sunday", Nothing) IsNot Nothing Then
                        If DayLabel.Text = "Sunday" Then
                            If RVh = System.DateTime.Now.Hour Then
                                If RVm = System.DateTime.Now.Minute Then
                                    If RVs = System.DateTime.Now.Second Then
                                        System.Threading.Thread.Sleep(1005)
                                        Dialog3.AlarmList.SelectedIndex = ALC.SelectedIndex
                                        ALACT()
                                        Dialog3.ACTION()
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                ALC.SelectedIndex = 0
            Else
                If My.Computer.Registry.CurrentUser.OpenSubKey("ALARMS\" & ALC.SelectedItem & "\Repeat\", False) Is Nothing Then
                    Dim RV = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem, "Date", Nothing)
                    Dim RVh = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem, "Hour", Nothing)
                    Dim RVm = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem, "Minute", Nothing)
                    Dim RVs = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem, "Second", Nothing)
                    If RV = DateLabel.Text Then
                        If RVh = System.DateTime.Now.Hour Then
                            If RVm = System.DateTime.Now.Minute Then
                                If RVs = System.DateTime.Now.Second Then
                                    System.Threading.Thread.Sleep(1050)
                                    Dialog3.AlarmList.SelectedIndex = ALC.SelectedIndex
                                    ALACT()
                                    Dialog3.ACTION()
                                End If
                            End If
                        End If
                    End If
                Else
                    Dim RVh = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem, "Hour", Nothing)
                    Dim RVm = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem, "Minute", Nothing)
                    Dim RVs = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem, "Second", Nothing)
                    If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Repeat\", "Monday", Nothing) IsNot Nothing Then
                        If DayLabel.Text = "Monday" Then
                            If RVh = System.DateTime.Now.Hour Then
                                If RVm = System.DateTime.Now.Minute Then
                                    If RVs = System.DateTime.Now.Second Then
                                        System.Threading.Thread.Sleep(1005)
                                        Dialog3.AlarmList.SelectedIndex = ALC.SelectedIndex
                                        ALACT()
                                        Dialog3.ACTION()
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Repeat\", "Tuesday", Nothing) IsNot Nothing Then
                        If DayLabel.Text = "Tuesday" Then
                            If RVh = System.DateTime.Now.Hour Then
                                If RVm = System.DateTime.Now.Minute Then
                                    If RVs = System.DateTime.Now.Second Then
                                        System.Threading.Thread.Sleep(1005)
                                        Dialog3.AlarmList.SelectedIndex = ALC.SelectedIndex
                                        ALACT()
                                        Dialog3.ACTION()
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Repeat\", "Wednesday", Nothing) IsNot Nothing Then
                        If DayLabel.Text = "Wednesday" Then
                            If RVh = System.DateTime.Now.Hour Then
                                If RVm = System.DateTime.Now.Minute Then
                                    If RVs = System.DateTime.Now.Second Then
                                        System.Threading.Thread.Sleep(1005)
                                        Dialog3.AlarmList.SelectedIndex = ALC.SelectedIndex
                                        ALACT()
                                        Dialog3.ACTION()
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Repeat\", "Thursday", Nothing) IsNot Nothing Then
                        If DayLabel.Text = "Thursday" Then
                            If RVh = System.DateTime.Now.Hour Then
                                If RVm = System.DateTime.Now.Minute Then
                                    If RVs = System.DateTime.Now.Second Then
                                        System.Threading.Thread.Sleep(1005)
                                        Dialog3.AlarmList.SelectedIndex = ALC.SelectedIndex
                                        ALACT()
                                        Dialog3.ACTION()
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Repeat\", "Friday", Nothing) IsNot Nothing Then
                        If DayLabel.Text = "Friday" Then
                            If RVh = System.DateTime.Now.Hour Then
                                If RVm = System.DateTime.Now.Minute Then
                                    If RVs = System.DateTime.Now.Second Then
                                        System.Threading.Thread.Sleep(1005)
                                        Dialog3.AlarmList.SelectedIndex = ALC.SelectedIndex
                                        ALACT()
                                        Dialog3.ACTION()
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Repeat\", "Saturday", Nothing) IsNot Nothing Then
                        If DayLabel.Text = "Saturday" Then
                            If RVh = System.DateTime.Now.Hour Then
                                If RVm = System.DateTime.Now.Minute Then
                                    If RVs = System.DateTime.Now.Second Then
                                        System.Threading.Thread.Sleep(1005)
                                        Dialog3.AlarmList.SelectedIndex = ALC.SelectedIndex
                                        ALACT()
                                        Dialog3.ACTION()
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Repeat\", "Sunday", Nothing) IsNot Nothing Then
                        If DayLabel.Text = "Sunday" Then
                            If RVh = System.DateTime.Now.Hour Then
                                If RVm = System.DateTime.Now.Minute Then
                                    If RVs = System.DateTime.Now.Second Then
                                        System.Threading.Thread.Sleep(1005)
                                        Dialog3.AlarmList.SelectedIndex = ALC.SelectedIndex
                                        ALACT()
                                        Dialog3.ACTION()
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                ALC.SelectedIndex += 1
            End If
        Else
            ALC.Items.Clear()
        End If
    End Sub
    Public alarm As New Media.SoundPlayer
    Public Sub ALACT()
        Dim RVM = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS" & ALC.SelectedItem, "MODE", Nothing)
        Dim RV16 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Mode\1\", "EnableRinging", Nothing)
        Dim RV17 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Mode\1\", "RingMode", Nothing)
        Dim RV18 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Mode\1\", "RingALIST", Nothing)
        Dim RV19 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Mode\1\", "RingPath", Nothing)
        Dim RV1a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Mode\1\", "RingSYS", Nothing)
        Dim RV1b = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Mode\1\", "RingRepeat", Nothing)
        Dim RV1c = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Mode\1\", "RingTimesRepeat", Nothing)
        Dim RV1d = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Mode\1\", "RingVol", Nothing)

        Dim RV117 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Mode\2\", "RingMode", Nothing)
        Dim RV118 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Mode\2\", "RingALIST", Nothing)
        Dim RV119 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Mode\2\", "RingPath", Nothing)
        Dim RV11a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Mode\2\", "RingSYS", Nothing)
        Dim RV11b = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Mode\2\", "RingRepeat", Nothing)
        Dim RV11c = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Mode\2\", "RingTimesRepeat", Nothing)
        Dim RV11d = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & ALC.SelectedItem & "\Mode\2\", "RingVol", Nothing)
        If RVM = "0" Then
            If RV16 = "1" Then
                If RV17 = "0" Then
                    If RV1a = "0" Then
                        Media.SystemSounds.Asterisk.Play()
                    ElseIf RV1a = "1" Then
                        Media.SystemSounds.Beep.Play()
                    ElseIf RV1a = "2" Then
                        Media.SystemSounds.Exclamation.Play()
                    ElseIf RV1a = "3" Then
                        Media.SystemSounds.Hand.Play()
                    ElseIf RV1a = "4" Then
                        Media.SystemSounds.Question.Play()
                    End If
                ElseIf RV17 = "1" Then
                    Dialog3.AlCol.Items.Clear()
                    Dim i As String
                    If My.Computer.FileSystem.DirectoryExists("C:\Windows\Media\Alarms\") Then
                        For Each i In My.Computer.FileSystem.GetFiles("C:\Windows\Media\Alarms\")
                            Dialog3.AlCol.Items.Add(i.Substring(i.LastIndexOf("\") + 1))
                        Next
                    End If
                    Dim alrpath As String = "C:\Windows\Media\Alarms\" & Dialog3.AlCol.Items.Item(RV18)
                    alarm = New Media.SoundPlayer(alrpath)
                    If RV1b = "1" Then
                        alarm.PlayLooping()
                    Else
                        'if RingTimesRepeat
                        alarm.Play()
                    End If
                ElseIf RV17 = "2" Then
                    If IO.File.Exists(RV19) = True Then
                        Dim alrpath As String = RV19
                        alarm = New Media.SoundPlayer(alrpath)
                        If RV1b = "1" Then
                            alarm.PlayLooping()
                        Else
                            'if RingTimesRepeat
                            alarm.Play()
                        End If
                    End If
                End If
            End If
        ElseIf RVM = "1" Then
            ALARMSTOPToolStripMenuItem.Visible = True
            ToolStripSeparator20.Visible = True
            If RV117 = "0" Then
                If RV11a = "0" Then
                    Media.SystemSounds.Asterisk.Play()
                ElseIf RV11a = "1" Then
                    Media.SystemSounds.Beep.Play()
                ElseIf RV11a = "2" Then
                    Media.SystemSounds.Exclamation.Play()
                ElseIf RV11a = "3" Then
                    Media.SystemSounds.Hand.Play()
                ElseIf RV11a = "4" Then
                    Media.SystemSounds.Question.Play()
                End If
            ElseIf RV117 = "1" Then
                Dialog3.AlCol.Items.Clear()
                Dim i As String
                If My.Computer.FileSystem.DirectoryExists("C:\Windows\Media\Alarms\") Then
                    For Each i In My.Computer.FileSystem.GetFiles("C:\Windows\Media\Alarms\")
                        Dialog3.AlCol.Items.Add(i.Substring(i.LastIndexOf("\") + 1))
                    Next
                End If
                Dim alrpath As String = "C:\Windows\Media\Alarms\" & Dialog3.AlCol.Items.Item(RV118)
                alarm = New Media.SoundPlayer(alrpath)
                If RV11b = "1" Then
                    alarm.PlayLooping()
                Else
                    'if RingTimesRepeat
                    alarm.Play()
                End If
            ElseIf RV117 = "2" Then
                If IO.File.Exists(RV119) = True Then
                    Dim alrpath As String = RV119
                    alarm = New Media.SoundPlayer(alrpath)
                    If RV11b = "1" Then
                        alarm.PlayLooping()
                    Else
                        'if RingTimesRepeat
                        alarm.Play()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TimeLabel_Click(sender As Object, e As MouseEventArgs) Handles TimeLabel.MouseUp, DayLabel.MouseUp, DateLabel.MouseUp
        If e.Button = MouseButtons.Left Then
            Dialog3.Show()
        ElseIf e.Button = MouseButtons.Right Then
            TaDCM.Show(MousePosition)
        End If
    End Sub

    Private Sub ALARMSTOPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ALARMSTOPToolStripMenuItem.Click
        On Error Resume Next
        Dialog3.ComboBox5.Enabled = True
        Dialog3.Button19.Enabled = True
        Dialog3.Button20.Enabled = True
        Dialog3.AlarmList.Enabled = True
        Dialog3.Button18.Enabled = True
        Dialog3.CheckBox15.Enabled = True
        Dialog3.alarm.Stop()
        alarm.Stop()
        Dialog3.Button17.Text = "TEST"
        ALARMSTOPToolStripMenuItem.Visible = False
        ToolStripSeparator20.Visible = False
    End Sub

    Private Sub ShowTimeAndDateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowTimeAndDateToolStripMenuItem.Click
        Dialog3.Show()
    End Sub

    Private Sub CMMAIN_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles CMMAIN.Opening
        CMMAIN.RightToLeft = RightToLeft.No
    End Sub
    Private Sub WithSecondsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithSecondsToolStripMenuItem.Click
        Dialog3.DateToCopy.Text = TimeLabel.Text
        Dialog3.DateToCopy.SelectAll()
        Dialog3.DateToCopy.Copy()
        Dialog3.ToolTip1.Active = True
        Dialog3.ToolTip1.Show("Successfully Copied!", Me, Me.DayLabel.Width / 2 + Me.DayLabel.Location.X, 2 + Me.DayLabel.Location.Y)
        Dialog3.ToolTipHider.Enabled = True
    End Sub
    Dim fixdayofweek As String
    Private Sub CopyCurrentDayToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyCurrentDayToolStripMenuItem.Click
        If DateTime.Now.DayOfWeek = 1 Then
            fixdayofweek = "Monday"
        ElseIf DateTime.Now.DayOfWeek = 2 Then
            fixdayofweek = "Tuesday"
        ElseIf DateTime.Now.DayOfWeek = 3 Then
            fixdayofweek = "Wednesday"
        ElseIf DateTime.Now.DayOfWeek = 4 Then
            fixdayofweek = "Thursday"
        ElseIf DateTime.Now.DayOfWeek = 5 Then
            fixdayofweek = "Friday"
        ElseIf DateTime.Now.DayOfWeek = 6 Then
            fixdayofweek = "Saturday"
        ElseIf DateTime.Now.DayOfWeek = 7 Then
            fixdayofweek = "Sunday"
        ElseIf DateTime.Now.DayOfWeek = 8 Then
            fixdayofweek = "DayOfErrors:)"
        ElseIf Not DateTime.Now.DayOfWeek >= 8 Then
            fixdayofweek = "Error On Us: <Date Error 404>"
        Else
            fixdayofweek = DateTime.Now.DayOfWeek
        End If
        Dialog3.DateToCopy.Text = fixdayofweek
        Dialog3.DateToCopy.SelectAll()
        Dialog3.DateToCopy.Copy()
        Dialog3.ToolTip1.Active = True
        Dialog3.ToolTip1.Show("Successfully Copied!", Me, DayLabel.Width / 2 + DayLabel.Location.X, 2 + DayLabel.Location.Y)
        Dialog3.ToolTipHider.Enabled = True
    End Sub

    Private Sub CopyEVERYTHINGToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyEVERYTHINGToolStripMenuItem.Click
        Dialog3.DateToCopy.Text = TimeLabel.Text & " " & DayLabel.Text & " " & DateLabel.Text
        Dialog3.DateToCopy.SelectAll()
        Dialog3.DateToCopy.Copy()
        Dialog3.ToolTip1.Active = True
        Dialog3.ToolTip1.Show("Successfully Copied!", Me, DayLabel.Width / 2 + DayLabel.Location.X, 2 + DayLabel.Location.Y)
        Dialog3.ToolTipHider.Enabled = True
    End Sub

    Private Sub CopyDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyDateToolStripMenuItem.Click
        Dialog3.DateToCopy.Text = DateLabel.Text
        Dialog3.DateToCopy.SelectAll()
        Dialog3.DateToCopy.Copy()
        Dialog3.ToolTip1.Active = True
        Dialog3.ToolTip1.Show("Successfully Copied!", Me, DayLabel.Width / 2 + DayLabel.Location.X, 2 + DayLabel.Location.Y)
        Dialog3.ToolTipHider.Enabled = True
    End Sub
    Dim fixsec As String
    Dim fixmin As String
    Dim fixhour As String
    Private Sub CopyTimeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyTimeToolStripMenuItem.Click
        If DateTime.Now.Second = 0 Then
            fixsec = "00"
        ElseIf DateTime.Now.Second = 1 Then
            fixsec = "01"
        ElseIf DateTime.Now.Second = 2 Then
            fixsec = "02"
        ElseIf DateTime.Now.Second = 3 Then
            fixsec = "03"
        ElseIf DateTime.Now.Second = 4 Then
            fixsec = "04"
        ElseIf DateTime.Now.Second = 5 Then
            fixsec = "05"
        ElseIf DateTime.Now.Second = 6 Then
            fixsec = "06"
        ElseIf DateTime.Now.Second = 7 Then
            fixsec = "07"
        ElseIf DateTime.Now.Second = 8 Then
            fixsec = "08"
        ElseIf DateTime.Now.Second = 9 Then
            fixsec = "09"
        Else
            fixsec = DateTime.Now.Second
        End If
        If DateTime.Now.Minute = 0 Then
            fixmin = "00"
        ElseIf DateTime.Now.Minute = 1 Then
            fixmin = "01"
        ElseIf DateTime.Now.Minute = 2 Then
            fixmin = "02"
        ElseIf DateTime.Now.Minute = 3 Then
            fixmin = "03"
        ElseIf DateTime.Now.Minute = 4 Then
            fixmin = "04"
        ElseIf DateTime.Now.Minute = 5 Then
            fixmin = "05"
        ElseIf DateTime.Now.Minute = 6 Then
            fixmin = "06"
        ElseIf DateTime.Now.Minute = 7 Then
            fixmin = "07"
        ElseIf DateTime.Now.Minute = 8 Then
            fixmin = "08"
        ElseIf DateTime.Now.Minute = 9 Then
            fixmin = "09"
        Else
            fixmin = DateTime.Now.Minute
        End If
        If DateTime.Now.Hour = 0 Then
            fixhour = "00"
        ElseIf DateTime.Now.Hour = 1 Then
            fixhour = "01"
        ElseIf DateTime.Now.Hour = 2 Then
            fixhour = "02"
        ElseIf DateTime.Now.Hour = 3 Then
            fixhour = "03"
        ElseIf DateTime.Now.Hour = 4 Then
            fixhour = "04"
        ElseIf DateTime.Now.Hour = 5 Then
            fixhour = "05"
        ElseIf DateTime.Now.Hour = 6 Then
            fixhour = "06"
        ElseIf DateTime.Now.Hour = 7 Then
            fixhour = "07"
        ElseIf DateTime.Now.Hour = 8 Then
            fixhour = "08"
        ElseIf DateTime.Now.Hour = 9 Then
            fixhour = "09"
        Else
            fixhour = DateTime.Now.Hour
        End If
        Dialog3.DateToCopy.Text = DateTime.Now.Hour & ":" & fixmin
        Dialog3.DateToCopy.SelectAll()
        Dialog3.DateToCopy.Copy()
        Dialog3.ToolTip1.Active = True
        Dialog3.ToolTip1.Show("Successfully Copied!", Me, DayLabel.Width / 2 + DayLabel.Location.X, 2 + DayLabel.Location.Y)
        Dialog3.ToolTipHider.Enabled = True
    End Sub

    Private Sub WithMilisecondsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithMilisecondsToolStripMenuItem.Click
        Dialog3.DateToCopy.Text = TimeLabel.Text & "." & DateTime.Now.Millisecond
        Dialog3.DateToCopy.SelectAll()
        Dialog3.DateToCopy.Copy()
        Dialog3.ToolTip1.Active = True
        Dialog3.ToolTip1.Show("Successfully Copied!", Me, DayLabel.Width / 2 + DayLabel.Location.X, 2 + DayLabel.Location.Y)
        Dialog3.ToolTipHider.Enabled = True
    End Sub

    Private Sub HoursToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HoursToolStripMenuItem.Click
        Dialog3.TimeToCopy.Text = DateTime.Now.Hour
        Dialog3.TimeToCopy.SelectAll()
        Dialog3.TimeToCopy.Copy()
        Dialog3.ToolTip1.Active = True
        Dialog3.ToolTip1.Show("Successfully Copied!", Me, DayLabel.Width / 2 + DayLabel.Location.X, 2 + DayLabel.Location.Y)
        Dialog3.ToolTipHider.Enabled = True
    End Sub

    Private Sub MinutesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MinutesToolStripMenuItem.Click
        Dialog3.TimeToCopy.Text = DateTime.Now.Minute
        Dialog3.TimeToCopy.SelectAll()
        Dialog3.TimeToCopy.Copy()
        Dialog3.ToolTip1.Active = True
        Dialog3.ToolTip1.Show("Successfully Copied!", Me, DayLabel.Width / 2 + DayLabel.Location.X, 2 + DayLabel.Location.Y)
        Dialog3.ToolTipHider.Enabled = True
    End Sub

    Private Sub SecondsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SecondsToolStripMenuItem.Click
        Dialog3.TimeToCopy.Text = DateTime.Now.Second
        Dialog3.TimeToCopy.SelectAll()
        Dialog3.TimeToCopy.Copy()
        Dialog3.ToolTip1.Active = True
        Dialog3.ToolTip1.Show("Successfully Copied!", Me, DayLabel.Width / 2 + DayLabel.Location.X, 2 + DayLabel.Location.Y)
        Dialog3.ToolTipHider.Enabled = True
    End Sub

    Private Sub MilisecondsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MilisecondsToolStripMenuItem.Click
        Dialog3.TimeToCopy.Text = DateTime.Now.Millisecond
        Dialog3.TimeToCopy.SelectAll()
        Dialog3.TimeToCopy.Copy()
        Dialog3.ToolTip1.Active = True
        Dialog3.ToolTip1.Show("Successfully Copied!", Me, DayLabel.Width / 2 + DayLabel.Location.X, 2 + DayLabel.Location.Y)
        Dialog3.ToolTipHider.Enabled = True
    End Sub

    Private Sub DayToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DayToolStripMenuItem.Click
        Dialog3.DateToCopy.Text = DateTime.Now.Day
        Dialog3.DateToCopy.SelectAll()
        Dialog3.DateToCopy.Copy()
        Dialog3.ToolTip1.Active = True
        Dialog3.ToolTip1.Show("Successfully Copied!", Me, DayLabel.Width / 2 + DayLabel.Location.X, 2 + DayLabel.Location.Y)
        Dialog3.ToolTipHider.Enabled = True
    End Sub

    Private Sub MonthToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MonthToolStripMenuItem.Click
        Dialog3.DateToCopy.Text = DateTime.Now.Month
        Dialog3.DateToCopy.SelectAll()
        Dialog3.DateToCopy.Copy()
        Dialog3.ToolTip1.Active = True
        Dialog3.ToolTip1.Show("Successfully Copied!", Me, DayLabel.Width / 2 + DayLabel.Location.X, 2 + DayLabel.Location.Y)
        Dialog3.ToolTipHider.Enabled = True
    End Sub

    Private Sub YearToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles YearToolStripMenuItem.Click
        Dialog3.DateToCopy.Text = DateTime.Now.Year
        Dialog3.DateToCopy.SelectAll()
        Dialog3.DateToCopy.Copy()
        Dialog3.ToolTip1.Active = True
        Dialog3.ToolTip1.Show("Successfully Copied!", Me, DayLabel.Width / 2 + DayLabel.Location.X, 2 + DayLabel.Location.Y)
        Dialog3.ToolTipHider.Enabled = True
    End Sub

    Private Sub TimeAndDateSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TimeAndDateSettingsToolStripMenuItem.Click
        System.Diagnostics.Process.Start("C:\Windows\System32\timedate.cpl")
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

            Dim XPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Microsoft\Windows\WinX\Group"
            If IO.Directory.Exists(XPath & 1) Then
                For Each i As String In My.Computer.FileSystem.GetFiles(XPath & 1)
                    Dim IIO As New IO.FileInfo(i)
                    Dim item As New ToolStripMenuItem(IIO.Name)

                    Dim FirstSpaceIndex As Integer = item.Text.Trim.IndexOf(" ")

                    Dim firstChar As Char = IIO.Name.Chars(0)

                    If Char.IsNumber(firstChar) Then
                        item.Text = IO.Path.GetFileNameWithoutExtension(IIO.FullName).Trim.Substring(FirstSpaceIndex + 3).TrimStart()
                    Else
                        item.Text = IO.Path.GetFileNameWithoutExtension(IIO.FullName).Trim
                    End If

                    item.Tag = IIO.FullName

                    Try
                        Dim ico As Icon = Desktop.GetFileIcon(IIO.FullName, False)
                        item.Image = ico.ToBitmap
                    Catch ex As Exception

                    End Try

                    AddHandler item.Click, AddressOf ActionCenterToolStripMenuItem_Click

                    If Not IIO.Name = "desktop.ini" Then
                        ActionCM.Items.Add(item)
                    End If
                Next
                ActionCM.Items.Add(ToolStripSeparator4)
            End If

            If IO.Directory.Exists(XPath & 2) Then
                For Each i As String In My.Computer.FileSystem.GetFiles(XPath & 2)
                    Dim IIO As New IO.FileInfo(i)
                    Dim item As New ToolStripMenuItem(IIO.Name)

                    Dim FirstSpaceIndex As Integer = item.Text.Trim.IndexOf(" ")

                    Dim firstChar As Char = IIO.Name.Chars(0)

                    If Char.IsNumber(firstChar) Then
                        item.Text = IO.Path.GetFileNameWithoutExtension(IIO.FullName).Trim.Substring(FirstSpaceIndex + 3).TrimStart()
                    Else
                        item.Text = IO.Path.GetFileNameWithoutExtension(IIO.FullName).Trim
                    End If

                    item.Tag = IIO.FullName

                    Try
                        Dim ico As Icon = Desktop.GetFileIcon(IIO.FullName, False)
                        item.Image = ico.ToBitmap
                    Catch ex As Exception

                    End Try

                    AddHandler item.Click, AddressOf ActionCenterToolStripMenuItem_Click

                    If Not IIO.Name = "desktop.ini" Then
                        ActionCM.Items.Add(item)
                    End If
                Next
                ActionCM.Items.Add(ToolStripSeparator5)
            End If
            If IO.Directory.Exists(XPath & 3) Then
                For Each i As String In My.Computer.FileSystem.GetFiles(XPath & 3)
                    Dim IIO As New IO.FileInfo(i)
                    Dim item As New ToolStripMenuItem(IIO.Name)

                    Dim FirstSpaceIndex As Integer = item.Text.Trim.IndexOf(" ")

                    Dim firstChar As Char = IIO.Name.Chars(0)

                    If Char.IsNumber(firstChar) Then
                        item.Text = IO.Path.GetFileNameWithoutExtension(IIO.FullName).Trim.Substring(FirstSpaceIndex + 3).TrimStart()
                    Else
                        item.Text = IO.Path.GetFileNameWithoutExtension(IIO.FullName).Trim
                    End If

                    item.Tag = IIO.FullName

                    Try
                        Dim ico As Icon = Desktop.GetFileIcon(IIO.FullName, False)
                        item.Image = ico.ToBitmap
                    Catch ex As Exception

                    End Try

                    AddHandler item.Click, AddressOf ActionCenterToolStripMenuItem_Click

                    If Not IIO.Name = "desktop.ini" Then
                        ActionCM.Items.Add(item)
                    End If
                Next

                ActionCM.Items.Add(ToolStripSeparator14)
                ActionCM.Items.Add(ClipboardToolStripMenuItem)
                ActionCM.Items.Add(PerformanceToolStripMenuItem)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ToolStripButton3_MouseUp(sender As Object, e As MouseEventArgs) Handles ToolStripButton3.MouseUp
        'Shell("RunDll32.exe shell32.dll,Control_RunDLL ncpa.cpl")
        MCM.Show(MousePosition)
    End Sub

    Private Sub AppBar_BackColorChanged(sender As Object, e As EventArgs) Handles Me.BackColorChanged
        If Me.BackColor.R >= 150 OrElse Me.BackColor.G >= 130 AndAlso Me.BackColor.B >= 130 Then
            Me.ForeColor = Color.Black
            Button2.FlatAppearance.BorderColor = Color.White
        Else
            Me.ForeColor = Color.White
            Button2.FlatAppearance.BorderColor = Color.Black
        End If
    End Sub

    Private Sub KillToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KillToolStripMenuItem.Click
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

            If MsgBox("Are you sure do you want to kill this process? [" & pr.ProcessName & "]", MsgBoxStyle.YesNo, "Confirm Box") = MsgBoxResult.Yes Then
                Try
                    pr.Kill()

                    LoadApps()
                Catch ex As Exception
                    MsgBox("Unable to kill the process. " & ex.Message, MsgBoxStyle.Critical, "Error")
                End Try

            End If
        End If
    End Sub

    Private Sub DestroyProgramToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DestroyWindowToolStripMenuItem.Click
        If MsgBox("!!!WARNING!!!" & Environment.NewLine & """Destroy a Window"" means, that everything that the selected Window has, including data on your memory will be ""destroyed"" which I don't recommend. Well this warning is to prevent it happening one time. Do you really want to continue?", MsgBoxStyle.YesNo, "Warning") = MsgBoxResult.Yes Then
            If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
                Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
                SendMessage(windowHandle, WM_DESTROY, 0, 0)
            End If
        End If
    End Sub

    Private Sub OpenProgramsLocationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenProgramsLocationToolStripMenuItem.Click
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

            If IO.File.Exists(pr.MainModule.FileName) = True Then
                Dim fi As New IO.FileInfo(Process.GetProcessById(TPCM.Tag).MainModule.FileName)
                Process.Start("explorer.exe", fi.DirectoryName)
            End If
        End If
    End Sub

    Private Sub OpenProgramsWorkingDirectoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenProgramsWorkingDirectoryToolStripMenuItem.Click
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

            If IO.Directory.Exists(pr.StartInfo.WorkingDirectory) = True Then
                Dim fi As New IO.DirectoryInfo(Process.GetProcessById(TPCM.Tag).StartInfo.WorkingDirectory)
                Process.Start("explorer.exe", fi.FullName)
            End If
        End If
    End Sub

    Private Sub AlwaysOnTopToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AlwaysOnTopToolStripMenuItem.Click
        On Error Resume Next
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)
            SetWindowPos(windowHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        End If
    End Sub

    Private Sub AlwaysOnTopToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AlwaysOnTopToolStripMenuItem1.Click
        On Error Resume Next
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)
            SetWindowPos(windowHandle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        End If
    End Sub

    Private Sub HideProcessFromAppbarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HideProcessFromAppbarToolStripMenuItem.Click
        On Error Resume Next
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(TPCM.Tag, IntPtr)
            Dim pr As Process = GetProcessFromWindowHandle(windowHandle)

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\HiddenProcesses\", pr.ProcessName, "")

            LoadApps()
            MainToolTip.Show("The """ & Process.GetProcessById(TPCM.Tag).ProcessName & """ has been successfully set to hidden from Appbar!", Me, Me.Width - ToolStrip1.Width / 2, Me.Location.Y + ToolStrip1.Height + 150, 5000)
        End If
    End Sub

    Private Sub BlockProcessToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BlockProcessToolStripMenuItem.Click
        If MessageBox.Show("Are you sure do you want to block [" & Process.GetProcessById(TPCM.Tag).ProcessName & "] process?" & Environment.NewLine & "This means, when this shell will be running, then that process you'll set as ""Blocked"" will no longer be working. To allow the process again: Right click on the Appbar, select properties and remove the app from Blocklist.", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
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

    Private Sub SoundControlToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SoundControlToolStripMenuItem.Click
        On Error Resume Next
        Process.Start("C:\Windows\System32\SndVol.exe")
    End Sub

    Private Sub AboutShellToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutShellToolStripMenuItem.Click
        AboutDialog.ShowDialog()
    End Sub

    Private Sub ClipboardViewerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClipboardViewerToolStripMenuItem.Click
        ClipboardViewer.Show()
    End Sub
    Public PinnedAppsDir As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Internet Explorer\Quick Launch"
    Public Sub LoadPinnedApps()
        AppStrip.Items.Clear()

        'i.Substring(i.LastIndexOf("\") + 1)
        Try
            For Each i As String In Directory.GetFiles(PinnedAppsDir.Trim)
                If Not i.Substring(i.LastIndexOf("\") + 1) = "desktop.ini" Then

                    Dim ico As Icon = Desktop.GetFileIcon(PinnedAppsDir.Trim & "\" & i.Substring(i.LastIndexOf("\") + 1), False)
                    Dim FI As New FileInfo(PinnedAppsDir.Trim & "\" & i.Substring(i.LastIndexOf("\") + 1))

                    Dim Item As New ToolStripMenuItem
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
        If Not TPCM.Tag Is Nothing AndAlso TypeOf TPCM.Tag Is IntPtr Then
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

    Private Sub PinnedAppMouseUp(ByVal sender As Object, e As MouseEventArgs)
        Dim item As ToolStripMenuItem = CType(sender, ToolStripMenuItem)

        If File.Exists(item.Tag) Then
            Select Case e.Button
                Case MouseButtons.Left
                    Process.Start(item.Tag)

                Case MouseButtons.Right
                    Try
                        Desktop.ShowShellContextMenu(item.Tag, MousePosition)
                    Catch ex As Exception
                        PIcm.Tag = item.Tag
                        PIcm.Show(MousePosition)
                    End Try

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
        Process.Start("explorer.exe", fi.DirectoryName)
    End Sub

    Private Sub RemoveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemoveToolStripMenuItem.Click
        On Error Resume Next
        If IO.File.Exists(PIcm.Tag) Then
            Dim fi As New FileInfo(PIcm.Tag)
            If MessageBox.Show("Are you sure do you want remove """ & fi.Name & """ Item from Pinned Bar?", "Remove Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                My.Computer.FileSystem.DeleteFile(PIcm.Tag, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)

                LoadPinnedApps()
            End If
        End If
    End Sub

    Private Sub RemoveEntirelyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemoveEntirelyToolStripMenuItem.Click
        On Error Resume Next
        If IO.File.Exists(PIcm.Tag) Then
            Dim fi As New FileInfo(PIcm.Tag)
            If MessageBox.Show("Are you sure do you want remove """ & fi.Name & """ Item from Pinned Bar entirely?", "Remove Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                My.Computer.FileSystem.DeleteFile(PIcm.Tag, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)

                LoadPinnedApps()
            End If
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        On Error Resume Next
        Process.Start("explorer.exe", PinnedAppsDir)
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        Dim ofd As New OpenFileDialog

        ofd.Title = "Select a program to be Pinned into Pinned Bar..."
        ofd.Multiselect = False
        ofd.AutoUpgradeEnabled = False
        ofd.CheckFileExists = True
        ofd.SupportMultiDottedExtensions = True
        ofd.Filter = "Programs (*.exe;*.pif;*.bat;*.cmd)|*.exe;*.pif;*.bat;*.cmd|All files (*.*)|*.*"

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

    Private Sub AppBar_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged, Me.Resize
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "SpaceForEmergeTray", "0") = 0 Then
            If Not Me.Width = SystemInformation.PrimaryMonitorSize.Width Then
                Me.Width = SystemInformation.PrimaryMonitorSize.Width
            End If
        ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "SpaceForEmergeTray", "0") = 1 Then
            If SizeLocked = True Then
                Dim ValueW As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "CustomWidth", SystemInformation.PrimaryMonitorSize.Width / 2 + SystemInformation.PrimaryMonitorSize.Width / 4)
                If Not Me.Width = ValueW Then
                    Me.Width = ValueW
                End If
            End If
        End If
        If Not Me.WindowState = FormWindowState.Normal Then
            Me.WindowState = FormWindowState.Normal
        End If
    End Sub

    Private Sub AppBar_StyleChanged(sender As Object, e As EventArgs) Handles Me.StyleChanged
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub Button1_MouseHover(sender As Object, e As EventArgs) Handles Button1.MouseHover
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 1 Then
            If Startmenu.Visible = False Then
                Try
                    Button1.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Hover", ""))
                Catch ex As Exception
                    Button1.BackgroundImage = My.Resources.StartRight
                End Try
            Else
                Try
                    Button1.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Pressed", ""))
                Catch ex As Exception
                    Button1.BackgroundImage = My.Resources.StartLeft
                End Try
            End If
        ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 2 Then
            'ORB Code here
        End If
    End Sub

    Private Sub Button1_MouseLeave(sender As Object, e As EventArgs) Handles Button1.MouseLeave
        Hidden = True
        MouseAway = True

        Startmenu.FST = False

        Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - 2)
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 1 Then
            If Startmenu.Visible = False Then
                Try
                    Button1.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Normal", ""))
                Catch ex As Exception
                    Button1.BackgroundImage = My.Resources.StartRight
                End Try
            Else
                Try
                    Button1.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Pressed", ""))
                Catch ex As Exception
                    Button1.BackgroundImage = My.Resources.StartLeft
                End Try
            End If
        ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 2 Then
            'ORB Code here
        End If
    End Sub

    Private Sub Button1_MouseEnter(sender As Object, e As EventArgs) Handles Button1.MouseEnter
        Hidden = False
        MouseAway = False

        If Startmenu.Visible = False Then
            Button1.Focus()
        End If

        Startmenu.FST = True

        Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - 2)
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 1 Then
            If Startmenu.Visible = False Then
                Try
                    Button1.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Hover", ""))
                Catch ex As Exception
                    Button1.BackgroundImage = My.Resources.StartRight
                End Try
            Else
                Try
                    Button1.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Pressed", ""))
                Catch ex As Exception
                    Button1.BackgroundImage = My.Resources.StartLeft
                End Try
            End If
        ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 2 Then
            'ORB Code here
        End If
    End Sub

    Private Sub ProcessStrip_MouseUp(sender As Object, e As MouseEventArgs) Handles ProcessStrip.MouseUp
        SaveToolStripOrder()

        IsReordering = True
        CheckForProcessUpdates()
    End Sub

    Private Sub ReloadAppsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReloadAppsToolStripMenuItem.Click
        LoadApps()
        ClockTray.SyncAppearance()

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

    Private Sub Panel6_Paint(sender As Object, e As PaintEventArgs) Handles Panel6.Paint

    End Sub
    Public SizeLocked As Boolean = True
    Private lastMousePos As Point
    Private Sub Panel6_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel6.MouseMove

        If Panel6.Visible = True Then
            If e.Button = MouseButtons.Left Then
                Panel6.Capture = True

                Me.Width = MousePosition.X

                SizeLocked = False

            ElseIf e.Button = MouseButtons.None Then
                Panel6.Capture = False

                SizeLocked = True

                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "CustomWidth", Me.Width, Microsoft.Win32.RegistryValueKind.DWord)
                Try
                    Dim hWnd As IntPtr = FindWindow("EmergeDesktopApplet", vbNullString)
                    If hWnd <> IntPtr.Zero Then
                        Dim newX As Integer = Me.Width
                        Dim newY As Integer = Me.Location.Y
                        Dim newWidth As Integer = SystemInformation.PrimaryMonitorSize.Width - Me.Width - ClockTray.Width
                        Dim newHeight As Integer = Me.Height

                        MoveAndResizeWindow(hWnd, newX, newY, newWidth, newHeight)
                    Else
                    End If
                Catch ex As Exception
                End Try
            End If
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
    Public Function GetWindowState(ByVal hWnd As IntPtr) As String
        Dim wp As WINDOWPLACEMENT
        wp.length = CType(Marshal.SizeOf(wp), UInteger)

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
    End Function

    Private Const GWL_EXSTYLE As Integer = -20

    ' Extended Window Styles
    Private Const WS_EX_TOPMOST As UInteger = &H8

    Private ReadOnly HWND_TOPMOST As IntPtr = New IntPtr(-1)
    Private ReadOnly HWND_NOTOPMOST As IntPtr = New IntPtr(-2)

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

    Private Function KeyboardHookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        If nCode >= 0 AndAlso wParam = CType(WM_KEYUP, IntPtr) Then ' Změněno na WM_KEYUP
            Dim hookStruct As KBDLLHOOKSTRUCT = Marshal.PtrToStructure(lParam, GetType(KBDLLHOOKSTRUCT))

            'Win key pressed Events
            If hookStruct.vkCode = VK_WIN_LEFT <> 0 Then
                If UsedMultipleKey = True Then
                    UsedMultipleKey = False
                    WinKeyPressed = False
                Else
                    WinKeyPressed = False
                    RaiseEvent HotkeyPressed(Nothing, EventArgs.Empty)
                End If
            ElseIf hookStruct.vkCode = VK_WIN_RIGHT <> 0 Then
                If UsedMultipleKey = True Then
                    UsedMultipleKey = False
                    WinKeyPressed = False
                Else
                    WinKeyPressed = False
                    RaiseEvent HotkeyPressed(Nothing, EventArgs.Empty)
                End If
            End If

            'Ctrl key pressed Events
            If hookStruct.vkCode = VK_CTRL_LEFT <> 0 Then
                If UsedMultipleKey = True Then
                    UsedMultipleKey = False
                Else
                    CtrlKeyPressed = False
                End If
            ElseIf hookStruct.vkCode = VK_CTRL_RIGHT <> 0 Then
                If UsedMultipleKey = True Then
                    UsedMultipleKey = False
                Else
                    CtrlKeyPressed = False
                End If
            End If

            'Shift key pressed Events
            If hookStruct.vkCode = VK_SHIFT_LEFT <> 0 Then
                If UsedMultipleKey = True Then
                    UsedMultipleKey = False
                Else
                    ShiftKeyPressed = False
                End If
            ElseIf hookStruct.vkCode = VK_SHIFT_RIGHT <> 0 Then
                If UsedMultipleKey = True Then
                    UsedMultipleKey = False
                Else
                    ShiftKeyPressed = False
                End If
            End If

            'Alt key pressed Events
            If hookStruct.vkCode = VK_ALT_LEFT <> 0 Then
                If UsedMultipleKey = True Then
                    UsedMultipleKey = False
                Else
                    AltKeyPressed = False
                End If
            ElseIf hookStruct.vkCode = VK_ALT_RIGHT <> 0 Then
                If UsedMultipleKey = True Then
                    UsedMultipleKey = False
                Else
                    AltKeyPressed = False
                    AltTab.Close()
                End If
            End If
        ElseIf nCode >= 0 AndAlso wParam = CType(WM_KEYDOWN, IntPtr) Then
            Dim hookStruct As KBDLLHOOKSTRUCT = Marshal.PtrToStructure(lParam, GetType(KBDLLHOOKSTRUCT))

            If hookStruct.vkCode = VK_WIN_LEFT <> 0 Then
                'RaiseEvent HotkeyPressed(Nothing, EventArgs.Empty)
                WinKeyPressed = True
            ElseIf hookStruct.vkCode = VK_WIN_RIGHT <> 0 Then
                WinKeyPressed = True
            End If

            If hookStruct.vkCode = VK_CTRL_LEFT <> 0 Then
                CtrlKeyPressed = True
            ElseIf hookStruct.vkCode = VK_CTRL_RIGHT <> 0 Then
                CtrlKeyPressed = True
            End If

            If hookStruct.vkCode = VK_SHIFT_LEFT <> 0 Then
                ShiftKeyPressed = True
            ElseIf hookStruct.vkCode = VK_SHIFT_RIGHT <> 0 Then
                ShiftKeyPressed = True
            End If

            If hookStruct.vkCode = 164 <> 0 Then
                AltKeyPressed = True
            ElseIf hookStruct.vkCode = 165 <> 0 Then
                AltKeyPressed = True
            End If

            'Actual key events:
            If WinKeyPressed = True Then

                If hookStruct.vkCode = 37 <> 0 Then
                    UsedMultipleKey = True
                    Dim hWnd As IntPtr = GetForegroundWindow()

                    If hWnd <> IntPtr.Zero Then
                        Dim processId As UInteger = 0
                        GetWindowThreadProcessId(hWnd, processId)
                        Dim process As Process = Process.GetProcessById(CInt(processId))
                        If Not process.ProcessName = "KrrShell" Then
                            Dim currentState As String = GetWindowState(hWnd)

                            Select Case currentState
                                Case "Normal"
                                    Dim newX As Integer = 0
                                    Dim newY As Integer = 0
                                    Dim newWidth As Integer = My.Computer.Screen.WorkingArea.Width / 2
                                    Dim newHeight As Integer = My.Computer.Screen.WorkingArea.Height
                                    MoveAndResizeWindow(hWnd, newX, newY, newWidth, newHeight)
                                Case "Max"
                                    RestoreWindow(hWnd)
                                    Dim newX As Integer = 0
                                    Dim newY As Integer = 0
                                    Dim newWidth As Integer = My.Computer.Screen.WorkingArea.Width / 2
                                    Dim newHeight As Integer = My.Computer.Screen.WorkingArea.Height
                                    MoveAndResizeWindow(hWnd, newX, newY, newWidth, newHeight)
                                Case "Min"
                                    RestoreWindow(hWnd)
                                    Dim newX As Integer = 0
                                    Dim newY As Integer = 0
                                    Dim newWidth As Integer = My.Computer.Screen.WorkingArea.Width / 2
                                    Dim newHeight As Integer = My.Computer.Screen.WorkingArea.Height
                                    MoveAndResizeWindow(hWnd, newX, newY, newWidth, newHeight)
                                Case Else

                            End Select
                        Else

                        End If
                    End If
                    'WinKeyPressed = False
                ElseIf hookStruct.vkCode = 38 <> 0 Then
                    UsedMultipleKey = True
                    Dim hWnd As IntPtr = GetForegroundWindow()

                    Dim processId As UInteger = 0
                    GetWindowThreadProcessId(hWnd, processId)
                    Dim process As Process = Process.GetProcessById(CInt(processId))
                    If Not process.ProcessName = "KrrShell" Then

                        If hWnd <> IntPtr.Zero Then
                            Dim currentState As String = GetWindowState(hWnd)

                            Select Case currentState
                                Case "Normal"
                                    MaximizeWindow(hWnd)
                                Case "Max"
                                    RestoreWindow(hWnd)
                                Case "Min"
                                    RestoreWindow(hWnd)
                                Case Else

                            End Select
                        Else

                        End If
                    End If
                    'WinKeyPressed = False
                ElseIf hookStruct.vkCode = 39 <> 0 Then
                    UsedMultipleKey = True
                    Dim hWnd As IntPtr = GetForegroundWindow()

                    Dim processId As UInteger = 0
                    GetWindowThreadProcessId(hWnd, processId)
                    Dim process As Process = Process.GetProcessById(CInt(processId))
                    If Not process.ProcessName = "KrrShell" Then

                        If hWnd <> IntPtr.Zero Then
                            Dim currentState As String = GetWindowState(hWnd)

                            Select Case currentState
                                Case "Normal"
                                    Dim newX As Integer = My.Computer.Screen.WorkingArea.Width / 2
                                    Dim newY As Integer = 0
                                    Dim newWidth As Integer = My.Computer.Screen.WorkingArea.Width / 2
                                    Dim newHeight As Integer = My.Computer.Screen.WorkingArea.Height
                                    MoveAndResizeWindow(hWnd, newX, newY, newWidth, newHeight)
                                Case "Max"
                                    RestoreWindow(hWnd)
                                    Dim newX As Integer = My.Computer.Screen.WorkingArea.Width / 2
                                    Dim newY As Integer = 0
                                    Dim newWidth As Integer = My.Computer.Screen.WorkingArea.Width / 2
                                    Dim newHeight As Integer = My.Computer.Screen.WorkingArea.Height
                                    MoveAndResizeWindow(hWnd, newX, newY, newWidth, newHeight)
                                Case "Min"
                                    RestoreWindow(hWnd)
                                    Dim newX As Integer = My.Computer.Screen.WorkingArea.Width / 2
                                    Dim newY As Integer = 0
                                    Dim newWidth As Integer = My.Computer.Screen.WorkingArea.Width / 2
                                    Dim newHeight As Integer = My.Computer.Screen.WorkingArea.Height
                                    MoveAndResizeWindow(hWnd, newX, newY, newWidth, newHeight)
                                Case Else

                            End Select
                        Else

                        End If
                    End If
                    'WinKeyPressed = False
                ElseIf hookStruct.vkCode = 40 <> 0 Then
                    UsedMultipleKey = True
                    Dim hWnd As IntPtr = GetForegroundWindow()

                    Dim processId As UInteger = 0
                    GetWindowThreadProcessId(hWnd, processId)
                    Dim process As Process = Process.GetProcessById(CInt(processId))
                    If Not process.ProcessName = "KrrShell" Then

                        If hWnd <> IntPtr.Zero Then
                            Dim currentState As String = GetWindowState(hWnd)

                            Select Case currentState
                                Case "Normal"
                                    MinimizeWindow(hWnd)
                                Case "Max"
                                    RestoreWindow(hWnd)
                                Case "Min"
                                    RestoreWindow(hWnd)
                                Case Else

                            End Select
                        Else

                        End If
                        'WinKeyPressed = False
                    End If

                ElseIf hookStruct.vkCode = 9 <> 0 Then
                    'WIN + TAB

                    UsedMultipleKey = True
                    Dim hWnd As IntPtr = GetForegroundWindow()

                    Dim processId As UInteger = 0
                    GetWindowThreadProcessId(hWnd, processId)
                    Dim process As Process = Process.GetProcessById(CInt(processId))
                    AltTab.ActiveProcessID = processId
                    AltTab.Show()
                    AltTab.Select()
                    WinKeyPressed = False

                ElseIf hookStruct.vkCode = 65 <> 0 Then
                    'WIN + A (TimeDate/Action Center)

                    UsedMultipleKey = True
                    Dialog3.Show()
                    WinKeyPressed = False

                ElseIf hookStruct.vkCode = 66 <> 0 Then
                    'WIN + B (Take Focus on Notify Icons)

                    UsedMultipleKey = True
                    AppBar.ToolStrip1.Focus()
                    AppBar.ToolStrip1.Select()
                    WinKeyPressed = False

                ElseIf hookStruct.vkCode = 67 <> 0 OrElse hookStruct.vkCode = 86 <> 0 Then
                    'WIN + C (Copilot, but here Clipboard Viewer)
                    'MsgBox("Sorry but Microsoft Copilot isn't here!") 'Little Easter egg!

                    UsedMultipleKey = True
                    ClipboardViewer.Show()
                    WinKeyPressed = False

                ElseIf hookStruct.vkCode = 68 <> 0 Then
                    'WIN + D (Show Desktop)

                    UsedMultipleKey = True
                    Desktop.BringToFront()
                    WinKeyPressed = False

                ElseIf hookStruct.vkCode = 69 <> 0 Then
                    'WIN + E (Launch File Explorer)

                    UsedMultipleKey = True
                    Process.Start("C:\Windows\explorer.exe", "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}")
                    WinKeyPressed = False

                ElseIf hookStruct.vkCode = 73 <> 0 Then
                    'WIN + I (Launch Settings, but because I hate it, Control Panel will be!)

                    UsedMultipleKey = True
                    Process.Start("C:\Windows\System32\control.exe")
                    WinKeyPressed = False

                ElseIf hookStruct.vkCode = 76 <> 0 Then
                    'WIN + L (Lock Work Station)

                    UsedMultipleKey = True
                    Process.Start("RunDll32.exe", "user32.dll,LockWorkStation")
                    WinKeyPressed = False

                ElseIf hookStruct.vkCode = 77 <> 0 Then
                    'WIN + M (Minimize all Windows)

                    UsedMultipleKey = True
                    For Each i As ToolStripButton In AppBar.ProcessStrip.Items
                        If Not i.Tag Is Nothing AndAlso TypeOf i.Tag Is IntPtr Then
                            Dim windowHandle As IntPtr = CType(i.Tag, IntPtr)
                            MinimizeWindow(windowHandle)
                        End If
                    Next
                    WinKeyPressed = False

                ElseIf hookStruct.vkCode = 82 <> 0 Then
                    'WIN + R (Run Dialog)

                    UsedMultipleKey = True
                    RunDialog.Show()
                    WinKeyPressed = False

                ElseIf hookStruct.vkCode = 83 <> 0 Then
                    'WIN + S (Windows Search (in progress...))

                    UsedMultipleKey = True
                    InputBox("(PROTOTYPE) Type something here to search!", "Search...")
                    WinKeyPressed = False

                ElseIf hookStruct.vkCode = 84 <> 0 Then
                    'WIN + T (Focuses the Taskbar)

                    UsedMultipleKey = True
                    AppBar.ProcessStrip.Focus()
                    AppBar.ProcessStrip.Select()
                    WinKeyPressed = False

                ElseIf hookStruct.vkCode = 88 <> 0 Then
                    'WIN + X (Shows WIN + X menu)

                    UsedMultipleKey = True
                    AppBar.ActionCM.Show(AppBar, AppBar.Location)
                    WinKeyPressed = False

                ElseIf hookStruct.vkCode = 89 <> 0 Then
                    'WIN + Y (Set focused Window on Top or disables TopMost) 'Not in Windows!

                    UsedMultipleKey = True
                    Dim hWnd As IntPtr = GetForegroundWindow()

                    If hWnd <> IntPtr.Zero Then
                        ToggleAlwaysOnTop(hWnd)
                    Else
                        MessageBox.Show("No active window has been found.", "Info")
                    End If
                    WinKeyPressed = False

                End If
            End If

            If ShiftKeyPressed = True Then
                If hookStruct.vkCode = 121 <> 0 Then
                    UsedMultipleKey = True
                    Process.Start("C:\Windows\System32\cmd.exe")
                    WinKeyPressed = False
                End If
            End If

            'Failed Alt+Tab key, because it is a system keybind and cannot be turned off...

            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "Enabled", "0") = 1 Then
                If AltKeyPressed = True Then
                    If hookStruct.vkCode = 9 <> 0 Then
                        UsedMultipleKey = True
                        AltTab.Show()
                        WinKeyPressed = False
                    End If
                End If
            End If
        End If

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