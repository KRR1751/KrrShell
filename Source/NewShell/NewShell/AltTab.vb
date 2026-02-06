Imports NewShell.AppBar
Imports System.ComponentModel
Imports System.Runtime.InteropServices

Public Class AltTab
    Public ActiveProcessID As Integer = 0
    <Runtime.InteropServices.DllImport("dwmapi.dll")> Public Shared Function DwmExtendFrameIntoClientArea(ByVal hWnd As IntPtr, ByRef pMarinset As side) As Integer
    End Function

    <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)> Public Structure side
        Public Left As Integer
        Public Right As Integer
        Public Top As Integer
        Public Bottom As Integer
    End Structure

    Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

    Sub GetWindowSize(ByVal hWnd As Long, Width As Long, Height As Long)
        Dim rc As RECT
        GetWindowRect(hWnd, rc)
        Width = rc.right - rc.left
        Height = rc.bottom - rc.top
    End Sub

    ' Window Capturing
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

    <DllImport("user32.dll")>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function

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
    Private ActiveWindowHandle As IntPtr = IntPtr.Zero

    Private ReadOnly IconCache As New Dictionary(Of String, Integer)

    Private ReadOnly ProcessIconCache As New Dictionary(Of String, Integer)
    Private ReadOnly WindowIconCache As New Dictionary(Of IntPtr, Integer)
#End Region

    Private ActiveItem As ToolStripMenuItem = Nothing
    Private initialMousePos As Point
    Private Sub AltTab_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MenuStrip1.Items.Clear()

        initialMousePos = Cursor.Position

        Dim currentActive As IntPtr = AppBar.ActiveWindowHandle

        Me.Activate()
        Me.Focus()
        Me.KeyPreview = True

        DisposeIconCache()

        Dim applications As List(Of TaskbarWindowInfo) = GetTaskbarApplications()

        Dim sortedApps As New List(Of TaskbarWindowInfo)

        Dim activeWin = applications.FirstOrDefault(Function(w) w.Handle = currentActive)

        If activeWin IsNot Nothing Then
            sortedApps.Add(activeWin)
            sortedApps.AddRange(applications.Where(Function(w) w.Handle <> currentActive))
        Else
            sortedApps = applications
        End If

        For Each windowInfo In sortedApps
            Try
                Dim p As Process = Process.GetProcessById(windowInfo.PID)
                Dim item As New ToolStripMenuItem()

                item.Text = windowInfo.Title
                item.Tag = windowInfo.Handle
                item.AutoSize = False
                item.Size = New Size(64, 64)
                item.DisplayStyle = ToolStripItemDisplayStyle.Image
                item.ImageScaling = ToolStripItemImageScaling.None
                item.Image = GetProcessIcon(windowInfo.Handle, Nothing)
                item.AutoToolTip = True
                item.ToolTipText = windowInfo.Title

                AddHandler item.MouseEnter, Sub(senderItem As Object, ea As EventArgs)
                                                Dim currentPos As Point = Cursor.Position

                                                Dim deltaX As Integer = Math.Abs(currentPos.X - initialMousePos.X)
                                                Dim deltaY As Integer = Math.Abs(currentPos.Y - initialMousePos.Y)

                                                If deltaX > 5 OrElse deltaY > 5 Then
                                                    DirectCast(senderItem, ToolStripMenuItem).Select()
                                                    UpdatePreview()
                                                End If
                                            End Sub

                If Not p.Responding Then item.BackColor = Color.Red

                MenuStrip1.Items.Add(item)
            Catch : End Try
        Next

        If MenuStrip1.Items.Count > 1 Then
            MenuStrip1.Items(1).Select()
        ElseIf MenuStrip1.Items.Count > 0 Then
            MenuStrip1.Items(0).Select()
        End If
    End Sub

    Private Sub AltTab_LoadOld(sender As Object, e As EventArgs) 'Handles MyBase.Load
        MenuStrip1.Items.Clear()
        Me.Activate()
        Me.Focus()
        Me.KeyPreview = True

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

                    Dim item As New ToolStripMenuItem()
                    item.Text = windowInfo.Title
                    item.Tag = windowInfo.Handle
                    item.Image = windowImage

                    item.ToolTipText = windowInfo.Title

                    item.AutoSize = False
                    item.AutoToolTip = True

                    item.ImageScaling = ToolStripItemImageScaling.None

                    item.DisplayStyle = ToolStripItemDisplayStyle.Image
                    item.Size = New Size(64, 64)


                    If p.Responding = False Then
                        item.BackColor = Color.Red
                    End If

                    MenuStrip1.Items.Add(item)

                    If windowInfo.Handle = AppBar.ActiveWindowHandle Then
                        ActiveItem = item
                    End If
                Else

                    For Each windowInfo In group
                        Dim item As New ToolStripMenuItem()
                        item.Text = windowInfo.Title
                        item.Tag = windowInfo.Handle

                        If windowInfo.Handle = AppBar.ActiveWindowHandle Then
                            ActiveItem = item
                        End If

                        item.AutoSize = False
                        item.ImageScaling = ToolStripItemImageScaling.None

                        item.DisplayStyle = ToolStripItemDisplayStyle.Image
                        item.Size = New Size(64, 64)

                        item.Image = GetProcessIcon(windowInfo.Handle, processImage)

                        If p.Responding = False Then
                            item.BackColor = Color.Yellow
                        End If

                        MenuStrip1.Items.Add(item)
                    Next
                End If

            Catch ex As Exception
                Dim errorItem As New ToolStripLabel()
                errorItem.Text = $"Error/Ended process (PID: {group.Key})"
                MenuStrip1.Items.Add(errorItem)
            End Try
        Next

        MenuStrip1.Select()
        MenuStrip1.Focus()
        If ActiveItem IsNot Nothing Then ActiveItem.Select()
    End Sub

    Private Sub MenuStrip1_Click(sender As Object, e As EventArgs) Handles Panel1.Click, Panel2.Click, PictureBox1.Click, MenuStrip1.Click
        Me.Close()
    End Sub
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

    <DllImport("user32.dll")>
    Private Shared Function IsIconic(hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function AttachThreadInput(ByVal idAttach As UInteger, ByVal idAttachTo As UInteger, ByVal fAttach As Boolean) As Boolean
    End Function

    Private Sub ForceForegroundWindow(ByVal hWnd As IntPtr)
        Dim foregroundThread As UInteger = GetWindowThreadProcessId(GetForegroundWindow(), IntPtr.Zero)
        Dim targetThread As UInteger = GetWindowThreadProcessId(hWnd, IntPtr.Zero)

        If foregroundThread <> targetThread Then
            AttachThreadInput(foregroundThread, targetThread, True)
            SetForegroundWindow(hWnd)
            AttachThreadInput(foregroundThread, targetThread, False)
        Else
            SetForegroundWindow(hWnd)
        End If

        Dim currentState As String = GetWindowState(hWnd)

        Select Case currentState
            Case "Normal"
                'normal

            Case "Maximized"
                'max

            Case "Minimized"
                ShowWindow(hWnd, SHOW_WINDOW.SW_RESTORE)

            Case Else
                ShowWindow(hWnd, SHOW_WINDOW.SW_RESTORE)

        End Select

        AppBar.ActiveWindowHandle = hWnd
    End Sub

    Private Sub AltTab_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyUp, MenuStrip1.KeyUp
        If e.KeyCode = Keys.LWin OrElse e.KeyCode = Keys.RWin Then

            Dim selectedItem As ToolStripMenuItem = MenuStrip1.Items.OfType(Of ToolStripMenuItem)().FirstOrDefault(Function(i) i.Selected)

            If selectedItem IsNot Nothing AndAlso selectedItem.Tag IsNot Nothing AndAlso TypeOf selectedItem.Tag Is IntPtr Then
                Dim windowHandle As IntPtr = CType(selectedItem.Tag, IntPtr)

                Try
                    ForceForegroundWindow(windowHandle)
                    AppBar.ActiveWindowHandle = windowHandle
                Catch ex As Exception
                    Debug.WriteLine("Activation failed: " & ex.Message)
                End Try
            End If

            Me.Close()

        ElseIf e.KeyCode = Keys.Tab Then
            UpdatePreview()
        End If
    End Sub

    Private Sub UpdatePreview()
        Dim selectedItem As ToolStripMenuItem = MenuStrip1.Items.OfType(Of ToolStripMenuItem)().FirstOrDefault(Function(i) i.Selected)

        If selectedItem IsNot Nothing AndAlso selectedItem.Tag IsNot Nothing Then
            Dim windowHandle As IntPtr = CType(selectedItem.Tag, IntPtr)

            Label1.Text = selectedItem.Text

            Try
                Dim previewImage As Image = RenderWindow(windowHandle, False)

                If previewImage IsNot Nothing Then
                    PictureBox1.Image = previewImage

                    Dim targetHeight As Integer = 165

                    If previewImage.Width > targetHeight Then
                        Dim aspectRatio As Double = CDbl(previewImage.Width) / CDbl(previewImage.Height)
                        Dim targetWidth As Integer = CInt(targetHeight * aspectRatio)

                        PictureBox1.Size = New Size(targetWidth + 10, targetHeight)
                        PictureBox1.Location = New Point((Panel1.Width - PictureBox1.Width) \ 2, 5)
                    Else
                        PictureBox1.Size = New Size(178, previewImage.Height)
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine("Preview render failed: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked
        If Not e.ClickedItem.Tag Is Nothing AndAlso TypeOf e.ClickedItem.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(e.ClickedItem.Tag, IntPtr)

            Try
                Dim currentState As String = GetWindowState(windowHandle)

                Select Case currentState
                    Case "Normal"
                        SetForegroundWindow(windowHandle)
                        AppBar.ActiveWindowHandle = windowHandle

                    Case "Maximized"
                        SetForegroundWindow(windowHandle)
                        AppBar.ActiveWindowHandle = windowHandle

                    Case "Minimized"
                        SetForegroundWindow(windowHandle)
                        AppBar.ActiveWindowHandle = windowHandle
                        ShowWindow(windowHandle, SHOW_WINDOW.SW_RESTORE)

                    Case Else
                        SetForegroundWindow(windowHandle)
                        AppBar.ActiveWindowHandle = windowHandle
                        ShowWindow(windowHandle, SHOW_WINDOW.SW_RESTORE)
                End Select
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Cannot focus a process")
            End Try
        End If
    End Sub

    Private Sub AltTab_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate, Me.LostFocus
        Me.Activate()
        MenuStrip1.Select()
    End Sub

    Public Sub LoadAppsOld()
        If ActiveProcessID = String.Empty Then
            For Each pr As Process In Process.GetProcesses
                If Not pr.MainWindowTitle = "" Then
                    Try
                        Dim ico As Icon = Icon.ExtractAssociatedIcon(pr.MainModule.FileName)
                        Dim item As New ToolStripMenuItem
                        With item
                            .AutoSize = False
                            .DisplayStyle = ToolStripItemDisplayStyle.Image
                            .Text = pr.MainWindowTitle
                            .Tag = pr.Id
                            .Image = ico.ToBitmap
                            .ImageScaling = ToolStripItemImageScaling.None
                            .Size = New Size(64, 64)
                        End With
                        MenuStrip1.Items.Add(item)
                    Catch ex As Exception
                        Dim item As New ToolStripMenuItem
                        With item
                            .AutoSize = False
                            .DisplayStyle = ToolStripItemDisplayStyle.Image
                            .Text = pr.MainWindowTitle
                            .Tag = pr.Id
                            .Image = My.Resources.ProgramMedium
                            .ImageScaling = ToolStripItemImageScaling.None
                            .Size = New Size(64, 64)
                        End With
                        MenuStrip1.Items.Add(item)
                    End Try
                End If
            Next
        Else
            Try
                Dim ActiveProcess As Process = Process.GetProcessById(ActiveProcessID)
                Try
                    Dim ico As Icon = Icon.ExtractAssociatedIcon(ActiveProcess.MainModule.FileName)
                    Dim item As New ToolStripMenuItem
                    With item
                        .AutoSize = False
                        .DisplayStyle = ToolStripItemDisplayStyle.Image
                        .Text = ActiveProcess.MainWindowTitle
                        .Tag = ActiveProcess.Id
                        .Image = ico.ToBitmap
                        .ImageScaling = ToolStripItemImageScaling.None
                        .Size = New Size(64, 64)
                    End With
                    MenuStrip1.Items.Add(item)
                Catch ex As Exception
                    Dim item As New ToolStripMenuItem
                    With item
                        .AutoSize = False
                        .DisplayStyle = ToolStripItemDisplayStyle.Image
                        .Text = ActiveProcess.MainWindowTitle
                        .Tag = ActiveProcess.Id
                        .Image = My.Resources.ProgramMedium
                        .ImageScaling = ToolStripItemImageScaling.None
                        .Size = New Size(64, 64)
                    End With
                    MenuStrip1.Items.Add(item)
                End Try
                For Each pr As Process In Process.GetProcesses
                    If Not pr.MainWindowTitle = "" Then
                        If Not ActiveProcessID = pr.Id Then
                            Try
                                Dim ico As Icon = Icon.ExtractAssociatedIcon(pr.MainModule.FileName)
                                Dim item As New ToolStripMenuItem
                                With item
                                    .AutoSize = False
                                    .DisplayStyle = ToolStripItemDisplayStyle.Image
                                    .Text = pr.MainWindowTitle
                                    .Tag = pr.Id
                                    .Image = ico.ToBitmap
                                    .ImageScaling = ToolStripItemImageScaling.None
                                    .Size = New Size(64, 64)
                                End With
                                MenuStrip1.Items.Add(item)
                            Catch ex As Exception
                                Dim item As New ToolStripMenuItem
                                With item
                                    .AutoSize = False
                                    .DisplayStyle = ToolStripItemDisplayStyle.Image
                                    .Text = pr.MainWindowTitle
                                    .Tag = pr.Id
                                    .Image = My.Resources.ProgramMedium
                                    .ImageScaling = ToolStripItemImageScaling.None
                                    .Size = New Size(64, 64)
                                End With
                                MenuStrip1.Items.Add(item)
                            End Try
                        End If
                    End If
                Next
            Catch ex As Exception
                For Each pr As Process In Process.GetProcesses
                    If Not pr.MainWindowTitle = "" Then
                        Try
                            Dim ico As Icon = Icon.ExtractAssociatedIcon(pr.MainModule.FileName)
                            Dim item As New ToolStripMenuItem
                            With item
                                .AutoSize = False
                                .DisplayStyle = ToolStripItemDisplayStyle.Image
                                .Text = pr.MainWindowTitle
                                .Tag = pr.Id
                                .Image = ico.ToBitmap
                                .ImageScaling = ToolStripItemImageScaling.None
                                .Size = New Size(64, 64)

                            End With
                            MenuStrip1.Items.Add(item)
                        Catch ex2 As Exception
                            Dim item As New ToolStripMenuItem
                            With item
                                .AutoSize = False
                                .DisplayStyle = ToolStripItemDisplayStyle.Image
                                .Text = pr.MainWindowTitle
                                .Tag = pr.Id
                                .Image = My.Resources.ProgramMedium
                                .ImageScaling = ToolStripItemImageScaling.None
                                .Size = New Size(64, 64)
                            End With
                            MenuStrip1.Items.Add(item)
                        End Try
                    End If
                Next
            End Try
        End If
    End Sub

    Private Sub AltTab_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        ActiveProcessID = 0
        MenuStrip1.Items.Clear()
    End Sub
End Class