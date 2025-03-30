Imports System.IO
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Public Class AppBar
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
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            'Me.BackColor = Color.Cyan
            Dim side As side = side
            side.Left = -1
            side.Right = -1
            side.Top = -1
            side.Bottom = -1
            Dim result As Integer = DwmExtendFrameIntoClientArea(Me.Handle, side)
        Catch ex As Exception

        End Try

        fBarRegistered = False
        RegisterBar()
        Me.BringToFront()
        Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - Me.Height)
        Me.Width = SystemInformation.PrimaryMonitorSize.Width
        Me.BackColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "BackColor", "#B77358"))

        'Process.Start("C:\Windows\NewShell\TrayIcons\emergeTray.exe")

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

        LockAppbarToolStripMenuItem.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Locked", Nothing)
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
        Button1.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "StartButton", True)
        StartButtonToolStripMenuItem.Checked = Button1.Visible
        Try
            Button1.Width = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer1", "50")
        Catch ex As Exception
            Button1.Width = 50
        End Try

        ' AppStrip
        Panel1.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "PinnedBar", True)
        PinnedBarToolStripMenuItem.Checked = Panel1.Visible
        Try
            Panel1.Width = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer2", "107")
        Catch ex As Exception
            Panel1.Width = 107
        End Try

        AppStrip.Items.Clear()
        For Each i As String In My.Computer.FileSystem.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Internet Explorer\Quick Launch")
            If Not i.Substring(i.LastIndexOf("\") + 1) = "desktop.ini" Then
                Dim Item As New ToolStripMenuItem
                Dim ico As Icon = Icon.ExtractAssociatedIcon(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Internet Explorer\Quick Launch\" & i.Substring(i.LastIndexOf("\") + 1))
                With Item
                    Item.DisplayStyle = ToolStripItemDisplayStyle.Image
                    Item.Text = i.Substring(i.LastIndexOf("\") + 1)
                    If ico IsNot Nothing Then
                        Item.Image = ico.ToBitmap
                    End If
                End With
                AppStrip.Items.Add(i.Substring(i.LastIndexOf("\") + 1), ico.ToBitmap)
            End If
        Next

        BlockingProcesses.Enabled = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "Enabled", "0")
        Try
            BlockingProcesses.Interval = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\BlockedProcesses", "Ticks", "1")
        Catch ex As Exception
            BlockingProcesses.Interval = 1
        End Try

        ' ProcessStrip
        Panel4.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "MenuBar", True)
        MenuBarToolStripMenuItem.Checked = Panel4.Visible
        Button2.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "ShowDesktopButton", True)
        Try
            Panel4.Width = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Layer3", "164")
        Catch ex As Exception
            Panel4.Width = 164
        End Try

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
                        AddHandler item.MouseUp, AddressOf ProcessStripItemClick
                        AddHandler item.MouseHover, AddressOf ProcessStripItemMouseOver
                        AddHandler item.MouseLeave, AddressOf ProcessStripItemMouseLeave
                    End With
                    'For Each i As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\HiddenProcesses").GetValueNames
                    'If Not item.ToolTipText = i Then
                    ProcessStrip.Items.Add(item)
                    'End If
                    'Next
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
                        AddHandler item.MouseUp, AddressOf ProcessStripItemClick
                        AddHandler item.MouseHover, AddressOf ProcessStripItemMouseOver
                        AddHandler item.MouseLeave, AddressOf ProcessStripItemMouseLeave
                    End With
                    For Each i As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\HiddenProcesses").GetValueNames
                        If Not item.ToolTipText = i Then
                            ProcessStrip.Items.Add(item)
                        End If
                    Next
                End Try
            End If
        Next

        Me.TopMost = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "OnTop", True)
        LockAppbarToolStripMenuItem.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "Locked", False)
        AutoHideToolStripMenuItem.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar", "AutoHide", 0)

        ' Desktop

        Desktop.Show()
        Desktop.SendToBack()

        ' Alarms

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
        Dim rvA = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\", "DisableAlarms", Nothing)
        If rvA = "0" Then
            DisableAlarmOption.Checked = False
            AlarmController.Enabled = True
        Else
            DisableAlarmOption.Checked = True
            AlarmController.Enabled = False
        End If
    End Sub

    Private Sub ToolStrip1_ItemAdded(sender As Object, e As ToolStripItemEventArgs) Handles AppStrip.ItemAdded
        e.Item.DisplayStyle = ToolStripItemDisplayStyle.Image
    End Sub

    Private Sub ToolStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles AppStrip.ItemClicked
        Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Internet Explorer\Quick Launch\" & e.ClickedItem.Text)
    End Sub
    Dim LCP As Integer = 0
    Dim LCP64 As Integer = 0
    Private Sub ProcessTimer_Tick(sender As Object, e As EventArgs) Handles ProcessTimer.Tick
        If Not Process.GetProcesses.Length = LCP OrElse Not Process.GetProcesses.LongLength = LCP64 Then
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
                            AddHandler item.MouseUp, AddressOf ProcessStripItemClick
                            AddHandler item.MouseHover, AddressOf ProcessStripItemMouseOver
                            AddHandler item.MouseLeave, AddressOf ProcessStripItemMouseLeave
                        End With
                        'For Each i As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\HiddenProcesses").GetValueNames
                        'If Not item.ToolTipText = i Then
                        ProcessStrip.Items.Add(item)
                        ' End If
                        'Next
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
                            AddHandler item.MouseUp, AddressOf ProcessStripItemClick
                            AddHandler item.MouseHover, AddressOf ProcessStripItemMouseOver
                            AddHandler item.MouseLeave, AddressOf ProcessStripItemMouseLeave
                        End With
                        'For Each i As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\HiddenProcesses").GetValueNames
                        'If Not item.ToolTipText = i Then
                        ProcessStrip.Items.Add(item)
                        'End If
                        'Next
                    End Try
                End If
            Next
        End If
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
    Private Sub ProcessStripItemClick(sender As Object, e As MouseEventArgs)
        If e.Button = Windows.Forms.MouseButtons.Left Then
            If CType(sender, ToolStripMenuItem).Checked = True Then
                Dim App As Process = Process.GetProcessById(CType(sender, ToolStripMenuItem).Tag)
                Dim foregroundHandle As IntPtr = GetForegroundWindow()
                Dim foregroundProcessId As Integer = 0
                GetWindowThreadProcessId(foregroundHandle, foregroundProcessId)
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
        ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
            TPCM.Tag = CType(sender, ToolStripMenuItem).Tag
            StartSameProcessToolStripMenuItem.ToolTipText = Process.GetProcessById(CType(sender, ToolStripMenuItem).Tag).MainModule.FileName
            TPCM.Show(Control.MousePosition)
        End If
    End Sub
    Public MouseAway As Boolean = False
    Private Sub ProcessStripItemMouseOver(sender As Object, e As EventArgs)
        WAT.Show()
        Try
            Dim pr As Process = Process.GetProcessById(CType(sender, ToolStripMenuItem).Tag)
            MouseAway = False
            WAT.PictureBox1.Image = RenderWindow(pr.MainWindowHandle, True)
        Catch ex As Exception
            WAT.PictureBox1.Image = Image.FromFile("C:\Windows\Resources\Themes\Nothing.png")
        End Try
        WAT.Label1.Text = CType(sender, ToolStripMenuItem).Text
        WAT.Location = New Point(Control.MousePosition.X - WAT.Width / 2, SystemInformation.WorkingArea.Height - WAT.Height)
    End Sub

    Private Sub ProcessStripItemMouseLeave(sender As Object, e As EventArgs)
        WAT.Close()
    End Sub

    Private Sub ProcessTimerDelay_Tick(sender As Object, e As EventArgs) Handles ProcessTimerDelay.Tick
        LCP = Process.GetProcesses.Length
        LCP64 = Process.GetProcesses.LongLength
    End Sub

    Private Sub ProcessStrip_ItemAdded(sender As Object, e As ToolStripItemEventArgs) Handles ProcessStrip.ItemAdded
        For Each i As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\HiddenProcesses\").GetValueNames
            If e.Item.ToolTipText = i Then
                e.Item.Visible = False
            End If
        Next
    End Sub

    Private Sub DefaultToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DefaultToolStripMenuItem.Click
        Dim App As Process = Process.GetProcessById(TPCM.Tag)
        If App.Id > 0 Then
            AppActivate(App.Id)
            ShowWindow(App.MainWindowHandle, SHOW_WINDOW.SW_SHOWDEFAULT)
        End If
    End Sub

    Private Sub MaximalizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MaximalizeToolStripMenuItem.Click
        Dim App As Process = Process.GetProcessById(TPCM.Tag)
        If App.Id > 0 Then
            AppActivate(App.Id)
            ShowWindow(App.MainWindowHandle, SHOW_WINDOW.SW_MAXIMIZE)
        End If
    End Sub

    Private Sub NormalizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NormalizeToolStripMenuItem.Click
        Dim App As Process = Process.GetProcessById(TPCM.Tag)
        If App.Id > 0 Then
            AppActivate(App.Id)
            ShowWindow(App.MainWindowHandle, SHOW_WINDOW.SW_NORMAL)
        End If
    End Sub

    Private Sub MinimalizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MinimalizeToolStripMenuItem.Click
        Dim App As Process = Process.GetProcessById(TPCM.Tag)
        If App.Id > 0 Then
            AppActivate(App.Id)
            ShowWindow(App.MainWindowHandle, SHOW_WINDOW.SW_MINIMIZE)
        End If
    End Sub

    Private Sub ForceMinimalizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ForceMinimalizeToolStripMenuItem.Click
        Dim App As Process = Process.GetProcessById(TPCM.Tag)
        If App.Id > 0 Then
            AppActivate(App.Id)
            ShowWindow(App.MainWindowHandle, SHOW_WINDOW.SW_FORCEMINIMIZE)
        End If
    End Sub

    Private Sub HIDEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HIDEToolStripMenuItem.Click
        If MsgBox("Are you sure do you want to hide this process? [" & Process.GetProcessById(TPCM.Tag).ProcessName & "]", MsgBoxStyle.YesNo, "Confirm Box") = MsgBoxResult.Yes Then
            Dim App As Process = Process.GetProcessById(TPCM.Tag)
            If App.Id > 0 Then
                AppActivate(App.Id)
                ShowWindow(App.MainWindowHandle, SHOW_WINDOW.SW_HIDE)
            End If
        End If
    End Sub

    Private Sub SwitchToToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SwitchToToolStripMenuItem.Click
        Dim App As Process = Process.GetProcessById(TPCM.Tag)
        If App.Id > 0 Then
            AppActivate(App.Id)
            ShowWindow(App.MainWindowHandle, SHOW_WINDOW.SW_SHOW)
        End If
    End Sub

    Private Sub SwitchButNotActiveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SwitchButNotActiveToolStripMenuItem.Click
        Dim App As Process = Process.GetProcessById(TPCM.Tag)
        If App.Id > 0 Then
            ShowWindow(App.MainWindowHandle, SHOW_WINDOW.SW_SHOWNOACTIVATE)
        End If
    End Sub

    Private Sub KillToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KillToolStripMenuItem.Click
        If MsgBox("Are you sure do you want to kill this process? [" & Process.GetProcessById(TPCM.Tag).ProcessName & "]", MsgBoxStyle.YesNo, "Confirm Box") = MsgBoxResult.Yes Then
            Try
                Process.GetProcessById(TPCM.Tag).Kill()
            Catch ex As Exception
                MsgBox("Unable to kill the process. " & ex.Message, MsgBoxStyle.Critical, "Error")
            End Try
        End If
    End Sub

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        Process.GetProcessById(TPCM.Tag).Close()
    End Sub

    Private Sub StartSameProcessToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartSameProcessToolStripMenuItem.Click
        Process.Start(Process.GetProcessById(TPCM.Tag).MainModule.FileName)
    End Sub

    Private Sub OpenProcessPathToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenProcessPathToolStripMenuItem.Click
        If IO.File.Exists(Process.GetProcessById(TPCM.Tag).MainModule.FileName) = True Then
            Dim fi As New IO.FileInfo(Process.GetProcessById(TPCM.Tag).MainModule.FileName)
            Process.Start(fi.DirectoryName)
        End If
    End Sub

    Private Sub TPCM_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TPCM.Opening
        For Each pr As Process In Process.GetProcesses
            If TPCM.Tag = pr.Id Then
                If pr.PriorityClass = ProcessPriorityClass.Idle Then
                    PRprcmb6.Checked = True
                    PRprcmb5.Checked = False
                    PRprcmb4.Checked = False
                    PRprcmb3.Checked = False
                    PRprcmb2.Checked = False
                    PRprcmb1.Checked = False
                    SetPriorityLevelToolStripMenuItem.Enabled = True
                ElseIf pr.PriorityClass = ProcessPriorityClass.BelowNormal Then
                    PRprcmb6.Checked = False
                    PRprcmb5.Checked = True
                    PRprcmb4.Checked = False
                    PRprcmb3.Checked = False
                    PRprcmb2.Checked = False
                    PRprcmb1.Checked = False
                    SetPriorityLevelToolStripMenuItem.Enabled = True
                ElseIf pr.PriorityClass = ProcessPriorityClass.Normal Then
                    PRprcmb6.Checked = False
                    PRprcmb5.Checked = False
                    PRprcmb4.Checked = True
                    PRprcmb3.Checked = False
                    PRprcmb2.Checked = False
                    PRprcmb1.Checked = False
                    SetPriorityLevelToolStripMenuItem.Enabled = True
                ElseIf pr.PriorityClass = ProcessPriorityClass.AboveNormal Then
                    PRprcmb6.Checked = False
                    PRprcmb5.Checked = False
                    PRprcmb4.Checked = False
                    PRprcmb3.Checked = True
                    PRprcmb2.Checked = False
                    PRprcmb1.Checked = False
                    SetPriorityLevelToolStripMenuItem.Enabled = True
                ElseIf pr.PriorityClass = ProcessPriorityClass.High Then
                    PRprcmb6.Checked = False
                    PRprcmb5.Checked = False
                    PRprcmb4.Checked = False
                    PRprcmb3.Checked = False
                    PRprcmb2.Checked = True
                    PRprcmb1.Checked = False
                    SetPriorityLevelToolStripMenuItem.Enabled = True
                ElseIf pr.PriorityClass = ProcessPriorityClass.RealTime Then
                    PRprcmb6.Checked = False
                    PRprcmb5.Checked = False
                    PRprcmb4.Checked = False
                    PRprcmb3.Checked = False
                    PRprcmb2.Checked = False
                    PRprcmb1.Checked = True
                    SetPriorityLevelToolStripMenuItem.Enabled = True
                Else
                    SetPriorityLevelToolStripMenuItem.Enabled = False
                End If
            End If
        Next
    End Sub

    Private Sub PRprcmb1_Click(sender As Object, e As EventArgs) Handles PRprcmb1.Click
        For Each pr As Process In Process.GetProcesses
            If pr.Id = TPCM.Tag Then
                pr.PriorityClass = ProcessPriorityClass.RealTime
                PRprcmb6.Checked = False
                PRprcmb5.Checked = False
                PRprcmb4.Checked = False
                PRprcmb3.Checked = False
                PRprcmb2.Checked = False
                PRprcmb1.Checked = True
            End If
        Next
    End Sub

    Private Sub PRprcmb2_Click(sender As Object, e As EventArgs) Handles PRprcmb2.Click
        For Each pr As Process In Process.GetProcesses
            If pr.Id = TPCM.Tag Then
                pr.PriorityClass = ProcessPriorityClass.High
                PRprcmb6.Checked = False
                PRprcmb5.Checked = False
                PRprcmb4.Checked = False
                PRprcmb3.Checked = False
                PRprcmb2.Checked = True
                PRprcmb1.Checked = False
            End If
        Next
    End Sub

    Private Sub PRprcmb3_Click(sender As Object, e As EventArgs) Handles PRprcmb3.Click
        For Each pr As Process In Process.GetProcesses
            If pr.Id = TPCM.Tag Then
                pr.PriorityClass = ProcessPriorityClass.AboveNormal
                PRprcmb6.Checked = False
                PRprcmb5.Checked = False
                PRprcmb4.Checked = False
                PRprcmb3.Checked = True
                PRprcmb2.Checked = False
                PRprcmb1.Checked = False
            End If
        Next
    End Sub

    Private Sub PRprcmb4_Click(sender As Object, e As EventArgs) Handles PRprcmb4.Click
        For Each pr As Process In Process.GetProcesses
            If pr.Id = TPCM.Tag Then
                pr.PriorityClass = ProcessPriorityClass.Normal
                PRprcmb6.Checked = False
                PRprcmb5.Checked = False
                PRprcmb4.Checked = True
                PRprcmb3.Checked = False
                PRprcmb2.Checked = False
                PRprcmb1.Checked = False
            End If
        Next
    End Sub

    Private Sub PRprcmb5_Click(sender As Object, e As EventArgs) Handles PRprcmb5.Click
        For Each pr As Process In Process.GetProcesses
            If pr.Id = TPCM.Tag Then
                pr.PriorityClass = ProcessPriorityClass.BelowNormal
                PRprcmb6.Checked = False
                PRprcmb5.Checked = True
                PRprcmb4.Checked = False
                PRprcmb3.Checked = False
                PRprcmb2.Checked = False
                PRprcmb1.Checked = False
            End If
        Next
    End Sub

    Private Sub PRprcmb6_Click(sender As Object, e As EventArgs) Handles PRprcmb6.Click
        For Each pr As Process In Process.GetProcesses
            If pr.Id = TPCM.Tag Then
                pr.PriorityClass = ProcessPriorityClass.Idle
                PRprcmb6.Checked = True
                PRprcmb5.Checked = False
                PRprcmb4.Checked = False
                PRprcmb3.Checked = False
                PRprcmb2.Checked = False
                PRprcmb1.Checked = False
            End If
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Startmenu.Visible = True Then
            Startmenu.Visible = False
            Button1.BackgroundImage = My.Resources._1
        Else
            Startmenu.Visible = True
            Button1.BackgroundImage = My.Resources.right
        End If
    End Sub

    Private Sub ToolStripButton1_MouseUp(sender As Object, e As MouseEventArgs) Handles ToolStripButton1.MouseUp
        If e.Button = MouseButtons.Left Then
            VolumeControl.Show(Me)
        End If
    End Sub

    Private Sub AppBar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown, ProcessStrip.KeyDown, AppStrip.KeyDown
        If e.KeyCode = Keys.LWin OrElse e.KeyCode = Keys.RWin Then
            If Startmenu.Visible = True Then
                Startmenu.Visible = False
                Button1.BackgroundImage = My.Resources._1
            Else
                Startmenu.Visible = True
                Button1.BackgroundImage = My.Resources.right
            End If
        End If
    End Sub

    Private Sub RunDialogToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RunDialogToolStripMenuItem.Click
        RunDialog.Show()
        RunDialog.Activate()
    End Sub

    Private Sub Panel4_Resize(sender As Object, e As EventArgs) Handles Panel4.Resize
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
    End Sub
    Public CanClose As Boolean = False
    Private Sub AppBar_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If CanClose = False Then
            e.Cancel = True
            SA.ShowDialog()
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
        Else
            Splitter1.Visible = True
            Splitter2.Visible = True
            Splitter3.Visible = True
        End If
    End Sub

    Private Sub TaskManagerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TaskManagerToolStripMenuItem.Click
        Process.Start("taskmgr.exe", "/d")
    End Sub

    Private Sub ReloadAppsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReloadAppsToolStripMenuItem.Click
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
                        AddHandler item.MouseUp, AddressOf ProcessStripItemClick
                        AddHandler item.MouseHover, AddressOf ProcessStripItemMouseOver
                        AddHandler item.MouseLeave, AddressOf ProcessStripItemMouseLeave
                    End With
                    'For Each i As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\HiddenProcesses").GetValueNames
                    'If Not item.ToolTipText = i Then
                    ProcessStrip.Items.Add(item)
                    ' End If
                    'Next
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
                        AddHandler item.MouseUp, AddressOf ProcessStripItemClick
                        AddHandler item.MouseHover, AddressOf ProcessStripItemMouseOver
                        AddHandler item.MouseLeave, AddressOf ProcessStripItemMouseLeave
                    End With
                    'For Each i As String In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\Appbar\HiddenProcesses").GetValueNames
                    'If Not item.ToolTipText = i Then
                    ProcessStrip.Items.Add(item)
                    ' End If
                    'Next
                End Try
            End If
        Next

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

    Private Sub ToolStrip1_ItemClicked_1(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub

    Private Sub Controller_Tick(sender As Object, e As EventArgs) Handles Controller.Tick
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


    Private Sub ProcessStrip_MouseLeave(sender As Object, e As EventArgs) Handles ProcessStrip.MouseLeave, Button1.MouseLeave, AppStrip.MouseLeave, TimeLabel.MouseLeave, DateLabel.MouseLeave, DayLabel.MouseLeave
        Hidden = True
        Me.Location = New Point(0, SystemInformation.PrimaryMonitorSize.Height - 2)
    End Sub

    Private Sub ToolStrip1_MouseEnter(sender As Object, e As EventArgs) Handles ToolStrip1.MouseEnter, ProcessStrip.MouseEnter, Button1.MouseEnter, AppStrip.MouseEnter, TimeLabel.MouseEnter, DateLabel.MouseEnter, DayLabel.MouseEnter
        Hidden = False
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
        On Error Resume Next
        ActionCM.Items.Clear()
        Dim XPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Microsoft\Windows\WinX\Group"
        If IO.Directory.Exists(XPath & 1) Then
            For Each i As String In My.Computer.FileSystem.GetFiles(XPath & 1)
                Dim IIO As New IO.FileInfo(i)
                Dim item As New ToolStripMenuItem(IIO.Name)
                With item
                    .Text = IO.Path.GetFileNameWithoutExtension(IIO.FullName)
                    .Tag = IIO.FullName
                    AddHandler item.Click, AddressOf ActionCenterToolStripMenuItem_Click
                End With
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
                With item
                    .Text = IO.Path.GetFileNameWithoutExtension(IIO.FullName)
                    .Tag = IIO.FullName
                    AddHandler item.Click, AddressOf ActionCenterToolStripMenuItem_Click
                End With
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
                With item
                    .Text = IO.Path.GetFileNameWithoutExtension(IIO.FullName)
                    .Tag = IIO.FullName
                    AddHandler item.Click, AddressOf ActionCenterToolStripMenuItem_Click
                End With
                If Not IIO.Name = "desktop.ini" Then
                    ActionCM.Items.Add(item)
                End If
            Next
            ActionCM.Items.Add(ToolStripSeparator14)
            ActionCM.Items.Add(ClipboardToolStripMenuItem)
            ActionCM.Items.Add(PerformanceToolStripMenuItem)
        End If
    End Sub

    Private Sub ToolStripButton3_MouseUp(sender As Object, e As MouseEventArgs) Handles ToolStripButton3.MouseUp
        If e.Button = MouseButtons.Left Then
            Shell("RunDll32.exe shell32.dll,Control_RunDLL ncpa.cpl")
        End If
    End Sub
End Class
