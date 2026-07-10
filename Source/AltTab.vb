Imports KrrShell.AppBar
Imports System.ComponentModel
Imports System.Runtime.InteropServices

Public Class AltTab
    Public ActiveProcessID As Integer = 0
    Public coolPeek As New AeroPeek

    Public ClassicMode As Boolean = True

    <DllImport("dwmapi.dll")>
    Public Shared Function DwmExtendFrameIntoClientArea(ByVal hWnd As IntPtr, ByRef pMarinset As side) As Integer
    End Function

    <StructLayout(LayoutKind.Sequential)> Public Structure side
        Public Left As Integer
        Public Right As Integer
        Public Top As Integer
        Public Bottom As Integer
    End Structure

    <DllImport("dwmapi.dll")>
    Public Shared Function DwmSetWindowAttribute(hwnd As IntPtr, attr As Integer, ByRef attrValue As Integer, attrSize As Integer) As Integer
    End Function

    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)

        Dim value As Integer = 1
        DwmSetWindowAttribute(Me.Handle, 2, value, Marshal.SizeOf(value))
    End Sub

    Public Sub New()
        InitializeComponent()

        Me.FormBorderStyle = FormBorderStyle.None
    End Sub

    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.Style = cp.Style Or &H80000   ' WS_SYSMENU '
            cp.Style = cp.Style Or &H800000  ' WS_BORDER '
            cp.ExStyle = cp.ExStyle Or WS_EX_TOPMOST
            Return cp
        End Get
    End Property

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetWindowRect(ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
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
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Public Shared Function DestroyIcon(ByVal hIcon As IntPtr) As Boolean
    End Function

    Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As EnumWindowsProc, ByVal lParam As IntPtr) As Boolean

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function GetWindowText(ByVal hWnd As IntPtr, ByVal lpString As System.Text.StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function GetWindowTextLength(ByVal hWnd As IntPtr) As Integer
    End Function

    <DllImport("user32.dll")>
    Public Shared Function IsWindowVisible(ByVal hwnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Public Shared Function GetAncestor(ByVal hwnd As IntPtr, ByVal gaFlags As Integer) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Public Shared Function GetParent(ByVal hwnd As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetAsyncKeyState(ByVal vKey As Integer) As Short
    End Function

    Private Const VK_LWIN As Integer = &H5B
    Private Const VK_RWIN As Integer = &H5C

    Public Const GA_ROOTOWNER As Integer = 3
    Public Const GA_ROOT As Integer = 2
    Public Const GWL_EXSTYLE As Integer = -20

    Public Const WS_EX_TOPMOST As Integer = &H8
    Public Const WS_EX_TOOLWINDOW As Integer = &H80
    Public Const WS_EX_APPWINDOW As Integer = &H40000

    <DllImport("user32.dll", EntryPoint:="GetWindowLong", SetLastError:=True)>
    Private Shared Function GetWindowLong32(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
    End Function

    <DllImport("user32.dll", EntryPoint:="GetWindowLongPtr", SetLastError:=True)>
    Private Shared Function GetWindowLongPtr64(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
    End Function

    Public Shared Function GetWindowLong(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
        If IntPtr.Size = 8 Then
            Return CInt(GetWindowLongPtr64(hWnd, nIndex).ToInt64())
        Else
            Return GetWindowLong32(hWnd, nIndex)
        End If
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

        Dim exStyle As Integer = GetWindowLong(hwnd, GWL_EXSTYLE)

        If (exStyle And WS_EX_TOOLWINDOW) = WS_EX_TOOLWINDOW AndAlso (exStyle And WS_EX_APPWINDOW) <> WS_EX_APPWINDOW Then
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
                Return Icon.FromHandle(hIcon).Clone()
            Catch
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
                Dim sb As New System.Text.StringBuilder(titleLength + 1)
                GetWindowText(hwnd, sb, sb.Capacity)

                Dim processId As Integer = 0
                GetWindowThreadProcessId(hwnd, processId)

                WindowList.Add(New TaskbarWindowInfo() With {
                    .Handle = hwnd,
                    .Title = sb.ToString(),
                    .PID = processId
                })
            End If
        End If
        Return True
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

#End Region

    Private ActiveItem As ToolStripMenuItem = Nothing
    Private initialMousePos As Point
    Private Sub AltTab_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FlowLayoutPanel1.Controls.Clear()
        initialMousePos = Cursor.Position
        Dim currentActive As IntPtr = AppBar.ActiveWindowHandle

        Me.Activate()
        Me.Focus()
        Me.KeyPreview = True
        Me.Opacity = 0

        DisposeIconCache()

        ' Theming
        ApplyTheme()
    End Sub

    Public Sub MoveNextItem()
        If FlowLayoutPanel1.Controls.Count = 0 Then Exit Sub

        Dim currentSelItem As Integer = -1

        For i As Integer = 0 To FlowLayoutPanel1.Controls.Count - 1
            If FlowLayoutPanel1.Controls(i).Focused Then
                currentSelItem = i
                Exit For
            End If
        Next

        Dim nextSelItem As Integer = currentSelItem + 1
        If nextSelItem > FlowLayoutPanel1.Controls.Count - 1 Then
            nextSelItem = 0
        End If

        Dim btnToSelect As Button = DirectCast(FlowLayoutPanel1.Controls(nextSelItem), Button)
        btnToSelect.Focus()

        FlowLayoutPanel1.ScrollControlIntoView(btnToSelect)

        UpdatePreview()
    End Sub

    Public Sub MovePreviousItem()
        If FlowLayoutPanel1.Controls.Count = 0 Then Exit Sub

        Dim currentSelItem As Integer = -1

        For i As Integer = 0 To FlowLayoutPanel1.Controls.Count - 1
            If FlowLayoutPanel1.Controls(i).Focused Then
                currentSelItem = i
                Exit For
            End If
        Next

        Dim nextSelItem As Integer = currentSelItem - 1
        If nextSelItem < 0 Then
            nextSelItem = FlowLayoutPanel1.Controls.Count - 1
        End If

        Dim btnToSelect As Button = DirectCast(FlowLayoutPanel1.Controls(nextSelItem), Button)
        btnToSelect.Select()

        FlowLayoutPanel1.ScrollControlIntoView(btnToSelect)
        UpdatePreview()
    End Sub

    Public Sub ApplyTheme()
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "ClassicMode", False) Then
            Me.ForeColor = SystemColors.ControlText
            Me.BackColor = SystemColors.ControlLightLight

            Me.FormBorderStyle = FormBorderStyle.Fixed3D

            PictureBox1.BorderStyle = BorderStyle.Fixed3D
            ClassicMode = True
        Else
            Me.ForeColor = AppBar.primaryForeColor
            Me.BackColor = GetAccentColor()

            Me.FormBorderStyle = FormBorderStyle.None

            PictureBox1.BorderStyle = BorderStyle.FixedSingle
            ClassicMode = False

            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "DWMAffect", False) Then
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
            End If
        End If
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

    <DllImport("user32.dll")>
    Private Shared Sub keybd_event(bVk As Byte, bScan As Byte, dwFlags As UInteger, dwExtraInfo As UIntPtr)
    End Sub

    <DllImport("user32.dll")>
    Private Shared Function BringWindowToTop(hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Sub SwitchToThisWindow(hWnd As IntPtr, fAltTab As Boolean)
    End Sub

    Private Const VK_MENU As Byte = &H12 ' ALT
    Private Const KEYEVENTF_KEYUP As UInteger = &H2

    Private Sub ForceForegroundWindow(ByVal hWnd As IntPtr)
        Dim foregroundThread As UInteger = GetWindowThreadProcessId(GetForegroundWindow(), IntPtr.Zero)
        Dim targetThread As UInteger = GetWindowThreadProcessId(hWnd, IntPtr.Zero)

        'keybd_event(VK_MENU, 0, 0, UIntPtr.Zero)
        'keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, UIntPtr.Zero)

        If foregroundThread <> targetThread Then
            AttachThreadInput(foregroundThread, targetThread, True)
            SetForegroundWindow(hWnd)
            BringWindowToTop(hWnd)
            AttachThreadInput(foregroundThread, targetThread, False)
        Else
            SetForegroundWindow(hWnd)
            BringWindowToTop(hWnd)
        End If

        Dim currentState As String = GetWindowState(hWnd)

        Select Case currentState
            Case "Normal"
                ShowWindow(hWnd, SHOW_WINDOW.SW_SHOWNORMAL)
                SetForegroundWindow(hWnd)
            Case "Maximized"
                ShowWindow(hWnd, SHOW_WINDOW.SW_MAXIMIZE)
                SetForegroundWindow(hWnd)
            Case "Minimized"
                ShowWindow(hWnd, SHOW_WINDOW.SW_RESTORE)
                SetForegroundWindow(hWnd)
            Case Else
                ShowWindow(hWnd, SHOW_WINDOW.SW_RESTORE)
                SetForegroundWindow(hWnd)
        End Select

        SwitchToThisWindow(hWnd, True)
        SetForegroundWindow(hWnd)

        AppBar.ActiveWindowHandle = hWnd
    End Sub

    Private Sub AltTab_KeyUp(sender As Object, e As KeyEventArgs) Handles MyBase.KeyUp ' Přidat případně Handles pro prvky, pokud neodchytává KeyPreview
        If e.KeyCode = Keys.LWin OrElse e.KeyCode = Keys.RWin OrElse e.KeyCode = Keys.Menu Then
            Dim selectedItem As Button = FlowLayoutPanel1.Controls.OfType(Of Button)().FirstOrDefault(Function(i) i.Focused)

            If selectedItem IsNot Nothing AndAlso selectedItem.Tag IsNot Nothing Then
                ActivateWindowByHandle(CType(selectedItem.Tag, IntPtr))
            End If

            Me.Opacity = 0
            Me.Hide()
            e.Handled = True
        End If
    End Sub

    Private Sub AltTab_MouseWheel(sender As Object, e As MouseEventArgs) Handles Me.MouseWheel, FlowLayoutPanel1.MouseWheel
        If My.Computer.Keyboard.CtrlKeyDown Then Exit Sub

        If e.Delta < 0 Then
            MoveNextItem()
        ElseIf e.Delta > 0 Then
            MovePreviousItem()
        End If

        Dim handeArgs As HandledMouseEventArgs = TryCast(e, HandledMouseEventArgs)
        If handeArgs IsNot Nothing Then
            handeArgs.Handled = True
        End If
    End Sub

    Private Sub UpdatePreview()
        Dim selectedItem As Button = FlowLayoutPanel1.Controls.OfType(Of Button)().FirstOrDefault(Function(i) i.Focused)

        If selectedItem IsNot Nothing AndAlso selectedItem.Tag IsNot Nothing Then
            Dim windowHandle As IntPtr = CType(selectedItem.Tag, IntPtr)

            Dim apps = GetTaskbarApplications()
            Dim appInfo = apps.FirstOrDefault(Function(w) w.Handle = windowHandle)
            If appInfo IsNot Nothing Then Label1.Text = appInfo.Title

            Try
                Dim windowRect As RECT
                GetWindowRect(windowHandle, windowRect)

                Dim targetHeight As Integer = 165
                If windowRect.Width > targetHeight Then
                    Dim aspectRatio As Double = CDbl(windowRect.Width) / CDbl(windowRect.Height)
                    Dim targetWidth As Integer = CInt(targetHeight * aspectRatio)
                    PictureBox1.Size = New Size(targetWidth + 10, targetHeight)
                Else
                    PictureBox1.Size = New Size(178, windowRect.Height)
                End If

                Dim centerX As Integer = (Panel1.Width - PictureBox1.Width) \ 2
                Dim centerY As Integer = (Panel1.Height - PictureBox1.Height) \ 2
                PictureBox1.Location = New Point(centerX, centerY)

                coolPeek.ShowControlPeek(windowHandle, Me, PictureBox1, False)

            Catch ex As Exception
                Debug.WriteLine("Preview render failed: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub ActivateWindowByHandle(ByVal windowHandle As IntPtr)
        Try
            ForceForegroundWindow(windowHandle)
            AppBar.ActiveWindowHandle = windowHandle
        Catch ex As Exception
            Debug.WriteLine("Activation failed: " & ex.Message)
        End Try
    End Sub

    Private Sub AltTab_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        ActiveProcessID = 0

        coolPeek.HidePeek()
    End Sub

    Private Function IsKeyDown(ByVal key As Keys) As Boolean
        Return (GetAsyncKeyState(CInt(key)) And &H8000) <> 0
    End Function

    Private Sub Controller_Tick(sender As Object, e As EventArgs) Handles Controller.Tick
        If Me.Visible Then
            If Not My.Computer.Keyboard.AltKeyDown AndAlso Not IsKeyDown(Keys.LWin) AndAlso Not IsKeyDown(Keys.RWin) Then
                ActivateSelectedWindow()
            End If
        End If
    End Sub

    Public Sub ActivateSelectedWindow()
        Dim selectedItem As Button = FlowLayoutPanel1.Controls.OfType(Of Button)().FirstOrDefault(Function(i) i.Focused)

        If selectedItem IsNot Nothing AndAlso selectedItem.Tag IsNot Nothing Then
            Dim windowHandle As IntPtr = CType(selectedItem.Tag, IntPtr)
            ActivateWindowByHandle(windowHandle)
        End If

        Me.Opacity = 0
        Me.Hide()
    End Sub

    Private Sub AltTab_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        If Me.Visible Then
            FlowLayoutPanel1.Controls.Clear()
            initialMousePos = Cursor.Position
            Dim currentActive As IntPtr = AppBar.ActiveWindowHandle

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

                    Dim btn As New Button With {
                        .Text = "",
                        .Tag = windowInfo.Handle,
                        .Size = New Size(64, 64),
                        .Image = GetProcessIcon(windowInfo.Handle, SystemIcons.Application.ToBitmap),
                        .ImageList = Nothing,
                        .FlatStyle = FlatStyle.Flat
                    }
                    btn.FlatAppearance.BorderSize = 0
                    btn.FlatAppearance.BorderColor = SystemColors.ActiveBorder
                    btn.FlatAppearance.CheckedBackColor = SystemColors.ButtonHighlight

                    Dim tt As New ToolTip()
                    tt.SetToolTip(btn, windowInfo.Title)

                    AddHandler btn.MouseEnter, Sub(senderItem As Object, ea As EventArgs)
                                                   Dim currentPos As Point = Cursor.Position
                                                   Dim deltaX As Integer = Math.Abs(currentPos.X - initialMousePos.X)
                                                   Dim deltaY As Integer = Math.Abs(currentPos.Y - initialMousePos.Y)

                                                   If deltaX > 5 OrElse deltaY > 5 Then
                                                       Dim targetBtn = DirectCast(senderItem, Button)
                                                       targetBtn.Focus()
                                                       FlowLayoutPanel1.ScrollControlIntoView(targetBtn)
                                                       UpdatePreview()
                                                   End If
                                               End Sub

                    AddHandler btn.GotFocus, Sub(senderItem As Object, ea As EventArgs)
                                                 Dim targetBtn = DirectCast(senderItem, Button)

                                                 targetBtn.FlatAppearance.BorderSize = 1
                                                 If ClassicMode Then targetBtn.BackColor = SystemColors.Highlight
                                             End Sub

                    AddHandler btn.LostFocus, Sub(senderItem As Object, ea As EventArgs)
                                                  Dim targetBtn = DirectCast(senderItem, Button)

                                                  targetBtn.FlatAppearance.BorderSize = 0
                                                  targetBtn.BackColor = Color.Transparent
                                              End Sub

                    AddHandler btn.Click, Sub(senderItem As Object, ea As EventArgs)
                                              Dim targetBtn = DirectCast(senderItem, Button)
                                              ActivateWindowByHandle(CType(targetBtn.Tag, IntPtr))
                                              Me.Opacity = 0
                                              Me.Hide()
                                          End Sub

                    If Not p.Responding Then btn.BackColor = Color.Red

                    FlowLayoutPanel1.Controls.Add(btn)
                Catch : End Try
            Next

            Dim btnCount As Integer = FlowLayoutPanel1.Controls.Count

            If btnCount > 0 Then
                Dim btnMargin As Padding = FlowLayoutPanel1.Controls(0).Margin
                Dim fullBtnWidth As Integer = 64 + btnMargin.Horizontal
                Dim fullBtnHeight As Integer = 64 + btnMargin.Vertical

                Dim targetColumns As Integer = 6
                Dim targetRows As Integer = CInt(Math.Ceiling(btnCount / CDbl(targetColumns)))
                Dim actualColumns As Integer = Math.Min(btnCount, targetColumns)

                Dim contentWidth As Integer = actualColumns * fullBtnWidth + FlowLayoutPanel1.Padding.Horizontal
                Dim contentHeight As Integer = targetRows * fullBtnHeight + FlowLayoutPanel1.Padding.Vertical

                If targetRows > 2 Then
                    contentHeight = 2 * fullBtnHeight + FlowLayoutPanel1.Padding.Vertical
                    contentWidth += SystemInformation.VerticalScrollBarWidth
                End If

                FlowLayoutPanel1.Size = New Size(contentWidth, contentHeight)

                Dim maxContentWidth As Integer = Math.Max(contentWidth, Math.Max(Panel1.Width, Panel2.Width))
                Dim finalFormWidth As Integer = maxContentWidth + (Me.Width - Me.ClientRectangle.Width)

                Dim windowBordersHeight As Integer = Me.Height - Me.ClientRectangle.Height
                Dim finalFormHeight As Integer = Panel1.Height + contentHeight + Panel2.Height + windowBordersHeight

                Me.Size = New Size(finalFormWidth, finalFormHeight)
            End If

        Else
            ActiveProcessID = 0
            coolPeek.HidePeek()
        End If
    End Sub

    Public Async Sub OpenWithDelay()
        Me.Opacity = 0

        ShowWindow(Me.Handle, SHOW_WINDOW.SW_SHOWNOACTIVATE)

        If FlowLayoutPanel1.Controls.Count > 1 Then
            Dim targetBtn As Button = DirectCast(FlowLayoutPanel1.Controls(1), Button)
            targetBtn.Select()
            FlowLayoutPanel1.ScrollControlIntoView(targetBtn)
        ElseIf FlowLayoutPanel1.Controls.Count > 0 Then
            Dim targetBtn As Button = DirectCast(FlowLayoutPanel1.Controls(0), Button)
            targetBtn.Select()
            FlowLayoutPanel1.ScrollControlIntoView(targetBtn)
        End If

        UpdatePreview()

        ' --- ASYNC GUI WAIT --- '
        Dim coolDelayObj As Object = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "ShowDelay", 0)
        Dim coolDelay As Integer = Convert.ToInt32(coolDelayObj)

        ApplyTheme()

        If coolDelay > 0 Then
            Await Task.Delay(coolDelay)
            If Me.Visible Then
                Me.Opacity = 1
            End If
        Else
            Me.Opacity = 1
        End If
    End Sub
End Class