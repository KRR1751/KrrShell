Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Public Class ScreenDrawer
    Private _engine As BrightnessEngine

    Public Sub New()
        InitializeComponent()
        Me.DoubleBuffered = True
    End Sub

    Private Sub ScreenDrawer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Me.IsHandleCreated Then
            InitOsd(Me)

            _engine = New BrightnessEngine()
        End If
    End Sub

    Public Sub OnVolumeChanged()
        Me.Invalidate()
        UltimateOverlay.ShowVolumeOsd()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim g = e.Graphics
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias

        Dim currentVolume As Integer = CInt(Math.Round(VolumeControl.GetCurrentVolume() * 100))
        Dim isMuted As Boolean = VolumeControl.IsSysMuted()

        ' Colors
        Dim mainColor As Color = If(isMuted, Color.Gray, Color.White)
        Dim glowColor As Color = Color.FromArgb(100, 2, 2, 2)
        Dim statusText As String = If(isMuted, "MUTED", "VOLUME: " & currentVolume & "%")

        Using font As New Font("Segoe UI", 36, FontStyle.Bold)
            For x As Integer = -2 To 2 Step 2
                For y As Integer = -2 To 2 Step 2
                    If x = 0 And y = 0 Then Continue For
                    Using glowBrush As New SolidBrush(glowColor)
                        g.DrawString(statusText, font, glowBrush, 20 + x, 15 + y)
                    End Using
                Next
            Next

            Using textBrush As New SolidBrush(mainColor)
                g.DrawString(statusText, font, textBrush, 20, 15)
            End Using
        End Using

        ' Slider
        Dim trackRect As New Rectangle(20, Me.Height - 40, Me.Width - 40, 10)
        g.FillRectangle(New SolidBrush(Color.FromArgb(50, 255, 255, 255)), trackRect)

        If Not isMuted Then
            Dim fillWidth As Integer = CInt((trackRect.Width / 100) * currentVolume)
            If fillWidth > 0 Then
                Dim fillRect As New Rectangle(20, Me.Height - 40, fillWidth, 10)
                g.FillRectangle(New SolidBrush(mainColor), fillRect)
            End If
        End If
    End Sub

    Private Sub ScreenDrawer_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        If Not Me.WindowState = FormWindowState.Maximized Then Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub ScreenDrawer_Activated(sender As Object, e As EventArgs) Handles Me.Activated, Me.GotFocus
        If AppBar.IsHandleCreated Then AppBar.Focus()
    End Sub

    Private Sub ScreenDrawer_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        StopOsdGuard()
    End Sub
End Class

Public Class BrightnessEngine
    <StructLayout(LayoutKind.Sequential)>
    Public Structure PHYSICAL_MONITOR
        Public hPhysicalMonitor As IntPtr
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)>
        Public szPhysicalMonitorDescription As String
    End Structure

    <DllImport("user32.dll")>
    Private Shared Function MonitorFromWindow(hwnd As IntPtr, dwFlags As UInteger) As IntPtr
    End Function

    <DllImport("dxva2.dll", SetLastError:=True)>
    Private Shared Function GetNumberOfPhysicalMonitorsFromHMONITOR(hMonitor As IntPtr, ByRef pdwNumberOfPhysicalMonitors As UInteger) As Boolean
    End Function

    <DllImport("dxva2.dll", SetLastError:=True)>
    Private Shared Function GetPhysicalMonitorsFromHMONITOR(hMonitor As IntPtr, dwPhysicalMonitorArraySize As UInteger, <Out> pPhysicalMonitorArray As PHYSICAL_MONITOR()) As Boolean
    End Function

    <DllImport("dxva2.dll", SetLastError:=True)>
    Private Shared Function GetMonitorBrightness(hMonitor As IntPtr, ByRef pdwMinimumBrightness As UInteger, ByRef pdwCurrentBrightness As UInteger, ByRef pdwMaximumBrightness As UInteger) As Boolean
    End Function

    <DllImport("dxva2.dll", SetLastError:=True)>
    Private Shared Function SetMonitorBrightness(hMonitor As IntPtr, dwNewBrightness As UInteger) As Boolean
    End Function

    <DllImport("dxva2.dll", SetLastError:=True)>
    Private Shared Function DestroyPhysicalMonitors(dwPhysicalMonitorArraySize As UInteger, pPhysicalMonitorArray As PHYSICAL_MONITOR()) As Boolean
    End Function

    Private Const MONITOR_DEFAULTTOPRIMARY As UInteger = &H1

    Public Shared Sub SetBrightness(windowHandle As IntPtr, newValue As Integer)
        If newValue < 0 Then newValue = 0
        If newValue > 100 Then newValue = 100

        Dim hMonitor As IntPtr = MonitorFromWindow(windowHandle, MONITOR_DEFAULTTOPRIMARY)
        Dim numMonitors As UInteger = 0

        If GetNumberOfPhysicalMonitorsFromHMONITOR(hMonitor, numMonitors) AndAlso numMonitors > 0 Then
            Dim physicalMonitors(numMonitors - 1) As PHYSICAL_MONITOR
            ' DŮLEŽITÉ: Získat, nastavit a OKAMŽITĚ zničit handle, aby se uvolnila sběrnice pro DWM
            If GetPhysicalMonitorsFromHMONITOR(hMonitor, numMonitors, physicalMonitors) Then
                SetMonitorBrightness(physicalMonitors(0).hPhysicalMonitor, CUInt(newValue))
                DestroyPhysicalMonitors(numMonitors, physicalMonitors)
            End If
        End If
    End Sub

    Public Shared Function GetBrightness(windowHandle As IntPtr) As Integer
        Dim currentBrightness As UInteger = 0
        Dim hMonitor As IntPtr = MonitorFromWindow(windowHandle, MONITOR_DEFAULTTOPRIMARY)
        Dim numMonitors As UInteger = 0

        If GetNumberOfPhysicalMonitorsFromHMONITOR(hMonitor, numMonitors) AndAlso numMonitors > 0 Then
            Dim physicalMonitors(numMonitors - 1) As PHYSICAL_MONITOR
            If GetPhysicalMonitorsFromHMONITOR(hMonitor, numMonitors, physicalMonitors) Then
                Dim min As UInteger, max As UInteger
                GetMonitorBrightness(physicalMonitors(0).hPhysicalMonitor, min, currentBrightness, max)
                DestroyPhysicalMonitors(numMonitors, physicalMonitors)
            End If
        End If
        Return CInt(currentBrightness)
    End Function
End Class

Module UltimateOverlay
    <DllImport("user32.dll")>
    Private Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Function SetWindowPos(hWnd As IntPtr, hWndInsertAfter As IntPtr, X As Integer, Y As Integer, cx As Integer, cy As Integer, uFlags As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Function SetWinEventHook(eventMin As UInteger, eventMax As UInteger, hmodWinEventProc As IntPtr, lpfnWinEventProc As WinEventDelegate, idProcess As UInteger, idThread As UInteger, dwFlags As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Function UnhookWinEvent(hWinEventHook As IntPtr) As Boolean
    End Function

    Private Const GWL_EXSTYLE As Integer = -20
    Private Const WS_EX_LAYERED As Integer = &H80000
    Private Const WS_EX_TRANSPARENT As Integer = &H20

    Private ReadOnly HWND_TOPMOST As New IntPtr(-1)
    Private Const SWP_NOMOVE As Integer = &H2
    Private Const SWP_NOSIZE As Integer = &H1
    Private Const SWP_NOACTIVATE As Integer = &H10
    Private Const SWP_SHOWWINDOW As Integer = &H40

    Private Const WINEVENT_OUTOFCONTEXT As Integer = &H0
    Private Const EVENT_SYSTEM_FOREGROUND As UInteger = &H3
    Private Const EVENT_OBJECT_LOCATIONCHANGE As UInteger = &H800B

    Private Const OBJID_WINDOW As Integer = 0
    Private Const CHILDID_SELF As Integer = 0

    Public Delegate Sub WinEventDelegate(hWinEventHook As IntPtr, eventType As UInteger, hwnd As IntPtr, idObject As Integer, idChild As Integer, dwEventThread As UInteger, dwmsEventTime As UInteger)
    Private _hook As IntPtr
    Private _procDelegate As WinEventDelegate
    Private _targetHwnd As IntPtr

    Public Sub MakeTransparent(hWnd As IntPtr)
        Dim currentStyle As Integer = GetWindowLong(hWnd, GWL_EXSTYLE)
        SetWindowLong(hWnd, GWL_EXSTYLE, currentStyle Or WS_EX_LAYERED Or WS_EX_TRANSPARENT)
    End Sub

    Public Sub ForceToTop(hWnd As IntPtr)
        SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE)
    End Sub

    Public Sub StartOsdGuard(targetHwnd As IntPtr)
        _targetHwnd = targetHwnd
        _procDelegate = New WinEventDelegate(AddressOf WinEventProc)

        _hook = SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_OBJECT_LOCATIONCHANGE, IntPtr.Zero, _procDelegate, 0, 0, WINEVENT_OUTOFCONTEXT)
    End Sub

    Public Sub StopOsdGuard()
        If _hook <> IntPtr.Zero Then
            UnhookWinEvent(_hook)
            _hook = IntPtr.Zero
        End If
    End Sub

    Private WithEvents _fadeTimer As New Windows.Forms.Timer With {.Interval = 50}
    Private WithEvents _displayTimer As New Windows.Forms.Timer With {.Interval = 3000}
    Private _parentForm As Form

    Public Sub InitOsd(f As Form)
        _parentForm = f
        MakeTransparent(f.Handle)
        StartOsdGuard(f.Handle)

        AddHandler _fadeTimer.Tick, AddressOf FadeTick
        AddHandler _displayTimer.Tick, AddressOf DisplayTick
    End Sub

    Public Sub ShowVolumeOsd()
        _parentForm.Opacity = 1.0
        _displayTimer.Stop()
        _fadeTimer.Stop()
        _displayTimer.Start()
        ForceToTop(_parentForm.Handle)
    End Sub

    Private Sub DisplayTick(sender As Object, e As EventArgs)
        _displayTimer.Stop()
        _fadeTimer.Start()
    End Sub

    Private Sub FadeTick(sender As Object, e As EventArgs)
        If _parentForm.Opacity > 0 Then
            _parentForm.Opacity -= 0.1
        Else
            _fadeTimer.Stop()
        End If
    End Sub

    Private Sub WinEventProc(hWinEventHook As IntPtr, eventType As UInteger, hwnd As IntPtr, idObject As Integer, idChild As Integer, dwEventThread As UInteger, dwmsEventTime As UInteger)
        If idObject = OBJID_WINDOW AndAlso idChild = CHILDID_SELF Then
            If hwnd <> _targetHwnd Then
                ForceToTop(_targetHwnd)
            End If
        End If
    End Sub
End Module