Imports System.Runtime.InteropServices
Imports System.Drawing

Public Class WAT
    Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Integer) As Integer

    Private Declare Function ShowWindow Lib "user32.dll" (
    ByVal hWnd As IntPtr,
    ByVal nCmdShow As SHOW_WINDOW
    ) As Boolean

    <Flags()>
    Private Enum SHOW_WINDOW As Integer
        SW_HIDE = 0
        SW_NORMAL = 1
        SW_MINIMIZED = 2
        SW_MAXIMIZE = 3
        SW_SHOWNOACTIVATE = 4
        SW_SHOW = 5
        SW_MINIMIZE = 6
        SW_MINIMIZENOACTIVE = 7
        SW_SHOWNOACTIVE = 8
        SW_RESTORE = 9
        SW_SHOWDEFAULT = 10
        SW_FORCEMINIMIZE = 11
        SW_MAX = 11
    End Enum

    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (
   ByVal hWnd As Long,
   ByVal wMsg As Long,
   ByVal wParam As Long,
   ByVal lParam As Long) As Long

    Const WM_CLOSE As Long = &H10

#Region " do NOT Activate the Window please. "
    Public Const WS_EX_TOOLWINDOW As Long = &H80L
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim value As CreateParams = MyBase.CreateParams
            'Don't allow the window to be activated.
            value.ExStyle = value.ExStyle Or NativeConstants.WS_EX_NOACTIVATE Or WS_EX_TOOLWINDOW
            Return value
        End Get
    End Property

    Public Overloads Sub ShowUnactivated()
        NativeMethods.ShowWindow(Me.Handle, NativeConstants.SW_SHOWNOACTIVATE)
    End Sub

    Public Overloads Sub ShowUnactivated(ByVal owner As Form)
        Me.Owner = owner
        Me.ShowUnactivated()
    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = NativeConstants.WM_MOUSEACTIVATE Then
            'Tell Windows not to activate the window on a mouse click.
            m.Result = CType(NativeConstants.MA_NOACTIVATE, IntPtr)
        Else
            MyBase.WndProc(m)
        End If
        'End If
    End Sub

    Friend NotInheritable Class NativeMethods
        <Runtime.InteropServices.DllImport("user32")>
        Public Shared Function ShowWindow(ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
        End Function
    End Class

    Friend NotInheritable Class NativeConstants
        Public Const WS_EX_NOACTIVATE As Integer = &H8000000
        Public Const WM_MOUSEACTIVATE As Integer = &H21
        Public Const MA_NOACTIVATE As Integer = 3
        Public Const SW_SHOWNOACTIVATE As Integer = 4
    End Class
#End Region

    <Runtime.InteropServices.DllImport("dwmapi.dll")> Public Shared Function DwmExtendFrameIntoClientArea(ByVal hWnd As IntPtr, ByRef pMarinset As side) As Integer
    End Function

    <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)> Public Structure side
        Public Left As Integer
        Public Right As Integer
        Public Top As Integer
        Public Bottom As Integer
    End Structure

    <DllImport("dwmapi.dll", PreserveSig:=False)>
    Public Shared Function DwmGetColorizationColor(ByRef pcrColorization As UInteger, ByRef pfOpaqueBlend As Boolean) As Integer
    End Function

    Private TargetBounds As Rectangle
    Private WithEvents AnimTimer As New Timer With {.Interval = 10}
    Private ReadOnly Smoothness As Double = 0.3
    Private ReadOnly Epsilon As Integer = 1

    Public Sub MoveTo(ByVal newBounds As Rectangle)
        If Me.IsDisposed OrElse Me.Disposing Then Return

        Dim maxWidth As Integer = 500
        Dim finalRect As Rectangle = newBounds

        If finalRect.Width > maxWidth Then
            Dim widthDifference As Integer = finalRect.Width - maxWidth
            Dim newX As Integer = finalRect.X + (widthDifference / 2)

            finalRect = New Rectangle(newX, finalRect.Y, maxWidth, finalRect.Height)
        End If
        ' ------------------------

        If TargetBounds = finalRect AndAlso Me.Visible Then
            Return
        End If

        TargetBounds = finalRect

        If Not Me.Visible Then
            Me.Opacity = 0
            Me.SetBounds(finalRect.X, finalRect.Y, finalRect.Width, finalRect.Height)
            MyBase.Show()
        End If

        AnimTimer.Start()
    End Sub

    Private Sub AnimTimer_Tick(sender As Object, e As EventArgs) Handles AnimTimer.Tick
        If Me.IsDisposed OrElse Me.Disposing Then
            AnimTimer.Stop()
            Return
        End If

        If Me.Opacity < 1.0 Then
            Me.Opacity = Math.Min(1.0, Me.Opacity + 0.1)
        End If

        Dim dX As Double = TargetBounds.X - Me.Left
        Dim dY As Double = TargetBounds.Y - Me.Top
        Dim dW As Double = TargetBounds.Width - Me.Width
        Dim dH As Double = TargetBounds.Height - Me.Height

        If Math.Abs(dX) < Epsilon AndAlso Math.Abs(dY) < Epsilon AndAlso
       Math.Abs(dW) < Epsilon AndAlso Math.Abs(dH) < Epsilon Then

            Me.Bounds = TargetBounds
            If Me.Opacity >= 1.0 Then AnimTimer.Stop()
        Else
            Dim nextL As Integer = CInt(Me.Left + (dX * Smoothness))
            Dim nextT As Integer = CInt(Me.Top + (dY * Smoothness))
            Dim nextW As Integer = CInt(Me.Width + (dW * Smoothness))
            Dim nextH As Integer = CInt(Me.Height + (dH * Smoothness))

            Me.SetBounds(nextL, nextT, nextW, nextH)
        End If
    End Sub

    Public Shadows Sub Show()
        Me.Opacity = 0
        MyBase.Show()
        FadeTimer.Start()
    End Sub

    Public Shadows Sub Hide()
        AnimTimer.Stop()
        If Not Me.IsDisposed Then
            Me.Opacity = 0
            MyBase.Hide()
        End If
    End Sub

    Private Sub FadeTimer_Tick(sender As Object, e As EventArgs) Handles FadeTimer.Tick
        If Me.Opacity < 1.0 Then
            Me.Opacity += 0.1
        Else
            FadeTimer.Stop()
        End If
    End Sub

    Private Sub WAT_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Me.Location = New Point(Control.MousePosition.X - Me.Width / 2, SystemInformation.WorkingArea.Height - Me.Height)

        Button2.BackColor = Color.Transparent
        Button3.BackColor = Color.Transparent

        Try
            Me.BackColor = Color.Black
            Label1.ForeColor = Color.White

            Dim accentColor As Color = WindowsColorSettings.GetAccentColor()
            Button4.BackColor = accentColor

            Dim side As side = side
            side.Left = -1
            side.Right = -1
            side.Top = -1
            side.Bottom = -1
            Dim result As Integer = DwmExtendFrameIntoClientArea(Me.Handle, side)
        Catch ex As Exception
            Try
                Dim accentColor As Color = WindowsColorSettings.GetAccentColor()
                Me.BackColor = accentColor
                Label1.ForeColor = SystemColors.ActiveCaptionText
            Catch ex1 As Exception
                Me.BackColor = SystemColors.ActiveCaption
                Label1.ForeColor = SystemColors.ActiveCaptionText
            End Try
        End Try

        'Button4.ContextMenuStrip = AppBar.TPCM
    End Sub

    Private Sub WAT_MouseHover(sender As Object, e As EventArgs) Handles Me.MouseLeave, Button4.MouseLeave, Button1.MouseLeave, Button2.MouseLeave, Button3.MouseLeave, Panel1.MouseLeave, Panel4.MouseLeave, Label1.MouseLeave
        Dim liveBar = Application.OpenForms.OfType(Of AppBar)().FirstOrDefault()

        liveBar?.Invoke(Sub()
                            liveBar.CloseTimer.Start()
                        End Sub)
    End Sub

    Private Sub Button4_MouseEnter1(sender As Object, e As EventArgs) Handles Button4.MouseEnter
        Me.Select()
        Button4.Select()
    End Sub

    Private Sub Button4_MouseEnter(sender As Object, e As EventArgs) Handles Button4.MouseEnter, Me.MouseEnter, Button1.MouseEnter, Button2.MouseEnter, Button3.MouseEnter, Panel1.MouseEnter, Panel4.MouseEnter, Label1.MouseEnter
        If AppBar.isAeroPeekEnabled Then TimerPeek.Start()

        Dim liveBar = Application.OpenForms.OfType(Of AppBar)().FirstOrDefault()

        liveBar?.Invoke(Sub()
                            liveBar.CloseTimer.Stop()
                        End Sub)

    End Sub

    Private Sub Button4_MouseLeave(sender As Object, e As EventArgs) Handles Button4.MouseLeave, Me.MouseLeave, Button1.MouseLeave, Button2.MouseLeave, Button3.MouseLeave, Panel1.MouseLeave, Panel4.MouseLeave, Label1.MouseLeave
        TimerPeek.Stop()
        AeroPeek.HidePeek()

        AeroPeekScreen.Hide()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Not Me.Tag Is Nothing AndAlso TypeOf Me.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(Me.Tag, IntPtr)

            Try
                Dim currentState As String = GetWindowState(windowHandle)

                Select Case currentState
                    Case "Normal"
                        ShowWindow(windowHandle, SHOW_WINDOW.SW_MINIMIZE)

                    Case "Maximized"
                        ShowWindow(windowHandle, SHOW_WINDOW.SW_MINIMIZE)

                    Case "Minimized"
                        SetForegroundWindow(windowHandle)
                        ShowWindow(windowHandle, SHOW_WINDOW.SW_RESTORE)

                    Case Else
                        SetForegroundWindow(windowHandle)
                        ShowWindow(windowHandle, SHOW_WINDOW.SW_RESTORE)
                End Select
            Catch ex As Exception

            End Try

            Me.Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not Me.Tag Is Nothing AndAlso TypeOf Me.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(Me.Tag, IntPtr)
            SendMessage(windowHandle, WM_CLOSE, 0, 0)

            Me.Close()
        End If
    End Sub

    <DllImport("user32.dll")>
    Private Shared Function IsIconic(hWnd As IntPtr) As Boolean
    End Function

    Private Sub Label1_MouseUp(sender As Object, e As MouseEventArgs) Handles Label1.MouseUp
        If e.Button = MouseButtons.Right Then
            AppBar.WindowStateToolStripMenuItem.DropDown.Show(AppBar, Cursor.Position)
        End If
    End Sub

    Private Sub Button2_MouseUp(sender As Object, e As MouseEventArgs) Handles Button2.MouseUp
        If Not Me.Tag Is Nothing AndAlso TypeOf Me.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(Me.Tag, IntPtr)

            Try
                Dim currentState As String = GetWindowState(windowHandle)

                Select Case currentState
                    Case "Normal"
                        SetForegroundWindow(windowHandle)
                        ShowWindow(windowHandle, SHOW_WINDOW.SW_MAXIMIZE)

                    Case "Maximized"
                        SetForegroundWindow(windowHandle)
                        ShowWindow(windowHandle, SHOW_WINDOW.SW_NORMAL)

                    Case "Minimized"
                        SetForegroundWindow(windowHandle)
                        ShowWindow(windowHandle, SHOW_WINDOW.SW_MAXIMIZE)

                    Case Else
                        SetForegroundWindow(windowHandle)
                        ShowWindow(windowHandle, SHOW_WINDOW.SW_RESTORE)
                End Select
            Catch ex As Exception

            End Try

            Me.Close()
        End If
    End Sub

    Private Sub WAT_BackColorChanged(sender As Object, e As EventArgs) Handles Me.BackColorChanged
        Label1.BackColor = Me.BackColor
        Panel1.BackColor = Me.BackColor
        Panel2.BackColor = Me.BackColor
        Panel3.BackColor = Me.BackColor
        Panel4.BackColor = Me.BackColor
        Button1.BackColor = Color.Red
    End Sub

    Private Sub TimerPeek_Tick(sender As Object, e As EventArgs) Handles TimerPeek.Tick
        TimerPeek.Stop()

        If Not AppBar.isAeroPeekEnabled Then Exit Sub

        If Me.Tag IsNot Nothing AndAlso TypeOf Me.Tag Is IntPtr Then
            Dim targetHWnd As IntPtr = CType(Me.Tag, IntPtr)
            Dim targetPos = AeroPeek.GetVisibleWindowRect(targetHWnd)

            AeroPeekScreen.Show()
            Me.BringToFront()

            AeroPeek.ShowPeek(targetHWnd, AeroPeekScreen.Handle, targetPos)
        End If
    End Sub

    Private Sub Button4_MouseUp(sender As Object, e As MouseEventArgs) Handles Button4.MouseUp, Label1.Click
        If e.Button = MouseButtons.Right Then
            AppBar.WindowStateToolStripMenuItem.DropDown.Show(AppBar, Cursor.Position)
            Exit Sub
        End If

        TimerPeek.Stop()
        AeroPeek.HidePeek()
        AeroPeekScreen.Hide()

        Me.Hide()

        If Not Me.Tag Is Nothing AndAlso TypeOf Me.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(Me.Tag, IntPtr)

            Try
                Dim currentState As String = GetWindowState(windowHandle)

                Select Case currentState
                    Case "Normal"
                        SetForegroundWindow(windowHandle)

                    Case "Maximized"
                        SetForegroundWindow(windowHandle)

                    Case "Minimized"
                        SetForegroundWindow(windowHandle)
                        ShowWindow(windowHandle, SHOW_WINDOW.SW_RESTORE)

                    Case Else
                        SetForegroundWindow(windowHandle)
                        ShowWindow(windowHandle, SHOW_WINDOW.SW_RESTORE)
                End Select
            Catch ex As Exception

            End Try
        End If
    End Sub
End Class

Public Class WindowsColorSettings

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

    Public Shared Function IsAccentColorAppliedToTitleBar() As Boolean
        Try
            Dim applies As Boolean = False

            Dim result As Integer = DwmGetWindowAttribute(Process.GetCurrentProcess().MainWindowHandle, DWMWA_CAPTION_COLOR, applies, CUInt(Marshal.SizeOf(applies)))
            If result = 0 Then
                Return applies
            End If
        Catch ex As Exception

        End Try

        Return False
    End Function
End Class

Public Class AeroPeek
    <DllImport("dwmapi.dll")>
    Private Shared Function DwmRegisterThumbnail(dest As IntPtr, src As IntPtr, ByRef thumb As IntPtr) As Integer
    End Function

    <DllImport("dwmapi.dll")>
    Private Shared Function DwmUnregisterThumbnail(hThumb As IntPtr) As Integer
    End Function

    <DllImport("dwmapi.dll")>
    Private Shared Function DwmUpdateThumbnailProperties(hThumb As IntPtr, ByRef props As DWM_THUMBNAIL_PROPERTIES) As Integer
    End Function

    <DllImport("dwmapi.dll")>
    Private Shared Function DwmGetWindowAttribute(hwnd As IntPtr, dwAttribute As Integer, ByRef pvAttribute As Rect, cbAttribute As Integer) As Integer
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DWM_THUMBNAIL_PROPERTIES
        Public dwFlags As UInteger
        Public rcDestination As Rect
        Public rcSource As Rect
        Public opacity As Byte
        Public fVisible As Boolean
        Public fSourceClientAreaOnly As Boolean
    End Structure

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

    Private Shared _currentThumbnail As IntPtr = IntPtr.Zero

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowPlacement(hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Private Structure WINDOWPLACEMENT
        Public length As Integer
        Public flags As Integer
        Public showCmd As Integer
        Public ptMinPosition As Point
        Public ptMaxPosition As Point
        Public rcNormalPosition As Rect
    End Structure

    Public Shared Function GetVisibleWindowRect(hWnd As IntPtr) As Rect
        Dim wp As New WINDOWPLACEMENT()
        wp.length = Marshal.SizeOf(wp)
        GetWindowPlacement(hWnd, wp)

        If wp.showCmd = 2 Then
            If (wp.flags And &H2) <> 0 Then
                Dim screenArea = Screen.FromHandle(hWnd).WorkingArea
                Return New Rect With {
    .Left = screenArea.X,
    .Top = screenArea.Y,
    .Right = screenArea.Right,
    .Bottom = screenArea.Bottom
}
            Else
                Return wp.rcNormalPosition
            End If
        Else
            Dim r As New Rect()
            DwmGetWindowAttribute(hWnd, 9, r, Marshal.SizeOf(GetType(Rect)))
            Return r
        End If
    End Function

    Public Shared Sub ShowPeek(hWndToPeek As IntPtr, destinationHandle As IntPtr, targetRect As RECT, Optional clientAreaOnly As Boolean = False)
        HidePeek()

        If DwmRegisterThumbnail(destinationHandle, hWndToPeek, _currentThumbnail) = 0 Then
            Dim props As New DWM_THUMBNAIL_PROPERTIES()
            props.dwFlags = &H1 Or &H4 Or &H8 Or &H10
            props.fVisible = True
            props.fSourceClientAreaOnly = clientAreaOnly
            props.opacity = 255

            props.rcDestination = targetRect

            DwmUpdateThumbnailProperties(_currentThumbnail, props)
        End If
    End Sub

    Public Shared Sub HidePeek()
        If _currentThumbnail <> IntPtr.Zero Then
            DwmUnregisterThumbnail(_currentThumbnail)
            _currentThumbnail = IntPtr.Zero
        End If
    End Sub
End Class