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
    Public MouseAway As Boolean = False

#Region " do NOT Activate the Window please. "
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim value As CreateParams = MyBase.CreateParams
            'Don't allow the window to be activated.
            value.ExStyle = value.ExStyle Or NativeConstants.WS_EX_NOACTIVATE
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
        'If CheckBox1.Checked = True Then
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
    End Sub

    Private Sub WAT_MouseHover(sender As Object, e As EventArgs) Handles Me.MouseLeave, Button4.MouseLeave, Button1.MouseLeave, Button2.MouseLeave, Button3.MouseLeave, Panel1.MouseLeave, Label1.MouseLeave
        AppBar.MouseAway = True
        MouseAway = False
        'AwayCheck.Enabled = True
    End Sub

    Private Sub Button4_MouseEnter(sender As Object, e As EventArgs) Handles Button4.MouseEnter, Me.MouseEnter, Button1.MouseEnter, Button2.MouseEnter, Button3.MouseEnter, Panel1.MouseEnter, Label1.MouseEnter
        MouseAway = False
        AppBar.MouseAway = True
        'AwayCheck.Enabled = True
    End Sub

    Private Sub Button4_MouseLeave(sender As Object, e As EventArgs) Handles Button4.MouseLeave, Me.MouseLeave, Button1.MouseLeave, Button2.MouseLeave, Button3.MouseLeave, Panel1.MouseLeave, Label1.MouseLeave
        MouseAway = True
        AppBar.MouseAway = True
        'AwayCheck.Enabled = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Not AppBar.TPCM.Tag Is Nothing AndAlso TypeOf AppBar.TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(AppBar.TPCM.Tag, IntPtr)

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
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not AppBar.TPCM.Tag Is Nothing AndAlso TypeOf AppBar.TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(AppBar.TPCM.Tag, IntPtr)
            SendMessage(windowHandle, WM_CLOSE, 0, 0)

            Me.Close()
        End If
    End Sub

    Private Sub AwayCheck_Tick(sender As Object, e As EventArgs) Handles AwayCheck.Tick
        If MouseAway = True AndAlso AppBar.MouseAway = True Then
            Me.Close()
            'AwayCheck.Enabled = False
        End If
    End Sub

    <DllImport("user32.dll")>
    Private Shared Function IsIconic(hWnd As IntPtr) As Boolean
    End Function
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Not AppBar.TPCM.Tag Is Nothing AndAlso TypeOf AppBar.TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(AppBar.TPCM.Tag, IntPtr)

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

        Me.Close()
    End Sub

    Private Sub Label1_MouseUp(sender As Object, e As MouseEventArgs) Handles Label1.MouseUp
        If e.Button = MouseButtons.Right Then
            AppBar.WindowStateToolStripMenuItem.DropDown.Show(AppBar, Cursor.Position)
        End If
    End Sub

    Private Sub Button2_MouseUp(sender As Object, e As MouseEventArgs) Handles Button2.MouseUp
        If Not AppBar.TPCM.Tag Is Nothing AndAlso TypeOf AppBar.TPCM.Tag Is IntPtr Then
            Dim windowHandle As IntPtr = CType(AppBar.TPCM.Tag, IntPtr)

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
        End If
    End Sub

    Private Sub WAT_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        If Button4.BackgroundImage IsNot Nothing Then
            Dim targetHeight As Integer = 140
            If Button4.BackgroundImage.Height > targetHeight Then
                Dim aspectRatio As Double = CDbl(Button4.BackgroundImage.Width) / CDbl(Button4.BackgroundImage.Height)
                Dim targetWidth As Integer = CInt(targetHeight * aspectRatio)
                Me.Size = New Size(targetWidth - 10, 140)
            Else
                Me.Size = New Size(178, Button4.BackgroundImage.Height)
            End If
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