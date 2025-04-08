Public Class WAT
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

    Private Sub WAT_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Location = New Point(Control.MousePosition.X - Me.Width / 2, SystemInformation.WorkingArea.Height - Me.Height)
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub WAT_MouseHover(sender As Object, e As EventArgs) Handles Me.MouseLeave, PictureBox1.MouseLeave, Button1.MouseLeave, Button2.MouseLeave, Button3.MouseLeave, Panel1.MouseLeave, Label1.MouseLeave
        AppBar.MouseAway = True
        MouseAway = True
        Me.Close()
    End Sub

    Private Sub PictureBox1_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox1.MouseEnter, Me.MouseEnter, Button1.MouseEnter, Button2.MouseEnter, Button3.MouseEnter, Panel1.MouseEnter, Label1.MouseEnter
        MouseAway = False
        AppBar.MouseAway = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim App As Process = Process.GetProcessById(AppBar.TPCM.Tag)
        If App.Id > 0 Then
            ShowWindow(App.MainWindowHandle, SHOW_WINDOW.SW_MINIMIZE)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim App As Process = Process.GetProcessById(AppBar.TPCM.Tag)
        If App.Id > 0 Then
            AppActivate(App.Id)
            ShowWindow(App.MainWindowHandle, SHOW_WINDOW.SW_NORMAL)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim App As Process = Process.GetProcessById(AppBar.TPCM.Tag)
        If App.Id > 0 Then
            App.CloseMainWindow()
        End If
    End Sub
End Class