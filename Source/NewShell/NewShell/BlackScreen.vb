Public Class BlackScreen
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
        <Runtime.InteropServices.DllImport("user32")> _
        Public Shared Function ShowWindow(ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
        End Function
    End Class

    Friend NotInheritable Class NativeConstants
        Public Const WS_EX_NOACTIVATE As Integer = &H8000000
        Public Const WM_MOUSEACTIVATE As Integer = &H21
        Public Const MA_NOACTIVATE As Integer = 3
        Public Const SW_SHOWNOACTIVATE As Integer = 4
    End Class
End Class