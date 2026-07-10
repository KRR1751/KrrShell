Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

Public Class BalloonPopup
    Inherits Form

    ' ============================================================
    '  Public events
    ' ============================================================

    ' Fired when the user clicks anywhere on the popup body (not the X button).
    Public Event BalloonClicked As EventHandler

    ' Fired when the popup closes for any reason (timeout, click, manual Close()).
    Public Event BalloonClosed As EventHandler

    ' ============================================================
    '  Public properties
    ' ============================================================

    <DefaultValue("")>
    Public Property Title As String = ""

    <DefaultValue("")>
    Public Property Body As String = ""

    ' Custom icon image. The setter makes a DEFENSIVE COPY so that the
    ' caller can dispose the original bitmap without affecting us.
    Public Property IconImage As Image
        Get
            Return _iconImage
        End Get
        Set(value As Image)
            ' Dispose any previous copy we held.
            If _iconImage IsNot Nothing Then
                _iconImage.Dispose()
                _iconImage = Nothing
            End If

            If value Is Nothing Then Return

            ' *** KEY FIX: make our own copy ***
            ' The source bitmap (e.g. info.Icon) may be disposed by another
            ' thread (async icon loader, NIM_MODIFY with NIF_ICON, etc).
            ' If we just store the reference, OnPaint will throw
            ' "Parameter is not valid." from GDI+ when it tries to draw a
            ' disposed bitmap. By making a copy NOW while the source is
            ' still alive, we become independent of its lifetime.
            Try
                _iconImage = CType(New Bitmap(value), Bitmap)
            Catch
                ' Source bitmap is already disposed or corrupt - fall back
                ' to the system icon based on ToolTipIcon.
                _iconImage = Nothing
            End Try
        End Set
    End Property
    Private _iconImage As Bitmap

    ' System icon type used when IconImage is Nothing.
    <DefaultValue(ToolTipIcon.None)>
    Public Property ToolTipIcon As ToolTipIcon = ToolTipIcon.None

    ' Screen-coordinates position where the popup should appear.
    Public Property PopupPosition As Point = Point.Empty

    ' How long to stay visible before auto-closing (ms). 0 = manual close only.
    <DefaultValue(10000)>
    Public Property Duration As Integer = 10000

    ' ============================================================
    '  Internal state
    ' ============================================================

    Private WithEvents autoCloseTimer As Timer
    Friend WithEvents FadeTimer As Timer
    Private components As IContainer
    Private _closeReason As String = "manual"

    ' ============================================================
    '  Constructor + UI setup
    ' ============================================================

    Public Sub New()
        Me.FormBorderStyle = FormBorderStyle.None
        Me.StartPosition = FormStartPosition.Manual
        Me.ShowInTaskbar = False
        Me.TopMost = True                          ' always on top, regardless of focus
        Me.BackColor = Color.FromArgb(255, 252, 220)   ' classic balloon cream
        Me.ForeColor = Color.Black
        Me.Font = New Font("Segoe UI", 9)
        Me.Size = New Size(280, 80)
        Me.DoubleBuffered = True
        Me.TopLevel = True
    End Sub

    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)

        ' Position the popup at PopupPosition (screen coords), clamped to screen.
        If PopupPosition <> Point.Empty Then
            Dim loc As Point = PopupPosition
            Dim scr As Screen = Screen.FromPoint(loc)
            ' Shift up-left if it would go off the right/bottom edge.
            If loc.X + Me.Width > scr.WorkingArea.Right Then
                loc.X = scr.WorkingArea.Right - Me.Width - 10
            End If
            If loc.Y + Me.Height > scr.WorkingArea.Bottom Then
                loc.Y = scr.WorkingArea.Bottom - Me.Height - 10
            End If
            Me.Location = loc
        End If

        ' Start auto-close timer.
        If Duration > 0 Then
            autoCloseTimer = New Timer() With {.Interval = Duration}
            AddHandler autoCloseTimer.Tick, Sub()
                                                _closeReason = "timeout"
                                                Me.Close()
                                            End Sub
            autoCloseTimer.Start()
        End If
    End Sub

    Protected Overrides Sub OnClosing(e As CancelEventArgs)
        ' Stop and dispose the timer first so it can't fire during close.
        If autoCloseTimer IsNot Nothing Then
            autoCloseTimer.Stop()
            autoCloseTimer.Dispose()
            autoCloseTimer = Nothing
        End If
        RaiseEvent BalloonClosed(Me, EventArgs.Empty)
        MyBase.OnClosing(e)
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            ' Dispose our private icon copy.
            If _iconImage IsNot Nothing Then
                _iconImage.Dispose()
                _iconImage = Nothing
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' ============================================================
    '  Custom painting - we draw title, body, icon, and a nice border
    ' ============================================================

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        ' *** DEFENSIVE try/catch ***
        ' If anything goes wrong with GDI+ (disposed bitmap, corrupt icon
        ' handle, etc), we MUST NOT crash the host application. Just skip
        ' painting for this frame and let the next paint try again.
        Try
            MyBase.OnPaint(e)
            PaintPopup(e.Graphics)
        Catch ex As Exception
            Debug.WriteLine("BalloonPopup.OnPaint failed: " & ex.Message)
            ' Try to at least fill the background so the popup isn't transparent.
            Try
                Using brush As New SolidBrush(Me.BackColor)
                    e.Graphics.FillRectangle(brush, Me.ClientRectangle)
                End Using
            Catch
            End Try
        End Try
    End Sub

    Private Sub PaintPopup(g As Graphics)
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

        Dim rect As New Rectangle(0, 0, Me.Width - 1, Me.Height - 1)

        ' Background fill (in case DoubleBuffered didn't clear it)
        Using bgBrush As New SolidBrush(Me.BackColor)
            g.FillRectangle(bgBrush, rect)
        End Using

        ' Border
        Using pen As New Pen(Color.FromArgb(180, 160, 120), 1)
            g.DrawRectangle(pen, rect)
        End Using

        ' Icon: custom IconImage preferred, else system icon from ToolTipIcon.
        ' We check IsNothing AND try-catch the draw call because the bitmap
        ' could become invalid between the check and the draw (very unlikely
        ' now that we keep our own copy, but defense in depth).
        Dim iconRect As New Rectangle(8, 10, 24, 24)
        If _iconImage IsNot Nothing Then
            Try
                g.DrawImage(_iconImage, iconRect)
            Catch ex As Exception
                ' Bitmap went bad - fall through to system icon.
                DrawSystemIcon(g, iconRect)
            End Try
        Else
            DrawSystemIcon(g, iconRect)
        End If

        ' Title (bold)
        Dim titleX As Integer = iconRect.Right + 8
        Using titleFont As New Font(Me.Font, FontStyle.Bold)
            Using titleBrush As New SolidBrush(Color.FromArgb(60, 60, 60))
                g.DrawString(If(Title, ""), titleFont, titleBrush, titleX, 8)
            End Using
        End Using

        ' Body (regular)
        Dim bodyRect As New Rectangle(titleX, 30, Me.Width - titleX - 12, Me.Height - 38)
        Using bodyBrush As New SolidBrush(Color.Black)
            ' Use a NEW StringFormat instead of the shared GenericDefault,
            ' which could theoretically be disposed by some other code.
            Using sf As New StringFormat()
                g.DrawString(If(Body, ""), Me.Font, bodyBrush, bodyRect, sf)
            End Using
        End Using

        ' Small "X" close button in top-right corner
        Dim closeRect As New Rectangle(Me.Width - 20, 4, 14, 14)
        Using closePen As New Pen(Color.FromArgb(120, 120, 120), 1)
            g.DrawLine(closePen, closeRect.X, closeRect.Y, closeRect.Right, closeRect.Bottom)
            g.DrawLine(closePen, closeRect.Right, closeRect.Y, closeRect.X, closeRect.Bottom)
        End Using
    End Sub

    Private Sub DrawSystemIcon(g As Graphics, iconRect As Rectangle)
        Dim sysIcon As Icon = GetSystemIcon(ToolTipIcon)
        If sysIcon IsNot Nothing Then
            Try
                g.DrawIcon(sysIcon, iconRect)
            Catch
            End Try
        End If
    End Sub

    Private Shared Function GetSystemIcon(icon As ToolTipIcon) As Icon
        Select Case icon
            Case ToolTipIcon.Info : Return SystemIcons.Information
            Case ToolTipIcon.Warning : Return SystemIcons.Warning
            Case ToolTipIcon.Error : Return SystemIcons.Error
            Case Else : Return Nothing
        End Select
    End Function

    ' ============================================================
    '  Click handling - the WHOLE popup is clickable
    ' ============================================================

    ' We override OnMouseClick (not OnClick) because MouseClick gives us
    ' the button that was pressed. OnClick fires for keyboard activation
    ' too which we don't want here.
    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        MyBase.OnMouseClick(e)

        If e.Button <> MouseButtons.Left Then Return

        Dim closeRect As New Rectangle(Me.Width - 20, 4, 14, 14)

        If closeRect.Contains(e.Location) Then
            ' User clicked the X close button.
            _closeReason = "manual"
            Me.Close()
        Else
            ' User clicked anywhere else on the popup body.
            _closeReason = "click"
            RaiseEvent BalloonClicked(Me, EventArgs.Empty)
            Me.Close()
        End If
    End Sub

    ' Prevent focus stealing on Show() - the host app keeps focus.
    Protected Overrides ReadOnly Property ShowWithoutActivation As Boolean
        Get
            Return True
        End Get
    End Property

    ' WndProc override: ensure we don't activate on mouse down.
    ' WM_MOUSEACTIVATE = &H21
    ' MA_NOACTIVATE = 3
    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = &H21 Then  ' WM_MOUSEACTIVATE
            m.Result = New IntPtr(3)  ' MA_NOACTIVATE
            Return
        End If
        MyBase.WndProc(m)
    End Sub

    Public ReadOnly Property CloseReasonValue As String
        Get
            Return _closeReason
        End Get
    End Property

    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.FadeTimer = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'FadeTimer
        '
        Me.FadeTimer.Interval = 5
        '
        'BalloonPopup
        '
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Name = "BalloonPopup"
        Me.ResumeLayout(False)

    End Sub
End Class
