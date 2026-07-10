Imports System.Runtime.InteropServices
Imports System.Text

Public Class TextStylizer
    Private Const HOTKEY_BOLD_ID As Integer = 1
    Private Const HOTKEY_ITALIC_ID As Integer = 2
    Private Const HOTKEY_UNDERLINE_ID As Integer = 3
    Private Const HOTKEY_STRIKEOUT_ID As Integer = 4

    Private Const WM_GETTEXT As Integer = &HD
    Private Const WM_GETTEXTLENGTH As Integer = &HE
    Private Const EM_GETSEL As Integer = &HB0
    Private Const EM_SETSEL As Integer = &HB1
    Private Const EM_REPLACESEL As Integer = &HC2

    Private Const COMBINING_STRIKEOUT As Integer = &H336
    Private Const COMBINING_UNDERLINE As Integer = &H332

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function RegisterHotKey(hWnd As IntPtr, id As Integer, fsModifiers As Integer, vk As Integer) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function UnregisterHotKey(hWnd As IntPtr, id As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, lpdwProcessId As IntPtr) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function AttachThreadInput(idAttach As Integer, idAttachTo As Integer, fAttach As Boolean) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetFocus() As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="SendMessageW", CharSet:=CharSet.Unicode)>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As Integer, lParam As Integer) As Integer
    End Function

    <DllImport("user32.dll", EntryPoint:="SendMessageW", CharSet:=CharSet.Unicode)>
    Private Shared Function SendMessageGetText(hWnd As IntPtr, Msg As Integer, wParam As Integer, lParam As StringBuilder) As Integer
    End Function

    <DllImport("user32.dll", EntryPoint:="SendMessageW", CharSet:=CharSet.Unicode)>
    Private Shared Function SendMessageReplaceSel(hWnd As IntPtr, Msg As Integer, wParam As Integer, lParam As String) As Integer
    End Function

    <DllImport("kernel32.dll")>
    Private Shared Function GetCurrentThreadId() As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetGUIThreadInfo(idThread As Integer, ByRef lpgti As GUITHREADINFO) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Private Structure GUITHREADINFO
        Public cbSize As Integer
        Public flags As Integer
        Public hwndActive As IntPtr
        Public hwndFocus As IntPtr
        Public hwndCapture As IntPtr
        Public hwndMenuOwner As IntPtr
        Public hwndMoveSize As IntPtr
        Public hwndCaret As IntPtr
        Public rcCaret_Left As Integer
        Public rcCaret_Top As Integer
        Public rcCaret_Right As Integer
        Public rcCaret_Bottom As Integer
    End Structure

    Public Shared Sub ProcessWin32SelectedText(styleId As Integer, targetHWnd As IntPtr)
        Dim foregroundHWnd As IntPtr = targetHWnd
        If foregroundHWnd = IntPtr.Zero Then Return

        Dim remoteThreadId As Integer = GetWindowThreadProcessId(foregroundHWnd, IntPtr.Zero)

        Dim gti As New GUITHREADINFO()
        gti.cbSize = Marshal.SizeOf(gti)

        Dim focusedControlHWnd As IntPtr = IntPtr.Zero

        If GetGUIThreadInfo(remoteThreadId, gti) Then
            focusedControlHWnd = gti.hwndFocus
        End If

        If focusedControlHWnd = IntPtr.Zero Then
            focusedControlHWnd = foregroundHWnd
        End If

        Dim sel As Integer = SendMessage(focusedControlHWnd, EM_GETSEL, 0, 0)
        Dim startPos As Integer = sel And &HFFFF
        Dim endPos As Integer = (sel >> 16) And &HFFFF

        If startPos = 0 AndAlso endPos = 0 AndAlso focusedControlHWnd <> foregroundHWnd Then
            focusedControlHWnd = foregroundHWnd
            sel = SendMessage(focusedControlHWnd, EM_GETSEL, 0, 0)
            startPos = sel And &HFFFF
            endPos = (sel >> 16) And &HFFFF
        End If

        If startPos = endPos Then Return

        Dim textLength As Integer = SendMessage(focusedControlHWnd, WM_GETTEXTLENGTH, 0, 0)
        Dim buffer As New StringBuilder(textLength + 1)
        SendMessageGetText(focusedControlHWnd, WM_GETTEXT, buffer.Capacity, buffer)

        Dim fullText As String = buffer.ToString()
        If startPos < 0 OrElse endPos > fullText.Length OrElse startPos > endPos Then Return

        Dim selectedText As String = fullText.Substring(startPos, endPos - startPos)
        Dim transformedText As String = TransformUnicodeText(selectedText, styleId)

        SendMessageReplaceSel(focusedControlHWnd, EM_REPLACESEL, 1, transformedText)

        Dim newEndPos As Integer = startPos + transformedText.Length
        SendMessage(focusedControlHWnd, EM_SETSEL, startPos, newEndPos)
    End Sub

    Private Shared Function TransformUnicodeText(input As String, actionId As Integer) As String
        Dim sb As New StringBuilder()
        Dim i As Integer = 0

        While i < input.Length
            Dim codePoint As Integer = 0
            Dim isSurrogate As Boolean = False

            If Char.IsHighSurrogate(input(i)) AndAlso i + 1 < input.Length AndAlso Char.IsLowSurrogate(input(i + 1)) Then
                codePoint = Char.ConvertToUtf32(input(i), input(i + 1))
                isSurrogate = True
                i += 2
            Else
                codePoint = Convert.ToInt32(input(i))
                i += 1
            End If

            Dim currentStyle As Integer = 0
            Dim baseChar As Integer = GetBaseCharAndStyle(codePoint, currentStyle)

            Dim hasUnderline As Boolean = False
            Dim hasStrikeout As Boolean = False

            While i < input.Length
                Dim nextChCode As Integer = Convert.ToInt32(input(i))
                If nextChCode = COMBINING_UNDERLINE Then
                    hasUnderline = True
                    i += 1
                ElseIf nextChCode = COMBINING_STRIKEOUT Then
                    hasStrikeout = True
                    i += 1
                Else
                    Exit While
                End If
            End While

            Dim newStyle As Integer = currentStyle
            Dim nextUnderlineState As Boolean = hasUnderline
            Dim nextStrikeoutState As Boolean = hasStrikeout

            Select Case actionId
                Case HOTKEY_BOLD_ID
                    If currentStyle = 1 Then newStyle = 0
                    If currentStyle = 3 Then newStyle = 2
                    If currentStyle = 0 Then newStyle = 1
                    If currentStyle = 2 Then newStyle = 3
                Case HOTKEY_ITALIC_ID
                    If currentStyle = 2 Then newStyle = 0
                    If currentStyle = 3 Then newStyle = 1
                    If currentStyle = 0 Then newStyle = 2
                    If currentStyle = 1 Then newStyle = 3
                Case HOTKEY_UNDERLINE_ID
                    nextUnderlineState = Not hasUnderline
                Case HOTKEY_STRIKEOUT_ID
                    nextStrikeoutState = Not hasStrikeout
            End Select

            If baseChar >= 65 AndAlso baseChar <= 90 Then
                Select Case newStyle
                    Case 0 : sb.Append(Convert.ToChar(baseChar))
                    Case 1 : sb.Append(Char.ConvertFromUtf32(&H1D5D4 + (baseChar - 65)))
                    Case 2 : sb.Append(Char.ConvertFromUtf32(&H1D608 + (baseChar - 65)))
                    Case 3 : sb.Append(Char.ConvertFromUtf32(&H1D63C + (baseChar - 65)))
                End Select
            ElseIf baseChar >= 97 AndAlso baseChar <= 122 Then
                Select Case newStyle
                    Case 0 : sb.Append(Convert.ToChar(baseChar))
                    Case 1 : sb.Append(Char.ConvertFromUtf32(&H1D5EE + (baseChar - 97)))
                    Case 2 : sb.Append(Char.ConvertFromUtf32(&H1D622 + (baseChar - 97)))
                    Case 3 : sb.Append(Char.ConvertFromUtf32(&H1D656 + (baseChar - 97)))
                End Select
            Else
                If isSurrogate Then
                    sb.Append(Char.ConvertFromUtf32(codePoint))
                Else
                    sb.Append(Convert.ToChar(codePoint))
                End If
            End If

            If nextUnderlineState Then
                sb.Append(Char.ConvertFromUtf32(COMBINING_UNDERLINE))
            End If
            If nextStrikeoutState Then
                sb.Append(Char.ConvertFromUtf32(COMBINING_STRIKEOUT))
            End If
        End While

        Return sb.ToString()
    End Function

    Private Shared Function GetBaseCharAndStyle(codePoint As Integer, ByRef style As Integer) As Integer
        ' Sans Bold A-Z (&H1D5D4 - &H1D5ED)
        If codePoint >= &H1D5D4 AndAlso codePoint <= &H1D5ED Then : style = 1 : Return codePoint - &H1D5D4 + 65 : End If
        ' Sans Bold a-z (&H1D5EE - &H1D607)
        If codePoint >= &H1D5EE AndAlso codePoint <= &H1D607 Then : style = 1 : Return codePoint - &H1D5EE + 97 : End If

        ' Sans Italic A-Z (&H1D608 - &H1D621)
        If codePoint >= &H1D608 AndAlso codePoint <= &H1D621 Then : style = 2 : Return codePoint - &H1D608 + 65 : End If
        ' Sans Italic a-z (&H1D622 - &H1D63B)
        If codePoint >= &H1D622 AndAlso codePoint <= &H1D63B Then : style = 2 : Return codePoint - &H1D622 + 97 : End If

        ' Sans Bold Italic A-Z (&H1D63C - &H1D655)
        If codePoint >= &H1D63C AndAlso codePoint <= &H1D655 Then : style = 3 : Return codePoint - &H1D63C + 65 : End If
        ' Sans Bold Italic a-z (&H1D656 - &H1D66F)
        If codePoint >= &H1D656 AndAlso codePoint <= &H1D66F Then : style = 3 : Return codePoint - &H1D656 + 97 : End If

        ' Normal text
        style = 0
        Return codePoint
    End Function

    Private Function ConvertToUnicodeBold(input As String) As String
        Dim sb As New StringBuilder()
        For Each ch As Char In input
            Dim code As Integer = Convert.ToInt32(ch)
            If code >= 65 AndAlso code <= 90 Then
                sb.Append(Char.ConvertFromUtf32(&H1D400 + (code - 65)))
            ElseIf code >= 97 AndAlso code <= 122 Then
                sb.Append(Char.ConvertFromUtf32(&H1D41A + (code - 97)))
            Else
                sb.Append(ch)
            End If
        Next
        Return sb.ToString()
    End Function

    Private Function ConvertToUnicodeItalic(input As String) As String
        Dim sb As New StringBuilder()
        For Each ch As Char In input
            Dim code As Integer = Convert.ToInt32(ch)
            If code >= 65 AndAlso code <= 90 Then
                sb.Append(Char.ConvertFromUtf32(&H1D434 + (code - 65)))
            ElseIf code >= 97 AndAlso code <= 122 Then
                sb.Append(Char.ConvertFromUtf32(&H1D44E + (code - 97)))
            Else
                sb.Append(ch)
            End If
        Next
        Return sb.ToString()
    End Function

    Private Function ConvertToUnicodeUnderline(input As String) As String
        Dim sb As New StringBuilder()
        For Each ch As Char In input
            sb.Append(ch)
            sb.Append(Char.ConvertFromUtf32(&H332))
        Next
        Return sb.ToString()
    End Function
End Class
