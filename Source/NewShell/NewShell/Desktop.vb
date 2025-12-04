Imports System.IO
Imports System.Linq
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports System.Windows.Forms.Menu
Imports Microsoft.Win32
Imports Shell32

Public Class Desktop
    Private Const REG_KEY_PATH As String = "DesktopBackground\Shell"

    <DllImport("shell32.dll", SetLastError:=True)>
    Private Shared Function SHGetStockIconInfo(
    siid As SHSTOCKICONID,
    uFlags As UInteger,
    ByRef psii As SHSTOCKICONINFO
) As Integer
    End Function

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Private Structure SHSTOCKICONINFO
        Public cbSize As UInteger
        Public hIcon As IntPtr
        Public iSysImageIndex As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public szPath As String
    End Structure

    Private Enum SHSTOCKICONID As Integer
        SIID_SHIELD = 77
        SIID_PROPERTIES = 124
        SIID_COPY = 236
        SIID_CUT = 233
        SIID_PASTE = 237
        SIID_DELETE = 234
        SIID_RENAME = 305
    End Enum

    Private Const SHGSI_ICON As UInteger = &H100
    Private Const SHGSI_SMALLICON As UInteger = &H1

    <DllImport("user32.dll")>
    Private Shared Function CallWindowProc(lpPrevWndFunc As IntPtr, hWnd As IntPtr, Msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    Private Const WM_MOUSEACTIVATE As Integer = &H21
    Private Const MA_NOACTIVATEANDEAT As Integer = &H4

    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = WM_MOUSEACTIVATE Then
            m.Result = CType(MA_NOACTIVATEANDEAT, IntPtr)
            Return
        End If
        MyBase.WndProc(m)
    End Sub

    '  Alt+Tab hotkey Disabling/Enabling

    Private Const WH_KEYBOARD_LL As Integer = 13
    Private Const WM_KEYDOWN As Integer = &H100
    Private Const WM_SYSKEYDOWN As Integer = &H104

    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Integer, ByVal lpfn As KeyboardHookProc, ByVal hMod As IntPtr, ByVal dwThreadId As Integer) As IntPtr
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As IntPtr) As Boolean
    Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As IntPtr, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As IntPtr
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Short

    Private hHook As IntPtr = IntPtr.Zero
    Private keyboardHookDelegate As KeyboardHookProc
    Private Delegate Function KeyboardHookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr

    Private Structure KBDLLHOOKSTRUCT
        Public vkCode As Integer
        Public scanCode As Integer
        Public flags As Integer
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure

    Public Function LowLevelKeyboardProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        If nCode >= 0 AndAlso (wParam = WM_KEYDOWN OrElse wParam = WM_SYSKEYDOWN) Then
            Dim keyboardStruct As KBDLLHOOKSTRUCT = CType(Marshal.PtrToStructure(lParam, GetType(KBDLLHOOKSTRUCT)), KBDLLHOOKSTRUCT)

            If keyboardStruct.vkCode = Keys.Tab AndAlso (GetAsyncKeyState(Keys.LMenu) OrElse GetAsyncKeyState(Keys.RMenu)) Then
                Return New IntPtr(1)
            End If
        End If

        Return CallNextHookEx(hHook, nCode, wParam, lParam)
    End Function

    <DllImport("shlwapi.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function SHLoadIndirectString(
        ByVal pszSource As String,
        ByVal pszOutBuf As StringBuilder,
        ByVal cchOutBuf As UInteger,
        ByVal ppvReserved As IntPtr) As Integer
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function ExtractIconEx(
    ByVal lpszFile As String,
    ByVal nIconIndex As Integer,
    ByVal phiconLarge() As IntPtr,
    ByVal phiconSmall() As IntPtr,
    ByVal nIcons As UInteger
) As UInteger
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function DestroyIcon(ByVal hIcon As IntPtr) As Boolean
    End Function

    Public Shared Function GetIcon(ByVal filePath As String, ByVal iconIndex As Integer, ByVal isSmallIcon As Boolean) As Icon
        Dim hIconLarge(0) As IntPtr
        Dim hIconSmall(0) As IntPtr
        Dim iconToReturn As Icon = Nothing

        Try
            Dim result As UInteger = ExtractIconEx(filePath, iconIndex, hIconLarge, hIconSmall, 1)

            If result > 0 Then
                Dim hIcon As IntPtr = If(isSmallIcon, hIconSmall(0), hIconLarge(0))

                If hIcon <> IntPtr.Zero Then
                    iconToReturn = Icon.FromHandle(hIcon).Clone()

                    DestroyIcon(hIcon)
                End If
            End If

        Catch ex As Exception
            Return Nothing
        End Try

        Return iconToReturn
    End Function

    Private Const CLSID_THISPC As String = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    Private Const CLSID_RECYCLEBIN As String = "::{645FF040-5081-101B-9F08-00AA002F954E}"
    Private Const CLSID_USERFILES As String = "::{59031a47-3f72-44a7-89c5-5595fe6b30ee}"
    Private Const CLSID_CONTROL_PANEL As String = "::{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}"
    Private Const CLSID_NETWORK As String = "::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}"

    Private Function IsShellIconVisible(clsidGUID As String) As Boolean
        Dim cleanGUID As String = clsidGUID.Replace("::", "")
        Dim keyPath As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"

        Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(keyPath)
            If key IsNot Nothing Then
                Dim val As Object = key.GetValue(cleanGUID)
                If val IsNot Nothing Then
                    Return CInt(val) = 0
                End If
            End If
        End Using

        If clsidGUID = CLSID_RECYCLEBIN Then Return True
        Return False
    End Function

    Private Function GetShellItemName(clsid As String) As String
        Try
            Dim id As Guid = New Guid(clsid.Replace("::", ""))
            Dim type As Type = Type.GetTypeFromCLSID(id)

            Select Case clsid
                Case CLSID_THISPC : Return "This PC"
                Case CLSID_RECYCLEBIN : Return "Recycle Bin"
                Case CLSID_USERFILES : Return Environment.UserName
                Case CLSID_NETWORK : Return "Network"
                Case CLSID_CONTROL_PANEL : Return "Control Panel"
                Case Else : Return "Unknown Item"
            End Select
        Catch ex As Exception
            Return "System Item"
        End Try
    End Function

    Private Sub Desktop_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.DoubleBuffered = True
        PictureBox1.AllowDrop = True

        BackUpdate()
        LoadDesktopIcons()

        FSW.Path = DesktopDir
        PFSW.Path = DesktopDirPublic

        Dim ShellIconsKey As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
        If Registry.CurrentUser.OpenSubKey(ShellIconsKey) Is Nothing Then
            Registry.CurrentUser.CreateSubKey(ShellIconsKey)

            Registry.SetValue(ShellIconsKey, CLSID_RECYCLEBIN, "0")
        End If

        ' Custom Alt+Tab

        Try

            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\AltTab", "Enabled", "0") = 1 Then
                keyboardHookDelegate = New KeyboardHookProc(AddressOf LowLevelKeyboardProc)

                Dim appModule As IntPtr = GetModuleHandle(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name & ".exe")
                hHook = SetWindowsHookEx(WH_KEYBOARD_LL, keyboardHookDelegate, appModule, 0)

            End If

        Catch ex As Exception
            hHook = IntPtr.Zero
            MessageBox.Show("Failed to implement Custom Alt+Tab hook. So we'll switching it back to Windows' Default.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public CanClose As Boolean = False
    Private Sub Desktop_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.TaskManagerClosing OrElse e.CloseReason = CloseReason.WindowsShutDown Then
            e.Cancel = False
            If GlobalKeyboardHook.Unhook() Then
                Debug.WriteLine("Hook ended.")
            Else
                Debug.WriteLine("Hook ending failed.")
            End If
        Else
            If CanClose = False Then
                e.Cancel = True
                SA.ShowDialog(AppBar)
            Else
                If GlobalKeyboardHook.Unhook() Then
                    Debug.WriteLine("Hook ended.")
                Else
                    Debug.WriteLine("Hook ending failed.")
                End If
            End If
        End If
    End Sub

    <StructLayout(LayoutKind.Sequential)>
    Private Structure SHFILEINFO
        Public hIcon As IntPtr
        Public iIcon As Integer
        Public dwAttributes As UInteger
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public szDisplayName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)>
        Public szTypeName As String
    End Structure

    Private Const SHGFI_ICON As UInteger = &H100
    Private Const SHGFI_SMALLICON As UInteger = &H1
    Private Const SHGFI_LARGEICON As UInteger = &H0
    Private Const SHGFI_USEFILEATTRIBUTES As UInteger = &H10
    Private Const SHGFI_DISPLAYNAME As UInteger = &H200
    Private Const SHGFI_PIDL As UInteger = &H8

    Private Const FILE_ATTRIBUTE_NORMAL As UInteger = &H80
    Private Const FILE_ATTRIBUTE_DIRECTORY As UInteger = &H10

    <DllImport("shell32.dll")>
    Private Shared Function SHGetFileInfo(
        ByVal pszPath As String,
        ByVal dwFileAttributes As UInteger,
        ByRef psfi As SHFILEINFO,
        ByVal cbFileInfo As UInteger,
        ByVal uFlags As UInteger) As IntPtr
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Unicode, ExactSpelling:=True, PreserveSig:=False)>
    Private Shared Function SHParseDisplayName(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszName As String, ByVal pbc As IntPtr, ByRef ppidl As IntPtr, ByVal sfgaoIn As UInteger, ByRef psfgaoOut As UInteger) As Integer
    End Function

    <DllImport("shell32.dll")>
    Private Shared Sub ILFree(ByVal pidl As IntPtr)
    End Sub

    Public Function GetFileIcon(ByVal filePath As String, ByVal isLargeIcon As Boolean) As Icon
        Dim shInfo As SHFILEINFO = New SHFILEINFO()
        Dim flags As UInteger = SHGFI_ICON

        If isLargeIcon Then
            flags = flags Or SHGFI_LARGEICON
        Else
            flags = flags Or SHGFI_SMALLICON
        End If

        Dim attributes As UInteger = FILE_ATTRIBUTE_NORMAL

        If String.IsNullOrEmpty(filePath) OrElse System.IO.Directory.Exists(filePath) Then
            If Not String.IsNullOrEmpty(filePath) AndAlso System.IO.Directory.Exists(filePath) Then
                attributes = FILE_ATTRIBUTE_DIRECTORY
            ElseIf String.IsNullOrEmpty(filePath) Then
                flags = flags Or SHGFI_USEFILEATTRIBUTES
                attributes = FILE_ATTRIBUTE_DIRECTORY
            End If
        End If

        Dim result As IntPtr = SHGetFileInfo(
            filePath,
            attributes,
            shInfo,
            CUInt(Marshal.SizeOf(shInfo)),
            flags)

        If Not result.Equals(IntPtr.Zero) AndAlso Not shInfo.hIcon.Equals(IntPtr.Zero) Then
            Try
                Dim iconResult As Icon = Icon.FromHandle(shInfo.hIcon)
                Dim finalIcon As Icon = CType(iconResult.Clone(), Icon)

                DestroyIcon(shInfo.hIcon)

                Return finalIcon
            Catch ex As Exception
                Return Nothing
            End Try
        End If

        Return Nothing
    End Function

    Public Function GetFolderIcon(ByVal folderPath As String, ByVal isLargeIcon As Boolean) As Icon
        If Not System.IO.Directory.Exists(folderPath) Then
            Return Nothing
        End If

        Dim shInfo As SHFILEINFO = New SHFILEINFO()

        Dim flags As UInteger = SHGFI_ICON

        If isLargeIcon Then
            flags = flags Or SHGFI_LARGEICON
        Else
            flags = flags Or SHGFI_SMALLICON
        End If

        Dim result As IntPtr = SHGetFileInfo(
            folderPath,
            FILE_ATTRIBUTE_DIRECTORY,
            shInfo,
            CUInt(Marshal.SizeOf(shInfo)),
            flags)

        If Not result.Equals(IntPtr.Zero) AndAlso Not shInfo.hIcon.Equals(IntPtr.Zero) Then
            Try
                Dim iconResult As Icon = Icon.FromHandle(shInfo.hIcon)
                Dim finalIcon As Icon = CType(iconResult.Clone(), Icon)

                DestroyIcon(shInfo.hIcon)

                Return finalIcon
            Catch ex As Exception
                Return Nothing
            End Try
        End If

        Return Nothing
    End Function

    Public Structure ShellItemInfo
        Public Icon As Bitmap
        Public DisplayName As String
    End Structure

    Public Shared Function GetShellItemInfo(shellPath As String) As ShellItemInfo
        Dim info As New ShellItemInfo()
        Dim pidl As IntPtr = IntPtr.Zero

        Try
            SHParseDisplayName(shellPath, IntPtr.Zero, pidl, 0, 0)

            If pidl <> IntPtr.Zero Then
                Dim shinfo As New SHFILEINFO()

                SHGetFileInfo(pidl, 0, shinfo, CUInt(Marshal.SizeOf(shinfo)), SHGFI_ICON Or SHGFI_LARGEICON Or SHGFI_DISPLAYNAME Or SHGFI_PIDL)

                info.DisplayName = shinfo.szDisplayName

                If shinfo.hIcon <> IntPtr.Zero Then

                    Using ico As Icon = Icon.FromHandle(shinfo.hIcon)
                        info.Icon = ico.ToBitmap()
                    End Using

                    DestroyIcon(shinfo.hIcon)
                End If
            End If
        Catch ex As Exception
            info.DisplayName = "Unknown Item"
            info.Icon = SystemIcons.Error.ToBitmap()
        Finally
            If pidl <> IntPtr.Zero Then ILFree(pidl)
        End Try

        Return info
    End Function
    Private currentDraggedButton As Button = Nothing
    Private mouseOffset As Point
    Private wasDragged As Boolean = False
    Private isDraggingOut As Boolean = False
    Private Sub Button_MouseMove(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.None Then

            currentDraggedButton = CType(sender, Button)
            mouseOffset = New Point(e.X, e.Y)
            currentDraggedButton.BringToFront()

            wasDragged = False

            isDraggingOut = False

        ElseIf e.Button = MouseButtons.Middle Then

            currentDraggedButton.Capture = True

            If currentDraggedButton IsNot Nothing Then
                Dim screenPosition As Point = Cursor.Position
                Dim parentLocation As Point = currentDraggedButton.Parent.PointToClient(screenPosition)

                currentDraggedButton.Location = New Point(parentLocation.X - mouseOffset.X, parentLocation.Y - mouseOffset.Y)
                wasDragged = True
            End If
        ElseIf e.Button = MouseButtons.Left Then
            If currentDraggedButton IsNot Nothing Then
                Dim screenPosition As Point = Cursor.Position
                Dim parentLocation As Point = currentDraggedButton.Parent.PointToClient(screenPosition)

                If e.Button = MouseButtons.Left AndAlso (Math.Abs(e.X) > 5 OrElse Math.Abs(e.Y) > 5) Then

                    Dim filePath As String = CStr(currentDraggedButton.Tag)

                    If String.IsNullOrEmpty(filePath) OrElse filePath.StartsWith("::{") Then
                        isDraggingOut = False
                    Else
                        ' Drag and Drop is for me an Advanced code, even with Gemini.
                        ' So I'll implement 2 options.

                        ' 1. It will be normal moving the icons on Middle button
                        ' 2. Drag&Drop will execute if you use Left click

                        ' For those who are reading this and you manage to fix it
                        ' to just use Left Click on the Drag&Drop and changing the
                        ' size, please let me know!

                        isDraggingOut = True

                        Dim fileList As New System.Collections.Specialized.StringCollection()
                        fileList.Add(filePath)

                        Dim data As New DataObject()
                        data.SetFileDropList(fileList)

                        Dim dropResult As DragDropEffects = currentDraggedButton.DoDragDrop(data, DragDropEffects.Move)

                        If dropResult = DragDropEffects.Move Then
                            If MovedIconPositions.ContainsKey(filePath) Then
                                MovedIconPositions.Remove(filePath)
                            End If
                        End If

                        currentDraggedButton.Capture = False
                        currentDraggedButton = Nothing
                        isDraggingOut = False

                        LoadDesktopIcons()
                        Return
                    End If
                End If

                If Not isDraggingOut Then
                    currentDraggedButton.Capture = True
                    currentDraggedButton.Location = New Point(parentLocation.X - mouseOffset.X, parentLocation.Y - mouseOffset.Y)
                    wasDragged = True
                End If
            End If
        End If
    End Sub
    Dim selectedPaths As New List(Of String)
    Private Sub Button_MouseUp(sender As Object, e As MouseEventArgs)
        DesktopCM.Close()
        If wasDragged Then
            If currentDraggedButton IsNot Nothing Then
                currentDraggedButton.Capture = False

                If MovedIconPositions.ContainsKey(currentDraggedButton.Tag.ToString()) Then
                    MovedIconPositions(currentDraggedButton.Tag.ToString()) = currentDraggedButton.Location
                Else
                    MovedIconPositions.Add(currentDraggedButton.Tag.ToString(), currentDraggedButton.Location)
                End If

                currentDraggedButton = Nothing
            End If

            wasDragged = False
            Return
        Else
            If e.Button = MouseButtons.Left Then
                If CType(sender, Button).Tag.StartsWith("::{") Then
                    Process.Start("explorer.exe", CType(sender, Button).Tag)
                Else
                    If File.Exists(CType(sender, Button).Tag) Then
                        Process.Start(CType(sender, Button).Tag)
                    ElseIf Directory.Exists(CType(sender, Button).Tag) Then
                        Process.Start("explorer.exe", CType(sender, Button).Tag)
                    End If
                End If
            ElseIf e.Button = MouseButtons.Right Then
                ShowShellContextMenu(CType(sender, Button).Tag, MousePosition)
                renamingButton = CType(sender, Button)
            End If
        End If
    End Sub
    Private MovedIconPositions As New Dictionary(Of String, Point)()
    Public Sub LoadDesktopIcons()
        Dim ICON_WIDTH As Integer = SystemInformation.IconSpacingSize.Width
        Dim ICON_HEIGHT As Integer = SystemInformation.IconSpacingSize.Height
        Dim SPACE As Integer = 5 'SystemInformation.IconSize.Height

        Dim targetPath As String = DesktopDir
        Dim targetPathPublic As String = DesktopDirPublic

        If Not Directory.Exists(targetPath) Then
            MessageBox.Show("Failed to load Desktop Items! Path to your desktop doesn't exist.", "Error")
            Return
        End If

        PictureBox1.Controls.Clear()

        Dim Items As New List(Of String)()

        If IsShellIconVisible(CLSID_THISPC) Then Items.Add(CLSID_THISPC)
        If IsShellIconVisible(CLSID_USERFILES) Then Items.Add(CLSID_USERFILES)
        If IsShellIconVisible(CLSID_NETWORK) Then Items.Add(CLSID_NETWORK)
        If IsShellIconVisible(CLSID_RECYCLEBIN) Then Items.Add(CLSID_RECYCLEBIN)
        If IsShellIconVisible(CLSID_CONTROL_PANEL) Then Items.Add(CLSID_CONTROL_PANEL)

        Dim directories() As String = Directory.GetDirectories(targetPath)
        Dim files() As String = Directory.GetFiles(targetPath)

        Dim publicdirectories() As String = Directory.GetDirectories(targetPathPublic)
        Dim publicfiles() As String = Directory.GetFiles(targetPathPublic)

        'For Folders and Files all together:
        'Dim items As String() = Directory.GetFileSystemEntries(targetPath)

        Items.AddRange(directories)
        Items.AddRange(publicdirectories)

        Items.AddRange(files)
        Items.AddRange(publicfiles)

        Dim OccupiedGridPoints As New Dictionary(Of String, Boolean)()

        Dim maxHeight As Integer = PictureBox1.Height

        Dim columnItemCount As Integer = Math.Floor(maxHeight / (ICON_HEIGHT + SPACE))

        If columnItemCount <= 0 Then columnItemCount = 1

        For i As Integer = 0 To Items.Count - 1
            Dim isShellItem As Boolean = Items(i).StartsWith("::{")
            Dim isDirectory As Boolean = Directory.Exists(Items(i))
            Dim isFile As Boolean = File.Exists(Items(i))

            If isFile = False AndAlso isDirectory = False AndAlso isShellItem = False Then Continue For

            Dim X As Integer
            Dim Y As Integer
            Dim itemPath As String = Items(i)

            If MovedIconPositions.ContainsKey(itemPath) Then
                Dim savedLocation As Point = MovedIconPositions(itemPath)
                X = savedLocation.X
                Y = savedLocation.Y
            Else
                Dim newPos As Point = FindNextFreePosition(OccupiedGridPoints, columnItemCount, ICON_WIDTH, ICON_HEIGHT, SPACE, Items.Count)
                X = newPos.X
                Y = newPos.Y

                'Dim line As Integer = i Mod columnItemCount
                'Dim column As Integer = i \ columnItemCount

                'X = (column * (ICON_WIDTH + SPACE))
                'Y = 5 + (line * (ICON_HEIGHT + SPACE))
            End If

            Dim DesktopItem As New ShadowButton()

            DesktopItem.Parent = PictureBox1

            DesktopItem.Size = New Size(ICON_WIDTH, ICON_HEIGHT)
            DesktopItem.Location = New Point(X, Y)

            If isFile = True Then
                Dim FI As New FileInfo(Items(i))

                DesktopItem.Image = GetFileIcon(Items(i), True).ToBitmap

                If FI.Extension.ToLower = ".lnk" Then
                    DesktopItem.Text = Path.GetFileNameWithoutExtension(FI.FullName)
                Else
                    DesktopItem.Text = FI.Name
                End If

                DesktopItem.Tag = Items(i)

                If DesktopItem.Image.Size.Width > DesktopItem.Size.Width AndAlso DesktopItem.Image.Size.Height > DesktopItem.Size.Height Then
                    DesktopItem.BackgroundImage = DesktopItem.Image
                    DesktopItem.Image = Nothing
                End If
            ElseIf isDirectory = True Then
                Dim DI As New DirectoryInfo(Items(i))

                DesktopItem.Image = GetFolderIcon(Items(i), True).ToBitmap
                DesktopItem.Text = DI.Name
                DesktopItem.Tag = Items(i)

                If DesktopItem.Image.Size.Width > DesktopItem.Size.Width AndAlso DesktopItem.Image.Size.Height > DesktopItem.Size.Height Then
                    DesktopItem.BackgroundImage = DesktopItem.Image
                    DesktopItem.Image = Nothing
                End If
            ElseIf isShellItem Then
                Dim info As ShellHelpers.ShellItemInfo = ShellHelpers.GetShellItemInfo(itemPath)

                DesktopItem.Text = info.DisplayName

                If info.Icon IsNot Nothing Then
                    DesktopItem.Image = info.Icon

                    If DesktopItem.Image.Size.Width > DesktopItem.Size.Width Then
                        DesktopItem.BackgroundImage = DesktopItem.Image
                        DesktopItem.Image = Nothing
                    End If
                End If

                DesktopItem.Tag = Items(i)

            End If

            ' Events
            'AddHandler DesktopItem.MouseDown, AddressOf Button_MouseDown
            AddHandler DesktopItem.MouseMove, AddressOf Button_MouseMove
            AddHandler DesktopItem.MouseUp, AddressOf Button_MouseUp

            ' Appearance
            DesktopItem.TextAlign = ContentAlignment.BottomCenter
            DesktopItem.TextImageRelation = TextImageRelation.ImageAboveText
            DesktopItem.ForeColor = SystemColors.HighlightText
            DesktopItem.BackColor = Color.Transparent
            DesktopItem.FlatAppearance.BorderSize = 0
            DesktopItem.FlatStyle = FlatStyle.Flat

            'DesktopItem.Font = New Font(Me.Font, FontStyle.Bold)
            'DesktopItem.ForeColor = Color.Black

            PictureBox1.Controls.Add(DesktopItem)
        Next
    End Sub

    Private Function FindNextFreePosition(ByRef OccupiedPoints As Dictionary(Of String, Boolean), ByVal columnItemCount As Integer, ByVal iconWidth As Integer, ByVal iconHeight As Integer, ByVal space As Integer, ByVal maxIndex As Integer) As Point

        Dim X As Integer = 0
        Dim Y As Integer = 0

        For gridIndex As Integer = 0 To maxIndex

            Dim line As Integer = gridIndex Mod columnItemCount
            Dim column As Integer = gridIndex \ columnItemCount

            Dim checkX As Integer = (column * (iconWidth + space))
            Dim checkY As Integer = 5 + (line * (iconHeight + space))

            Dim gridKey As String = $"{checkX}:{checkY}"

            If Not OccupiedPoints.ContainsKey(gridKey) Then
                OccupiedPoints.Add(gridKey, True)
                Return New Point(checkX, checkY)
            End If
        Next

        Dim finalLine As Integer = maxIndex Mod columnItemCount
        Dim finalColumn As Integer = maxIndex \ columnItemCount

        X = (finalColumn * (iconWidth + space))
        Y = 5 + (finalLine * (iconHeight + space))

        Return New Point(X, Y)
    End Function

    Private WithEvents renameTextBox As TextBox = Nothing
    Private renamingButton As Button = Nothing

    Public Sub StartRenaming(btn As Button)
        If btn Is Nothing Then Return

        renamingButton = btn

        renamingButton.Enabled = False

        renameTextBox = New TextBox()
        renameTextBox.Parent = PictureBox1

        Dim textHeight As Integer = 40
        Dim boxWidth As Integer = btn.Width

        renameTextBox.Size = New Size(boxWidth, 20)
        renameTextBox.Location = New Point(btn.Left + (btn.Width - boxWidth) / 2, btn.Bottom - renameTextBox.Height)

        renameTextBox.Text = btn.Text
        renameTextBox.TextAlign = HorizontalAlignment.Center
        renameTextBox.BorderStyle = BorderStyle.FixedSingle
        renameTextBox.BringToFront()

        renameTextBox.Focus()
        renameTextBox.SelectAll()
    End Sub

    Private Sub renameTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles renameTextBox.KeyDown
        If e.KeyCode = Keys.Enter Then
            hasFinishedRenaming = True

            FinishRenaming(True)
            e.SuppressKeyPress = True
        ElseIf e.KeyCode = Keys.Escape Then
            hasFinishedRenaming = True
            FinishRenaming(False)
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub renameTextBox_LostFocus(sender As Object, e As EventArgs) Handles renameTextBox.LostFocus
        If hasFinishedRenaming Then
            hasFinishedRenaming = False
            Return
        End If

        FinishRenaming(True)
    End Sub
    Private hasFinishedRenaming As Boolean = False
    Private Sub FinishRenaming(saveChanges As Boolean)
        If renameTextBox Is Nothing OrElse renamingButton Is Nothing Then Return

        Dim newName As String = renameTextBox.Text.Trim()
        Dim originalName As String = renamingButton.Text
        Dim itemPath As String = CStr(renamingButton.Tag)

        renameTextBox.Dispose()
        renameTextBox = Nothing

        If saveChanges AndAlso newName <> originalName AndAlso Not String.IsNullOrEmpty(newName) Then
            Dim success As Boolean = RenameItem(itemPath, newName)

            If success Then
                renamingButton.Text = newName
            Else
                MessageBox.Show("Item renaming failed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                StartRenaming(renamingButton)
            End If
        End If

        renamingButton.Enabled = True

        renamingButton = Nothing
    End Sub

    Private Function RenameItem(currentPath As String, newName As String) As Boolean
        Try
            ' 1. SHELL ITEMS (::{GUID}...)
            If currentPath.StartsWith("::{") Then
                Dim clsid As String = currentPath.Replace("::", "")
                Return RenameShellItemInRegistry(clsid, newName)

                ' 2. FOLDERS
            ElseIf Directory.Exists(currentPath) Then
                Dim parentDir As String = Path.GetDirectoryName(currentPath)

                My.Computer.FileSystem.RenameDirectory(currentPath, newName)

                renamingButton.Tag = Path.Combine(parentDir, newName)
                Return True

                ' 3. FILES
            ElseIf File.Exists(currentPath) Then
                Dim parentDir As String = Path.GetDirectoryName(currentPath)

                My.Computer.FileSystem.RenameFile(currentPath, newName)

                renamingButton.Tag = Path.Combine(parentDir, newName)
                Return True
            End If

            Return False
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function RenameShellItemInRegistry(clsid As String, newName As String) As Boolean
        Try
            Dim keyPath As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\" & clsid

            Using key As RegistryKey = Registry.CurrentUser.CreateSubKey(keyPath)
                If key IsNot Nothing Then
                    key.SetValue("", newName)
                    Return True
                End If
            End Using
            Return False
        Catch ex As Exception
            Debug.WriteLine("Registry Write Error: " & ex.Message)
            Return False
        End Try
    End Function
    Dim DesktopDir As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
    Dim DesktopDirPublic As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory)
    Private Function IsCLSIDPath(ByVal path As String) As Boolean
        If String.IsNullOrWhiteSpace(path) Then Return False
        Return path.StartsWith("::{", StringComparison.OrdinalIgnoreCase) AndAlso path.EndsWith("}")
    End Function

    Public Sub ShowShellContextMenu(ByVal targetPath As String, ByVal displayPoint As Point)

        If String.IsNullOrWhiteSpace(targetPath) Then Return

        targetPath = targetPath.Trim()
        Dim isVirtual As Boolean = IsCLSIDPath(targetPath)

        If Not isVirtual AndAlso Not File.Exists(targetPath) AndAlso Not Directory.Exists(targetPath) Then
            MessageBox.Show("Path is invalid or doesn't exist: " & targetPath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim shell As Shell32.Shell = Nothing
        Dim folder As Shell32.Folder = Nothing
        Dim item As Shell32.FolderItem = Nothing

        Try
            shell = New Shell32.Shell()
            If shell Is Nothing Then Throw New Exception("Cannot iniciate Shell.Application (Shell32.Shell).")

            If isVirtual Then
                folder = shell.NameSpace(targetPath)
                If folder IsNot Nothing Then
                    item = folder.Self
                End If
            Else
                Dim directoryPath As String = Path.GetDirectoryName(targetPath)
                Dim fileName As String = Path.GetFileName(targetPath)

                If Directory.Exists(targetPath) AndAlso (String.IsNullOrEmpty(directoryPath) OrElse directoryPath = targetPath) Then
                    folder = shell.NameSpace(targetPath)
                    If folder IsNot Nothing Then
                        item = folder.Self
                    End If
                Else
                    folder = shell.NameSpace(directoryPath)
                    If folder IsNot Nothing Then
                        item = folder.ParseName(fileName)
                    End If
                End If
            End If

            If item Is Nothing Then item = shell.NameSpace(targetPath)
            If item Is Nothing Then Throw New Exception("Failed to get Shell FolderItem for path. Check COM Reference.")

            Dim verbsSource = item.Verbs()
            If verbsSource Is Nothing Then Return

            Dim cms As New ContextMenuStrip()
            Dim verbsCount As Integer = 0
            Dim groupBreakIndex As Integer = 3
            Dim bmpIcon As Image = Nothing

            For Each v As Object In verbsSource
                Dim verb As Shell32.FolderItemVerb = CType(v, Shell32.FolderItemVerb)
                Dim verbName As String = verb.Name

                Dim cleanDisplayName As String = verbName.Replace("&", "").Trim()

                If String.IsNullOrEmpty(cleanDisplayName) Then Continue For

                If verbsCount > 0 AndAlso verbsCount = groupBreakIndex Then
                    cms.Items.Add(New ToolStripSeparator())
                End If

                Dim tsmi As New ToolStripMenuItem(cleanDisplayName) With {.Tag = verb}

                If verbsCount = 0 Then
                    tsmi.Font = New Font(tsmi.Font, FontStyle.Bold)
                End If

                AddHandler tsmi.Click, Sub(s, ea)
                                           Try
                                               CType(DirectCast(s, ToolStripMenuItem).Tag, Shell32.FolderItemVerb).DoIt()
                                           Catch ex As Exception
                                               If ex.HResult = -2147467263 OrElse ex.Message.ToLower().Contains("implement") Then
                                                   StartRenaming(renamingButton)
                                               Else
                                                   MessageBox.Show("Failed to execute command: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                               End If
                                           End Try
                                       End Sub
                cms.Items.Add(tsmi)
                verbsCount += 1
            Next

            If cms.Items.Count > 0 Then
                cms.Show(displayPoint)
            End If

        Catch ex As Exception
            MessageBox.Show("Failed to make Shell Context Menu: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If Not item Is Nothing Then Marshal.ReleaseComObject(item)
            If Not folder Is Nothing Then Marshal.ReleaseComObject(folder)
            If Not shell Is Nothing Then Marshal.ReleaseComObject(shell)
        End Try
    End Sub

    Private Structure ContextMenuItem
        Public Name As String
        Public Text As String
        Public Position As String ' Top, Middle, Bottom
        Public IconPath As String
        Public IsSubMenu As Boolean
        Public IsSeparator As Boolean
        Public Command As String
    End Structure

    Private Sub LoadAndGenerateMenuItems(ByVal menuItemsCollection As ToolStripItemCollection, ByVal keyPath As String)
        Dim desktopShellKey As RegistryKey = Nothing

        Try
            desktopShellKey = Registry.ClassesRoot.OpenSubKey(keyPath)

            If desktopShellKey IsNot Nothing Then
                Dim subKeyNames() As String = desktopShellKey.GetSubKeyNames()
                Dim menuItems As New List(Of ContextMenuItem)()

                For Each subKeyName As String In subKeyNames
                    Dim subKey As RegistryKey = desktopShellKey.OpenSubKey(subKeyName)

                    If subKey IsNot Nothing Then
                        menuItems.Add(ReadRegistryKey(subKey, subKeyName))
                        subKey.Close()
                    End If
                Next

                GenerateMenu(menuItemsCollection, menuItems)

            End If
        Catch ex As Exception
            MessageBox.Show("Error while reading Registry: " & keyPath & vbCrLf & "Chyba: " & ex.Message, "Chyba registru", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If desktopShellKey IsNot Nothing Then desktopShellKey.Close()
        End Try
    End Sub

    Private Function ReadRegistryKey(ByVal subKey As RegistryKey, ByVal subKeyName As String) As ContextMenuItem
        Dim item As New ContextMenuItem()
        item.Name = subKeyName

        Dim textValue As Object = subKey.GetValue("MUIVerb")

        If textValue Is Nothing Then
            textValue = subKey.GetValue("")
        ElseIf textValue.ToString.Substring("0").Contains("@"c) Then

        End If

        item.Text = If(textValue IsNot Nothing, textValue.ToString(), subKeyName)

        Dim posValue As Object = subKey.GetValue("Position")
        item.Position = If(posValue IsNot Nothing, posValue.ToString(), String.Empty)

        Dim iconValue As Object = subKey.GetValue("Icon")
        item.IconPath = If(iconValue IsNot Nothing, iconValue.ToString(), String.Empty)

        item.IsSeparator = subKey.GetValueNames().Contains("Separator")

        item.IsSubMenu = subKey.GetValueNames().Contains("SubCommands")

        Dim commandKey As RegistryKey = subKey.OpenSubKey("command")
        If commandKey IsNot Nothing Then
            item.Command = If(commandKey.GetValue("") IsNot Nothing, commandKey.GetValue("").ToString(), String.Empty)
            commandKey.Close()
        End If

        Return item
    End Function

    Private Sub GenerateMenu(ByVal menuItemsCollection As ToolStripItemCollection, ByVal items As List(Of ContextMenuItem))

        Dim sortedItems = items.OrderBy(Function(i)
                                            Select Case i.Position.ToLower()
                                                Case "top" : Return -2
                                                Case "bottom" : Return 2
                                                Case Else : Return 0
                                            End Select
                                        End Function).ThenBy(Function(i) i.Text).ToList()

        Dim lastPosition As String = "top"

        Dim hasProcessedTopOrMiddle As Boolean = False

        For Each item As ContextMenuItem In sortedItems

            Dim currentGroupPosition As String
            Select Case item.Position.ToLower()
                Case "top" : currentGroupPosition = "top"
                Case "bottom" : currentGroupPosition = "bottom"
                Case Else : currentGroupPosition = "middle"
            End Select

            If (currentGroupPosition = "bottom" And (lastPosition = "top" Or lastPosition = "middle")) Then
                menuItemsCollection.Add(New ToolStripSeparator())
            End If

            If currentGroupPosition = "middle" Then
                lastPosition = "middle"
            ElseIf currentGroupPosition = "bottom" Then
                lastPosition = "bottom"
            End If

            If item.Text.Contains("@"c) = True Then
                Dim indirectString As String = item.Text.Trim()

                If String.IsNullOrEmpty(indirectString) Then
                    item.Text = "Invalid Text."
                    Continue For
                End If

                'Const MAX_PATH As UInteger = 260
                Const BUFFER_SIZE As UInteger = 1024
                Dim resultBuffer As New StringBuilder(CInt(BUFFER_SIZE))

                Dim hr As Integer = SHLoadIndirectString(
                    indirectString,
                    resultBuffer,
                    BUFFER_SIZE,
                    IntPtr.Zero)

                If hr = 0 Then
                    item.Text = resultBuffer.ToString()
                Else
                    item.Text = $"Error by loading text (HRESULT: {hr:X})"
                End If
            End If

            Dim menuItem As New ToolStripMenuItem(item.Text)

            If item.IconPath IsNot String.Empty Then
                Dim inputText As String = item.IconPath.Trim()

                If String.IsNullOrEmpty(inputText) Then
                    MessageBox.Show("Invalid Icon path to load for: " & item.Text, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                Dim filePath As String = ""
                Dim iconIndex As Integer = 0

                Dim parts() As String = inputText.Split(New Char() {","c}, StringSplitOptions.RemoveEmptyEntries)

                If parts.Length > 0 Then
                    filePath = parts(0).Trim()

                    filePath = Environment.ExpandEnvironmentVariables(filePath)

                    If parts.Length > 2 Then
                        MessageBox.Show("Invalid Icon path format for: """ & item.Text & """. It supposed to be like this: 'file.dll,-index'.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    ElseIf parts.Length = 2 Then
                        If Not Integer.TryParse(parts(1).Trim(), iconIndex) Then
                            MessageBox.Show("Icon index is not valid for: " & item.Text, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Return
                        End If
                    Else
                        iconIndex = 0
                    End If
                End If

                If iconIndex = -1 Then iconIndex = 0

                Dim extractedIcon As Icon = GetIcon(filePath, iconIndex, isSmallIcon:=True)

                If extractedIcon IsNot Nothing Then
                    menuItem.Image?.Dispose()
                    menuItem.Image = extractedIcon.ToBitmap()
                Else
                    menuItem.Image?.Dispose()
                    menuItem.Image = Nothing
                    Console.WriteLine($"Failed to get icon from: '{filePath}' with index of: {iconIndex}.", "Error while extracting", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If

            If item.IsSubMenu Then
                LoadAndGenerateMenuItems(menuItem.DropDownItems, REG_KEY_PATH & "\" & item.Name & "\Shell")
            Else
                menuItem.Tag = item.Command
                AddHandler menuItem.Click, AddressOf MenuItem_Click
            End If

            menuItemsCollection.Add(menuItem)

            If item.IsSeparator Then menuItemsCollection.Add(New ToolStripSeparator())
        Next
    End Sub

    Private Sub MenuItem_Click(sender As Object, e As EventArgs)
        Dim clickedItem As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim commandToExecute As String = clickedItem.Tag.ToString()

        If Not String.IsNullOrEmpty(commandToExecute) Then

            Dim inputString As String = commandToExecute
            Dim fileName As String = ""
            Dim arguments As String = ""

            Dim firstSpace As Integer = inputString.IndexOf(" ")

            If firstSpace = -1 Then
                fileName = inputString
                arguments = String.Empty
            Else
                fileName = inputString.Substring(0, firstSpace)
                arguments = inputString.Substring(firstSpace + 1)
            End If

            Try
                Dim startInfo As New ProcessStartInfo(fileName, arguments)

                Process.Start(startInfo)

            Catch ex As Exception
                MessageBox.Show($"Error while loading process: {ex.Message}", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DesktopCM.Opening
        DesktopCM.Items.Clear()

        DesktopCM.Items.Add(ViewToolStripMenuItem)
        DesktopCM.Items.Add(RefreshToolStripMenuItem)
        DesktopCM.Items.Add(New ToolStripSeparator())

        DesktopCM.Items.Add(PasteToolStripMenuItem)
        PasteToolStripMenuItem.Enabled = Clipboard.ContainsFileDropList

        DesktopCM.Items.Add(NewToolStripMenuItem)
        DesktopCM.Items.Add(New ToolStripSeparator())

        LoadAndGenerateMenuItems(DesktopCM.Items, REG_KEY_PATH)

        DesktopCM.Items.Add(New ToolStripSeparator())
        DesktopCM.Items.Add(DesktopOptionsToolStripMenuItem)
    End Sub

    Private Sub Desktop_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp, PictureBox1.MouseUp
        If e.Button = MouseButtons.Right Then
            'ShowShellContextMenu(DesktopDir, MousePosition)
            DesktopCM.Show(MousePosition)
        ElseIf e.Button = MouseButtons.Left Then
            DesktopCM.Close()

            selectRectangle = Rectangle.Empty
            PictureBox1.Invalidate()
        End If
    End Sub

    Private Sub PictureBox1_DragEnter(sender As Object, e As DragEventArgs) Handles PictureBox1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Move
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub PictureBox1_DragDrop(sender As Object, e As DragEventArgs) Handles PictureBox1.DragDrop
        Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())

        If files IsNot Nothing AndAlso files.Length > 0 Then

            Dim dropLocation As Point = PictureBox1.PointToClient(New Point(e.X - SystemInformation.IconSpacingSize.Width / 2, e.Y - SystemInformation.IconSpacingSize.Height / 2))

            For Each filePath As String In files
                Dim fileName As String = Path.GetFileName(filePath)
                Dim newPath As String = Path.Combine(DesktopDir, fileName)

                Dim sourceDirectory As String = Path.GetDirectoryName(filePath)

                Try
                    If Not String.Equals(sourceDirectory, DesktopDir, StringComparison.OrdinalIgnoreCase) Then

                        If Directory.Exists(filePath) Then
                            My.Computer.FileSystem.MoveDirectory(filePath, newPath, True)
                        ElseIf File.Exists(filePath) Then
                            My.Computer.FileSystem.MoveFile(filePath, newPath, True)
                        End If

                    Else

                    End If


                    If Not MovedIconPositions.ContainsKey(newPath) Then
                        MovedIconPositions.Add(newPath, dropLocation)
                    Else
                        MovedIconPositions(newPath) = dropLocation
                    End If

                Catch ex As Exception
                    MessageBox.Show($"Chyba při přesunu '{fileName}': {ex.Message}", "Chyba DragDrop", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            Next

            LoadDesktopIcons()
        End If
    End Sub

    Private Sub Desktop_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        Me.SendToBack()
    End Sub

    Private Sub WallpaperRefleshToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem.Click
        BackUpdate()
        LoadDesktopIcons()
    End Sub

    Public Sub BackUpdate()
        If IO.File.Exists(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", Nothing)) = True Then
            Me.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", Nothing))
            PictureBox1.Image = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", Nothing))
            PictureBox1.Visible = True

            ' Applies the "Style" of the Background
            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", Nothing) = 0 AndAlso My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "TileWallpaper", Nothing) = 0 Then
                Me.BackgroundImageLayout = ImageLayout.Center
                PictureBox1.Visible = True
                PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
            ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", Nothing) = 0 AndAlso My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "TileWallpaper", Nothing) = 1 Then
                Me.BackgroundImageLayout = ImageLayout.Tile
                PictureBox1.Visible = True
            ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", Nothing) = 2 OrElse My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", Nothing) = 10 Then
                Me.BackgroundImageLayout = ImageLayout.Stretch
                PictureBox1.Visible = True
                PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
            ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", Nothing) = 6 Then
                Me.BackgroundImageLayout = ImageLayout.Zoom
                PictureBox1.Visible = True
                PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            Else
                Me.BackgroundImageLayout = ImageLayout.None
                PictureBox1.SizeMode = PictureBoxSizeMode.Normal
            End If
        Else
            Me.BackgroundImage = Nothing
            PictureBox1.Visible = False
        End If
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        If Clipboard.ContainsFileDropList() = True Then

            For Each filePath As String In Clipboard.GetFileDropList()
                If Directory.Exists(filePath) = True Then
                    Dim DI As New DirectoryInfo(filePath)
                    My.Computer.FileSystem.MoveDirectory(filePath, DesktopDir & "\" & DI.Name, True)
                ElseIf File.Exists(filePath) Then
                    Dim FI As New FileInfo(filePath)
                    My.Computer.FileSystem.MoveFile(filePath, DesktopDir & "\" & FI.Name, True)
                End If

            Next

            Clipboard.Clear()

        End If
    End Sub

    Private Sub DesktopOptionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DesktopOptionsToolStripMenuItem.Click
        DesktopProperties.ShowDialog(Me)
    End Sub
    Private selectRectangle As Rectangle = Rectangle.Empty
    Private startPoint As Point
    Private isDragging As Boolean = False
    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        If e.Button = MouseButtons.None Then
            currentDraggedButton = Nothing
            startPoint = e.Location

            If isDragging = True Then isDragging = False
            PictureBox1.Capture = False

        ElseIf e.Button = MouseButtons.Left Then
            isDragging = True
            PictureBox1.Capture = True

            Dim currentPoint As Point = e.Location

            Dim left As Integer = Math.Min(startPoint.X, currentPoint.X)
            Dim top As Integer = Math.Min(startPoint.Y, currentPoint.Y)
            Dim width As Integer = Math.Abs(startPoint.X - currentPoint.X)
            Dim height As Integer = Math.Abs(startPoint.Y - currentPoint.Y)

            selectRectangle = New Rectangle(left, top, width, height)

            PictureBox1.Invalidate()
        End If
    End Sub

    Private Sub NewToolStripMenuItem_DropDownOpening(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.DropDownOpening
        NewFileToolStripMenuItem.Image = GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\imageres.dll", 2, True).ToBitmap
        NewFolderToolStripMenuItem.Image = GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\imageres.dll", 3, True).ToBitmap
        NewLinkToolStripMenuItem.Image = GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\imageres.dll", 154, True).ToBitmap
    End Sub

    Private Sub PictureBox1_Paint(sender As Object, e As PaintEventArgs) Handles PictureBox1.Paint
        If Not selectRectangle.IsEmpty Then
            Dim fillColor As Color = Color.FromArgb(80, Color.LightBlue)

            Using brush As New SolidBrush(fillColor)
                e.Graphics.FillRectangle(brush, selectRectangle)
            End Using

            Using pen As New Pen(Color.Blue, 1)
                pen.DashStyle = Drawing2D.DashStyle.Solid
                e.Graphics.DrawRectangle(pen, selectRectangle)
            End Using
        End If
    End Sub

    Private Sub NewFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewFolderToolStripMenuItem.Click
        NFD.PATH = DesktopDir
        NFD.isFileDialog = False

        If NFD.ShowDialog(Me) = DialogResult.OK Then
            LoadDesktopIcons()
        End If
    End Sub

    Private Sub NewFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewFileToolStripMenuItem.Click
        NFD.PATH = DesktopDir
        NFD.isFileDialog = True

        If NFD.ShowDialog(Me) = DialogResult.OK Then
            LoadDesktopIcons()
        End If
    End Sub

    Private Sub NewLinkToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewLinkToolStripMenuItem.Click
        LinkCreator.PATH = DesktopDir

        If LinkCreator.ShowDialog(Me) = DialogResult.OK Then
            LoadDesktopIcons()
        End If
    End Sub

    Private Sub FSW_Changed(sender As Object, e As FileSystemEventArgs) Handles FSW.Changed, FSW.Created, PFSW.Changed, PFSW.Created
        LoadDesktopIcons()
    End Sub

    Private Sub FSW_Renamed(sender As Object, e As RenamedEventArgs) Handles FSW.Renamed, PFSW.Renamed
        Me.Invoke(Sub()
                      Dim oldFileName As String = Path.GetFileName(e.OldName)
                      Dim newFileName As String = Path.GetFileName(e.Name)
                      Dim buttonToUpdate As Control = Nothing

                      For Each ctrl As Control In PictureBox1.Controls
                          If ctrl.Text = oldFileName Then
                              buttonToUpdate = ctrl
                              Exit For
                          End If
                      Next

                      If buttonToUpdate IsNot Nothing Then
                          buttonToUpdate.Name = newFileName
                          buttonToUpdate.Text = newFileName
                          buttonToUpdate.Tag = e.FullPath
                      Else
                          Console.WriteLine($"Cannot find an item called: {oldFileName}")
                      End If
                  End Sub)
    End Sub

    Private Sub FSW_Deleted(sender As Object, e As FileSystemEventArgs) Handles FSW.Deleted, PFSW.Deleted
        Me.Invoke(Sub()
                      Dim FileName As String = Path.GetFileName(e.Name)
                      Dim buttonToUpdate As Control = Nothing

                      For Each ctrl As Control In PictureBox1.Controls
                          If ctrl.Text = e.Name Then
                              buttonToUpdate = ctrl
                              Exit For
                          End If
                      Next

                      If buttonToUpdate IsNot Nothing Then
                          buttonToUpdate.Dispose()
                      Else
                          Console.WriteLine($"Cannot find an deleted item called: {FileName}")
                      End If
                  End Sub)
    End Sub

    Private Sub DesktopCM_LostFocus(sender As Object, e As EventArgs) Handles DesktopCM.LostFocus
        DesktopCM.Close()
    End Sub

    Private Sub AutoArrangeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutoArrangeToolStripMenuItem.Click
        MovedIconPositions.Clear()

        LoadDesktopIcons()
    End Sub

    Private Sub DesktopCM_GotFocus(sender As Object, e As EventArgs) Handles DesktopCM.GotFocus
        DesktopCM.BringToFront()
    End Sub
End Class

Public Class ShellHelpers

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure SHFILEINFO
        Public hIcon As IntPtr
        Public iIcon As Integer
        Public dwAttributes As UInteger
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public szDisplayName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)>
        Public szTypeName As String
    End Structure

    Private Const SHGFI_ICON As UInteger = &H100
    Private Const SHGFI_LARGEICON As UInteger = &H0
    Private Const SHGFI_DISPLAYNAME As UInteger = &H200
    Private Const SHGFI_PIDL As UInteger = &H8
    Private Const SHGFI_USEFILEATTRIBUTES As UInteger = &H10

    Private Const CSIDL_DRIVES As Integer = &H11       ' My Computer/This PC)
    Private Const CSIDL_NETWORK As Integer = &H12      ' Network
    Private Const CSIDL_BITBUCKET As Integer = &HA     ' Recycle Bin
    Private Const CSIDL_CONTROLS As Integer = &H3      ' Control Panel

    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SHGetFileInfo(ByVal ppidl As IntPtr, ByVal dwFileAttributes As UInteger, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As UInteger, ByVal uFlags As UInteger) As IntPtr
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function SHParseDisplayName(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszName As String, ByVal pbc As IntPtr, ByRef ppidl As IntPtr, ByVal sfgaoIn As UInteger, ByRef psfgaoOut As UInteger) As Integer
    End Function

    <DllImport("shell32.dll")>
    Private Shared Function SHGetSpecialFolderLocation(ByVal hwnd As IntPtr, ByVal csidl As Integer, ByRef ppidl As IntPtr) As Integer
    End Function

    <DllImport("shell32.dll")>
    Private Shared Sub ILFree(ByVal pidl As IntPtr)
    End Sub

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function DestroyIcon(ByVal hIcon As IntPtr) As Boolean
    End Function


    Public Structure ShellItemInfo
        Public Icon As Bitmap
        Public DisplayName As String
    End Structure

    Public Shared Function GetShellItemInfo(shellPath As String) As ShellItemInfo
        Dim info As New ShellItemInfo()
        Dim pidl As IntPtr = IntPtr.Zero
        Dim success As Boolean = False

        Try
            Dim guidOnly As String = shellPath.Replace("::", "").ToUpper()

            Select Case guidOnly
                Case "{645FF040-5081-101B-9F08-00AA002F954E}" ' Recycle Bin
                    SHGetSpecialFolderLocation(IntPtr.Zero, CSIDL_BITBUCKET, pidl)
                Case "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" ' This PC
                    SHGetSpecialFolderLocation(IntPtr.Zero, CSIDL_DRIVES, pidl)
                Case "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}" ' Network
                    SHGetSpecialFolderLocation(IntPtr.Zero, CSIDL_NETWORK, pidl)
                Case Else
                    SHParseDisplayName(shellPath, IntPtr.Zero, pidl, 0, 0)
            End Select

            If pidl <> IntPtr.Zero Then
                Dim shinfo As New SHFILEINFO()
                Dim res As IntPtr = SHGetFileInfo(pidl, 0, shinfo, CUInt(Marshal.SizeOf(shinfo)), SHGFI_ICON Or SHGFI_LARGEICON Or SHGFI_DISPLAYNAME Or SHGFI_PIDL)

                If res <> IntPtr.Zero Then
                    info.DisplayName = shinfo.szDisplayName
                    If shinfo.hIcon <> IntPtr.Zero Then
                        Using ico As Icon = Icon.FromHandle(shinfo.hIcon)
                            info.Icon = ico.ToBitmap()
                        End Using
                        DestroyIcon(shinfo.hIcon)
                        success = True
                    End If
                End If
            End If

        Catch ex As Exception

        Finally
            If pidl <> IntPtr.Zero Then ILFree(pidl)
        End Try

        If String.IsNullOrEmpty(info.DisplayName) OrElse info.DisplayName = "Unknown Item" Then
            info = GetInfoFromRegistry(shellPath, info)
        End If

        If String.IsNullOrEmpty(info.DisplayName) Then info.DisplayName = "System Item"
        If info.Icon Is Nothing Then info.Icon = SystemIcons.Application.ToBitmap()

        Return info
    End Function

    Private Shared Function GetInfoFromRegistry(shellPath As String, currentInfo As ShellItemInfo) As ShellItemInfo
        Try
            Dim clsid As String = shellPath.Replace("::", "")
            Dim keyPath As String = "CLSID\" & clsid

            Using key As RegistryKey = Registry.ClassesRoot.OpenSubKey(keyPath)
                If key IsNot Nothing Then

                    If String.IsNullOrEmpty(currentInfo.DisplayName) Then
                        Dim name As Object = key.GetValue("LocalizedString")
                        If name Is Nothing Then name = key.GetValue("") ' Default value

                        If name IsNot Nothing AndAlso Not name.ToString().StartsWith("@") Then
                            currentInfo.DisplayName = name.ToString()
                        ElseIf name IsNot Nothing AndAlso name.ToString().StartsWith("@") Then
                            currentInfo.DisplayName = GetStaticName(clsid)
                        End If
                    End If

                End If
            End Using
        Catch
        End Try

        If String.IsNullOrEmpty(currentInfo.DisplayName) Then
            currentInfo.DisplayName = GetStaticName(shellPath.Replace("::", ""))
        End If

        Return currentInfo
    End Function

    Private Shared Function GetStaticName(clsid As String) As String
        Select Case clsid.ToUpper()
            Case "{645FF040-5081-101B-9F08-00AA002F954E}" : Return "Recycle Bin"
            Case "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" : Return "This PC"
            Case "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}" : Return "Network"
            Case "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}" : Return "Control Panel"
            Case Else : Return "Unknown Item"
        End Select
    End Function
End Class

Public Class ShadowButton
    Inherits Button

    Private _OriginalText As String = String.Empty
    Public Property OriginalText() As String
        Get
            Return _OriginalText
        End Get

        Set(ByVal value As String)
            _OriginalText = value
        End Set
    End Property

    Protected Overrides Sub OnPaint(ByVal pevent As PaintEventArgs)
        MyBase.OnPaint(pevent)

        Me.OriginalText = Me.Text

        Dim g As Graphics = pevent.Graphics
        Dim textToDraw As String = Me.Text
        Dim font As Font = New Font(Me.Font, FontStyle.Regular)
        Dim textColor As Color = Color.White
        Dim shadowColor As Color = Color.FromArgb(255, 0, 0, 0)

        Dim format As New StringFormat()
        format.Alignment = StringAlignment.Center
        format.LineAlignment = StringAlignment.Center
        Dim textRect As New Rectangle(0, 0, Me.Width - 3, Me.Height + 30)

        Using shadowBrush As New SolidBrush(shadowColor)
            'g.DrawString(textToDraw, font, shadowBrush, textRect.X + 1, textRect.Y + 1, format)
        End Using

        Using textBrush As New SolidBrush(textColor)
            ' Trying to add Shadow to the Icons... it is failing.

            'g.DrawString(textToDraw, font, textBrush, textRect, format)
        End Using

        format.Dispose()
    End Sub

End Class