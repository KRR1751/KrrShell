Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports Microsoft.Win32
Imports System.Collections.Specialized

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
Public Structure SHFILEINFO
    Public hIcon As IntPtr
    Public iIcon As Integer
    Public dwAttributes As Integer
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
    Public szDisplayName As String
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)>
    Public szTypeName As String
End Structure

<ComImport()>
<Guid("46EB5926-582E-4017-9FDF-E8998DAA0950")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Interface IImageList
    Function Add(hbmImage As IntPtr, hbmMask As IntPtr, ByRef pi As Integer) As Integer
    Function ReplaceIcon(i As Integer, hicon As IntPtr, ByRef pi As Integer) As Integer
    Function SetOverlayImage(iImage As Integer, iOverlay As Integer) As Integer
    Function Replace(i As Integer, hbmImage As IntPtr, hbmMask As IntPtr) As Integer
    Function AddMasked(hbmImage As IntPtr, crMask As Integer, ByRef pi As Integer) As Integer
    Function Draw(ByRef pimldp As IMAGELISTDRAWPARAMS) As Integer
    Function Remove(i As Integer) As Integer
    Function GetIcon(i As Integer, flags As Integer, ByRef picon As IntPtr) As Integer
End Interface

<StructLayout(LayoutKind.Sequential)>
Public Structure IMAGELISTDRAWPARAMS
    Public cbSize As Integer
    Public himl As IntPtr
    Public i As Integer
    Public hdcDst As IntPtr
    Public x As Integer
    Public y As Integer
    Public cx As Integer
    Public cy As Integer
    Public xBitmap As Integer
    Public yBitmap As Integer
    Public rgbBk As Integer
    Public rgbFg As Integer
    Public fStyle As Integer
    Public dwRop As Integer
    Public fState As Integer
    Public Frame As Integer
    Public crEffect As Integer
End Structure

Public Class Desktop
    Private Const REG_KEY_PATH As String = "DesktopBackground\Shell"

    Private Const SHGFI_ICON As UInteger = &H100
    Private Const SHGFI_SYSICONINDEX As UInteger = &H4000
    Private Const SHGFI_LARGEICON As UInteger = &H0
    Private Const SHGFI_SMALLICON As UInteger = &H1

    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SHGetFileInfo(pszPath As String, dwFileAttributes As UInteger, ByRef psfi As SHFILEINFO, cbFileInfo As UInteger, uFlags As UInteger) As IntPtr
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Shared Function ExtractIconEx(libName As String, iconIndex As Integer, largeIconPtr() As IntPtr, smallIconPtr() As IntPtr, nIcons As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
    End Function

    Private BaseIconSize As Size = SystemInformation.IconSize
    Private ItemWidth As Integer = SystemInformation.IconSpacingSize.Width
    Private ItemHeight As Integer = SystemInformation.IconSpacingSize.Height
    Private ZoomLevel As Double = 1.0
    Private Icons As New List(Of DesktopIcon)
    Private SelectionRect As Rectangle = Rectangle.Empty
    Private IsSelecting As Boolean = False
    Private IsDragging As Boolean = False
    Private DragStart As Drawing.Point
    Private LastMousePos As Drawing.Point

    Public Const WS_EX_TOOLWINDOW As Long = &H80L
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            ' Apply the ToolWindow style before the window Is actually created
            cp.ExStyle = cp.ExStyle Or WS_EX_TOOLWINDOW
            Return cp
        End Get
    End Property

    Public Async Function BackUpdate() As Task
        Try
            Dim wallpaperPath As String = CStr(Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", ""))
            Dim styleVal As String = CStr(Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", "0"))
            Dim tileVal As String = CStr(Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "TileWallpaper", "0"))

            If String.IsNullOrEmpty(wallpaperPath) OrElse Not File.Exists(wallpaperPath) Then Return

            Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
            Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height
            Dim finalImage As Image = Nothing

            Await Task.Run(Sub()
                               Try
                                   Dim bytes As Byte() = File.ReadAllBytes(wallpaperPath)
                                   Dim ms As New MemoryStream(bytes)
                                   Dim tempImg As Image = Image.FromStream(ms)

                                   Dim isGif As Boolean = tempImg.RawFormat.Equals(Drawing.Imaging.ImageFormat.Gif)

                                   If isGif Then
                                       finalImage = tempImg
                                   Else
                                       Dim needsResize As Boolean = (tempImg.Width > screenWidth OrElse tempImg.Height > screenHeight) AndAlso
                                                              (styleVal = "2" OrElse styleVal = "6" OrElse styleVal = "10" OrElse styleVal = "22")

                                       If needsResize Then
                                           Dim bmp As New Bitmap(screenWidth, screenHeight)
                                           Using g As Graphics = Graphics.FromImage(bmp)
                                               g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                                               g.DrawImage(tempImg, 0, 0, screenWidth, screenHeight)
                                           End Using
                                           finalImage = bmp
                                       Else
                                           finalImage = New Bitmap(tempImg)
                                       End If

                                       tempImg.Dispose()
                                       ms.Dispose()
                                   End If
                               Catch ex As Exception
                                   Debug.WriteLine("Hybrid Update Error: " & ex.Message)
                               End Try
                           End Sub)

            If finalImage IsNot Nothing Then
                If PictureBoxWallpaper.Image IsNot Nothing Then
                    Dim oldImg = PictureBoxWallpaper.Image
                    PictureBoxWallpaper.Image = Nothing
                    oldImg.Dispose()
                End If

                PictureBoxWallpaper.Image = finalImage
                ApplyWallpaperStyle(styleVal, tileVal)
                PictureBoxWallpaper.Visible = True
            End If

        Catch ex As Exception
            Debug.WriteLine("BackUpdate Main Error: " & ex.Message)
        End Try
    End Function

    Private Sub ApplyWallpaperStyle(style As String, tile As String)
        ' 0 = Center (Tile=0), 2 = Stretch, 6 = Fit (Zoom), 10 = Fill (Center + Zoom), 22 = Span

        Select Case style
            Case "0"
                If tile = "1" Then
                    PictureBoxWallpaper.SizeMode = PictureBoxSizeMode.Normal
                Else
                    PictureBoxWallpaper.SizeMode = PictureBoxSizeMode.CenterImage
                End If
            Case "2"
                PictureBoxWallpaper.SizeMode = PictureBoxSizeMode.StretchImage
            Case "6"
                PictureBoxWallpaper.SizeMode = PictureBoxSizeMode.Zoom
            Case "10", "22" ' Fill
                PictureBoxWallpaper.SizeMode = PictureBoxSizeMode.Zoom
            Case Else
                PictureBoxWallpaper.SizeMode = PictureBoxSizeMode.Normal
        End Select
    End Sub

    Public Function GetFileIcon(path As String) As Image
        Dim largeIcon(0) As IntPtr
        ExtractIconEx(path, 0, largeIcon, Nothing, 1)
        If largeIcon(0) <> IntPtr.Zero Then
            Return Icon.FromHandle(largeIcon(0)).ToBitmap()
        End If
        Return SystemIcons.Application.ToBitmap()
    End Function

    Private Function GetOptimizedIcon(path As String, targetSize As Integer) As Image
        If path.ToLower.StartsWith("shell:::") Then
            Dim info As ShellHelpers.ShellItemInfo = ShellHelpers.GetShellItemInfo(path.Replace("shell:::", ""))

            If info.Icon IsNot Nothing Then
                Return info.Icon
            Else
                Return SystemIcons.Application.ToBitmap
            End If
        End If

        Dim targetPath As String = path

        ' If fatal on Shortcuts, this will gather the real
        ' Path of a file/folder instead.
        If IO.Path.GetExtension(path).ToLower() = ".lnk" Then
            'Try : targetPath = GetShortcutTarget(path) : Catch : targetPath = path : End Try
        End If

        If IsValidImage(path) Then
            Try
                Using fs As New IO.FileStream(path, IO.FileMode.Open, IO.FileAccess.Read)
                    Using tempImg = Image.FromStream(fs)

                        Dim thumb As New Bitmap(targetSize, targetSize)
                        Using g = Graphics.FromImage(thumb)
                            g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic

                            Dim ratio = Math.Max(targetSize / tempImg.Width, targetSize / tempImg.Height)
                            Dim newW = CInt(tempImg.Width * ratio)
                            Dim newH = CInt(tempImg.Height * ratio)
                            Dim posX = (targetSize - newW) \ 2
                            Dim posY = (targetSize - newH) \ 2
                            g.DrawImage(tempImg, posX, posY, newW, newH)
                        End Using
                        Return thumb
                    End Using
                End Using
            Catch
            End Try
        End If

        Try
            Dim shinfo As New SHFILEINFO()
            SHGetFileInfo(targetPath, 0, shinfo, Marshal.SizeOf(shinfo), SHGFI_SYSICONINDEX)


            Dim iconSizeFlag As Integer
            If targetSize <= 16 Then
                iconSizeFlag = SHIL_SMALL
            ElseIf targetSize <= 32 Then
                iconSizeFlag = SHIL_LARGE
            ElseIf targetSize <= 48 Then
                iconSizeFlag = SHIL_EXTRALARGE
            Else
                iconSizeFlag = SHIL_JUMBO
            End If

            Dim iidImageList As New Guid("46EB5926-582E-4017-9FDF-E8998DAA0950")
            Dim iml As IImageList = Nothing
            SHGetImageList(iconSizeFlag, iidImageList, iml)

            Dim hIcon As IntPtr = IntPtr.Zero
            If iml IsNot Nothing Then
                iml.GetIcon(shinfo.iIcon, 1, hIcon)
            End If

            If hIcon <> IntPtr.Zero Then
                Dim bmp As Bitmap = Icon.FromHandle(hIcon).ToBitmap()
                DestroyIcon(hIcon)
                Return bmp
            End If
        Catch ex As Exception
            Return Icon.ExtractAssociatedIcon(path).ToBitmap()
        End Try

        Return Icon.ExtractAssociatedIcon(targetPath).ToBitmap()

    End Function

    Private Shared ReadOnly HWND_BOTTOM As IntPtr = New IntPtr(1)
    Private Const SWP_NOSIZE As UInt32 = &H1
    Private Const SWP_NOMOVE As UInt32 = &H2
    Private Const SWP_NOACTIVATE As UInt32 = &H10

    <DllImport("user32.dll")>
    Private Shared Function SetWindowPos(hWnd As IntPtr, hWndInsertAfter As IntPtr, X As Integer, Y As Integer, cx As Integer, cy As Integer, uFlags As UInt32) As Boolean
    End Function

    Private Sub SendToBottom()
        SetWindowPos(Me.Handle, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE)
    End Sub

    Public IsShowDesktopMode As Boolean = False

    Private Sub FormDesktop_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate
        If IsShowDesktopMode Then
            IsShowDesktopMode = False
            Me.TopMost = False
        Else
            SendToBottom()
        End If
    End Sub

    Public Sub ShowDesktop()
        IsShowDesktopMode = True
        Me.WindowState = FormWindowState.Maximized
        Me.BringToFront()
        Me.Activate()
    End Sub

    Private Const WM_WINDOWPOSCHANGING As Integer = &H46
    <StructLayout(LayoutKind.Sequential)>
    Private Structure WINDOWPOS
        Public hwnd As IntPtr
        Public hwndInsertAfter As IntPtr
        Public x As Integer
        Public y As Integer
        Public cx As Integer
        Public cy As Integer
        Public flags As Integer
    End Structure

    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = WM_WINDOWPOSCHANGING AndAlso Not IsShowDesktopMode Then
            Dim wp As WINDOWPOS = CType(Marshal.PtrToStructure(m.LParam, GetType(WINDOWPOS)), WINDOWPOS)
            wp.hwndInsertAfter = New IntPtr(1)
            wp.flags = wp.flags Or &H10

            Marshal.StructureToPtr(wp, m.LParam, True)
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub PictureBoxWallpaper_DoubleClick(sender As Object, e As EventArgs) Handles PictureBoxWallpaper.DoubleClick

        Dim mousePos As Drawing.Point = PictureBoxWallpaper.PointToClient(Cursor.Position)

        Dim hitIcon = Icons.FirstOrDefault(Function(i) i.Bounds.Contains(mousePos))

        If hitIcon IsNot Nothing Then
            Try
                Process.Start(New ProcessStartInfo(hitIcon.FilePath) With {.UseShellExecute = True})
            Catch ex As Exception
                MessageBox.Show("Failed to open file: " & ex.Message)
            End Try
        End If
    End Sub

    Private WithEvents HScrollIcons As New HScrollBar()
    Private TotalWidth As Integer = 0

    Private Sub InitScrollbar()
        HScrollIcons.Location = New Point(0, SystemInformation.WorkingArea.Height - HScrollIcons.Height)
        HScrollIcons.Size = New Size(SystemInformation.WorkingArea.Size.Width, HScrollIcons.Height)
        HScrollIcons.SmallChange = 20
        HScrollIcons.LargeChange = 100
        HScrollIcons.Visible = False
        Me.Controls.Add(HScrollIcons)
    End Sub

    Private Sub HScrollIcons_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollIcons.Scroll
        PictureBoxWallpaper.Invalidate()
    End Sub

    Private Sub PictureBoxWallpaper_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBoxWallpaper.MouseDown
        Dim virtualX = e.X + HScrollIcons.Value
        Dim virtualLocation As New Point(virtualX, e.Y)

        DragStart = virtualLocation
        LastMousePos = e.Location

        Dim hitIcon As DesktopIcon = Icons.FirstOrDefault(Function(i) i.Bounds.Contains(virtualLocation))

        If hitIcon IsNot Nothing Then
            If Not hitIcon.IsSelected Then
                If (ModifierKeys And Keys.Control) <> Keys.Control Then
                    Icons.ForEach(Sub(i) i.IsSelected = False)
                End If
                hitIcon.IsSelected = True
            End If
            IsDragging = True
        Else
            IsSelecting = True

            ' Start SelectionRect
            SelectionRect = New Rectangle(virtualLocation.X, virtualLocation.Y, 0, 0)
            Icons.ForEach(Sub(i) i.IsSelected = False)
        End If

        If IsDragging Or IsSelecting Then
            PictureBoxWallpaper.Capture = True
        End If

        PictureBoxWallpaper.Invalidate()
    End Sub

    Private Sub PictureBoxWallpaper_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBoxWallpaper.MouseMove
        Dim redrawNeeded As Boolean = False

        If IsSelecting Then

            Dim virtualX = e.X + HScrollIcons.Value
            SelectionRect = New Rectangle(
    Math.Min(DragStart.X, virtualX),
    Math.Min(DragStart.Y, e.Y),
    Math.Abs(DragStart.X - virtualX),
    Math.Abs(DragStart.Y - e.Y)
)
            For Each item In Icons
                item.IsSelected = SelectionRect.IntersectsWith(item.Bounds)
            Next
            redrawNeeded = True

        ElseIf IsDragging Then
            If e.Button = MouseButtons.Middle Then
                If Math.Abs(e.X - DragStart.X) > SystemInformation.DragSize.Width OrElse
                Math.Abs(e.Y - DragStart.Y) > SystemInformation.DragSize.Height Then
                    StartExternalDrag()
                End If

                redrawNeeded = True

            ElseIf e.Button = MouseButtons.Left Then
                Dim deltaX = e.X - LastMousePos.X
                Dim deltaY = e.Y - LastMousePos.Y

                Dim maxW = PictureBoxWallpaper.Width + HScrollIcons.Value
                Dim maxH = PictureBoxWallpaper.Height

                For Each item In Icons.Where(Function(i) i.IsSelected)
                    Dim newX = item.Bounds.X + deltaX
                    Dim newY = item.Bounds.Y + deltaY

                    If newX < 0 Then newX = 0
                    If newX + item.Bounds.Width > maxW Then newX = maxW - item.Bounds.Width

                    If newY < 0 Then newY = 0
                    If newY + item.Bounds.Height > maxH Then newY = maxH - item.Bounds.Height

                    item.Bounds = New Rectangle(newX, newY, item.Bounds.Width, item.Bounds.Height)
                Next

                ' EXTERNAL DRAG
                Dim mousePos As Point = e.Location
                If Not PictureBoxWallpaper.ClientRectangle.Contains(mousePos) Then
                    StartExternalDrag()
                End If

                redrawNeeded = True
            End If
        Else

            For Each item In Icons
                Dim wasHovered = item.IsHovered
                item.IsHovered = item.Bounds.Contains(e.Location)
                If wasHovered <> item.IsHovered Then redrawNeeded = True
            Next
        End If

        LastMousePos = e.Location
        If redrawNeeded Then PictureBoxWallpaper.Invalidate()
    End Sub

    Private Sub StartExternalDrag()
        Dim selectedFiles = Icons.Where(Function(i) i.IsSelected).Select(Function(i) i.FilePath).ToArray()

        If selectedFiles.Length > 0 Then

            IsDragging = False
            PictureBoxWallpaper.Capture = False

            Dim data As New DataObject(DataFormats.FileDrop, selectedFiles)

            PictureBoxWallpaper.DoDragDrop(data, DragDropEffects.Move Or DragDropEffects.Copy)
        End If
    End Sub

    Public Const CLSID_THISPC As String = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    Public Const CLSID_RECYCLEBIN As String = "::{645FF040-5081-101B-9F08-00AA002F954E}"
    Public Const CLSID_USERFILES As String = "::{59031a47-3f72-44a7-89c5-5595fe6b30ee}"
    Public Const CLSID_CONTROL_PANEL As String = "::{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}"
    Public Const CLSID_NETWORK As String = "::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}"

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

        Icons.Clear()

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

        Items.AddRange(directories)
        Items.AddRange(publicdirectories)

        Items.AddRange(files)
        Items.AddRange(publicfiles)

        For i As Integer = 0 To Items.Count - 1
            Dim isShellItem As Boolean = Items(i).StartsWith("::{")
            Dim isDirectory As Boolean = Directory.Exists(Items(i))
            Dim isFile As Boolean = File.Exists(Items(i))

            If isFile = False AndAlso isDirectory = False AndAlso isShellItem = False Then Continue For

            Dim DesktopItem As New DesktopIcon

            DesktopItem.Bounds = New Rectangle(0, 0, 100, 100)

            If isFile = True Then
                Dim FI As New FileInfo(Items(i))

                Try
                    DesktopItem.IconImage = GetOptimizedIcon(FI.FullName, BaseIconSize.Width + 10)
                Catch ex As Exception
                    DesktopItem.IconImage = SystemIcons.Application.ToBitmap
                End Try

                If FI.Extension.ToLower = ".lnk" Then
                    DesktopItem.Name = Path.GetFileNameWithoutExtension(FI.FullName)
                Else
                    DesktopItem.Name = FI.Name
                End If

                DesktopItem.FilePath = Items(i)

            ElseIf isDirectory = True Then
                Dim DI As New DirectoryInfo(Items(i))

                Try
                    DesktopItem.IconImage = GetOptimizedIcon(DI.FullName, BaseIconSize.Width + 10)
                Catch ex As Exception
                    DesktopItem.IconImage = SystemIcons.Question.ToBitmap
                End Try

                DesktopItem.Name = DI.Name
                DesktopItem.FilePath = Items(i)


            ElseIf isShellItem Then
                Dim info As ShellHelpers.ShellItemInfo = ShellHelpers.GetShellItemInfo(Items(i))

                DesktopItem.Name = info.DisplayName

                If info.Icon IsNot Nothing Then
                    DesktopItem.IconImage = info.Icon
                Else
                    DesktopItem.IconImage = SystemIcons.Application.ToBitmap
                End If

                DesktopItem.FilePath = "shell:" & Items(i)
            End If

            Icons.Add(DesktopItem)
        Next

        ArrangeIcons()
    End Sub

    Private Sub PictureBoxWallpaper_DragEnter(sender As Object, e As DragEventArgs) Handles PictureBoxWallpaper.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Move
        End If
    End Sub

    Private Function IsValidImage(filePath As String) As Boolean
        Try
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                Using img As Image = Image.FromStream(fs)
                    Return True
                End Using
            End Using
        Catch
            Return False
        End Try
    End Function

    Private Sub PictureBoxWallpaper_DragOver(sender As Object, e As DragEventArgs) Handles PictureBoxWallpaper.DragOver
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim clientPoint = PictureBoxWallpaper.PointToClient(New Point(e.X, e.Y))
            Dim target = Icons.FirstOrDefault(Function(i) i.Bounds.Contains(clientPoint) AndAlso IO.Directory.Exists(i.FilePath))


            If target IsNot Nothing Then
                e.Effect = DragDropEffects.Move
            Else
                e.Effect = DragDropEffects.Copy
            End If
        End If
    End Sub
    Public DesktopDir As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "DesktopDir", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory))
    Public DesktopDirPublic As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "DesktopDirPublic", Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory))
    Private Sub PictureBoxWallpaper_DragDrop(sender As Object, e As DragEventArgs) Handles PictureBoxWallpaper.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Dim desktopPath = DesktopDir
            Dim userFilesPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)

            Dim clientPoint = PictureBoxWallpaper.PointToClient(New Point(e.X, e.Y))
            Dim targetIcon = Icons.FirstOrDefault(Function(i) i.Bounds.Contains(clientPoint))

            For Each sourcePath In files
                Dim fileName = IO.Path.GetFileName(sourcePath)
                Dim finalDest As String = ""
                Dim isRecycleAction As Boolean = False

                If targetIcon IsNot Nothing Then

                    Dim targetExt = IO.Path.GetExtension(targetIcon.FilePath).ToLower()
                    If targetExt = ".exe" OrElse targetExt = ".lnk" OrElse targetExt = ".bat" Then
                        Try

                            Dim psi As New ProcessStartInfo(targetIcon.FilePath) With {
                            .Arguments = $"""{sourcePath}""",
                            .UseShellExecute = True
                        }
                            Process.Start(psi)
                            Continue For
                        Catch ex As Exception
                            Debug.WriteLine("Failed to open program: " & ex.Message)
                        End Try
                    End If


                    Select Case targetIcon.FilePath
                        Case "shell:" & CLSID_RECYCLEBIN
                            isRecycleAction = True
                        Case "shell:" & CLSID_USERFILES
                            finalDest = IO.Path.Combine(userFilesPath, fileName)
                        Case Else
                            If IO.Directory.Exists(targetIcon.FilePath) Then
                                finalDest = IO.Path.Combine(targetIcon.FilePath, fileName)
                            Else
                                finalDest = IO.Path.Combine(desktopPath, fileName)
                            End If
                    End Select
                Else
                    finalDest = IO.Path.Combine(desktopPath, fileName)
                End If


                Try
                    If isRecycleAction Then
                        If IO.Directory.Exists(sourcePath) Then
                            My.Computer.FileSystem.DeleteDirectory(sourcePath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                        Else
                            My.Computer.FileSystem.DeleteFile(sourcePath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                        End If
                        Continue For
                    End If

                    If sourcePath.Equals(finalDest, StringComparison.OrdinalIgnoreCase) Then Continue For

                    Dim safeDest = finalDest
                    If IO.File.Exists(safeDest) OrElse IO.Directory.Exists(safeDest) Then
                        safeDest = GetNewFileName(finalDest)
                    End If

                    If IO.Directory.Exists(sourcePath) Then
                        IO.Directory.Move(sourcePath, safeDest)
                    Else
                        IO.File.Move(sourcePath, safeDest)
                    End If
                Catch ex As Exception
                    Debug.WriteLine("DragDrop Error: " & ex.Message)
                End Try
            Next
            PictureBoxWallpaper.Invalidate()
        End If
    End Sub

    Private Function GetNewFileName(path As String) As String
        Dim count = 1
        Dim dir = IO.Path.GetDirectoryName(path)
        Dim name = IO.Path.GetFileNameWithoutExtension(path)
        Dim ext = IO.Path.GetExtension(path)
        Dim newPath = path
        While IO.File.Exists(newPath)
            count += 1
            newPath = IO.Path.Combine(dir, $"{name} ({count}){ext}")
        End While
        Return newPath
    End Function
    Dim DragToGrid As Boolean = True

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure SHELLEXECUTEINFO
        Public cbSize As Integer
        Public fMask As UInteger
        Public hwnd As IntPtr
        <MarshalAs(UnmanagedType.LPWStr)> Public lpVerb As String
        <MarshalAs(UnmanagedType.LPWStr)> Public lpFile As String
        <MarshalAs(UnmanagedType.LPWStr)> Public lpParameters As String
        <MarshalAs(UnmanagedType.LPWStr)> Public lpDirectory As String
        Public nShow As Integer
        Public hInstApp As IntPtr
        Public lpIDList As IntPtr
        <MarshalAs(UnmanagedType.LPWStr)> Public lpClass As String
        Public hkeyClass As IntPtr
        Public dwHotKey As UInteger
        Public hIcon As IntPtr
        Public hProcess As IntPtr
    End Structure

    <DllImport("shell32.dll", EntryPoint:="ShellExecuteExW", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function ShellExecuteEx(ByRef lpExecInfo As SHELLEXECUTEINFO) As Boolean
    End Function

    Private Const SEE_MASK_INVOKEIDLIST As UInteger = &HC
    Private Const SW_SHOWNORMAL As Integer = 1

    Public Sub ShowFileProperties(ByVal Filename As String)
        Dim sei As New SHELLEXECUTEINFO()
        sei.cbSize = Marshal.SizeOf(sei)
        sei.lpVerb = "properties"
        sei.lpFile = Filename
        sei.fMask = SEE_MASK_INVOKEIDLIST
        sei.nShow = SW_SHOWNORMAL

        If Not ShellExecuteEx(sei) Then
            Dim err As Integer = Marshal.GetLastWin32Error()
            MessageBox.Show("Failed to open properties. Error code: " & err)
        End If
    End Sub

    Public Sub ShowCustomContextMenu(displayPoint As Point)
        Dim selectedIcons = Icons.Where(Function(i) i.IsSelected).ToList()

        If selectedIcons.Count = 0 Then
            Dim desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            RegistryMenuManager.ShowRegistryMenu({desktopPath}, displayPoint)
            Return
        End If

        Dim paths As String() = selectedIcons.Select(Function(x) x.FilePath).ToArray()
        Dim mainPath As String = paths(0)

        Dim cms As New ContextMenuStrip()
        cms.RenderMode = ToolStripRenderMode.System

        RegistryMenuManager.ShowRegistryMenu(paths, displayPoint)
    End Sub

    Private Sub PictureBoxWallpaper_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBoxWallpaper.MouseUp
        If IsDragging AndAlso DragToGrid Then
            Dim gridX As Integer = CInt(ItemWidth * ZoomLevel) + 10
            Dim gridY As Integer = CInt(ItemHeight * ZoomLevel) + 10

            For Each item In Icons.Where(Function(i) i.IsSelected)
                Dim snappedX As Integer = CInt(Math.Round(item.Bounds.X / gridX) * gridX)
                Dim snappedY As Integer = CInt(Math.Round(item.Bounds.Y / gridY) * gridY)

                item.Bounds = New Rectangle(snappedX + 10, snappedY + 10, item.Bounds.Width, item.Bounds.Height)
            Next

            UpdateTotalWidth()
        End If

        If e.Button = MouseButtons.Right Then
            Dim movement = Math.Abs(e.X - DragStart.X) + Math.Abs(e.Y - DragStart.Y)

            If movement < 5 Then
                Dim hitIcon = Icons.FirstOrDefault(Function(i) i.Bounds.Contains(e.Location))
                If hitIcon IsNot Nothing Then
                    RenamingIcon = hitIcon

                    'Dim paths As String() = {hitIcon.FilePath}
                    'RegistryMenuManager.ShowRegistryMenu(paths, MousePosition)
                    ShowCustomContextMenu(MousePosition)
                Else
                    DesktopCM.Show(PictureBoxWallpaper, e.Location)
                    DesktopCM.BringToFront()
                End If
            End If
        End If

        FinishRename(False)

        IsSelecting = False
        IsDragging = False
        SelectionRect = Rectangle.Empty
        PictureBoxWallpaper.Invalidate()
    End Sub

    Private Sub UpdateTotalWidth()
        Dim maxRight As Integer = 0

        For Each item In Icons
            If item.Bounds.Right > maxRight Then
                maxRight = item.Bounds.Right
            End If
        Next

        TotalWidth = maxRight

        If TotalWidth > PictureBoxWallpaper.Width Then
            HScrollIcons.Maximum = TotalWidth
            HScrollIcons.LargeChange = PictureBoxWallpaper.Width
            HScrollIcons.Visible = True
        Else
            HScrollIcons.Visible = False
            HScrollIcons.Value = 0
        End If

        PictureBoxWallpaper.Invalidate()
    End Sub

    Private Async Sub Desktop_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint Or ControlStyles.DoubleBuffer, True)
        PictureBoxWallpaper.AllowDrop = True

        Dim deskIcon As Icon = AppBar.GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\imageres.dll", 105, False)
        If deskIcon IsNot Nothing Then
            Me.Icon = deskIcon
        End If

        Await BackUpdate()
        LoadDesktopIcons()

        Dim ShellIconsKey As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
        If Registry.CurrentUser.OpenSubKey(ShellIconsKey) Is Nothing Then
            Registry.CurrentUser.CreateSubKey(ShellIconsKey)

            Registry.SetValue(ShellIconsKey, CLSID_RECYCLEBIN, "0")
        End If

        LoadSettings()

        FSW = New IO.FileSystemWatcher()
        FSW.Path = DesktopDir
        FSW.EnableRaisingEvents = True

        InitScrollbar()

        ArrangeIcons()
    End Sub

    Public Sub LoadSettings()
        Dim defView As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Desktop", "DefaultView", 2)

        Select Case defView
            Case 0 ' Mini icons
                ZoomLevel = 0.5

            Case 1 ' Small icons
                ZoomLevel = 0.7

            Case 2 ' Medium icons
                ZoomLevel = 1

            Case 3 ' Large icons
                ZoomLevel = 2

            Case 4 ' Extra Large icons
                ZoomLevel = 3
        End Select

        DragToGrid = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Desktop", "DragToGrid", True)
    End Sub

    Private Sub DrawGlowText(g As Graphics, text As String, font As Font, rect As Rectangle, sf As StringFormat)

        Dim glowColor As Color = Color.FromArgb(200, 0, 0, 0)
        Using b As New SolidBrush(glowColor)

            For x As Integer = -1 To 1
                For y As Integer = 0 To 2
                    If x = 0 And y = 0 Then Continue For
                    g.DrawString(text, font, b, New Rectangle(rect.X + x, rect.Y + y, rect.Width, rect.Height), sf)
                Next
            Next
        End Using


        g.DrawString(text, font, Brushes.White, rect, sf)
    End Sub

    Private WithEvents FSW As IO.FileSystemWatcher

    Private Sub FSW_Created(sender As Object, e As IO.FileSystemEventArgs) Handles FSW.Created
        If Me.IsHandleCreated AndAlso Not Me.IsDisposed Then
            Me.Invoke(Sub()
                          Dim newItem = New DesktopIcon With {
                          .Name = IO.Path.GetFileNameWithoutExtension(e.FullPath),
                          .FilePath = e.FullPath,
                          .IconImage = GetOptimizedIcon(e.FullPath, CInt(BaseIconSize.Width * ZoomLevel))
                      }
                          Icons.Add(newItem)
                          ArrangeIcons()
                      End Sub)
        End If
    End Sub

    Private Sub FSW_Renamed(sender As Object, e As IO.RenamedEventArgs) Handles FSW.Renamed
        If Me.IsHandleCreated AndAlso Not Me.IsDisposed Then
            Me.Invoke(Sub()
                          Dim item = Icons.FirstOrDefault(Function(i) i.FilePath.Equals(e.OldFullPath, StringComparison.OrdinalIgnoreCase))
                          If item IsNot Nothing Then
                              item.Name = IO.Path.GetFileNameWithoutExtension(e.FullPath)
                              item.FilePath = e.FullPath
                              PictureBoxWallpaper.Invalidate()
                          End If
                      End Sub)
        End If
    End Sub

    Private Sub FSW_Deleted(sender As Object, e As IO.FileSystemEventArgs) Handles FSW.Deleted
        If Me.IsHandleCreated AndAlso Not Me.IsDisposed Then
            Me.Invoke(Sub()
                          Icons.RemoveAll(Function(i) i.FilePath.Equals(e.FullPath, StringComparison.OrdinalIgnoreCase))
                          PictureBoxWallpaper.Invalidate()
                      End Sub)
        End If
    End Sub

    Public Function GetShortcutTarget(shortcutPath As String) As String
        Try
            If Not File.Exists(shortcutPath) Then Return String.Empty

            Dim fileBytes As Byte() = File.ReadAllBytes(shortcutPath)

            If fileBytes.Length < &H4C OrElse fileBytes(0) <> &H4C Then Return String.Empty

            Dim flags As UInteger = BitConverter.ToUInt32(fileBytes, &H14)
            If (flags And &H2) = 0 Then Return String.Empty ' HasLinkInfo flag

            Dim offset As Integer = &H4C
            If (flags And &H1) <> 0 Then ' HasLinkTargetIDList flag
                Dim idListSize As Short = BitConverter.ToInt16(fileBytes, offset)
                offset += idListSize + 2
            End If

            Dim localBaseNameOffset As Integer = BitConverter.ToInt32(fileBytes, offset + &H10)
            Dim finalOffset As Integer = offset + localBaseNameOffset

            Dim pathBuilder As New StringBuilder()
            While finalOffset < fileBytes.Length AndAlso fileBytes(finalOffset) <> 0
                pathBuilder.Append(Chr(fileBytes(finalOffset)))
                finalOffset += 1
            End While

            Return pathBuilder.ToString()
        Catch
            Return String.Empty
        End Try
    End Function


    Private Sub PictureBoxWallpaper_Paint(sender As Object, e As PaintEventArgs) Handles PictureBoxWallpaper.Paint
        Dim g = e.Graphics
        Dim offsetX As Integer = HScrollIcons.Value

        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

        ' Scrollbar Offset
        g.TranslateTransform(-offsetX, 0)

        Dim currentIconSize As Integer = CInt(BaseIconSize.Width * ZoomLevel)

        For Each item In Icons
            If item.Bounds.Right < offsetX OrElse item.Bounds.Left > offsetX + PictureBoxWallpaper.Width Then
                Continue For
            End If

            ' Hover
            If item.IsSelected Then
                g.FillRectangle(New SolidBrush(Color.FromArgb(100, 51, 153, 255)), item.Bounds)
                g.DrawRectangle(New Pen(Color.DodgerBlue, 1), item.Bounds)
            ElseIf item.IsHovered Then
                g.FillRectangle(New SolidBrush(Color.FromArgb(60, 51, 153, 255)), item.Bounds)
                g.DrawRectangle(New Pen(Color.FromArgb(150, 51, 153, 255), 1), item.Bounds)
            End If

            ' Icon
            Dim imgRect As New Rectangle(
            item.Bounds.X + (item.Bounds.Width - currentIconSize) \ 2,
            item.Bounds.Y + 5,
            currentIconSize,
            currentIconSize
        )
            g.DrawImage(item.IconImage, imgRect)

            ' Text
            Dim fontSize As Single = CSng(9 * ZoomLevel)
            Using f As New Font("Segoe UI", fontSize, FontStyle.Regular)
                Dim textRect As New Rectangle(item.Bounds.X, imgRect.Bottom + 5, item.Bounds.Width, 40)
                Dim sf As New StringFormat With {.Alignment = StringAlignment.Center, .Trimming = StringTrimming.EllipsisCharacter}
                DrawGlowText(g, item.Name, f, textRect, sf)
            End Using
        Next

        If IsSelecting Then
            Using br As New SolidBrush(Color.FromArgb(100, 51, 153, 255))
                g.FillRectangle(br, SelectionRect)
            End Using
            Using p As New Pen(Color.FromArgb(220, 51, 153, 255))
                g.DrawRectangle(p, SelectionRect)
            End Using
        End If

        g.ResetTransform()
    End Sub

    Private LastIconState As Integer = -1

    Private Sub PictureBoxWallpaper_MouseWheel(sender As Object, e As MouseEventArgs) Handles PictureBoxWallpaper.MouseWheel
        If (ModifierKeys And Keys.Shift) = Keys.Shift Then
            Dim newValue As Integer = HScrollIcons.Value - (e.Delta)

            If newValue < 0 Then newValue = 0
            If newValue > HScrollIcons.Maximum - HScrollIcons.LargeChange Then
                newValue = HScrollIcons.Maximum - HScrollIcons.LargeChange
            End If

            HScrollIcons.Value = newValue
            PictureBoxWallpaper.Invalidate()
        End If

        If (ModifierKeys And Keys.Control) = Keys.Control Then
            Dim oldZoom = ZoomLevel
            If e.Delta > 0 Then
                ZoomLevel = Math.Min(ZoomLevel + 0.1, 3.0)
            Else
                ZoomLevel = Math.Max(ZoomLevel - 0.1, 0.5)
            End If

            Dim currentSize As Integer = CInt(BaseIconSize.Width * ZoomLevel)


            Dim currentState As Integer
            If currentSize <= 20 Then currentState = 0 Else If currentSize <= 36 Then currentState = 1 Else If currentSize <= 52 Then currentState = 2 Else currentState = 3


            If currentState <> LastIconState Then
                For Each item In Icons
                    If item.IconImage IsNot Nothing Then item.IconImage.Dispose()
                    item.IconImage = GetOptimizedIcon(item.FilePath, currentSize)
                Next
                LastIconState = currentState
            End If

            ArrangeIcons()
        End If
    End Sub

    Private Sub LargeIconsToolStripMenuItem_Click(sender As Object, e As EventArgs) 'Handles LargeIconsToolStripMenuItem.Click
        ZoomLevel = 1.5
        ArrangeIcons()
    End Sub

    Private Sub AutoArrangeToolStripMenuItem_Click(sender As Object, e As EventArgs) 'Handles AutoArrangeToolStripMenuItem.Click
        ArrangeIcons()
    End Sub

    Private Sub ArrangeIcons()
        Dim currentX As Integer = 10
        Dim currentY As Integer = 10
        Dim spacing As Integer = 10
        Dim iconWidth As Integer = ItemWidth * ZoomLevel
        Dim iconHeight As Integer = ItemHeight * ZoomLevel

        Dim maxRight As Integer = 0

        For Each item In Icons
            item.Bounds = New Rectangle(currentX, currentY, iconWidth, iconHeight)

            If item.Bounds.Right > maxRight Then
                maxRight = item.Bounds.Right
            End If

            currentY += iconHeight + spacing

            If currentY + iconHeight > PictureBoxWallpaper.Height Then
                currentY = 10
                currentX += iconWidth + spacing
            End If
        Next

        TotalWidth = maxRight

        If TotalWidth > PictureBoxWallpaper.Width Then
            HScrollIcons.Location = New Point(0, SystemInformation.WorkingArea.Height - HScrollIcons.Height)
            HScrollIcons.Size = New Size(SystemInformation.WorkingArea.Size.Width, HScrollIcons.Height)
            HScrollIcons.Maximum = TotalWidth
            HScrollIcons.LargeChange = PictureBoxWallpaper.Width
            HScrollIcons.Visible = True
            HScrollIcons.BringToFront()
        Else
            HScrollIcons.Visible = False
            HScrollIcons.Value = 0
        End If

        PictureBoxWallpaper.Invalidate()
    End Sub

    Private WithEvents RenameBox As TextBox
    Public RenamingIcon As DesktopIcon = Nothing

    Public Sub StartRename(item As DesktopIcon)
        If item Is Nothing Then Return

        RenamingIcon = item
        RenameBox = New TextBox()

        RenameBox.Text = item.Name
        RenameBox.TextAlign = HorizontalAlignment.Center
        RenameBox.Multiline = True

        Dim rect = item.Bounds
        RenameBox.Bounds = New Rectangle(rect.X, rect.Y + (rect.Height - 30), rect.Width, 30)

        PictureBoxWallpaper.Controls.Add(RenameBox)
        RenameBox.Focus()
        RenameBox.SelectAll()
    End Sub

    Private Sub RenameBox_KeyDown(sender As Object, e As KeyEventArgs) Handles RenameBox.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            FinishRename(True)
        ElseIf e.KeyCode = Keys.Escape Then
            FinishRename(False)
        End If
    End Sub

    Private Sub RenameBox_LostFocus(sender As Object, e As EventArgs) Handles RenameBox.LostFocus
        FinishRename(True)
    End Sub

    Private Sub FinishRename(save As Boolean)
        If RenameBox Is Nothing Then Return

        If RenamingIcon Is Nothing Then
            CleanupRenameBox()
            Return
        End If

        If save Then
            Dim newBaseName = RenameBox.Text.Trim()
            Dim selectedIcons = Icons.Where(Function(i) i.IsSelected).ToList()

            If selectedIcons.Count = 0 Then selectedIcons.Add(RenamingIcon)

            If newBaseName <> "" Then
                Dim useIndex As Boolean = selectedIcons.Count > 1
                Dim counter As Integer = 1

                For Each item In selectedIcons
                    Try
                        Dim oldPath = item.FilePath

                        If oldPath.StartsWith("shell:::") Then
                            Dim clsid As String = oldPath.Replace("shell:::", "")

                            If RenameShellItemInRegistry(clsid, newBaseName) Then
                                item.Name = newBaseName
                                PictureBoxWallpaper.Invalidate()
                            End If
                            Continue For
                        End If

                        Dim directory = IO.Path.GetDirectoryName(oldPath)
                        Dim extension = IO.Path.GetExtension(oldPath)

                        Dim finalName As String = If(useIndex, $"{newBaseName} ({counter})", newBaseName)
                        Dim newPath = IO.Path.Combine(directory, finalName & extension)

                        If Not oldPath.Equals(newPath, StringComparison.OrdinalIgnoreCase) Then
                            If oldPath.StartsWith("shell:::") Then
                                Dim clsid As String = oldPath.Replace("shell:::", "")
                                RenameShellItemInRegistry(clsid, newBaseName)
                                Continue For
                            End If

                            If IO.File.Exists(newPath) OrElse IO.Directory.Exists(newPath) Then
                                newPath = GetNewFileName(newPath)
                            End If

                            If IO.Directory.Exists(oldPath) Then
                                IO.Directory.Move(oldPath, newPath)
                            Else
                                IO.File.Move(oldPath, newPath)
                            End If
                        End If
                        counter += 1
                    Catch ex As Exception
                        Debug.WriteLine("Chyba přejmenování: " & ex.Message)
                    End Try
                Next
            End If
        End If

        CleanupRenameBox()
    End Sub

    Private Sub CleanupRenameBox()
        If RenameBox IsNot Nothing Then
            RemoveHandler RenameBox.LostFocus, AddressOf RenameBox_LostFocus

            PictureBoxWallpaper.Controls.Remove(RenameBox)
            RenameBox.Dispose()
            RenameBox = Nothing
        End If
        RenamingIcon = Nothing
        PictureBoxWallpaper.Focus()
    End Sub

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

    Private Sub PictureBoxWallpaper_Resize(sender As Object, e As EventArgs) Handles PictureBoxWallpaper.Resize, PictureBoxWallpaper.LocationChanged
        ArrangeIcons()

        If Me.WindowState <> FormWindowState.Maximized Then
            Me.WindowState = FormWindowState.Maximized
        End If

        If Me.FormBorderStyle <> FormBorderStyle.None Then
            Me.FormBorderStyle = FormBorderStyle.None
        End If
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
        DesktopCM.Items.Add(PropertiesToolStripMenuItem)
    End Sub

    <DllImport("shlwapi.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function SHLoadIndirectString(
        ByVal pszSource As String,
        ByVal pszOutBuf As StringBuilder,
        ByVal cchOutBuf As UInteger,
        ByVal ppvReserved As IntPtr) As Integer
    End Function

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

                Dim extractedIcon As Icon = AppBar.GetIcon(filePath, iconIndex, isSmallIcon:=True)

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
        Try
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
                    Dim startInfo As New ProcessStartInfo(fileName, arguments) With {.UseShellExecute = True}

                    Shell(inputString)
                    'Process.Start(startInfo)
                Catch ex As Exception
                    MessageBox.Show($"Error while loading process: {ex.Message}", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Async Sub RefreshToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem.Click
        Await BackUpdate()
        LoadDesktopIcons()
        ArrangeIcons()
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ArrangeIconsToolStripMenuItem.Click
        ArrangeIcons()
    End Sub

    Private Sub GridSnappingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GridSnappingToolStripMenuItem.Click
        If DragToGrid = True Then DragToGrid = False Else DragToGrid = True

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Desktop", "DragToGrid", DragToGrid, RegistryValueKind.DWord)
    End Sub

    Private Sub ViewToolStripMenuItem_DropDownOpening(sender As Object, e As EventArgs) Handles ViewToolStripMenuItem.DropDownOpening
        GridSnappingToolStripMenuItem.Checked = DragToGrid

        ExtraLargeToolStripMenuItem.Checked = False
        LargeIconsToolStripMenuItem.Checked = False
        MediumIconsToolStripMenuItem.Checked = False
        SmallIconsToolStripMenuItem.Checked = False
        MiniIconsToolStripMenuItem.Checked = False

        PasteToolStripMenuItem.Enabled = Clipboard.ContainsFileDropList

        Select Case ZoomLevel
            Case 3 ' Extra Large
                ExtraLargeToolStripMenuItem.CheckState = CheckState.Indeterminate

            Case 2 ' Large
                LargeIconsToolStripMenuItem.CheckState = CheckState.Indeterminate

            Case 1 ' Medium
                MediumIconsToolStripMenuItem.CheckState = CheckState.Indeterminate

            Case 0.7 ' Small
                SmallIconsToolStripMenuItem.CheckState = CheckState.Indeterminate

            Case 0.5 ' Mini
                MiniIconsToolStripMenuItem.CheckState = CheckState.Indeterminate

        End Select
    End Sub

    Private Sub ExtraLargeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExtraLargeToolStripMenuItem.Click
        ZoomLevel = 3
        ArrangeIcons()

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Desktop", "DefaultView", 4, RegistryValueKind.DWord)
    End Sub

    Private Sub LargeIconsToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles LargeIconsToolStripMenuItem.Click
        ZoomLevel = 2
        ArrangeIcons()

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Desktop", "DefaultView", 3, RegistryValueKind.DWord)
    End Sub

    Private Sub MediumIconsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MediumIconsToolStripMenuItem.Click
        ZoomLevel = 1
        ArrangeIcons()

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Desktop", "DefaultView", 2, RegistryValueKind.DWord)
    End Sub

    Private Sub SmallIconsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SmallIconsToolStripMenuItem.Click
        ZoomLevel = 0.7
        ArrangeIcons()

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Desktop", "DefaultView", 1, RegistryValueKind.DWord)
    End Sub

    Private Sub MiniIconsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MiniIconsToolStripMenuItem.Click
        ZoomLevel = 0.5
        ArrangeIcons()

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Desktop", "DefaultView", 0, RegistryValueKind.DWord)
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        Dim data As IDataObject = Clipboard.GetDataObject()

        If data IsNot Nothing AndAlso data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files As String() = DirectCast(data.GetData(DataFormats.FileDrop), String())

            Dim dropEffect As IO.MemoryStream = DirectCast(data.GetData("Preferred DropEffect"), IO.MemoryStream)

            Dim moveFiles As Boolean = False

            If dropEffect IsNot Nothing Then
                Dim reader As New IO.BinaryReader(dropEffect)
                Dim effect As Integer = reader.ReadInt32()

                ' flag: 1 = Copy, 2 = Move, 4 = Link
                If (effect And 2) = 2 Then
                    moveFiles = True
                End If
            End If

            For Each sourcePath In files
                Try
                    Dim fileName As String = Path.GetFileName(sourcePath)
                    Dim destPath As String = Path.Combine(DesktopDir, fileName)

                    If File.Exists(sourcePath) Then
                        If moveFiles Then
                            If File.Exists(destPath) Then
                                If MessageBox.Show($"A file called ""{destPath}"" already exists. Do you want overwrite the current file with the new one?", "File Exists", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                                    File.Delete(destPath)
                                    File.Move(sourcePath, destPath)
                                End If
                            Else
                                File.Move(sourcePath, destPath)
                            End If
                        Else
                            If File.Exists(destPath) Then
                                If MessageBox.Show($"A file called ""{destPath}"" already exists. Do you want overwrite the current file with the new one?", "File Exists", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                                    File.Copy(sourcePath, destPath, True)
                                End If
                            Else
                                File.Copy(sourcePath, destPath, False)
                            End If
                        End If

                    ElseIf Directory.Exists(sourcePath) Then
                        If moveFiles Then
                            My.Computer.FileSystem.MoveDirectory(sourcePath, destPath, True)
                        Else
                            My.Computer.FileSystem.CopyDirectory(sourcePath, destPath, True)
                        End If
                    End If

                Catch ex As Exception
                    MessageBox.Show($"Failed to process path {sourcePath}: {ex.Message}")
                End Try
            Next

            If moveFiles Then Clipboard.Clear()
        End If
    End Sub

    Private Sub NewToolStripMenuItem_DropDownOpening(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.DropDownOpening
        NewFileToolStripMenuItem.Image = AppBar.GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\imageres.dll", 2, True).ToBitmap
        NewFolderToolStripMenuItem.Image = AppBar.GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\imageres.dll", 3, True).ToBitmap
        NewLinkToolStripMenuItem.Image = AppBar.GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\imageres.dll", 154, True).ToBitmap
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

    Private Sub PropertiesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PropertiesToolStripMenuItem.Click
        DesktopProperties.ShowDialog(Me)
    End Sub

    Private Sub PictureBoxWallpaper_KeyUp(sender As Object, e As KeyEventArgs) Handles PictureBoxWallpaper.KeyUp, Me.KeyUp
        If e.Control AndAlso e.KeyCode = Keys.A Then
            For Each desktopIcon In Icons
                desktopIcon.IsSelected = True
            Next
            PictureBoxWallpaper.Invalidate() ' Překreslíme plochu, aby byl vidět výběr
            e.Handled = True
            Return
        End If

        Dim selectedIcons = Icons.Where(Function(i) i.IsSelected).ToList()
        If selectedIcons.Count = 0 Then Return

        ' --- CTRL + C (Copy) ---
        If e.Control AndAlso e.KeyCode = Keys.C Then
            Dim paths As New StringCollection()
            paths.AddRange(selectedIcons.Select(Function(i) i.FilePath).ToArray())
            Clipboard.SetFileDropList(paths)
            e.Handled = True

            ' --- CTRL + X (Cut) ---
        ElseIf e.Control AndAlso e.KeyCode = Keys.X Then
            Dim paths As New StringCollection()
            paths.AddRange(selectedIcons.Select(Function(i) i.FilePath).ToArray())
            ' Tady by se mohl přidat ten Preferred DropEffect, ale pro teď:
            Clipboard.SetFileDropList(paths)
            e.Handled = True

            ' --- DELETE / SHIFT + DELETE ---
        ElseIf e.KeyCode = Keys.Delete Then
            ' Shift+Delete = Permanentní smazání (bez koše)
            Dim recycleBin As FileIO.RecycleOption = If(e.Shift, FileIO.RecycleOption.DeletePermanently, FileIO.RecycleOption.SendToRecycleBin)
            Dim msg = If(e.Shift, "Are you sure do you want permanently delete those selected items?", "Move items into the Recycle Bin?")

            If MessageBox.Show(msg, "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                For Each item In selectedIcons
                    Try
                        If Directory.Exists(item.FilePath) Then
                            My.Computer.FileSystem.DeleteDirectory(item.FilePath, FileIO.UIOption.OnlyErrorDialogs, recycleBin)
                        Else
                            My.Computer.FileSystem.DeleteFile(item.FilePath, FileIO.UIOption.OnlyErrorDialogs, recycleBin)
                        End If
                    Catch ex As Exception
                        Debug.WriteLine("Error while removing items with Keyboard: " & ex.Message)
                    End Try
                Next
            End If
            e.Handled = True

            ' --- F2 (Rename) ---
        ElseIf e.KeyCode = Keys.F2 AndAlso selectedIcons.Count = 1 Then
            StartRename(selectedIcons(0))
            e.Handled = True

            ' --- ENTER (Open) ---
        ElseIf e.KeyCode = Keys.Enter Then
            For Each item In selectedIcons
                Try
                    Process.Start(New ProcessStartInfo(item.FilePath) With {.UseShellExecute = True})
                Catch : End Try
            Next
            e.Handled = True
        End If
    End Sub
End Class

Public Class DesktopIcon
    Public Property Name As String
    Public Property FilePath As String
    Public Property IconImage As Image
    Public Property Bounds As Rectangle
    Public Property IsHovered As Boolean = False
    Public Property IsSelected As Boolean = False
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
            ' best-effort fallback
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

Public Module ShellIconHelper
    <ComImport()>
    <Guid("46EB5926-582E-4017-9FDF-E8998DAA0950")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Interface IImageList
        Function GetIcon(i As Integer, flags As Integer, ByRef picon As IntPtr) As Integer
    End Interface

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure SHFILEINFO
        Public hIcon As IntPtr
        Public iIcon As Integer
        Public dwAttributes As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public szDisplayName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)>
        Public szTypeName As String
    End Structure

    Public Const SHGFI_SYSICONINDEX As UInteger = &H4000
    Public Const SHIL_LARGE As Integer = &H0      ' 32x32
    Public Const SHIL_SMALL As Integer = &H1      ' 16x16
    Public Const SHIL_EXTRALARGE As Integer = &H2 ' 48x48
    Public Const SHIL_JUMBO As Integer = &H4      ' 256x256

    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Function SHGetFileInfo(pszPath As String, dwFileAttributes As UInteger, ByRef psfi As SHFILEINFO, cbFileInfo As UInteger, uFlags As UInteger) As IntPtr
    End Function

    <DllImport("shell32.dll", EntryPoint:="#727")>
    Public Function SHGetImageList(iImageList As Integer, ByRef riid As Guid, ByRef ppv As IImageList) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Function DestroyIcon(hIcon As IntPtr) As Boolean
    End Function
End Module