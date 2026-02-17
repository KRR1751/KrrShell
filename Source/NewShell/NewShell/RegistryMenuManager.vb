Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports Microsoft.Win32

Public Class ShellItemRules
    Private Const SFGAO_CANCOPY As UInteger = &H1
    Private Const SFGAO_CANMOVE As UInteger = &H2
    Private Const SFGAO_CANRENAME As UInteger = &H10
    Private Const SFGAO_CANDELETE As UInteger = &H20
    Private Const SHGFI_ATTRIBUTES As UInteger = &H800

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Private Structure SHFILEINFO
        Public hIcon As IntPtr
        Public iIcon As Integer
        Public dwAttributes As UInteger
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> Public szDisplayName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)> Public szTypeName As String
    End Structure

    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SHGetFileInfo(pszPath As String, dwFileAttributes As UInteger, ByRef psfi As SHFILEINFO, cbFileInfo As Integer, uFlags As UInteger) As IntPtr
    End Function

    Public Shared Function CanDelete(path As String) As Boolean
        Dim sfi As New SHFILEINFO()
        SHGetFileInfo(path, 0, sfi, Marshal.SizeOf(sfi), SHGFI_ATTRIBUTES)
        Return (sfi.dwAttributes And SFGAO_CANDELETE) <> 0
    End Function

    Public Shared Function CanRename(path As String) As Boolean
        Dim sfi As New SHFILEINFO()
        SHGetFileInfo(path, 0, sfi, Marshal.SizeOf(sfi), SHGFI_ATTRIBUTES)
        Return (sfi.dwAttributes And SFGAO_CANRENAME) <> 0
    End Function

    Public Shared Function CanMove(path As String) As Boolean
        Dim sfi As New SHFILEINFO()
        SHGetFileInfo(path, 0, sfi, Marshal.SizeOf(sfi), SHGFI_ATTRIBUTES)
        Return (sfi.dwAttributes And SFGAO_CANMOVE) <> 0
    End Function

    Public Shared Function CanCopy(path As String) As Boolean
        Dim sfi As New SHFILEINFO()
        SHGetFileInfo(path, 0, sfi, Marshal.SizeOf(sfi), SHGFI_ATTRIBUTES)
        Return (sfi.dwAttributes And SFGAO_CANCOPY) <> 0
    End Function

    Public Shared Function CanCreateNew(directoryPath As String) As Boolean
        Try
            If String.IsNullOrEmpty(directoryPath) OrElse Not Directory.Exists(directoryPath) Then Return False

            Dim tempFile As String = Path.Combine(directoryPath, Path.GetRandomFileName())
            Using fs = File.Create(tempFile)

            End Using

            File.Delete(tempFile)
            Return True
        Catch
            Return False
        End Try
    End Function

    Public Shared Function CanPaste(directoryPath As String) As Boolean
        Try
            Dim di As New IO.DirectoryInfo(directoryPath)
            Dim acl = di.GetAccessControl()
            Return CanCreateNew(directoryPath)
        Catch
            Return False
        End Try
    End Function
End Class

Public Class RegistryMenuManager

    ' --- WIN32 API ---
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

    Public Structure ShellItemInfo
        Public Icon As Bitmap
        Public DisplayName As String
    End Structure

    <DllImport("shlwapi.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function SHLoadIndirectString(ByVal pszSource As String, ByVal pszOutBuf As StringBuilder, ByVal cchOutBuf As UInteger, ByVal ppvReserved As IntPtr) As Integer
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function ExtractIconEx(ByVal lpszFile As String, ByVal nIconIndex As Integer, ByRef phiconLarge As IntPtr, ByRef phiconSmall As IntPtr, ByVal nIcons As UInteger) As UInteger
    End Function

    <DllImport("user32.dll")>
    Private Shared Function DestroyIcon(ByVal hIcon As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function ReleaseCapture() As Boolean
    End Function

    Private Function GetClipboardDropEffect() As DragDropEffects
        Dim data As IDataObject = Clipboard.GetDataObject()
        If data IsNot Nothing AndAlso data.GetDataPresent("Preferred DropEffect") Then
            Using stream As IO.MemoryStream = DirectCast(data.GetData("Preferred DropEffect"), IO.MemoryStream)
                Dim reader As New IO.BinaryReader(stream)
                Dim effect As Integer = reader.ReadInt32()

                If effect = 2 Then Return DragDropEffects.Move
            End Using
        End If
        Return DragDropEffects.Copy
    End Function

    Private Class ContextMenuItem
        Public Name As String = ""
        Public Text As String = ""
        Public Command As String = ""
        Public IconPath As String = ""
        Public Position As String = "middle"
        Public IsSeparator As Boolean = False
        Public IsSubMenu As Boolean = False
        Public IsShellEx As Boolean = False
        Public SubItems As New List(Of ContextMenuItem)()
        Public Property ItemName As String
    End Class

    Public Shared Function GetIconByExtension(ByVal extension As String, ByVal isLarge As Boolean) As Icon
        Dim flags As UInteger = &H100 ' SHGFI_ICON
        If isLarge Then
            flags = flags Or &H0 ' SHGFI_LARGEICON
        Else
            flags = flags Or &H1 ' SHGFI_SMALLICON
        End If
        flags = flags Or &H10 ' SHGFI_USEFILEATTRIBUTES

        ' Oprava: Použijte správnou strukturu SHFILEINFO z této třídy, ne z ShellApi.
        Dim shfi As New SHFILEINFO()
        SHGetFileInfo(extension, &H80, shfi, Marshal.SizeOf(shfi), flags) ' &H80 = FILE_ATTRIBUTE_NORMAL

        If shfi.hIcon <> IntPtr.Zero Then
            Return Icon.FromHandle(shfi.hIcon)
        End If
        Return Nothing
    End Function

    Public Shared Sub ShowRegistryMenu(ByVal paths As String(), ByVal pos As Drawing.Point, Optional isTreeViewContextMenu As Boolean = False)
        Dim cms As New ContextMenuStrip()
        cms.RenderMode = ToolStripRenderMode.System
        Dim items As New List(Of ContextMenuItem)()

        Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim isBackground As Boolean = (paths IsNot Nothing AndAlso paths.Length = 1)
        Dim isThisPC As Boolean = (paths Is Nothing OrElse paths.Length = 0)

        If isThisPC OrElse isBackground Then
            If Not isTreeViewContextMenu Then
                cms.Items.Add(New ToolStripSeparator())
            End If
        End If

        If isThisPC Then

            LoadFromRegistry(items, "CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell")
            GenerateMenu(cms.Items, items, paths, True, "", cms)

        ElseIf isBackground Then

            Dim currentDir As String = paths(0)

            LoadFromRegistry(items, "Directory\Background\shell")
            LoadFromRegistry(items, "Directory\shell")
            LoadFromRegistry(items, "Folder\shell")

            If currentDir.Equals(desktopPath, StringComparison.OrdinalIgnoreCase) Then
                LoadFromRegistry(items, "DesktopBackground\Shell")
            End If

            items.RemoveAll(Function(i)
                                Dim n = i.Name.ToLower()
                                Dim t = i.Text.ToLower()
                                Return n = "open"
                            End Function)

            GenerateMenu(cms.Items, items, paths, True, "", cms)
            'AddBackgroundTools(cms, currentDir, isTreeViewContextMenu)

        Else
            Dim mainPath As String = paths(0)
            Dim extension As String = Path.GetExtension(mainPath).ToLower()
            Dim isDir As Boolean = Directory.Exists(mainPath)
            Dim isDrive As Boolean = isDir AndAlso (mainPath.Length <= 3 OrElse mainPath.EndsWith(":\"))

            If isDrive Then
                LoadFromRegistry(items, "Drive\shell", "drive")
            ElseIf isDir Then
                LoadFromRegistry(items, "Directory\shell")
                LoadFromRegistry(items, "Folder\shell")
            Else
                LoadFromRegistry(items, extension & "\shell")
                Dim progId = Registry.ClassesRoot.OpenSubKey(extension)?.GetValue("")?.ToString()
                If Not String.IsNullOrEmpty(progId) Then LoadFromRegistry(items, progId & "\shell")

                Dim category = GetFileCategory(extension)
                If Not String.IsNullOrEmpty(category) Then LoadFromRegistry(items, "SystemFileAssociations\" & category & "\shell")
                LoadFromRegistry(items, "SystemFileAssociations\" & extension & "\shell")
            End If

            Dim hasOpen = items.Any(Function(i) i.Name.ToLower.Contains("open"))
            If Not hasOpen Then
                items.Add(New ContextMenuItem() With {.Name = "open", .Text = "Open"})
            End If

            If Not isDir AndAlso Registry.ClassesRoot.OpenSubKey("Applications")?.GetValue("NoOpenWith") Is Nothing Then
                items.Add(New ContextMenuItem() With {.Name = "openwith", .Text = "Open With...", .Command = "rundll32.exe shell32.dll,OpenAs_RunDLL", .Position = "top"})
            End If

            LoadFromRegistry(items, "*\shell")
            LoadFromRegistry(items, "AllFilesystemObjects\shell")

            GenerateMenu(cms.Items, items, paths, True, "", cms)

            'AddStandardFileTools(cms, paths, isTreeViewContextMenu)
        End If

        cms.Items.Add(New ToolStripSeparator())
        cms.Items.Add(New ToolStripMenuItem("Properties", Nothing, Sub()
                                                                       If isBackground Then
                                                                           'ShowFileProperties(paths(0))
                                                                       End If
                                                                   End Sub))
        'ReleaseCapture()
        cms.Show(pos)

    End Sub

    Private Shared Function GetNewMenuPrototypes() As List(Of ContextMenuItem)
        Dim results As New List(Of ContextMenuItem)()

        results.Add(New ContextMenuItem() With {
        .Text = "Folder",
        .Name = "new_folder",
        .Command = "folder",
        .Position = "0",
        .IconPath = "shell32.dll,3"
    })

        results.Add(New ContextMenuItem() With {
        .Text = "Shortcut",
        .Name = "new_link",
        .Command = ".lnk",
        .Position = "1",
        .IconPath = "shell32.dll,263"
    })

        Using classesKey = Registry.ClassesRoot
            For Each subKeyName In classesKey.GetSubKeyNames()
                If subKeyName.StartsWith(".") Then
                    Using extKey = classesKey.OpenSubKey(subKeyName)
                        Dim snKey = extKey.OpenSubKey("ShellNew")
                        If snKey IsNot Nothing Then
                            Dim progId = extKey.GetValue("")?.ToString()
                            Dim friendlyName As String = ""
                            Dim itemName As String = snKey.GetValue("ItemName")?.ToString()

                            If Not String.IsNullOrEmpty(progId) Then
                                Using pKey = classesKey.OpenSubKey(progId)
                                    friendlyName = pKey?.GetValue("")?.ToString()
                                    If friendlyName.StartsWith("@") Then friendlyName = TranslateMui(friendlyName)
                                End Using
                            End If

                            If String.IsNullOrEmpty(friendlyName) Then friendlyName = subKeyName.Substring(1).ToUpper() & " File"

                            results.Add(New ContextMenuItem() With {
                            .Text = friendlyName,
                            .Name = "new_" & subKeyName,
                            .Command = subKeyName,
                            .ItemName = itemName,
                            .Position = "2"
                        })
                        End If
                    End Using
                End If
            Next
        End Using

        Return results.OrderBy(Function(x) x.Position).ThenBy(Function(x) x.Text).ToList()
    End Function

    Private Shared Function CreateFileFromExtension(proto As ContextMenuItem, folder As String)
        Try
            Dim ext As String = proto.Command
            Dim fullPath As String = ""
            Dim baseName As String = ""

            If ext = "folder" Then
                fullPath = Path.Combine(folder, "New Folder")
                Dim i = 2
                While Directory.Exists(fullPath)
                    fullPath = Path.Combine(folder, $"New Folder ({i})")
                    i += 1
                End While
                Directory.CreateDirectory(fullPath)
            Else
                If Not String.IsNullOrEmpty(proto.ItemName) Then
                    baseName = TranslateMui(proto.ItemName)
                Else
                    baseName = "New " & proto.Text
                End If

                For Each invalidChar In Path.GetInvalidFileNameChars()
                    baseName = baseName.Replace(invalidChar, "")
                Next

                fullPath = Path.Combine(folder, baseName & ext)

                Dim i = 2
                While File.Exists(fullPath)
                    fullPath = Path.Combine(folder, $"{baseName} ({i}){ext}")
                    i += 1
                End While

                File.WriteAllBytes(fullPath, New Byte() {})
            End If

            Return fullPath
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Shared Sub GenerateMenu(ByVal collection As ToolStripItemCollection, ByVal items As List(Of ContextMenuItem), ByVal paths As String(), ByVal isMainMenu As Boolean, Optional ext As String = "", Optional ContextMenu As ContextMenuStrip = Nothing)
        Dim addedCmds As New HashSet(Of String)()
        Dim addedTexts As New HashSet(Of String)()

        Dim sortedItems = items.OrderBy(Function(item)
                                            Dim rank = GetSortRank(item, ext)
                                            If rank = 50 OrElse rank = 100 Then
                                                Return rank.ToString("D5") & "_" & item.Text.ToLower()
                                            Else
                                                Return rank.ToString("D5")
                                            End If
                                        End Function).ToList()

        Dim lastRank As Integer = -1
        Dim isFirst As Boolean = True
        Dim isMainPath As Boolean = False
        If paths Is Nothing OrElse paths.Length = 1 Then isMainMenu = True

        For Each item In sortedItems
            If String.IsNullOrEmpty(item.Text) Then Continue For

            item.Text = Char.ToUpper(item.Text(0)) & item.Text.Substring(1)

            If item.Command Is Nothing Then item.Command = ""
            Dim normCmd = If(item.Command, "").Replace("""", "").Replace(" ", "").ToLower().Trim()
            Dim normText = item.Text.ToLower().Trim()
            Dim internalName = item.Name.ToLower()

            Dim isSystemAction = (internalName = "open" Or internalName = "explore" Or internalName = "edit" Or internalName = "runas" Or internalName = "refresh")

            If Not isSystemAction Then
                If (Not String.IsNullOrEmpty(normCmd) AndAlso addedCmds.Contains(normCmd)) OrElse addedTexts.Contains(normText) Then
                    Continue For
                End If
            End If

            addedCmds.Add(normCmd)
            addedTexts.Add(normText)

            Dim currentRank = GetSortRank(item, ext)
            Dim currentGroup = currentRank \ 10

            If Not isFirst AndAlso currentGroup <> (lastRank \ 10) Then
                collection.Add(New ToolStripSeparator())
            End If

            Dim menuItem As New ToolStripMenuItem(item.Text)
            menuItem.Name = item.Name

            If Not String.IsNullOrEmpty(item.IconPath) Then
                menuItem.Image = ExtractIconToImage(item.IconPath)
            ElseIf internalName = "runas" Then
                menuItem.Image = ExtractIconToImage("imageres.dll,73")
            End If

            If isFirst AndAlso isMainMenu Then
                menuItem.Font = New Font(menuItem.Font, FontStyle.Bold)
            End If

            If item.IsSubMenu Then
                GenerateMenu(menuItem.DropDownItems, item.SubItems, paths, False)
            Else
                menuItem.Tag = item.Command

                Dim DoAction = Sub()
                                   Dim btn As ToolStripMenuItem = menuItem

                                   If paths IsNot Nothing Then
                                       For Each i As String In paths
                                           If Directory.Exists(i) Then
                                               If btn.Name.ToLower = "open" Then

                                               Else
                                                   ExecuteCommand(btn.Tag.ToString(), i)
                                               End If
                                           ElseIf File.Exists(i) Then
                                               If btn.Name.ToLower = "open" Then
                                                   Dim PSI As New ProcessStartInfo(i)
                                                   PSI.UseShellExecute = True
                                                   Process.Start(PSI)
                                               Else
                                                   ExecuteCommand(btn.Tag.ToString(), i)
                                               End If
                                           Else
                                               '...
                                           End If
                                       Next
                                   Else
                                       Try
                                           If Not String.IsNullOrEmpty(Command) Then
                                               Dim cleanCmd = Command.Replace("%1", "").Trim()
                                               Process.Start(New ProcessStartInfo(cleanCmd) With {.UseShellExecute = True})
                                           End If
                                       Catch ex As Exception

                                       End Try
                                   End If

                               End Sub


                AddHandler menuItem.Click, DoAction
                AddHandler menuItem.MouseUp, Sub(sender As Object, e As MouseEventArgs)
                                                 If e.Button = MouseButtons.Right Then
                                                     DoAction.Invoke()

                                                     ContextMenu.Close()
                                                 End If
                                             End Sub

            End If

            collection.Add(menuItem)
            lastRank = currentRank
            isFirst = False
        Next
    End Sub

    Private Shared Function GetSortRank(i As ContextMenuItem, extension As String) As Integer
        Dim n = i.Name.ToLower()
        Dim t = i.Text.ToLower()
        Dim category = GetFileCategory(extension)

        ' --- GROUP 1: Open ---
        If n = "open" Or n = "refresh" Then Return 10

        If n = "openwith" Then
            Return If(category = "audio" Or category = "video", 5, 11)
        End If

        If n.ToLower.Contains("open") Or t.ToLower.Contains("open") Then Return 12

        If n.Contains("setwallpaper") Then Return 13

        If n = "explore" Then Return 14

        ' --- GROUP 2: RunAs ---
        If n = "runas" Then Return 20
        If n.Contains("run") Or n = "print" Or n = "printto" Then Return 21

        ' --- GROUP 3: Edit ---
        If n.Contains("edit") Or t.Contains("edit") Then Return 30

        ' --- GROUP 4: Multimedia ---
        If n.Contains("play") Or n = "preview" Then Return 40

        If i.Position = "drive" Then Return 50

        Return 100 ' Anything else
    End Function

    Private Shared Sub ExecuteCommand(ByVal cmdLine As String, ByVal filePath As String)
        Try
            If String.IsNullOrEmpty(cmdLine) Then Return

            If cmdLine.StartsWith("shell:::") Then
                Process.Start("explorer.exe", $"/e,/root,""{filePath}""")
                Return
            End If

            Dim expandedCmd = Environment.ExpandEnvironmentVariables(cmdLine)
            Dim quotedPath As String = ""
            Dim finalCmd As String = expandedCmd

            If finalCmd.ToLower.Contains("rundll32.exe") OrElse finalCmd.ToLower.Contains("explorer.exe") Then
                quotedPath = filePath
            Else
                If Not finalCmd.Contains("""%1""") Then
                    finalCmd = finalCmd.Replace("%1", """%1""")
                    quotedPath = filePath
                ElseIf Not finalCmd.Contains("""%L""") Then
                    finalCmd = finalCmd.Replace("%L", """%L""")
                    quotedPath = filePath
                ElseIf Not finalCmd.Contains("""%V""") Then
                    finalCmd = finalCmd.Replace("%V", """%V""")
                    quotedPath = filePath
                Else
                    quotedPath = filePath
                End If
            End If

            Dim exePath As String = ""
            Dim args As String = ""

            If finalCmd.StartsWith("""") Then
                Dim endQuote = finalCmd.IndexOf("""", 1)
                exePath = finalCmd.Substring(1, endQuote - 1)
                args = finalCmd.Substring(endQuote + 1).Trim()
            Else
                Dim spaceIdx = finalCmd.IndexOf(" ")
                If spaceIdx > 0 Then
                    exePath = finalCmd.Substring(0, spaceIdx)
                    args = finalCmd.Substring(spaceIdx + 1).Trim()
                Else
                    exePath = finalCmd
                End If
            End If

            If Not Path.IsPathRooted(exePath) Then
                exePath = GetFullExePath(exePath)
            End If

            Dim finalArgs As String = ""
            If args.Contains("%1") Or args.Contains("%L") Or args.Contains("%V") Then
                finalArgs = args.Replace("%1", quotedPath).Replace("%L", quotedPath).Replace("%V", quotedPath)
            Else
                If finalCmd.ToLower.Contains("rundll32.exe") OrElse finalCmd.ToLower.Contains("explorer.exe") Then
                    finalArgs = args & " " & quotedPath
                Else
                    finalArgs = args & " """ & quotedPath & """"
                End If

            End If

            Process.Start(New ProcessStartInfo(exePath, finalArgs) With {.UseShellExecute = True})

        Catch ex As Exception
            MessageBox.Show("Chyba: " & ex.Message)
        End Try
    End Sub

    Private Shared Function GetFullExePath(ByVal exeName As String) As String
        If Not exeName.ToLower().EndsWith(".exe") Then exeName &= ".exe"

        Dim registryPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & exeName

        Using key = Registry.LocalMachine.OpenSubKey(registryPath)
            If key IsNot Nothing Then Return key.GetValue("")?.ToString()
        End Using

        Using key = Registry.CurrentUser.OpenSubKey(registryPath)
            If key IsNot Nothing Then Return key.GetValue("")?.ToString()
        End Using

        Return exeName
    End Function
    Private Shared Sub LoadFromRegistry(ByVal list As List(Of ContextMenuItem), ByVal path As String, Optional forcePosition As String = "")
        Try
            Using key = Registry.ClassesRoot.OpenSubKey(path)
                If key Is Nothing Then Return
                For Each subName In key.GetSubKeyNames()
                    Using subKey = key.OpenSubKey(subName)
                        Dim item As New ContextMenuItem()
                        item.Name = subName

                        Dim mui = subKey.GetValue("MUIVerb")
                        Dim rawText = If(mui IsNot Nothing, mui.ToString(), subKey.GetValue("")?.ToString())
                        If String.IsNullOrEmpty(rawText) Then rawText = subName

                        item.Text = TranslateMui(Environment.ExpandEnvironmentVariables(rawText))
                        item.IconPath = subKey.GetValue("Icon")?.ToString()
                        item.Position = If(String.IsNullOrEmpty(forcePosition), subKey.GetValue("Position")?.ToString(), forcePosition)

                        Using cmdKey = subKey.OpenSubKey("command")
                            If cmdKey IsNot Nothing Then
                                item.Command = cmdKey.GetValue("")?.ToString()

                                If String.IsNullOrEmpty(item.Command) Then
                                    Dim delegateGuid = cmdKey.GetValue("DelegateExecute")?.ToString()
                                    If Not String.IsNullOrEmpty(delegateGuid) Then
                                        item.Command = "shell:::{52205fd8-5dfb-447d-801a-d0b52f2e83e1}" ' CLSID for File Explorer
                                    End If
                                End If
                            End If
                        End Using

                        If subKey.OpenSubKey("shell") IsNot Nothing Then
                            item.IsSubMenu = True
                            LoadFromRegistry(item.SubItems, path & "\" & subName & "\shell")
                        End If

                        Dim lowerName = subName.ToLower()
                        If String.IsNullOrEmpty(item.Command) AndAlso (lowerName = "explore" Or lowerName = "open") Then
                            item.Command = "explorer.exe" ' fallback
                        End If

                        If Not String.IsNullOrEmpty(item.Command) Or item.IsSubMenu Then
                            list.Add(item)
                        End If
                    End Using
                Next

            End Using
        Catch : End Try
    End Sub

    Private Shared Sub LoadShellExHandlers(ByVal items As List(Of ContextMenuItem), ByVal keyPath As String)
        Try
            Using rootKey As RegistryKey = Registry.ClassesRoot.OpenSubKey(keyPath & "\shellex\ContextMenuHandlers")
                If rootKey IsNot Nothing Then
                    For Each handlerName In rootKey.GetSubKeyNames()
                        Using subKey = rootKey.OpenSubKey(handlerName)
                            Dim clsid As String = subKey.GetValue("")?.ToString()
                            If String.IsNullOrEmpty(clsid) Then clsid = handlerName

                            If clsid.StartsWith("{") Then
                                Dim friendlyName As String = GetFriendlyNameFromCLSID(clsid)

                                If Not String.IsNullOrEmpty(friendlyName) Then
                                    items.Add(New ContextMenuItem() With {
                                    .Text = friendlyName,
                                    .Name = "shellex_" & handlerName,
                                    .Command = "CLSID:" & clsid,
                                    .IsShellEx = True
                                })
                                End If
                            End If
                        End Using
                    Next
                End If
            End Using
        Catch : End Try
    End Sub

    Private Shared Function GetFriendlyNameFromCLSID(ByVal clsid As String) As String
        Try
            Using key = Registry.ClassesRoot.OpenSubKey("CLSID\" & clsid)
                Dim name = key?.GetValue("")?.ToString()
                If name IsNot Nothing AndAlso name.StartsWith("@") Then
                    name = SHLoadIndirectString(name, New StringBuilder(1024), 1024, IntPtr.Zero).ToString()
                End If
                Return name
            End Using
        Catch : Return "" : End Try
    End Function

    Private Shared Function TranslateMui(ByVal input As String) As String
        If Not input.StartsWith("@") Then Return input
        Dim sb As New StringBuilder(1024)
        If SHLoadIndirectString(input, sb, CUInt(sb.Capacity), IntPtr.Zero) = 0 Then Return sb.ToString()
        Return input
    End Function

    Private Shared Function GetFileCategory(ByVal ext As String) As String
        Select Case ext
            Case ".jpg", ".jpeg", ".png", ".gif", ".bmp" : Return "image"
            Case ".mp4", ".mkv", ".avi", ".mov" : Return "video"
            Case ".txt", ".doc", ".docx", ".pdf" : Return "text"
            Case ".mp3", ".wav" : Return "audio"
            Case Else : Return ""
        End Select
    End Function

    Public Shared Function ExtractIconToImage(ByVal iconSource As String) As Image
        Try
            Dim cleanSource = Environment.ExpandEnvironmentVariables(iconSource.Replace("""", "").Trim())
            Dim file As String = "" : Dim index As Integer = 0
            If cleanSource.Contains(",") Then
                Dim parts = cleanSource.Split(","c)
                file = parts(0).Trim() : Integer.TryParse(parts(1).Trim(), index)
            Else : file = cleanSource : End If
            Dim hL As IntPtr, hS As IntPtr
            ExtractIconEx(file, index, hL, hS, 1)
            If hS <> IntPtr.Zero Then
                Dim bmp = Icon.FromHandle(hS).ToBitmap()
                DestroyIcon(hL) : DestroyIcon(hS)
                Return bmp
            End If
        Catch : End Try
        Return Nothing
    End Function

    Public Shared Function GetLocalizedName(ByVal resourcePath As String) As String
        If String.IsNullOrWhiteSpace(resourcePath) Then Return ""

        If resourcePath.StartsWith("@") Then
            Dim sb As New StringBuilder(1024)
            If SHLoadIndirectString(resourcePath, sb, CUInt(sb.Capacity), IntPtr.Zero) = 0 Then
                Return sb.ToString()
            End If
        End If

        Return resourcePath.Split(","c)(0).Replace("@", "")
    End Function
End Class