Imports System.Drawing.Imaging
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Security.Policy
Imports System.Web
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports Microsoft.Win32
Public Class Startmenu
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

    Private Declare Auto Function ExtractIconEx Lib "shell32.dll" (
        ByVal lpszFile As String,
        ByVal nIconIndex As Integer,
        ByVal phiconLarge() As IntPtr,
        ByVal phiconSmall() As IntPtr,
        ByVal nIcons As UInteger
    ) As UInteger

    Private Declare Auto Function DestroyIcon Lib "user32.dll" (
        ByVal hIcon As IntPtr
    ) As Boolean

    Public Shared Function GetIcon(ByVal filePath As String, ByVal iconIndex As Integer, ByVal isSmallIcon As Boolean) As Icon
        Dim hIconLarge(0) As IntPtr
        Dim hIconSmall(0) As IntPtr
        Dim iconToReturn As Icon = Nothing

        Try
            Dim result As Integer = ExtractIconEx(filePath, iconIndex, hIconLarge, hIconSmall, 1)

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
    Private Const FILE_ATTRIBUTE_NORMAL As UInteger = &H80
    Private Const FILE_ATTRIBUTE_DIRECTORY As UInteger = &H10

    Private Const GENERIC_FILE_INDEX As Integer = 0
    Private Const GENERIC_FOLDER_INDEX As Integer = 1

    <DllImport("shell32.dll")>
    Private Shared Function SHGetFileInfo(
        ByVal pszPath As String,
        ByVal dwFileAttributes As UInteger,
        ByRef psfi As SHFILEINFO,
        ByVal cbFileInfo As UInteger,
        ByVal uFlags As UInteger) As IntPtr
    End Function

    Private Function GetFileIconIndex(ByVal filePath As String, ByRef imageList As ImageList, Optional ByVal useLargeIcons As Boolean = False) As Integer

        Dim isDirectory As Boolean = Directory.Exists(filePath)
        Dim cacheKey As String = ""
        Dim fileAttributes As UInteger = FILE_ATTRIBUTE_NORMAL
        Dim isLink As Boolean = False

        If isDirectory Then
            cacheKey = "folder_" & Path.GetFileName(filePath)
            fileAttributes = FILE_ATTRIBUTE_DIRECTORY
        Else
            cacheKey = Path.GetExtension(filePath).ToLower()

            If cacheKey = ".lnk" Then
                isLink = True
                cacheKey = filePath.ToLower()
            ElseIf String.IsNullOrEmpty(cacheKey) Then
                cacheKey = "no_ext"
            End If
        End If

        cacheKey = If(useLargeIcons, "large_", "small_") & cacheKey

        If imageList.Images.ContainsKey(cacheKey) Then
            Return imageList.Images.IndexOfKey(cacheKey)
        End If

        Dim shfi As New SHFILEINFO()

        Dim iconSizeFlag As UInteger = If(useLargeIcons, SHGFI_LARGEICON, SHGFI_SMALLICON)
        Dim flags As UInteger = SHGFI_ICON Or iconSizeFlag

        If Not isDirectory And Not isLink Then
            flags = flags Or SHGFI_USEFILEATTRIBUTES
        End If

        Dim result As IntPtr = SHGetFileInfo(
            filePath,
            fileAttributes,
            shfi,
            CUInt(Marshal.SizeOf(shfi)),
            flags)

        If result <> IntPtr.Zero Then
            Try
                Dim icon As Icon = Icon.FromHandle(shfi.hIcon)
                imageList.Images.Add(cacheKey, icon)
                DestroyIcon(shfi.hIcon)
                Return imageList.Images.IndexOfKey(cacheKey)
            Catch
                Return If(isDirectory, GENERIC_FOLDER_INDEX, GENERIC_FILE_INDEX)
            End Try
        Else
            Return If(isDirectory, GENERIC_FOLDER_INDEX, GENERIC_FILE_INDEX)
        End If
    End Function


    Private Sub Startmenu_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Me.Visible = True
    End Sub
    Public FST As Boolean = False
    Private Sub Startmenu_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        If FST = False Then
            Me.Visible = False
            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 0 Then
                AppBar.Button1.BackgroundImage = My.Resources.StartRight
            ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 1 Then
                Try
                    AppBar.Button1.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Normal", ""))
                Catch ex As Exception
                    AppBar.Button1.BackgroundImage = My.Resources.StartRight
                End Try
            ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0") = 2 Then
                Try
                    AppBar.Button1.BackgroundImage = AppBar.OrbNormal
                Catch ex As Exception
                    AppBar.Button1.BackgroundImage = My.Resources.StartRight
                End Try
            End If
        End If
    End Sub
    Dim canClose As Boolean = False
    Private Sub Startmenu_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not e.CloseReason = CloseReason.TaskManagerClosing OrElse e.CloseReason = CloseReason.WindowsShutDown OrElse canClose = True Then
            canClose = False
            e.Cancel = True
        Else
            e.Cancel = False
        End If
    End Sub

    Private Sub Startmenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "Mode", "0")
                Case 0
                    Me.BackColor = Color.Black
                    TreeView2.BackColor = Color.Black
                    Panel2.BackColor = Color.Black
                    Dim side As side = side
                    side.Left = -1
                    side.Right = -1
                    side.Top = -1
                    side.Bottom = -1
                    Dim result As Integer = DwmExtendFrameIntoClientArea(Me.Handle, side)
                Case 1
                    Panel2.BackColor = SystemColors.ControlLight
                    Me.BackColor = SystemColors.Control
                    TreeView2.BackColor = SystemColors.MenuBar
                Case Else
                    Panel2.BackColor = Color.Transparent
            End Select

            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "CustomTransparency", "1") = 1 Then
                Me.Opacity = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "TransparencyLevel", "90") / 100
            End If
        Catch ex As Exception
            MsgBox("Failed to apply Start menu theme! " & ex.Message)
        End Try

        ' TreeView

        Me.Location = New Point(0 - SystemInformation.BorderSize.Width, My.Computer.Screen.Bounds.Height - Me.Height - AppBar.Height + SystemInformation.BorderSize.Height)

        If TreeView2.Nodes.Count = 0 Then
            Select Case ComboBox1.SelectedIndex
                Case 0
                    LoadItemsIntoTreeView(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\")
                Case Else
                    LoadItemsIntoTreeView("C:\ProgramData\Microsoft\Windows\Start Menu\Programs")
            End Select
        End If

        ComboBox1.SelectedIndex = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "DefaultFolder", "0")

        Try
            Me.Width = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\", "StartLastWidth", 500)
            Me.Height = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\", "StartLastHeight", 560)
        Catch ex As Exception
            Me.Size = New Size(500, 560)
        End Try

        ' Tiles

        ReloadTiles()

        'Account Picture Detection

        Try
            Button2.BackgroundImage = Image.FromFile(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\AppData\Local\Temp\" & Environment.UserName & ".bmp")
        Catch ex As Exception
            Button2.BackgroundImage = My.Resources.DefaultUserPicture
        End Try

        ' Buttons

        Button27.BackgroundImage = GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\system32\imageres.dll", 95, False).ToBitmap
        Button1.BackgroundImage = GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\system32\shell32.dll", 112, False).ToBitmap
        Button26.BackgroundImage = GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\system32\imageres.dll", 22, False).ToBitmap
    End Sub
    Dim PinnedAppsPath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned"
    Public Sub ReloadTiles()
        ToolStrip1.Items.Clear()

        Try
            If Directory.Exists(PinnedAppsPath.Trim) = True Then
                Dim DirectoryInformation As New DirectoryInfo(PinnedAppsPath)

                For Each i As DirectoryInfo In DirectoryInformation.GetDirectories
                    Dim item As New ToolStripMenuItem(i.Name)
                    With item
                        .AutoSize = False
                        .Text = i.Name
                        .TextAlign = ContentAlignment.BottomCenter
                        .TextImageRelation = TextImageRelation.Overlay
                        .ImageAlign = ContentAlignment.MiddleCenter
                        If Me.BackColor.R >= 150 OrElse Me.BackColor.G >= 130 AndAlso Me.BackColor.B >= 130 Then
                            .ForeColor = Color.Black
                        Else
                            .ForeColor = Color.White
                        End If

                        Try
                            .ToolTipText = IO.File.ReadAllText(DirectoryInformation.FullName & "\" & i.Name & "\Program")
                        Catch ex As Exception
                            .ToolTipText = "Unknown Program."
                        End Try

                        Try
                            Dim ico As Icon = Icon.ExtractAssociatedIcon(IO.File.ReadAllText(DirectoryInformation.FullName & "\" & i.Name & "\Program"))
                            .Image = ico.ToBitmap
                        Catch ex As Exception
                            .Image = My.Resources.Program
                        End Try

                        Try
                            If IO.File.ReadAllText(DirectoryInformation.FullName & "\" & i.Name & "\SizeType") = 1 Then
                                .Size = New Size(96, 96)
                            ElseIf IO.File.ReadAllText(DirectoryInformation.FullName & "\" & i.Name & "\SizeType") = 2 Then
                                .Size = New Size(192, 96)
                            ElseIf IO.File.ReadAllText(DirectoryInformation.FullName & "\" & i.Name & "\SizeType") = 3 Then
                                .Size = New Size(192, 192)
                            Else
                                .Size = New Size(96, 96)
                            End If
                        Catch ex As Exception
                            .Size = New Size(96, 96)
                        End Try

                        .Tag = IO.File.ReadAllText(DirectoryInformation.FullName & "\" & i.Name & "\SizeType")
                        AddHandler item.MouseUp, AddressOf TileClick
                    End With

                    ToolStrip1.Items.Add(item)
                Next

            Else
                ' Action when the Startmenu will not detect a directory.
                ' In my case, we will create it.

                Directory.CreateDirectory(Environment.GetFolderPath(PinnedAppsPath))

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TileClick(ByVal sender As System.Object, ByVal e As Windows.Forms.MouseEventArgs)

        Dim DirectoryInformation As New DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned")

        If e.Button = MouseButtons.Right Then
            TileCM.Show(Control.MousePosition)
            TileCM.Tag = CType(sender, ToolStripMenuItem).Text
        ElseIf e.Button = MouseButtons.Left Then
            If IO.File.Exists(DirectoryInformation.FullName & "\" & CType(sender, ToolStripMenuItem).Text & "\Arguments") = True Then
                Dim SI As New ProcessStartInfo(IO.File.ReadAllText(DirectoryInformation.FullName & "\" & CType(sender, ToolStripMenuItem).Text & "\Program")) With {
                    .Arguments = IO.File.ReadAllText(DirectoryInformation.FullName & "\" & CType(sender, ToolStripMenuItem).Text & "\Arguments")
                }
                Process.Start(SI)
            Else
                Process.Start(IO.File.ReadAllText(DirectoryInformation.FullName & "\" & CType(sender, ToolStripMenuItem).Text & "\Program"))
            End If
        End If
    End Sub

    Private Sub PopulateTreeView(ByVal directoryPath As String, ByVal targetNodeCollection As TreeNodeCollection)

        Try
            Dim subDirectories() As String = Directory.GetDirectories(directoryPath)

            For Each dirPath As String In subDirectories
                Dim directoryInfo As New DirectoryInfo(dirPath)
                Dim folderNode As New TreeNode(directoryInfo.Name)

                Dim iconIndex As Integer = GetFileIconIndex(dirPath, treeImageList, True)

                folderNode.ImageIndex = iconIndex
                folderNode.SelectedImageIndex = iconIndex
                folderNode.Tag = dirPath
                folderNode.ToolTipText = directoryInfo.FullName

                targetNodeCollection.Add(folderNode)
                folderNode.Nodes.Add("...")
            Next

        Catch ex As UnauthorizedAccessException
            Dim errorNode As New TreeNode($"Access Denied: {directoryPath}")
            targetNodeCollection.Add(errorNode)
        Catch ex As Exception

        End Try


        Try
            Dim files() As String = Directory.GetFiles(directoryPath)

            For Each filePath As String In files
                Dim fileInfo As New FileInfo(filePath)
                Dim fileNode As New TreeNode(fileInfo.Name)

                If Not fileInfo.Name = "desktop.ini" Then
                    Dim iconIndex As Integer = GetFileIconIndex(fileInfo.FullName, treeImageList, True)

                    fileNode.ImageIndex = iconIndex
                    fileNode.SelectedImageIndex = iconIndex
                    fileNode.Tag = filePath
                    fileNode.ToolTipText = fileInfo.FullName

                    targetNodeCollection.Add(fileNode)
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Sub LoadItemsIntoTreeView(ByVal rootPath As String)

        If Not Directory.Exists(rootPath) Then
            MessageBox.Show("Path doesn't exist! """ & rootPath & """. Please check if the directory is not missing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        TreeView2.Nodes.Clear()
        treeImageList.Images.Clear()

        PopulateTreeView(rootPath, TreeView2.Nodes)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        System.Diagnostics.Process.Start("C:\Windows\explorer.exe", "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}")

        canClose = True
        Me.Close()
    End Sub

    Private Sub Startmenu_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged
        If Not Me.Location = New Point(0 - SystemInformation.BorderSize.Width, My.Computer.Screen.Bounds.Height - Me.Height - AppBar.Height + SystemInformation.BorderSize.Height) Then
            Me.Location = New Point(0 - SystemInformation.BorderSize.Width, My.Computer.Screen.Bounds.Height - Me.Height - AppBar.Height + SystemInformation.BorderSize.Height)
        End If
    End Sub

    Private Sub Startmenu_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        If Me.Visible = True Then
            If TreeView2.Nodes.Count = 0 Then
                Select Case ComboBox1.SelectedIndex
                    Case 0
                        LoadItemsIntoTreeView(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\")
                    Case Else
                        LoadItemsIntoTreeView("C:\ProgramData\Microsoft\Windows\Start Menu\Programs")
                End Select
            End If

            Me.Location = New Point(0, My.Computer.Screen.Bounds.Height - Me.Height - AppBar.Height)

            Try
                Select Case My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "Mode", "0")
                    Case 0
                        Me.BackColor = Color.Black
                        TreeView2.BackColor = Color.Black
                        Panel2.BackColor = Color.Black

                        Dim side As side = side
                        side.Left = -1
                        side.Right = -1
                        side.Top = -1
                        side.Bottom = -1
                        Dim result As Integer = DwmExtendFrameIntoClientArea(Me.Handle, side)
                    Case 1
                        Panel2.BackColor = SystemColors.ControlLight
                        Me.BackColor = SystemColors.Control
                        TreeView2.BackColor = SystemColors.MenuBar

                        Dim side As side = side
                        side.Left = 0
                        side.Right = 0
                        side.Top = 0
                        side.Bottom = 0
                        Dim result As Integer = DwmExtendFrameIntoClientArea(Me.Handle, side)
                    Case Else
                        Panel2.BackColor = Color.Transparent


                End Select
            Catch ex As Exception
                MsgBox("Failed to apply Start menu theme! " & ex.Message)
            End Try
        Else
            ' When Startmenu become UnVisible
        End If
    End Sub

    Private Sub FCM_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles FCM.Opening

        If TreeView2.SelectedNode IsNot Nothing Then
            If Directory.Exists(TreeView2.SelectedNode.Tag) Then
                If TreeView2.SelectedNode.Nodes.Count > 0 Then
                    ts1.Visible = True
                    ToolStripSeparator4.Visible = True

                    Select Case TreeView2.SelectedNode.IsExpanded
                        Case True
                            ts1.Text = "Collapse"
                            ts2.Visible = False

                        Case False
                            ts1.Text = "Expand"
                            ts2.Visible = True
                    End Select
                Else
                    ts1.Visible = False
                    ts2.Visible = False
                    ToolStripSeparator4.Visible = False
                End If

                ToolStripSeparator5.Visible = True

                RefreshToolStripMenuItem.Visible = True
                ts6.Visible = True
                ts3.Visible = True
                ts4.Visible = True
                ts5.Visible = True
            ElseIf File.Exists(TreeView2.SelectedNode.Tag) Then
                ts1.Text = "Open"
                ts1.Visible = True

                ts2.Visible = False
                ToolStripSeparator5.Visible = False
                ToolStripSeparator4.Visible = False
                RefreshToolStripMenuItem.Visible = False
                ts3.Visible = False
                ts4.Visible = True
                ts5.Visible = True
                ts6.Visible = False
            End If
        Else
            ts1.Visible = False
            ts2.Visible = False
            ToolStripSeparator5.Visible = False
            ToolStripSeparator4.Visible = False
            RefreshToolStripMenuItem.Visible = True
            ts3.Visible = False
            ts4.Visible = False
            ts5.Visible = False
            ts6.Visible = True
        End If
    End Sub

    Private Sub NewGroupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewGroupToolStripMenuItem.Click
        Dim targetPath As String = String.Empty
        Select Case ComboBox1.SelectedIndex
            Case 0
                targetPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\"
            Case 1
                targetPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
            Case Else
                targetPath = TreeView2.SelectedNode.Tag & "\"
        End Select

        Dim newFolder As New TreeNode("New Folder")
        With newFolder
            .ImageIndex = 0
            .Tag = "[NewFolder]"
        End With

        If TreeView2.SelectedNode IsNot Nothing Then
            If Directory.Exists(TreeView2.SelectedNode.Tag) Then
                TreeView2.SelectedNode.Expand()
                TreeView2.SelectedNode.Nodes.Add(newFolder)

                lastNodeTag = TreeView2.SelectedNode.Tag

                TreeView2.SelectedNode = newFolder
                TreeView2.SelectedNode.BeginEdit()
            End If
        End If
    End Sub

    Public Sub NewFolderCreationOld()
        If TreeView2.SelectedNode IsNot Nothing Then
            NFD.PATH = TreeView2.SelectedNode.Tag & "\"
        Else
            NFD.PATH = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\"
        End If
        If NFD.ShowDialog = Windows.Forms.DialogResult.OK Then
            Select Case ComboBox1.SelectedIndex
                Case 0
                    LoadItemsIntoTreeView(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\")
                Case Else
                    LoadItemsIntoTreeView("C:\ProgramData\Microsoft\Windows\Start Menu\Programs")
            End Select
        End If
    End Sub

    Private Sub ts5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ts5.Click
        If MessageBox.Show("Are you sure do you want remove " & TreeView2.SelectedNode.Tag & " ?", "Delete Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
            Try
                If IO.Directory.Exists(TreeView2.SelectedNode.Tag) = True Then
                    My.Computer.FileSystem.DeleteDirectory(TreeView2.SelectedNode.Tag, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                    TreeView2.SelectedNode.Remove()
                ElseIf IO.File.Exists(TreeView2.SelectedNode.Tag) = True Then
                    My.Computer.FileSystem.DeleteFile(TreeView2.SelectedNode.Tag, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin, FileIO.UICancelOption.DoNothing)
                    TreeView2.SelectedNode.Remove()
                Else
                    MessageBox.Show("Failed to delete: " & TreeView2.SelectedNode.Tag & ". Check if the File/Directory still exists.", "File/Directory not found", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If

            Catch uex As UnauthorizedAccessException
                MsgBox("Access is denied to remove this item. If you will run this Shell as Administrator, this error will not be appearing.", MsgBoxStyle.Exclamation, "Access is denied.")

            Catch ex As Exception
                MsgBox(ex.Message)

            End Try

        End If
    End Sub

    Private Sub RefleshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        TreeView2.Update()
        Select Case ComboBox1.SelectedIndex
            Case 0
                LoadItemsIntoTreeView(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\")
            Case Else
                LoadItemsIntoTreeView("C:\ProgramData\Microsoft\Windows\Start Menu\Programs")
        End Select
    End Sub

    Private Sub Button3_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button3.MouseUp
        If e.Button = Windows.Forms.MouseButtons.Right Then
            ContextMenu1.Show(Me, Control.MousePosition, LeftRightAlignment.Left)
        End If
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        RunDialog.Show(AppBar)
        RunDialog.Activate()

        canClose = True
        Me.Close()
    End Sub
    Private Sub ts1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ts1.Click
        Select Case ts1.Text
            Case "Expand"
                TreeView2.SelectedNode.Expand()
            Case "Collapse"
                TreeView2.SelectedNode.Collapse()
            Case "Open"
                Try
                    Process.Start(TreeView2.SelectedNode.Tag)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical)
                End Try
        End Select
    End Sub

    Private Sub ts2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ts2.Click
        TreeView2.SelectedNode.ExpandAll()
    End Sub

    Private Sub ts3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ts3.Click
        Process.Start("explorer.exe", "/select," & TreeView2.SelectedNode.Tag)
    End Sub

    Private Sub ts4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ts4.Click
        TreeView2.SelectedNode.BeginEdit()
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Dim DirectoryInformation As New DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned")

        If IO.File.Exists(DirectoryInformation.FullName & "\" & TileCM.Tag & "\Arguments") = True Then
            Dim SI As New ProcessStartInfo(IO.File.ReadAllText(DirectoryInformation.FullName & "\" & TileCM.Tag & "\Program")) With {
                    .Arguments = IO.File.ReadAllText(DirectoryInformation.FullName & "\" & TileCM.Tag & "\Arguments")
                }
            Process.Start(SI)
        Else
            Process.Start(IO.File.ReadAllText(DirectoryInformation.FullName & "\" & TileCM.Tag & "\Program"))
        End If
    End Sub

    Private Sub Startmenu_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        If Not Me.Location = New Point(0, My.Computer.Screen.Bounds.Height - Me.Height - AppBar.Height) Then
            Me.Location = New Point(0, My.Computer.Screen.Bounds.Height - Me.Height - AppBar.Height)
        End If
    End Sub

    Private Sub TileSizeToolStripMenuItem_DropDownOpening(sender As Object, e As EventArgs) Handles TileSizeToolStripMenuItem.DropDownOpening
        If My.Computer.FileSystem.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & TileCM.Tag & "\SizeType") = 1 Then
            Small9696ToolStripMenuItem.Checked = True
            Medium19296ToolStripMenuItem.Checked = False
            Big192192ToolStripMenuItem.Checked = False
        ElseIf My.Computer.FileSystem.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & TileCM.Tag & "\SizeType") = 2 Then
            Small9696ToolStripMenuItem.Checked = False
            Medium19296ToolStripMenuItem.Checked = True
            Big192192ToolStripMenuItem.Checked = False
        ElseIf My.Computer.FileSystem.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & TileCM.Tag & "\SizeType") = 3 Then
            Small9696ToolStripMenuItem.Checked = False
            Medium19296ToolStripMenuItem.Checked = False
            Big192192ToolStripMenuItem.Checked = True
        End If
    End Sub

    Private Sub Small9696ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Small9696ToolStripMenuItem.Click
        IO.File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & TileCM.Tag & "\SizeType", 1)
        ReloadTiles()
    End Sub

    Private Sub Medium19296ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Medium19296ToolStripMenuItem.Click
        IO.File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & TileCM.Tag & "\SizeType", 2)
        ReloadTiles()
    End Sub

    Private Sub Big192192ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Big192192ToolStripMenuItem.Click
        IO.File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & TileCM.Tag & "\SizeType", 3)
        ReloadTiles()
    End Sub

    Private Sub RefleshToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RefleshToolStripMenuItem1.Click
        ReloadTiles()
        Me.Invalidate()
    End Sub

    Private Sub LocateInExplorerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LocateInExplorerToolStripMenuItem.Click
        Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\")
    End Sub

    Private Sub ShortcutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShortcutToolStripMenuItem.Click

        If TreeView2.SelectedNode IsNot Nothing Then
            LinkCreator.PATH = TreeView2.SelectedNode.Tag & "\"
        Else
            LinkCreator.PATH = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\"
        End If

        If LinkCreator.ShowDialog = Windows.Forms.DialogResult.OK Then
            Select Case ComboBox1.SelectedIndex
                Case 0
                    LoadItemsIntoTreeView(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\")
                Case Else
                    LoadItemsIntoTreeView("C:\ProgramData\Microsoft\Windows\Start Menu\Programs")
            End Select
            TreeView2.Nodes.Find(LinkCreator.Result, True)
        End If
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        Shell("Control.exe")

        canClose = True
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        SA.ShowDialog(AppBar)
    End Sub

    Private Sub LockScreenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LockScreenToolStripMenuItem.Click
        On Error Resume Next
        Process.Start("RunDll32.exe", "user32.dll,LockWorkStation")
    End Sub

    Private Sub SwitchUserToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SwitchUserToolStripMenuItem.Click
        On Error Resume Next
        Process.Start("tsdiscon.exe")
    End Sub

    Private Sub LogoffToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogoffToolStripMenuItem.Click
        On Error Resume Next
        If MessageBox.Show("This will log off your Windows session, Are you sure do you want log off?", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            Process.Start("shutdown.exe", "/l")
        End If
    End Sub

    Private Sub AccountSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AccountSettingsToolStripMenuItem.Click
        On Error Resume Next
        Process.Start("control", "/name Microsoft.UserAccounts")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        LoginMenuS.Show(Control.MousePosition)
    End Sub

    Private Sub ShutdownActionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShutdownActionsToolStripMenuItem.Click
        On Error Resume Next
        Shell("start C:\Windows\NewShell\Shortcuts\shutdown.vbs")
    End Sub

    Private Sub ShutdownWithFastBootToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShutdownWithFastBootToolStripMenuItem.Click
        On Error Resume Next
        Shell("shutdown.exe /s /hybrid /t 0")
    End Sub

    Private Sub HibernateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HibernateToolStripMenuItem.Click
        Shell("shutdown.exe /h")
    End Sub

    Private Sub SleepToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SleepToolStripMenuItem.Click
        Shell("RunDll32.exe powrprof.dll,SetSuspendState")
    End Sub

    Private Sub RestartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestartToolStripMenuItem.Click
        On Error Resume Next
        Shell("shutdown.exe /r /t 0")
    End Sub

    Private Sub PowerSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PowerSettingsToolStripMenuItem.Click
        Shell("control /name Microsoft.PowerOptions")
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        If MessageBox.Show("Are you sure do you want delete this Tile """ & TileCM.Tag & """?", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            On Error Resume Next
            IO.Directory.Delete(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & TileCM.Tag, True)
            ReloadTiles()
        End If
    End Sub

    Private Sub RenameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RenameToolStripMenuItem.Click
        On Error Resume Next
        Dim s As String = Microsoft.VisualBasic.InputBox("Type a new name for the tile to be renamed.", "Rename Dialog")
        My.Computer.FileSystem.RenameDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & TileCM.Tag, s)
    End Sub

    Private Sub DefaultToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DefaultToolStripMenuItem.Click
        Dim FI As New IO.FileInfo(IO.File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & TileCM.Tag & "\Program"))
        Process.Start("explorer.exe", " /select," & FI.FullName)
    End Sub

    Private Sub InStartmenuToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InStartmenuToolStripMenuItem.Click
        Dim FI As New IO.FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & TileCM.Tag & "\Program")
        Process.Start("explorer.exe", "/select," & FI.FullName)
    End Sub

    Dim lastNodeTag As String = String.Empty
    Private Sub TreeView2_AfterLabelEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) Handles TreeView2.AfterLabelEdit
        If Not e.Node.Tag = "[NewFolder]" Then
            If Not e.Label = String.Empty Then
                Try
                    If Directory.Exists(e.Node.Tag) Then
                        My.Computer.FileSystem.RenameDirectory(e.Node.Tag, e.Label)
                    ElseIf File.Exists(e.Node.Tag) Then
                        My.Computer.FileSystem.RenameFile(e.Node.Tag, e.Label)
                    End If

                    e.Node.Tag = e.Node.LastNode.Tag & "\" & e.Label

                Catch ex As Exception
                    MsgBox(ex.Message)
                    e.CancelEdit = True
                End Try
            Else
                e.CancelEdit = True
            End If
        Else
            If Not e.Label = String.Empty Then
                Try

                    If lastNodeTag IsNot Nothing Then
                        e.Node.Tag = lastNodeTag
                    Else
                        e.Node.Tag = TreeView2.Tag
                    End If

                Catch ex As Exception
                    MsgBox("Failed to make a folder name! " & ex.Message)
                    e.Node.Remove()
                End Try

                Try
                    My.Computer.FileSystem.CreateDirectory(e.Node.Tag & "\" & e.Label)
                Catch uex As UnauthorizedAccessException
                    MsgBox("Access is denied to create this folder. If you will run this Shell as Administrator, this error will not be appearing.", MsgBoxStyle.Exclamation, "Access is denied.")
                    e.Node.Remove()
                Catch ex As Exception
                    MsgBox(ex.Message)
                    e.Node.Remove()
                End Try
            Else
                e.Node.Remove()
            End If
        End If

    End Sub

    Private Sub TreeView2_BeforeExpand(ByVal sender As Object, ByVal e As TreeViewCancelEventArgs) Handles TreeView2.BeforeExpand
        Dim nodeToExpand As TreeNode = e.Node

        If nodeToExpand.Nodes.Count = 1 AndAlso nodeToExpand.Nodes(0).Text = "..." Then
            nodeToExpand.Nodes(0).Remove()

            Dim currentPath As String = nodeToExpand.Tag.ToString()

            Try
                PopulateTreeView(currentPath, nodeToExpand.Nodes)
            Catch ex As Exception
                nodeToExpand.Nodes.Add($"Error Loading: {ex.Message}")
            End Try
        End If
    End Sub

    Private Sub TreeView2_AfterCollapse(ByVal sender As Object, ByVal e As TreeViewEventArgs) Handles TreeView2.AfterCollapse
        e.Node.Nodes.Clear()
        e.Node.Nodes.Add("...")
    End Sub

    Private Sub TreeView2_DoubleClick(ByVal sender As Object, ByVal e As TreeNodeMouseClickEventArgs) 'Handles TreeView2.NodeMouseDoubleClick
        ' Enable it for having DoubleClick instead SingleClick :)

        If e.Button = MouseButtons.Left AndAlso File.Exists(e.Node.Tag.ToString) Then
            Try
                Process.Start(e.Node.Tag.ToString())
            Catch ex As Exception
                ex = Nothing
            End Try
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "DefaultFolder", ComboBox1.SelectedIndex, RegistryValueKind.DWord)
        Select Case ComboBox1.SelectedIndex
            Case 0
                LoadItemsIntoTreeView(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\")
                TreeView2.Tag = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\"
            Case Else
                LoadItemsIntoTreeView("C:\ProgramData\Microsoft\Windows\Start Menu\Programs")
                TreeView2.Tag = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
        End Select
    End Sub

    Private Sub Startmenu_BackColorChanged(sender As Object, e As EventArgs) Handles Me.BackColorChanged
        If Me.BackColor.R >= 150 OrElse Me.BackColor.G >= 130 AndAlso Me.BackColor.B >= 130 Then
            Me.ForeColor = Color.Black
            Label2.ForeColor = Color.Black
            Label3.ForeColor = Color.Black
            TreeView2.ForeColor = Color.Black
        Else
            Me.ForeColor = Color.White
            Label2.ForeColor = Color.White
            Label3.ForeColor = Color.White
            TreeView2.ForeColor = Color.White
        End If
    End Sub

    Public Function ColorMixer(Color1 As Color, Color2 As Color) As Color
        Dim r1 As Integer = Color1.R
        Dim g1 As Integer = Color1.G
        Dim b1 As Integer = Color1.B

        Dim r2 As Integer = Color2.R
        Dim g2 As Integer = Color2.G
        Dim b2 As Integer = Color2.B

        Dim smichanaR As Integer = (r1 + r2) \ 2
        Dim smichanaG As Integer = (g1 + g2) \ 2
        Dim smichanaB As Integer = (b1 + b2) \ 2

        Return Color.FromArgb(smichanaR, smichanaG, smichanaB)
    End Function

    Private Sub PropertiesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PropertiesToolStripMenuItem.Click
        StartmenuProperties.ShowDialog(Me)
    End Sub

    Public startColor As Color = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "BackColor", "#000000"))
    Public endColor As Color = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "BackColor2", "#000000"))

    Private Sub Startmenu_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "Mode", "2") = 2 Then

            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "UseSystemColor", "1") = 1 Then
                Dim accentColor As Color = WindowsColorSettings.GetAccentColor

                startColor = accentColor
                endColor = accentColor
            Else

                startColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "BackColor", "#000000"))
                endColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "BackColor2", "#000000"))
            End If

            Dim rv1 As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "ColorPos", "0")
            Dim rv2 As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "ColorPos2", "0")

            If rv1 = 0 Then
                If rv2 = 0 Then
                    Using gradientBrush As New System.Drawing.Drawing2D.LinearGradientBrush(
                            New Point(Me.ClientSize.Width, 0),
                            New Point(0, 0),
                            startColor,
                            endColor)
                        e.Graphics.FillRectangle(gradientBrush, Me.ClientRectangle)
                    End Using

                ElseIf rv2 = 1 Then

                    Using gradientBrush As New System.Drawing.Drawing2D.LinearGradientBrush(
                            New Point(0, 0),
                            New Point(Me.ClientSize.Width, 0),
                            startColor,
                            endColor)
                        e.Graphics.FillRectangle(gradientBrush, Me.ClientRectangle)
                    End Using
                End If
            ElseIf rv1 = 1 Then
                If rv2 = 0 Then
                    Using gradientBrush As New System.Drawing.Drawing2D.LinearGradientBrush(
              New Point(Me.ClientSize.Width, Me.ClientSize.Height),
              New Point(0, 0),
              startColor,
              endColor)
                        e.Graphics.FillRectangle(gradientBrush, Me.ClientRectangle)
                    End Using

                ElseIf rv2 = 1 Then

                    Using gradientBrush As New System.Drawing.Drawing2D.LinearGradientBrush(
                            New Point(0, 0),
                            New Point(Me.ClientSize.Width, Me.ClientSize.Height),
                            startColor,
                            endColor)
                        e.Graphics.FillRectangle(gradientBrush, Me.ClientRectangle)
                    End Using
                End If
            End If

            Me.BackColor = ColorMixer(startColor, endColor)
            TreeView2.BackColor = ColorMixer(startColor, endColor)
        End If
    End Sub

    Private Sub SelectControl_Tick(sender As Object, e As EventArgs) Handles SelectControl.Tick
        Try
            If IO.File.Exists(TreeView2.SelectedNode.Tag) Then
                Dim ico As Icon = Icon.ExtractAssociatedIcon(TreeView2.SelectedNode.Tag)
                Button2.BackgroundImage = ico.ToBitmap
            Else
                Try
                    Button2.BackgroundImage = Image.FromFile(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\AppData\Local\Temp\" & Environment.UserName & ".bmp")
                Catch ex1 As Exception
                    Button2.BackgroundImage = My.Resources.DefaultUserPicture
                End Try
            End If
        Catch ex As Exception
            Try
                Button2.BackgroundImage = Image.FromFile(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\AppData\Local\Temp\" & Environment.UserName & ".bmp")
            Catch ex1 As Exception
                Button2.BackgroundImage = My.Resources.DefaultUserPicture
            End Try
        End Try

        SelectControl.Enabled = False
    End Sub

    Private Sub CreateNewTileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateNewTileToolStripMenuItem.Click
        TileCreator.ShowDialog(Me)
    End Sub

    Private Sub TreeView2_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView2.NodeMouseClick
        If e.Button = MouseButtons.Right Then
            TreeView2.SelectedNode = e.Node
            FCM.Show(MousePosition)
        Else
            If Directory.Exists(e.Node.Tag) Then
                If e.Node.Nodes.Count > 0 Then
                    If e.Node.IsExpanded = True Then
                        e.Node.Collapse()
                    Else
                        e.Node.Expand()
                    End If
                End If
            ElseIf File.Exists(e.Node.Tag) Then

                SelectControl.Enabled = True

                Try
                    Dim pr As New Process()

                    pr.StartInfo.FileName = e.Node.Tag

                    pr.Start()

                    pr.WaitForInputIdle()

                    canClose = True
                    Me.Visible = False
                    Me.Close()
                Catch ex As Exception

                End Try
            End If
        End If
    End Sub

    Private Sub TreeView2_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView2.AfterSelect
        If File.Exists(e.Node.Tag) Then SelectControl.Enabled = True
    End Sub

    Private Sub Startmenu_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        If Not Me.WindowState = FormWindowState.Normal Then Me.WindowState = FormWindowState.Normal
    End Sub
End Class