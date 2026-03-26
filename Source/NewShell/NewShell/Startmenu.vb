Imports System.Diagnostics.Eventing.Reader
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Security.Policy
Imports System.Text
Imports System.Web
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports Microsoft.Win32
Public Class Startmenu

    Private Const WS_EX_TOOLWINDOW As Integer = &H80
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            ' Apply the ToolWindow style before the window is actually created
            cp.ExStyle = cp.ExStyle Or WS_EX_TOOLWINDOW
            Return cp
        End Get
    End Property

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

    Public Shadows Sub Show()
        Me.Opacity = 0
        MyBase.Show()
        FadeTimer.Start()
    End Sub

    Public Shadows Sub Hide()
        FadeTimer.Stop()
        Me.Opacity = 0
        MyBase.Hide()
    End Sub

    Private Sub FadeTimer_Tick(sender As Object, e As EventArgs) Handles FadeTimer.Tick
        If Me.Opacity < 1.0 Then
            Me.Opacity += 0.1
        Else
            FadeTimer.Stop()
        End If
    End Sub

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

        Dim targetSize As Integer = If(useLargeIcons, 48, 16)

        Try
            Dim bitmapIcon As Bitmap = DirectCast(Desktop.GetOptimizedIcon(filePath, targetSize), Bitmap)

            If bitmapIcon IsNot Nothing Then
                Dim hIcon As IntPtr = bitmapIcon.GetHicon()

                Dim newIcon As Icon = Icon.FromHandle(hIcon)

                imageList.Images.Add(cacheKey, newIcon)

                DestroyIcon(hIcon)

                newIcon.Dispose()
                bitmapIcon.Dispose()

                Return imageList.Images.IndexOfKey(cacheKey)
            Else
                Return If(isDirectory, GENERIC_FOLDER_INDEX, GENERIC_FILE_INDEX)
            End If
        Catch ex As Exception
            Return If(isDirectory, GENERIC_FOLDER_INDEX, GENERIC_FILE_INDEX)
        End Try
    End Function
    Public FST As Boolean = False
    Private Sub Startmenu_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        If FST = False Then

            Dim defValue As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Type", "0")

            Dim liveBar = Application.OpenForms.OfType(Of AppBar)().FirstOrDefault()

            liveBar?.Invoke(Sub()
                                Select Case defValue

                                    Case 1 : Try
                                            Dim customImage As Image = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Appbar\StartButton", "Normal", ""))
                                            If customImage IsNot Nothing Then
                                                liveBar.Start.BackgroundImage = customImage
                                            Else
                                                liveBar.Start.BackgroundImage = My.Resources.StartRight
                                            End If

                                        Catch ex As Exception
                                            liveBar.Start.BackgroundImage = My.Resources.StartRight
                                        End Try

                                    Case 2 : Try
                                            liveBar.Start.BackgroundImage = liveBar.OrbNormal
                                        Catch ex As Exception
                                            liveBar.Start.BackgroundImage = My.Resources.StartRight
                                        End Try

                                    Case Else
                                        liveBar.Start.BackgroundImage = My.Resources.StartRight

                                End Select
                            End Sub) : End If
            Me.Hide()

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

    <DllImport("shell32.dll", EntryPoint:="#262", CharSet:=CharSet.Unicode, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function GetUserTilePath(ByVal username As String, ByVal flags As UInteger, ByVal path As StringBuilder, ByVal pathLength As UInteger) As Integer
    End Function

    Public Shared Function GetCurrentUserImagePath() As String
        Dim accountPics As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Windows", "AccountPictures")
        If Directory.Exists(accountPics) Then
            For Each f In Directory.EnumerateFiles(accountPics)
                Dim ext = Path.GetExtension(f).ToLowerInvariant()
                If ext = ".png" Or ext = ".jpg" Or ext = ".jpeg" Or ext = ".bmp" Then
                    Return f
                End If
            Next
        End If

        Dim programDataPic As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Microsoft", "User Account Pictures", "user.png")
        If File.Exists(programDataPic) Then Return programDataPic

        Dim tempPic As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "AppData", "Local", "Temp", Environment.UserName & ".bmp")
        If File.Exists(tempPic) Then Return tempPic

        Return String.Empty
    End Function
    Public StartMenuFolder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs")
    Public StartMenuFolderPublic As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs")
    Private Sub Startmenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Theme
        ApplyStartTheme()

        Try ' Size
            Dim newWidth As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\", "StartLastWidth", 500)
            Dim newHeight As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\", "StartLastHeight", 560)

            Dim result As Integer

            If Integer.TryParse(newWidth, result) Then newWidth = result Else newWidth = 500
            If Integer.TryParse(newHeight, result) Then newHeight = result Else newHeight = 560

            Me.Size = New Size(newWidth, newHeight) : Catch ex As Exception

            ' On Error:
            Me.Size = New Size(500, 560)
        End Try

        ' Location
        Me.Location = StartDefaultLocation

        ' TreeView
        ComboBox1.SelectedIndex = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "DefaultFolder", "0")

        If TreeView2.Nodes.Count = 0 Then
            Select Case ComboBox1.SelectedIndex
                Case 0
                    LoadItemsIntoTreeView(StartMenuFolder)
                Case Else
                    LoadItemsIntoTreeView(StartMenuFolderPublic)
            End Select
        End If

        ' Tiles

        ReloadTiles()

        'Account Picture Detection

        Try
            Dim imagePath As String = GetCurrentUserImagePath()
            usr_Picture.BackgroundImage = Image.FromFile(imagePath)

        Catch ex As Exception
            Debug.WriteLine("Error loading profile image: " & ex.Message)
            usr_Picture.BackgroundImage = My.Resources.DefaultUserPicture
        End Try

        ' Buttons

        Try : cmd_btn4.BackgroundImage = GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\system32\imageres.dll", 95, False).ToBitmap : Catch ex As Exception : cmd_btn4.BackgroundImage = My.Resources.RunDialogIcon : End Try
        Try : If AppBar.UseExplorerFM = True Then : cmd_btn3.BackgroundImage = GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\explorer.exe", 1, False).ToBitmap : Else : If File.Exists(AppBar.CustomFMPath) Then : cmd_btn3.BackgroundImage = Icon.ExtractAssociatedIcon(AppBar.CustomFMPath).ToBitmap : Else : cmd_btn3.BackgroundImage = My.Resources.FileExplorerSmall : End If : End If : Catch ex As Exception : cmd_btn3.BackgroundImage = My.Resources.FileExplorerSmall : End Try
        Try : cmd_btn2.BackgroundImage = GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\system32\imageres.dll", 22, False).ToBitmap : Catch ex As Exception : cmd_btn2.BackgroundImage = My.Resources.Settings : End Try
        Try : cmd_btn1.BackgroundImage = GetIcon(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\system32\shell32.dll", 112, False).ToBitmap : Catch ex As Exception : cmd_btn1.BackgroundImage = My.Resources.ActionCenter : End Try

        Dim mainclr As Color = Color.FromArgb(20, 255, 255, 255)
        Dim mainclr1 As Color = Color.FromArgb(20, 210, 210, 210)

        'cmd_btn1.BackColor = mainclr
        'cmd_btn2.BackColor = mainclr1
        'cmd_btn3.BackColor = mainclr
        'cmd_btn4.BackColor = mainclr1

        ' Normal

        SetForegroundWindow(Me.Handle)
    End Sub

    Public Sub ApplyStartTheme()
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
    End Sub

    Public Function GetUserAccountPicture() As Image
        Try
            Dim imagePath As String = GetCurrentUserImagePath()

            If Not String.IsNullOrEmpty(imagePath) AndAlso File.Exists(imagePath) Then
                Return Image.FromFile(imagePath)
            Else
                Dim defaultPath As String = "C:\ProgramData\Microsoft\User Account Pictures\user.png"
                If File.Exists(defaultPath) Then
                    Return Image.FromFile(defaultPath)
                Else
                    Return Image.FromFile(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\AppData\Local\Temp\" & Environment.UserName & ".bmp")
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine("Error loading profile image: " & ex.Message)
            Return My.Resources.DefaultUserPicture
        End Try
    End Function

    Public Sub ReloadTiles()
        ToolStrip1.Items.Clear()

        Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles", True)

        If myKey IsNot Nothing Then
            Try
                For Each i As String In myKey.GetSubKeyNames

                    ' Registry values
                    Dim fullPath As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu\Tiles\" & i, "Program", String.Empty)
                    Dim args As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu\Tiles\" & i, "Arguments", String.Empty)

                    Dim order As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu\Tiles\" & i, "Order", 0)

                    Dim sizeModeW As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu\Tiles\" & i, "SizeTypeW", 1)
                    Dim sizeModeH As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu\Tiles\" & i, "SizeTypeH", 1)

                    'When I'll know how.
                    'Dim borderType As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu\Tiles\" & i, "BorderType", 0)

                    Dim backColor As Color = Color.Transparent
                    Dim imageLayout = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu\Tiles\" & i, "ImageLayout", 1)

                    Try
                        backColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu\Tiles\" & i, "BackColor", Color.Transparent))
                    Catch ex As Exception
                        backColor = Color.Transparent
                    End Try

                    ' Where it is all doing!
                    If Not String.IsNullOrEmpty(fullPath) OrElse File.Exists(fullPath) Then

                        Dim FI As New FileInfo(fullPath)
                        Dim PI As New ToolStripButton(i)

                        With PI ' Settings and Tile appearance

                            .ToolTipText = FI.FullName
                            .Tag = args

                            ' Style
                            .AutoSize = False
                            .TextAlign = ContentAlignment.BottomCenter
                            .TextImageRelation = TextImageRelation.Overlay
                            .ImageAlign = ContentAlignment.MiddleCenter

                            If Me.BackColor.R >= 150 OrElse Me.BackColor.G >= 130 AndAlso Me.BackColor.B >= 130 Then .ForeColor = Color.Black Else .ForeColor = Color.White

                            .BackColor = backColor

                            ' Image/Icon
                            Try
                                Dim bmp As Bitmap = Desktop.GetOptimizedIcon(FI.FullName, 48)

                                If bmp IsNot Nothing Then .Image = bmp Else .Image = My.Resources.ProgramMedium
                            Catch ex As Exception
                                .Image = My.Resources.ProgramMedium
                            End Try

                            If TypeOf imageLayout IsNot Integer Then
                                ' Image Layout....
                            End If

                            ' Sizing
                            Dim SMextraSmall As Integer = 48
                            Dim SMSmall As Integer = 96
                            Dim SMMedium As Integer = 144
                            Dim SMLarge As Integer = 192

                            Select Case sizeModeW ' Width
                                Case 0 ' Extra Small
                                    .Width = SMextraSmall
                                    .DisplayStyle = ToolStripItemDisplayStyle.Image
                                Case 1 ' Small
                                    .Width = SMSmall
                                    .DisplayStyle = ToolStripItemDisplayStyle.ImageAndText
                                Case 2 ' Medium
                                    .Width = SMMedium
                                    .DisplayStyle = ToolStripItemDisplayStyle.ImageAndText
                                Case 3 ' Large
                                    .Width = SMLarge
                                    .DisplayStyle = ToolStripItemDisplayStyle.ImageAndText

                            End Select : Select Case sizeModeH ' Height

                                Case 0 ' Extra Small
                                    .Height = SMextraSmall
                                    .DisplayStyle = ToolStripItemDisplayStyle.Image
                                Case 1 ' Small
                                    .Height = SMSmall
                                    .DisplayStyle = ToolStripItemDisplayStyle.ImageAndText
                                Case 2 ' Medium
                                    .Height = SMMedium
                                    .DisplayStyle = ToolStripItemDisplayStyle.ImageAndText
                                Case 3 ' Large
                                    .Height = SMLarge
                                    .DisplayStyle = ToolStripItemDisplayStyle.ImageAndText

                            End Select

                            ' Handlers
                            AddHandler .MouseUp, AddressOf TileClick

                        End With

                        ' Adding the items

                        ToolStrip1.Items.Add(PI)
                    End If
                Next
            Finally
                myKey.Close()
            End Try
        End If
    End Sub

    Private Sub TileClick(ByVal sender As System.Object, ByVal e As Windows.Forms.MouseEventArgs)

        Dim regName As String = CType(sender, ToolStripButton).Text
        Dim program As String = CType(sender, ToolStripButton).ToolTipText
        Dim args As String = CType(sender, ToolStripButton).Tag

        Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles\" & regName, True)

        If myKey IsNot Nothing Then
            If e.Button = MouseButtons.Right Then

                TileCM.Tag = regName
                TileCM.Show(Control.MousePosition)

            ElseIf e.Button = MouseButtons.Left Then

                If File.Exists(program) = True Then

                    Dim SI As New ProcessStartInfo(program) With {
                        .Arguments = args,
                        .UseShellExecute = True
                    }

                    Process.Start(SI)
                Else
                    Process.Start(program)
                End If
            End If
        End If
    End Sub

    Private Sub PopulateTreeView(ByVal directoryPath As String, ByVal targetNodeCollection As TreeNodeCollection)
        Dim cutoff As Integer = 20

        Try ' Directories
            Dim subDirectories() As String = Directory.GetDirectories(directoryPath)

            For Each dirPath As String In subDirectories
                Dim directoryInfo As New DirectoryInfo(dirPath)
                Dim folderNode As New TreeNode(directoryInfo.Name)

                If directoryInfo.Name.Length > cutoff Then
                    folderNode.Text = directoryInfo.Name.Substring(0, cutoff - 3) & "..."
                End If

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

        Try ' Files
            Dim files() As String = Directory.GetFiles(directoryPath)

            For Each filePath As String In files
                Dim fileInfo As New FileInfo(filePath)
                Dim fileNode As New TreeNode(fileInfo.Name)

                If Not fileInfo.Name = "desktop.ini" Then
                    Dim iconIndex As Integer = GetFileIconIndex(fileInfo.FullName, treeImageList, True)

                    If fileInfo.Extension.ToLower = ".lnk" Then
                        fileNode.Text = Path.GetFileNameWithoutExtension(filePath)

                        Try
                            fileNode.ToolTipText = AppBar.GetShortcutTarget(fileInfo.FullName)
                        Catch ex As Exception
                            fileNode.ToolTipText = fileInfo.FullName
                        End Try
                    Else
                        fileNode.ToolTipText = fileInfo.FullName
                    End If

                    If fileInfo.Name.Length > cutoff Then
                        fileNode.Text = fileInfo.Name.Substring(0, cutoff - 3) & "..."
                    End If

                    fileNode.ImageIndex = iconIndex
                    fileNode.SelectedImageIndex = iconIndex
                    fileNode.Tag = filePath

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

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_btn3.Click
        On Error Resume Next

        If AppBar.UseExplorerFM = True Then
            Process.Start("explorer.exe", "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}")
        Else
            If Not String.IsNullOrWhiteSpace(AppBar.CustomFMPath) AndAlso File.Exists(AppBar.CustomFMPath) Then
                Process.Start(AppBar.CustomFMPath)
            End If
        End If

        canClose = True
        Me.Close()
    End Sub
    Public StartDefaultLocation As New Point(0 - SystemInformation.BorderSize.Width, My.Computer.Screen.Bounds.Height - Me.Height - AppBar.Height + SystemInformation.BorderSize.Height)
    Private Sub Startmenu_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged
        If Not Me.Location = StartDefaultLocation Then
            Me.Location = StartDefaultLocation
        End If
    End Sub

    Private Sub Startmenu_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        If Not Me.WindowState = FormWindowState.Normal Then Me.WindowState = FormWindowState.Normal

        StartDefaultLocation = New Point(0 - SystemInformation.BorderSize.Width, My.Computer.Screen.Bounds.Height - Me.Height - AppBar.Height + SystemInformation.BorderSize.Height)
    End Sub

    Private Sub Startmenu_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        If Me.Visible = True Then

            If TreeView2.Nodes.Count = 0 Then
                Select Case ComboBox1.SelectedIndex
                    Case 0
                        LoadItemsIntoTreeView(StartMenuFolder)
                    Case Else
                        LoadItemsIntoTreeView(StartMenuFolderPublic)
                End Select
            End If

            usr_Picture.BackgroundImage = GetUserAccountPicture()
            ApplyStartTheme()
        Else
            ' When Startmenu is not visible
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
                targetPath = StartMenuFolder
            Case 1
                targetPath = StartMenuFolderPublic
            Case Else
                targetPath = TreeView2.SelectedNode.Tag & "\"
        End Select

        Dim newFolder As New TreeNode("New Folder")
        With newFolder
            .ImageIndex = 0
            .Tag = "[NewFolder]"
        End With

        If targetPath IsNot Nothing AndAlso Directory.Exists(targetPath) Then
            TreeView2.SelectedNode.Expand()
            TreeView2.SelectedNode.Nodes.Add(newFolder)

            lastNodeTag = targetPath

            TreeView2.SelectedNode = newFolder
            TreeView2.SelectedNode.BeginEdit()
        End If
    End Sub

    Private Sub ts5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ts5.Click
        If MessageBox.Show("Are you sure do you want remove " & TreeView2.SelectedNode.Tag & " ?", "Delete Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then

            Dim targetPath As String = TreeView2.SelectedNode.Tag

            Try
                If IO.Directory.Exists(targetPath) = True Then
                    My.Computer.FileSystem.DeleteDirectory(targetPath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                    TreeView2.SelectedNode.Remove()
                ElseIf IO.File.Exists(targetPath) = True Then
                    My.Computer.FileSystem.DeleteFile(targetPath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin, FileIO.UICancelOption.DoNothing)
                    TreeView2.SelectedNode.Remove()
                Else
                    MessageBox.Show("Failed to delete: """ & targetPath & """. Check if the File/Directory still exists.", "File/Directory not found", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
                LoadItemsIntoTreeView(StartMenuFolder)
            Case Else
                LoadItemsIntoTreeView(StartMenuFolderPublic)
        End Select
    End Sub

    Private Sub Button3_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles cmd_btn3.MouseUp
        If e.Button = Windows.Forms.MouseButtons.Right Then
            ContextMenu1.Show(Me, Control.MousePosition, LeftRightAlignment.Left)
        End If
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles cmd_btn4.Click
        On Error Resume Next
        RunDialog.Show(Me)

        Task.Delay(100).ContinueWith(Sub() Me.Hide())
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

        Dim destPath As String = TreeView2.SelectedNode.Tag

        If AppBar.UseExplorerFM = True Then
            Process.Start("explorer.exe", "/select," & destPath)
        Else
            If Not String.IsNullOrWhiteSpace(AppBar.CustomFMPath) AndAlso File.Exists(AppBar.CustomFMPath) Then

                If File.Exists(destPath) Then
                    Dim FI As New FileInfo(destPath)
                    Process.Start(AppBar.CustomFMPath, FI.DirectoryName)

                ElseIf Directory.Exists(destPath) Then
                    Process.Start(AppBar.CustomFMPath, """" & destPath & """")
                End If
            End If
        End If
    End Sub

    Private Sub ts4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ts4.Click
        TreeView2.SelectedNode.BeginEdit()
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click

        Dim regName As String = TileCM.Tag

        Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles\" & regName, True)

        If myKey IsNot Nothing Then

            Dim program As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu\Tiles\" & regName, "Program", String.Empty)
            Dim args As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu\Tiles\" & regName, "Arguments", String.Empty)

            If File.Exists(program) = True Then

                Dim SI As New ProcessStartInfo(program) With {
                        .Arguments = args,
                        .UseShellExecute = True
                    }

                Process.Start(SI)
            Else
                Process.Start(program)
            End If
        End If
    End Sub

    Private Sub RefleshToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RefleshToolStripMenuItem1.Click
        ReloadTiles()
        Me.Invalidate()
    End Sub

    Private Sub LocateInExplorerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LocateInExplorerToolStripMenuItem.Click
        On Error Resume Next

        Dim targetPath As String
        If ComboBox1.SelectedIndex = 0 Then targetPath = StartMenuFolder Else targetPath = StartMenuFolderPublic

        Dim pr As New ProcessStartInfo("", targetPath)
        If AppBar.UseExplorerFM Then
            pr.FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "explorer.exe")
        Else
            pr.FileName = AppBar.CustomFMPath
        End If

        Process.Start(pr)
    End Sub

    Private Sub ShortcutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShortcutToolStripMenuItem.Click

        If TreeView2.SelectedNode IsNot Nothing Then
            LinkCreator.PATH = TreeView2.SelectedNode.Tag & "\"
        Else
            If ComboBox1.SelectedIndex = 0 Then LinkCreator.PATH = StartMenuFolder Else LinkCreator.PATH = StartMenuFolderPublic
        End If

        If LinkCreator.ShowDialog = Windows.Forms.DialogResult.OK Then
            Select Case ComboBox1.SelectedIndex
                Case 0
                    LoadItemsIntoTreeView(StartMenuFolder)
                Case Else
                    LoadItemsIntoTreeView(StartMenuFolderPublic)
            End Select

            TreeView2.Nodes.Find(LinkCreator.Result, True)
        End If
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles cmd_btn2.Click
        On Error Resume Next

        Dim settingsPath As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "Settings", "control.exe")

        If Not String.IsNullOrWhiteSpace(settingsPath) AndAlso File.Exists(settingsPath) Then
            Process.Start(New ProcessStartInfo(settingsPath) With {.UseShellExecute = True})
        End If

        canClose = True
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles cmd_btn1.Click
        On Error Resume Next

        SA.ShowDialog(AppBar)
    End Sub

    'If MessageBox.Show("This will log off your Windows session, Are you sure do you want log off?", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then

    Private Sub AccountSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AccountSettingsToolStripMenuItem.Click
        On Error Resume Next
        Process.Start(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "Settings", "control"), "/name Microsoft.UserAccounts")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles usr_Picture.Click
        LoginMenuS.Show(Control.MousePosition)
    End Sub

    <DllImport("shell32.dll", EntryPoint:="#60", CharSet:=CharSet.Auto)>
    Private Shared Function ShowClassicShutdownDialog(hwndOwner As IntPtr) As UInt32
    End Function

    Private Sub ShutdownActionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShutdownActionsToolStripMenuItem.Click
        ShowClassicShutdownDialog(AppBar.Handle)
    End Sub

    Private Sub PowerSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PowerSettingsToolStripMenuItem.Click
        On Error Resume Next
        Process.Start(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "Settings", "control"), "/name Microsoft.PowerOptions")
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        If MessageBox.Show("Are you sure do you want delete this Tile """ & TileCM.Tag & """?", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then

            Dim regName As String = TileCM.Tag
            Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles", True)

            If myKey IsNot Nothing Then
                myKey.DeleteSubKey(regName)

                ReloadTiles()
            End If

            'IO.Directory.Delete(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & TileCM.Tag, True)
        End If
    End Sub

    Private Sub RenameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RenameToolStripMenuItem.Click
        On Error Resume Next

        Dim regName As String = TileCM.Tag

        Dim s As String = InputBox("Type a new name for the tile to be renamed.", "Rename Dialog", regName)
        If String.IsNullOrWhiteSpace(s) Then Exit Sub

        Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles\" & regName, True)

        Dim oldPath As String = "Software\Shell\StartMenu\Tiles\" & regName
        Dim newPath As String = "Software\Shell\StartMenu\Tiles\" & s

        If myKey IsNot Nothing Then

            Using newKey As RegistryKey = myKey.CreateSubKey(s)
                CopyKey(Registry.CurrentUser, oldPath, newPath)
            End Using

            Registry.CurrentUser.DeleteSubKeyTree(oldPath)

            ReloadTiles()
        End If

        'My.Computer.FileSystem.RenameDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned\" & TileCM.Tag, s)
    End Sub

    Private Sub CopyKey(root As RegistryKey, sourcePath As String, destPath As String)
        Using sourceKey As RegistryKey = root.OpenSubKey(sourcePath)
            Using destKey As RegistryKey = root.CreateSubKey(destPath)
                For Each valueName As String In sourceKey.GetValueNames()
                    destKey.SetValue(valueName, sourceKey.GetValue(valueName), sourceKey.GetValueKind(valueName))
                Next

                For Each subKeyName As String In sourceKey.GetSubKeyNames()
                    CopyKey(sourceKey, subKeyName, destPath & "\" & subKeyName)
                Next
            End Using
        End Using
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
                LoadItemsIntoTreeView(StartMenuFolder)
                TreeView2.Tag = StartMenuFolder
            Case Else
                LoadItemsIntoTreeView(StartMenuFolderPublic)
                TreeView2.Tag = StartMenuFolderPublic
        End Select
    End Sub

    Private Sub Startmenu_BackColorChanged(sender As Object, e As EventArgs) Handles Me.BackColorChanged
        Dim luminance As Double = (0.299 * Me.BackColor.R + 0.587 * Me.BackColor.G + 0.114 * Me.BackColor.B) / 255

        Dim contrastColor As Color = If(luminance > 0.5, Color.Black, Color.White)

        For Each ctrl As Control In Me.Controls
            ctrl.ForeColor = contrastColor
        Next
    End Sub

    Public Function ColorMixer(Color1 As Color, Color2 As Color) As Color
        Dim r1 As Integer = Color1.R
        Dim g1 As Integer = Color1.G
        Dim b1 As Integer = Color1.B

        Dim r2 As Integer = Color2.R
        Dim g2 As Integer = Color2.G
        Dim b2 As Integer = Color2.B

        Dim newR As Integer = (r1 + r2) \ 2
        Dim newG As Integer = (g1 + g2) \ 2
        Dim newB As Integer = (b1 + b2) \ 2

        Return Color.FromArgb(newR, newG, newB)
    End Function

    Private Sub PropertiesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PropertiesToolStripMenuItem.Click
        StartmenuProperties.ShowDialog(Me)
    End Sub

    Public startColor As Color = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "BackColor", "#000000"))
    Public endColor As Color = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "BackColor2", "#000000"))

    Private Sub Startmenu_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        Dim mode As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "Mode", "2")
        If mode <> 2 Then Exit Sub

        Dim useSystemColor As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "UseSystemColor", "1")
        If useSystemColor = 1 Then
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
    End Sub

    Private Sub SelectControl_Tick(sender As Object, e As EventArgs) Handles SelectControl.Tick
        Try
            Dim targetPath As String = TreeView2.SelectedNode.Tag

            If File.Exists(targetPath) Then
                Dim bitmapIcon As Bitmap = Desktop.GetOptimizedIcon(targetPath, 52)

                If bitmapIcon Is Nothing Then
                    bitmapIcon = Icon.ExtractAssociatedIcon(targetPath).ToBitmap
                End If

                usr_Picture.BackgroundImage = bitmapIcon

                'bitmapIcon.Dispose()
            Else
                usr_Picture.BackgroundImage = GetUserAccountPicture()
            End If
        Catch ex As Exception
            usr_Picture.BackgroundImage = GetUserAccountPicture()
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
            If Directory.Exists(e.Node.Tag) AndAlso e.Node.Nodes.Count > 0 Then

                If e.Node.IsExpanded = True Then e.Node.Collapse() Else e.Node.Expand()

            ElseIf File.Exists(e.Node.Tag) Then

                SelectControl.Enabled = True

                Try
                    Dim pr As New ProcessStartInfo(e.Node.Tag)

                    Dim prr As Process = Process.Start(pr)
                    prr.WaitForInputIdle()

                    Me.Visible = False
                Catch ex As Exception

                End Try
            End If
        End If
    End Sub

    Private Sub TreeView2_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView2.AfterSelect
        If File.Exists(e.Node.Tag) Then SelectControl.Enabled = True
    End Sub

    Private Sub OpenFileLocationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenFileLocationToolStripMenuItem.Click

        Dim regName As String = TileCM.Tag

        Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles\" & regName, True)

        If myKey IsNot Nothing Then

            Dim program As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu\Tiles\" & regName, "Program", String.Empty)

            If File.Exists(program) = True Then
                If AppBar.UseExplorerFM = True Then
                    Process.Start("explorer.exe", " /select," & program)
                Else
                    If Not String.IsNullOrWhiteSpace(AppBar.CustomFMPath) AndAlso File.Exists(AppBar.CustomFMPath) Then
                        Dim FI As New FileInfo(program)

                        Process.Start(AppBar.CustomFMPath, """" & FI.DirectoryName & """")
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub CustomToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CustomToolStripMenuItem.Click
        Dim regName As String = TileCM.Tag

        Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles\" & regName, True)

        If myKey IsNot Nothing Then

            Dim backColor As Color = Color.Transparent

            Try
                backColor = ColorTranslator.FromHtml(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu\Tiles\" & regName, "BackColor", Color.Transparent))
            Catch ex As Exception
                backColor = Color.Transparent
            End Try

            Dim CD As New ColorDialog With {
            .AllowFullOpen = True,
            .AnyColor = True,
            .FullOpen = True
        }

            If Not backColor = Color.Transparent Then
                CD.Color = backColor
            End If

            If CD.ShowDialog(Me) = DialogResult.OK Then
                myKey.SetValue("BackColor", ColorTranslator.ToHtml(CD.Color))

                ReloadTiles()
            End If

            myKey.Close()
        End If
    End Sub

    Private Sub SetToTransparentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SetToTransparentToolStripMenuItem.Click
        Dim regName As String = TileCM.Tag

        Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles\" & regName, True)

        If myKey IsNot Nothing Then
            myKey.DeleteValue("BackColor")

            ReloadTiles()

            myKey.Close()
        End If
    End Sub

    Private Sub ExtraSmallToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExtraSmallToolStripMenuItem.Click
        Dim regName As String = TileCM.Tag

        Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles\" & regName, True)

        If myKey IsNot Nothing Then
            myKey.SetValue("SizeTypeW", 0, RegistryValueKind.DWord)

            ReloadTiles()

            myKey.Close()
        End If
    End Sub

    Private Sub Small9696ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Small9696ToolStripMenuItem.Click
        Dim regName As String = TileCM.Tag

        Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles\" & regName, True)

        If myKey IsNot Nothing Then
            myKey.SetValue("SizeTypeW", 1, RegistryValueKind.DWord)

            ReloadTiles()

            myKey.Close()
        End If
    End Sub

    Private Sub Medium19296ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Medium19296ToolStripMenuItem.Click
        Dim regName As String = TileCM.Tag

        Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles\" & regName, True)

        If myKey IsNot Nothing Then
            myKey.SetValue("SizeTypeW", 2, RegistryValueKind.DWord)

            ReloadTiles()

            myKey.Close()
        End If
    End Sub

    Private Sub Big192192ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Big192192ToolStripMenuItem.Click
        Dim regName As String = TileCM.Tag

        Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles\" & regName, True)

        If myKey IsNot Nothing Then
            myKey.SetValue("SizeTypeW", 3, RegistryValueKind.DWord)

            ReloadTiles()

            myKey.Close()
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        Dim regName As String = TileCM.Tag

        Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles\" & regName, True)

        If myKey IsNot Nothing Then
            myKey.SetValue("SizeTypeH", 0, RegistryValueKind.DWord)

            ReloadTiles()

            myKey.Close()
        End If
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        Dim regName As String = TileCM.Tag

        Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles\" & regName, True)

        If myKey IsNot Nothing Then
            myKey.SetValue("SizeTypeH", 1, RegistryValueKind.DWord)

            ReloadTiles()

            myKey.Close()
        End If
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Dim regName As String = TileCM.Tag

        Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles\" & regName, True)

        If myKey IsNot Nothing Then
            myKey.SetValue("SizeTypeH", 2, RegistryValueKind.DWord)

            ReloadTiles()

            myKey.Close()
        End If
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        Dim regName As String = TileCM.Tag

        Dim myKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\StartMenu\Tiles\" & regName, True)

        If myKey IsNot Nothing Then
            myKey.SetValue("SizeTypeH", 3, RegistryValueKind.DWord)

            ReloadTiles()

            myKey.Close()
        End If
    End Sub

    Private Sub LoginMenuS_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles LoginMenuS.Opening
        LoginMenuS.Items.Clear()

        Dim sessionActions As String() = {"Lock", "Switch User", "Log Off"}

        For Each act In sessionActions
            Dim item As New ToolStripMenuItem(act)
            AddHandler item.Click, AddressOf SessionButtonPressed
            LoginMenuS.Items.Add(item)
        Next

        ' Final items
        LoginMenuS.Items.Add(New ToolStripSeparator)
        LoginMenuS.Items.Add(AccountSettingsToolStripMenuItem)
    End Sub

    Private Sub PowerMenuS_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles PowerMenuS.Opening
        PowerMenuS.Items.Clear()

        Dim powerActions = PowerManager.GetAvailablePowerActions

        For Each act In powerActions
            Dim item As New ToolStripMenuItem(act)
            AddHandler item.Click, AddressOf PowerButtonPressed
            PowerMenuS.Items.Add(item)
        Next

        ' Final items
        PowerMenuS.Items.Add(New ToolStripSeparator)
        PowerMenuS.Items.Add(ShutdownActionsToolStripMenuItem)
        PowerMenuS.Items.Add(New ToolStripSeparator)
        PowerMenuS.Items.Add(PowerSettingsToolStripMenuItem)
    End Sub

    Private Sub PowerButtonPressed(sender As Object, e As EventArgs)
        Dim item = DirectCast(sender, ToolStripMenuItem)

        Dim isBrutalMode As Boolean = False
        Dim isForceMode As Boolean = False

        If (ModifierKeys And Keys.Control) = Keys.Control Then
            isBrutalMode = True ' EWX_FORCE
        ElseIf SA.chkForceMode.Checked Then
            isForceMode = True ' EWX_FORCEIFHUNG
        End If

        PowerManager.ExecuteAction(item.Text, isForceMode, 0, PowerManager.IsFastStartupEnabled, SA.chkDisableWakeEvent.Checked, isBrutalMode)
    End Sub

    Private Sub SessionButtonPressed(sender As Object, e As EventArgs)
        Dim item = DirectCast(sender, ToolStripMenuItem)
        PowerManager.ExecuteSessionAction(item.Text)
    End Sub
End Class