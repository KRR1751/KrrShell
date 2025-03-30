Imports System.Drawing.Imaging
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Policy
Imports System.Web
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
    Private Sub Startmenu_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        'AppBar.Button1.BackgroundImage = AppBar.Label10.Image
        Me.Visible = True
    End Sub
    Public FST As Boolean = False
    Private Sub Startmenu_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        If FST = False Then
            'AppBar.Button1.BackgroundImage = AppBar.Label8.Image
            Me.Visible = False
            AppBar.Button1.BackgroundImage = My.Resources._1
        End If
    End Sub

    Private Sub Startmenu_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.None OrElse e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
        End If
    End Sub

    Private Sub Startmenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.BackColor = color.Black
            Dim side As side = side
            side.Left = -1
            side.Right = -1
            side.Top = -1
            side.Bottom = -1
            Dim result As Integer = DwmExtendFrameIntoClientArea(Me.Handle, side)
        Catch ex As Exception

        End Try

        Dim colorr As UInteger
        Dim opaqueBlend As Boolean
        DwmGetColorizationColor(colorr, opaqueBlend)
        Dim systemColor As Color = ColorTranslator.FromWin32(colorr)
        TreeView1.BackColor = systemColor
        TreeView2.BackColor = systemColor

        Me.Location = New Point(0, My.Computer.Screen.Bounds.Height - Me.Height - AppBar.Height)
        LoadRootDrives()
        LoadRootDrivesAllUsers()
        ComboBox1.SelectedIndex = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "DefaultFolder", "0")
        Try
            Me.Width = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\", "StartLastWidth", Nothing)
            Me.Height = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\", "StartLastHeight", Nothing)
        Catch ex As Exception
            Me.Size = New Size(500, 560)
        End Try
        ToolStrip1.Items.Clear()
        Dim DirectoryInformation As New DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned")
        For Each i As DirectoryInfo In DirectoryInformation.GetDirectories
            Dim item As New ToolStripMenuItem(i.Name)
            With item
                .AutoSize = False
                .Text = i.Name
                .TextAlign = ContentAlignment.BottomCenter
                .TextImageRelation = TextImageRelation.Overlay
                .ImageAlign = ContentAlignment.MiddleCenter
                .ForeColor = Color.White
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

        'Account Picture Detection
        PictureBox1.Image = Image.FromFile(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\AppData\Local\Temp\" & Environment.UserName & ".bmp")

    End Sub

    Public Sub ReloadTiles()
        ToolStrip1.Items.Clear()
        Dim DirectoryInformation As New DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\Pinned")
        For Each i As DirectoryInfo In DirectoryInformation.GetDirectories
            Dim item As New ToolStripMenuItem(i.Name)
            With item
                .AutoSize = False
                .Text = i.Name
                .TextAlign = ContentAlignment.BottomCenter
                .TextImageRelation = TextImageRelation.Overlay
                .ImageAlign = ContentAlignment.MiddleCenter
                .ForeColor = Color.White
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
            'Shell(IO.File.ReadAllText(DirectoryInformation.FullName & "\" & CType(sender, ToolStripMenuItem).Text & "\Program"))
        End If
    End Sub

    Public Sub LoadRootDrives()
        TreeView1.Nodes.Clear()
        Dim DirectoryInformation As New DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\")
        For Each d As DirectoryInfo In DirectoryInformation.GetDirectories 'DriveInfo In DriveInfo.GetDrives
            Dim node As New TreeNode(d.Name)
            With node
                .Tag = d.FullName
                .ImageKey = "folder"
                .ImageIndex = 4
                .SelectedImageIndex = 5
                .Nodes.Add("Empty")
            End With

            If Not node.Text = "Pinned" Then
                TreeView1.Nodes.Add(node)
            End If
        Next

        For Each f As FileInfo In DirectoryInformation.GetFiles 'DriveInfo In DriveInfo.GetDrives
            Dim node As New TreeNode(f.Name)
            With node
                .Tag = f.FullName
                .ImageKey = "file"
                .ImageIndex = 3
                .SelectedImageIndex = 3
            End With

            If Not node.Text = "desktop.ini" Then
                TreeView1.Nodes.Add(node)
            End If
        Next
    End Sub

    Public Sub LoadChildren(ByVal nd As TreeNode, ByVal dir As String)
        Dim DirectoryInformation As New DirectoryInfo(dir)

        'ListView1.Items.Clear()
        Dim SubItems() As ListViewItem.ListViewSubItem
        Dim Item As ListViewItem = Nothing

        Try
            'Load all sub folders into the node
            For Each d As DirectoryInfo In DirectoryInformation.GetDirectories

                If Not (d.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then
                    Dim folder As New TreeNode(d.Name)

                    With folder
                        .Tag = d.FullName
                        .ImageIndex = 5
                        .SelectedImageIndex = 4
                        .Nodes.Add("*Empty*")
                    End With

                    nd.Nodes.Add(folder)

                    Item = New ListViewItem(d.Name, 4)
                    SubItems = New ListViewItem.ListViewSubItem() {New ListViewItem.ListViewSubItem(Item, 0), New ListViewItem.ListViewSubItem(Item, "Directory"), New ListViewItem.ListViewSubItem(Item, d.LastAccessTime.ToShortDateString())}

                    Item.SubItems.AddRange(SubItems)
                    'ListView1.Items.Add(Item)
                End If

            Next

            'load all files into the child nodes
            For Each file As FileInfo In DirectoryInformation.GetFiles
                If Not (file.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then
                    Dim FN As New TreeNode(file.Name)
                    With FN
                        .Tag = file.FullName
                        .ImageIndex = 3
                        .SelectedImageKey = 3
                    End With
                    nd.Nodes.Add(FN)

                    Dim file_type As String = "File"
                    Dim file_icon As Integer = 3
                    Select Case file.FullName.Split(".").ToString
                        'System Files
                        Case "dll"
                            file_type = "Dynamic Link Library"
                            file_icon = 10
                        Case "sys"
                            file_type = "System File"
                            file_icon = 10
                        Case "exe"
                            file_type = "Executable"
                            file_icon = 2
                        Case "jar"
                            file_type = "Executable"
                            file_icon = 2
                        Case "dat"
                            file_type = "Data File"

                        Case "ini"
                            file_type = "INI File"
                            file_icon = 6
                        Case "bat"
                            file_type = "Batch File"
                            file_icon = 13
                        Case "cmd"
                            file_type = "Command File"
                            file_icon = 14

                            'Text Files
                        Case "txt"
                            file_type = "Document"
                            file_icon = 1
                        Case "html"
                            file_type = "Document"
                            file_icon = 1
                        Case "css"
                            file_type = "Document"
                            file_icon = 1
                        Case "rtf"
                            file_type = "Document"
                            file_icon = 1
                        Case "text"
                            file_type = "Document"
                            file_icon = 1
                        Case "log"
                            file_type = "Document"
                            file_icon = 1
                        Case "yml"
                            file_type = "Document"
                            file_icon = 1
                        Case "xml"
                            file_type = "Document"
                            file_icon = 1

                            'Compression Formats
                        Case "zip"
                            file_type = "Compressed File"
                            file_icon = 0
                        Case "rar"
                            file_type = "Compressed File"
                            file_icon = 15
                        Case "7z"
                            file_type = "Compressed File"
                            file_icon = 0
                        Case "pak"
                            file_type = "Compressed File"
                            file_icon = 0
                        Case "rpf"
                            file_type = "Compressed File"
                            file_icon = 0

                            'Image Formats
                        Case "bin"
                            file_type = "System Image"
                            file_icon = 9
                        Case "iso"
                            file_type = "System Image"
                            file_icon = 9
                        Case "img"
                            file_type = "System Image"
                            file_icon = 9
                        Case "dmg"
                            file_type = "System Image"
                            file_icon = 9

                            'Picture Formats
                        Case "bmp"
                            file_type = "Image"
                            file_icon = 8
                        Case "jpg"
                            file_type = "Image"
                            file_icon = 8
                        Case "png"
                            file_type = "Image"
                            file_icon = 8
                        Case "gif"
                            file_type = "Image"
                            file_icon = 8
                        Case "tiff"
                            file_type = "Image"
                            file_icon = 8
                        Case "jpeg"
                            file_type = "Image"
                            file_icon = 8
                        Case "ico"
                            file_type = "Image"
                            file_icon = 8
                        Case "jfif"
                            file_type = "Image"
                            file_icon = 8

                            'Video Formats
                        Case "mp4"
                            file_type = "Video"
                            file_icon = 12
                        Case "webm"
                            file_type = "Video"
                            file_icon = 12
                        Case "3gp"
                            file_type = "Video"
                            file_icon = 12
                        Case "m4v"
                            file_type = "Video"
                            file_icon = 12
                        Case "flv"
                            file_type = "Video"
                            file_icon = 12
                        Case "mpeg"
                            file_type = "Video"
                            file_icon = 12
                        Case "mpe"
                            file_type = "Video"
                            file_icon = 12
                        Case "mov"
                            file_type = "Video"
                            file_icon = 12
                        Case "swf"
                            file_type = "Video"
                            file_icon = 12
                        Case "wmv"
                            file_type = "video"
                            file_icon = 12

                            'Music Formats
                        Case "mp1"
                            file_type = "Music"
                            file_icon = 7
                        Case "mp2"
                            file_type = "Music"
                            file_icon = 7
                        Case "mp3"
                            file_type = "Music"
                            file_icon = 7
                        Case "mp4"
                            file_type = "Music"
                            file_icon = 7
                        Case "wav"
                            file_type = "Music"
                            file_icon = 7
                        Case "m4a"
                            file_type = "Music"
                            file_icon = 7
                        Case "flac"
                            file_type = "Music"
                            file_icon = 7
                        Case "wma"
                            file_type = "Music"
                            file_icon = 7
                        Case "ogg"
                            file_type = "Music"
                            file_icon = 7

                            'Font Formats
                        Case "ttf"
                            file_type = "Font File"
                            file_icon = 11
                        Case "ufo"
                            file_type = "Font File"
                            file_icon = 11
                        Case "fnt"
                            file_type = "Font File"
                            file_icon = 11
                        Case Else
                            file_type = "File"
                            file_icon = 11
                    End Select

                    Item = New ListViewItem(file.Name, file_icon)

                    Dim filesize As String = (file.Length / 1000000).ToString("N0") + " Mb"

                    SubItems = New ListViewItem.ListViewSubItem() {New ListViewItem.ListViewSubItem(Item, filesize.ToString), New ListViewItem.ListViewSubItem(Item, file_type.ToString()), New ListViewItem.ListViewSubItem(Item, file.LastAccessTime.ToShortDateString())}

                    Item.SubItems.AddRange(SubItems)
                    'ListView1.Items.Add(Item)
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub

    Private Sub TreeView1_AfterLabelEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) Handles TreeView1.AfterLabelEdit
        If Not e.Label = String.Empty Then
            If Directory.Exists(e.Node.Tag) Then
                My.Computer.FileSystem.RenameDirectory(e.Node.Tag, e.Label)
            ElseIf File.Exists(e.Node.Tag) Then
                My.Computer.FileSystem.RenameFile(e.Node.Tag, e.Label)
            End If
            e.Node.Tag = e.Node.LastNode.Tag & "\" & e.Label
            'LoadRootDrives()
        Else
            e.CancelEdit = True
        End If
    End Sub

    Private Sub TreeView1_BeforeExpand(ByVal sender As Object, ByVal e As TreeViewCancelEventArgs) Handles TreeView1.BeforeExpand

        If (e.Node.Tag <> "Desktop" AndAlso Not e.Node.Tag.Contains(":\")) OrElse Directory.Exists(e.Node.Tag) Then
            e.Node.Nodes.Clear()
            LoadChildren(e.Node, e.Node.Tag.ToString)

        ElseIf e.Node.ImageKey = "Dekstop" Then
            e.Node.Nodes.Clear()
        Else
            e.Cancel = True
            'MessageBox.Show("Error drive is empty, " + e.Node.ImageKey.ToString())
        End If

    End Sub

    Private Sub TreeView1_AfterSelect(ByVal sender As Object, ByVal e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        TreeView1.Tag = e.Node.Tag.ToString()
    End Sub

    Private Sub TreeView1_AfterCollapse(ByVal sender As Object, ByVal e As TreeViewEventArgs) Handles TreeView1.AfterCollapse
        e.Node.Nodes.Clear()
        e.Node.Nodes.Add("Empty")
    End Sub

    Private Sub TreeView1_DoubleClick(ByVal sender As Object, ByVal e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseDoubleClick
        If e.Button = MouseButtons.Left AndAlso File.Exists(e.Node.Tag.ToString) Then
            Try
                Process.Start(e.Node.Tag.ToString())
            Catch ex As Exception
                ex = Nothing
            End Try
        End If
    End Sub

    Private Sub TreeView1_NodeMouseClick(ByVal sender As Object, ByVal e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        Dim DirectoryInformation As New DirectoryInfo(e.Node.Tag.ToString())

        'ListView1.Items.Clear()
        Dim SubItems() As ListViewItem.ListViewSubItem
        Dim Item As ListViewItem = Nothing

        Try
            'Load all sub folders into the node
            For Each d As DirectoryInfo In DirectoryInformation.GetDirectories

                If Not (d.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then
                    Item = New ListViewItem(d.Name, 4)
                    SubItems = New ListViewItem.ListViewSubItem() {New ListViewItem.ListViewSubItem(Item, 0), New ListViewItem.ListViewSubItem(Item, "Directory"), New ListViewItem.ListViewSubItem(Item, d.LastAccessTime.ToShortDateString())}

                    Item.SubItems.AddRange(SubItems)
                    'ListView1.Items.Add(Item)
                End If

            Next

            'load all files into the child nodes
            For Each file As FileInfo In DirectoryInformation.GetFiles
                If Not (file.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then
                    Dim file_type As String = "File"
                    Dim file_icon As Integer = 3
                    Select Case file.FullName.Split(".").ToString
                        'System Files
                        Case "dll"
                            file_type = "Dynamic Link Library"
                            file_icon = 10
                        Case "sys"
                            file_type = "System File"
                            file_icon = 10
                        Case "exe"
                            file_type = "Executable"
                            file_icon = 2
                        Case "jar"
                            file_type = "Executable"
                            file_icon = 2
                        Case "dat"
                            file_type = "Data File"

                        Case "ini"
                            file_type = "INI File"
                            file_icon = 6
                        Case "bat"
                            file_type = "Batch File"
                            file_icon = 13
                        Case "cmd"
                            file_type = "Command File"
                            file_icon = 14

                            'Text Files
                        Case "txt"
                            file_type = "Document"
                            file_icon = 1
                        Case "html"
                            file_type = "Document"
                            file_icon = 1
                        Case "css"
                            file_type = "Document"
                            file_icon = 1
                        Case "rtf"
                            file_type = "Document"
                            file_icon = 1
                        Case "text"
                            file_type = "Document"
                            file_icon = 1
                        Case "log"
                            file_type = "Document"
                            file_icon = 1
                        Case "yml"
                            file_type = "Document"
                            file_icon = 1
                        Case "xml"
                            file_type = "Document"
                            file_icon = 1

                            'Compression Formats
                        Case "zip"
                            file_type = "Compressed File"
                            file_icon = 0
                        Case "rar"
                            file_type = "Compressed File"
                            file_icon = 15
                        Case "7z"
                            file_type = "Compressed File"
                            file_icon = 0
                        Case "pak"
                            file_type = "Compressed File"
                            file_icon = 0
                        Case "rpf"
                            file_type = "Compressed File"
                            file_icon = 0

                            'Image Formats
                        Case "bin"
                            file_type = "System Image"
                            file_icon = 9
                        Case "iso"
                            file_type = "System Image"
                            file_icon = 9
                        Case "img"
                            file_type = "System Image"
                            file_icon = 9
                        Case "dmg"
                            file_type = "System Image"
                            file_icon = 9

                            'Picture Formats
                        Case "bmp"
                            file_type = "Image"
                            file_icon = 8
                        Case "jpg"
                            file_type = "Image"
                            file_icon = 8
                        Case "png"
                            file_type = "Image"
                            file_icon = 8
                        Case "gif"
                            file_type = "Image"
                            file_icon = 8
                        Case "tiff"
                            file_type = "Image"
                            file_icon = 8
                        Case "jpeg"
                            file_type = "Image"
                            file_icon = 8
                        Case "ico"
                            file_type = "Image"
                            file_icon = 8
                        Case "jfif"
                            file_type = "Image"
                            file_icon = 8

                            'Video Formats
                        Case "mp4"
                            file_type = "Video"
                            file_icon = 12
                        Case "webm"
                            file_type = "Video"
                            file_icon = 12
                        Case "3gp"
                            file_type = "Video"
                            file_icon = 12
                        Case "m4v"
                            file_type = "Video"
                            file_icon = 12
                        Case "flv"
                            file_type = "Video"
                            file_icon = 12
                        Case "mpeg"
                            file_type = "Video"
                            file_icon = 12
                        Case "mpe"
                            file_type = "Video"
                            file_icon = 12
                        Case "mov"
                            file_type = "Video"
                            file_icon = 12
                        Case "swf"
                            file_type = "Video"
                            file_icon = 12
                        Case "wmv"
                            file_type = "video"
                            file_icon = 12

                            'Music Formats
                        Case "mp1"
                            file_type = "Music"
                            file_icon = 7
                        Case "mp2"
                            file_type = "Music"
                            file_icon = 7
                        Case "mp3"
                            file_type = "Music"
                            file_icon = 7
                        Case "mp4"
                            file_type = "Music"
                            file_icon = 7
                        Case "wav"
                            file_type = "Music"
                            file_icon = 7
                        Case "m4a"
                            file_type = "Music"
                            file_icon = 7
                        Case "flac"
                            file_type = "Music"
                            file_icon = 7
                        Case "wma"
                            file_type = "Music"
                            file_icon = 7
                        Case "ogg"
                            file_type = "Music"
                            file_icon = 7

                            'Font Formats
                        Case "ttf"
                            file_type = "Font File"
                            file_icon = 11
                        Case "ufo"
                            file_type = "Font File"
                            file_icon = 11
                        Case "fnt"
                            file_type = "Font File"
                            file_icon = 11
                        Case Else
                            file_type = "File"
                            file_icon = 11
                    End Select
                    Item = New ListViewItem(file.Name, file_icon)

                    Dim filesize As String = (file.Length / 1000000).ToString("N0") + " Mb"

                    SubItems = New ListViewItem.ListViewSubItem() {New ListViewItem.ListViewSubItem(Item, filesize.ToString), New ListViewItem.ListViewSubItem(Item, file_type.ToString()), New ListViewItem.ListViewSubItem(Item, file.LastAccessTime.ToShortDateString())}

                    Item.SubItems.AddRange(SubItems)
                    'ListView1.Items.Add(Item)
                End If
            Next
        Catch ex As Exception
            ex = Nothing
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        System.Diagnostics.Process.Start("C:\Windows\explorer.exe", "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}")
    End Sub

    Private Sub Startmenu_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged
        If Not Me.Location = New Point(0, My.Computer.Screen.Bounds.Height - Me.Height - AppBar.Height) Then
            Me.Location = New Point(0, My.Computer.Screen.Bounds.Height - Me.Height - AppBar.Height)
        End If
    End Sub

    Private Sub Startmenu_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged

    End Sub

    Private Sub Startmenu_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        If Me.Visible = True Then
            LoadRootDrives()
            Me.Location = New Point(0, My.Computer.Screen.Bounds.Height - Me.Height - AppBar.Height)
        Else

        End If
    End Sub

    Private Sub FCM_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles FCM.Opening
        If TreeView1.SelectedNode IsNot Nothing Then
            If Not TreeView1.SelectedNode.Nodes.Count = 0 Then
                ts1.Visible = True
                ToolStripSeparator4.Visible = True
                If TreeView1.SelectedNode.IsExpanded = True Then
                    ts1.Text = "Collapse"
                    ts2.Visible = False
                Else
                    ts1.Text = "Expand"
                    ts2.Visible = True
                End If
            Else
                ts1.Visible = False
                ts2.Visible = False
                ToolStripSeparator4.Visible = False
            End If
            ToolStripSeparator5.Visible = True
            ts6.Visible = True
            ts3.Visible = True
            ts4.Visible = True
            ts5.Visible = True
        Else
            ts1.Visible = False
            ts2.Visible = False
            ToolStripSeparator5.Visible = False
            ToolStripSeparator4.Visible = False
            ts3.Visible = False
            ts4.Visible = False
            ts5.Visible = False
            ts6.Visible = True
        End If
    End Sub

    Private Sub NewGroupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewGroupToolStripMenuItem.Click
        If TreeView1.SelectedNode IsNot Nothing Then
            NFD.PATH = TreeView1.SelectedNode.Tag & "\"
        Else
            NFD.PATH = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\"
        End If
        If NFD.ShowDialog = Windows.Forms.DialogResult.OK Then
            LoadRootDrives()
        End If
    End Sub

    Private Sub ts5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ts5.Click
        If MessageBox.Show("Are you sure do you want remove " & TreeView1.SelectedNode.Tag & " ?", "Delete Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then

            If IO.Directory.Exists(TreeView1.SelectedNode.Tag) = True Then
                IO.Directory.Delete(TreeView1.SelectedNode.Tag, True)
                'My.Computer.FileSystem.DeleteDirectory(TreeView1.SelectedNode.Tag, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                LoadRootDrives()
            ElseIf IO.File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\" & TreeView1.SelectedNode.Tag) = True Then
                My.Computer.FileSystem.DeleteFile(TreeView1.SelectedNode.Tag, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin, FileIO.UICancelOption.DoNothing)
                LoadRootDrives()
            Else
                MessageBox.Show("Failed to delete: " & TreeView1.SelectedNode.Tag & ". Check if the File/Directory still exists.", "File/Directory not found", MessageBoxButtons.OK, MessageBoxIcon.Error)
                LoadRootDrives()
            End If
        End If
    End Sub

    Private Sub RefleshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefleshToolStripMenuItem.Click
        TreeView1.Update()
        LoadRootDrives()
    End Sub

    Private Sub Button3_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button3.MouseUp
        If e.Button = Windows.Forms.MouseButtons.Right Then
            ContextMenu1.Show(Me, Control.MousePosition, LeftRightAlignment.Left)
        End If
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        RunDialog.Show(AppBar)
        RunDialog.Activate()
    End Sub
    Private Sub ts1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ts1.Click
        If ts1.Text = "Expand" Then
            TreeView1.SelectedNode.Expand()
        Else
            TreeView1.SelectedNode.Collapse()
        End If
    End Sub

    Private Sub ts2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ts2.Click
        TreeView1.SelectedNode.ExpandAll()
    End Sub

    Private Sub ts3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ts3.Click
        Dim SI As New ProcessStartInfo("explorer.exe")
        SI.Arguments = "/select," & TreeView1.SelectedNode.Tag
        Process.Start(SI)
    End Sub

    Private Sub ts4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ts4.Click
        TreeView1.SelectedNode.BeginEdit()
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
    End Sub

    Private Sub LocateInExplorerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LocateInExplorerToolStripMenuItem.Click
        Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\")
    End Sub

    Private Sub ShortcutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShortcutToolStripMenuItem.Click
        If TreeView1.SelectedNode IsNot Nothing Then
            LinkCreator.PATH = TreeView1.SelectedNode.Tag & "\"
        Else
            LinkCreator.PATH = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Start Menu\Programs\"
        End If
        If LinkCreator.ShowDialog = Windows.Forms.DialogResult.OK Then
            LoadRootDrives()
            TreeView1.Nodes.Find(LinkCreator.Result, True)
        End If
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        Shell("Control.exe")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        SA.ShowDialog(Me)
    End Sub

    Private Sub LockScreenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LockScreenToolStripMenuItem.Click
        On Error Resume Next
        Shell("RunDll32.exe user32.dll,LockWorkStation")
    End Sub

    Private Sub SwitchUserToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SwitchUserToolStripMenuItem.Click
        On Error Resume Next
        Shell("tsdiscon.exe")
    End Sub

    Private Sub LogoffToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogoffToolStripMenuItem.Click
        On Error Resume Next
        If MessageBox.Show("This will log off your Windows session, Are you sure do you want log off?", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            Shell("shutdown.exe /l")
        End If
    End Sub

    Private Sub AccountSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AccountSettingsToolStripMenuItem.Click
        On Error Resume Next
        Shell("control /name Microsoft.UserAccounts")
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
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

    Private Sub TreeView2_AfterLabelEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) Handles TreeView2.AfterLabelEdit
        If Not e.Label = String.Empty Then
            If Directory.Exists(e.Node.Tag) Then
                My.Computer.FileSystem.RenameDirectory(e.Node.Tag, e.Label)
            ElseIf File.Exists(e.Node.Tag) Then
                My.Computer.FileSystem.RenameFile(e.Node.Tag, e.Label)
            End If
            e.Node.Tag = e.Node.LastNode.Tag & "\" & e.Label
        Else
            e.CancelEdit = True
        End If
    End Sub

    Private Sub TreeView2_BeforeExpand(ByVal sender As Object, ByVal e As TreeViewCancelEventArgs) Handles TreeView2.BeforeExpand

        If (e.Node.Tag <> "Desktop" AndAlso Not e.Node.Tag.Contains(":\")) OrElse Directory.Exists(e.Node.Tag) Then
            e.Node.Nodes.Clear()
            LoadChildren(e.Node, e.Node.Tag.ToString)

        ElseIf e.Node.ImageKey = "Dekstop" Then
            e.Node.Nodes.Clear()
        Else
            e.Cancel = True
        End If

    End Sub

    Private Sub TreeView2_AfterSelect(ByVal sender As Object, ByVal e As TreeViewEventArgs) Handles TreeView2.AfterSelect
        TreeView2.Tag = e.Node.Tag.ToString()
    End Sub

    Private Sub TreeView2_AfterCollapse(ByVal sender As Object, ByVal e As TreeViewEventArgs) Handles TreeView2.AfterCollapse
        e.Node.Nodes.Clear()
        e.Node.Nodes.Add("Empty")
    End Sub

    Private Sub TreeView2_DoubleClick(ByVal sender As Object, ByVal e As TreeNodeMouseClickEventArgs) Handles TreeView2.NodeMouseDoubleClick
        If e.Button = MouseButtons.Left AndAlso File.Exists(e.Node.Tag.ToString) Then
            Try
                Process.Start(e.Node.Tag.ToString())
            Catch ex As Exception
                ex = Nothing
            End Try
        End If
    End Sub

    Public Sub LoadRootDrivesAllUsers()
        TreeView2.Nodes.Clear()
        Dim DirectoryInformation As New DirectoryInfo("C:\ProgramData\Microsoft\Windows\Start Menu\Programs")
        For Each d As DirectoryInfo In DirectoryInformation.GetDirectories 'DriveInfo In DriveInfo.GetDrives
            Dim node As New TreeNode(d.Name)
            With node
                .Tag = d.FullName
                .ImageKey = "folder"
                .ImageIndex = 4
                .SelectedImageIndex = 5
                .Nodes.Add("Empty")
            End With

            If Not node.Text = "Pinned" Then
                TreeView2.Nodes.Add(node)
            End If
        Next

        For Each f As FileInfo In DirectoryInformation.GetFiles 'DriveInfo In DriveInfo.GetDrives
            Dim node As New TreeNode(f.Name)
            With node
                .Tag = f.FullName
                .ImageKey = "file"
                .ImageIndex = 3
                .SelectedImageIndex = 3
            End With

            If Not node.Text = "desktop.ini" Then
                TreeView2.Nodes.Add(node)
            End If
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\StartMenu", "DefaultFolder", ComboBox1.SelectedIndex, RegistryValueKind.DWord)
        If ComboBox1.SelectedIndex = 0 Then
            TreeView1.Visible = True
            TreeView2.Visible = False
        Else
            TreeView1.Visible = False
            TreeView2.Visible = True
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        If ToolStripMenuItem1.Text = "Expand" Then
            TreeView2.SelectedNode.Expand()
        Else
            TreeView2.SelectedNode.Collapse()
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        TreeView2.SelectedNode.ExpandAll()
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        TreeView2.Update()
        LoadRootDrivesAllUsers()
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Dim SI As New ProcessStartInfo("explorer.exe")
        SI.Arguments = "/select," & TreeView2.SelectedNode.Tag
        Process.Start(SI)
    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        If TreeView1.SelectedNode IsNot Nothing Then
            NFD.PATH = TreeView2.SelectedNode.Tag & "\"
        Else
            NFD.PATH = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\"
        End If
        If NFD.ShowDialog = Windows.Forms.DialogResult.OK Then
            LoadRootDrivesAllUsers()
        End If
    End Sub

    Private Sub ToolStripMenuItem7_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem7.Click
        If TreeView1.SelectedNode IsNot Nothing Then
            LinkCreator.PATH = TreeView2.SelectedNode.Tag & "\"
        Else
            LinkCreator.PATH = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\"
        End If
        If LinkCreator.ShowDialog = Windows.Forms.DialogResult.OK Then
            LoadRootDrivesAllUsers()
            TreeView2.Nodes.Find(LinkCreator.Result, True)
        End If
    End Sub

    Private Sub ToolStripMenuItem8_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem8.Click
        TreeView2.SelectedNode.BeginEdit()
    End Sub

    Private Sub ToolStripMenuItem9_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem9.Click
        'For safety I didn't include this :)
    End Sub

    Private Sub AllFCM_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles AllFCM.Opening
        If TreeView2.SelectedNode IsNot Nothing Then
            If Not TreeView2.SelectedNode.Nodes.Count = 0 Then
                ToolStripMenuItem1.Visible = True
                ToolStripSeparator10.Visible = True
                If TreeView2.SelectedNode.IsExpanded = True Then
                    ToolStripMenuItem1.Text = "Collapse"
                    ToolStripMenuItem2.Visible = False
                Else
                    ToolStripMenuItem1.Text = "Expand"
                    ToolStripMenuItem2.Visible = True
                End If
            Else
                ToolStripMenuItem1.Visible = False
                ToolStripMenuItem2.Visible = False
                ToolStripSeparator10.Visible = False
            End If
            ToolStripSeparator11.Visible = True
            ToolStripMenuItem6.Visible = True
            ToolStripMenuItem3.Visible = True
            ToolStripMenuItem4.Visible = True
            ToolStripMenuItem5.Visible = True
        Else
            ToolStripMenuItem1.Visible = False
            ToolStripMenuItem2.Visible = False
            ToolStripSeparator11.Visible = False
            ToolStripSeparator10.Visible = False
            ToolStripMenuItem3.Visible = False
            ToolStripMenuItem4.Visible = False
            ToolStripMenuItem5.Visible = False
            ToolStripMenuItem6.Visible = True
        End If
    End Sub
End Class