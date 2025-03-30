Imports Microsoft.Win32

Public Class Desktop

    Private Sub Desktop_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MinimumSize = SystemInformation.PrimaryMonitorSize
        Me.MaximumSize = SystemInformation.PrimaryMonitorSize
        Me.Size = SystemInformation.PrimaryMonitorSize
        Me.Location = New Point(0, 0)
        Me.SendToBack()
        If IO.File.Exists(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", Nothing)) = True Then
            ' Applies the Desktop Background
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
                PictureBox1.Visible = False
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
            PictureBox1.Visible = False
        End If

        'Dim FV As New FolderView()
        'FV.MdiParent = Me
        'FV.Show()

        'FolderView.Show()
    End Sub
    Public CanClose As Boolean = False
    Private Sub Desktop_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If CanClose = False Then
            e.Cancel = True
            SA.ShowDialog(Me)
        End If
    End Sub

    Private Sub Desktop_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
        If Not Me.Location = New Point(0, 0) Then
            Me.Location = New Point(0, 0)
        End If
    End Sub

    Private Sub Desktop_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Me.SendToBack()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DesktopCM.Opening
        Me.SendToBack()
        DesktopCM.BringToFront()
        DesktopCM.Items.Clear()

        Dim TopItems As New List(Of String)
        Dim CenterItems As New List(Of String)
        Dim BottomItems As New List(Of String)

        For Each i As String In My.Computer.Registry.ClassesRoot.OpenSubKey("DesktopBackground\Shell").GetSubKeyNames
            If My.Computer.Registry.GetValue("HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & i, "Position", Nothing) = "Top" Then
                Dim item As New ToolStripMenuItem(My.Computer.Registry.GetValue("HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & i, "MUIVerb", i).ToString)
                With item
                    Text = My.Computer.Registry.GetValue("HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & i, "MUIVerb", i).ToString
                    Tag = "HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & i
                End With
                AddHandler item.Click, AddressOf ContextMenuItemClick
                DesktopCM.Items.Add(item)
            Else
                DesktopCM.Items.Add(WallpaperRefleshToolStripMenuItem)
                DesktopCM.Items.Add(ToolStripSeparator1)
            End If
        Next

        For Each i As String In My.Computer.Registry.ClassesRoot.OpenSubKey("DesktopBackground\Shell").GetSubKeyNames
            If My.Computer.Registry.GetValue("HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & i, "Position", Nothing) = "Center" Then
                Dim item As New ToolStripMenuItem(My.Computer.Registry.GetValue("HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & i, "MUIVerb", i).ToString)
                With item
                    Tag = "HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & i
                End With
                AddHandler item.Click, AddressOf ContextMenuItemClick
                DesktopCM.Items.Add(item)
            Else
                DesktopCM.Items.Add(PasteToolStripMenuItem)
                DesktopCM.Items.Add(ToolStripSeparator2)
            End If
        Next

        For Each i As String In My.Computer.Registry.ClassesRoot.OpenSubKey("DesktopBackground\Shell").GetSubKeyNames
            If My.Computer.Registry.GetValue("HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & i, "Position", Nothing) = "Bottom" Then
                If i = "Personalize" Then
                    Dim ico As Icon = Icon.ExtractAssociatedIcon("C:\Windows\System32\themecpl.dll")
                    Dim item As New ToolStripMenuItem(My.Computer.Registry.GetValue("HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & i, "MUIVerb", i).ToString)
                    With item
                        .Image = ico.ToBitmap
                        Tag = "HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & i
                    End With
                    AddHandler item.Click, AddressOf ContextMenuItemClick
                    DesktopCM.Items.Add(item)
                ElseIf i = "Display" Then
                    Dim ico As Icon = Icon.ExtractAssociatedIcon("C:\Windows\System32\display.dll")
                    Dim item As New ToolStripMenuItem(My.Computer.Registry.GetValue("HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & i, "MUIVerb", i).ToString)
                    With item
                        .Image = ico.ToBitmap
                        Tag = "HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & i
                    End With
                    AddHandler item.Click, AddressOf ContextMenuItemClick
                    DesktopCM.Items.Add(item)
                ElseIf i = "Gadgets" Then
                    Dim ico As Icon = Icon.ExtractAssociatedIcon("C:\Program Files\Windows Sidebar\dwmapi.dll")
                    Dim item As New ToolStripMenuItem(My.Computer.Registry.GetValue("HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & i, "MUIVerb", i).ToString)
                    With item
                        .Image = ico.ToBitmap
                        Tag = "HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & i
                    End With
                    AddHandler item.Click, AddressOf ContextMenuItemClick
                    DesktopCM.Items.Add(item)
                Else
                    Dim item As New ToolStripMenuItem(My.Computer.Registry.GetValue("HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & i, "MUIVerb", i).ToString)
                    With item
                        Tag = "HKEY_CLASSES_ROOT\DesktopBackground\Shell\" & i
                    End With
                    AddHandler item.Click, AddressOf ContextMenuItemClick
                    DesktopCM.Items.Add(item)
                End If
            End If
        Next
    End Sub

    Private Sub Desktop_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp, PictureBox1.MouseUp
        If e.Button = MouseButtons.Right Then
            DesktopCM.Show(Me, MousePosition)
        End If
    End Sub

    Private Sub Desktop_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        Me.SendToBack()
    End Sub

    Private Sub ContextMenuItemClick(sender As Object, e As EventArgs)
        Try
            If CType(sender, ToolStripMenuItem).Text = "Personalize" Then
                Process.Start("explorer", "::{26EE0668-A00A-44D7-9371-BEB064C98683}\1\::{ED834ED6-4B5A-4BFE-8F11-A626DCB6A921}")
            ElseIf CType(sender, ToolStripMenuItem).Text = "Display" Then
                Process.Start("ms-settings:display")
            ElseIf CType(sender, ToolStripMenuItem).Text = "Gadgets" Then
                Shell("C:\Program Files\Windows Sidebar\sidebar.exe /showGadgets")
            ElseIf CType(sender, ToolStripMenuItem).Text = "Appearance" Then
                DesktopProperties.ShowDialog(Me)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub WallpaperRefleshToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WallpaperRefleshToolStripMenuItem.Click
        'MsgBox(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", Nothing))
        BackUpdate()
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
                PictureBox1.Visible = False
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
            Me.BackgroundImage = Image.FromFile("C:\Windows\Resources\Themes\Nothing.png")
            PictureBox1.Visible = False
        End If
    End Sub
End Class