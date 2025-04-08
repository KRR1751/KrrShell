Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Windows.Forms

Public Class DesktopProperties

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DesktopProperties_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Desktop.SendToBack()
        Me.BringToFront()

        ComboBox2.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", Nothing)

        If IO.File.Exists(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", Nothing)) = True Then
            ComboBox1.SelectedIndex = 0
        Else
            ComboBox1.SelectedIndex = 1
        End If

        ImageList1.Images.Clear()
        ListView1.Items.Clear()

        For i As Integer = 0 To 100000
            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers\", "BackgroundHistoryPath" & i, Nothing) = Nothing Then
                Exit For
            Else
                If IO.File.Exists(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers\", "BackgroundHistoryPath" & i, Nothing)) Then
                    Dim FI As New IO.FileInfo(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers\", "BackgroundHistoryPath" & i, Nothing))
                    Dim item As New ListViewItem
                    Try
                        ImageList1.Images.Add(Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers\", "BackgroundHistoryPath" & i, Nothing)))
                    Catch ex As Exception

                    End Try
                    With item
                        .Text = FI.Name
                        .ToolTipText = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers\", "BackgroundHistoryPath" & i, Nothing)
                        .ImageIndex = i
                    End With

                    ListView1.Items.Add(item)
                End If
            End If
        Next

        If IO.File.Exists(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", Nothing)) = True Then
            BackPreview.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", Nothing))
            PictureBox1.Image = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", Nothing))
            PictureBox1.Visible = True

            ' Applies the "Style" of the Background
            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", Nothing) = 0 AndAlso My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "TileWallpaper", Nothing) = 0 Then
                RadioButton2.Checked = True
                BackPreview.BackgroundImageLayout = ImageLayout.Center
                PictureBox1.Visible = True
                PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
            ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", Nothing) = 0 AndAlso My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "TileWallpaper", Nothing) = 1 Then
                RadioButton5.Checked = True
                BackPreview.BackgroundImageLayout = ImageLayout.Tile
                PictureBox1.Visible = False
            ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", Nothing) = 2 OrElse My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", Nothing) = 10 Then
                RadioButton3.Checked = True
                BackPreview.BackgroundImageLayout = ImageLayout.Stretch
                PictureBox1.Visible = True
                PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
            ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", Nothing) = 6 Then
                RadioButton4.Checked = True
                BackPreview.BackgroundImageLayout = ImageLayout.Zoom
                PictureBox1.Visible = True
                PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            Else
                RadioButton1.Checked = True
                BackPreview.BackgroundImageLayout = ImageLayout.None
                PictureBox1.SizeMode = PictureBoxSizeMode.Normal
            End If
        Else
            BackPreview.BackgroundImage = Image.FromFile("C:\Windows\Resources\Themes\Nothing.png")
            PictureBox1.Visible = False
        End If
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim OFD As New OpenFileDialog
        OFD.Title = "Select a Background Image to appear on your Desktop!"
        OFD.Filter = "Picture formats (*.png;*.pneg;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff)|*.png;*.pneg;*.jpg;*.jpeg;*.bmp;*.gif|All files (*.*)|*.*"
        OFD.FileName = ""
        If OFD.ShowDialog = DialogResult.OK Then
            ComboBox2.Text = OFD.FileName
            BackUpdate()
        End If
    End Sub

    Public Sub BackUpdate()
        If IO.File.Exists(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", Nothing)) = True Then
            BackPreview.BackgroundImage = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", Nothing))
            PictureBox1.Image = Image.FromFile(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", Nothing))
            PictureBox1.Visible = True

            ' Applies the "Style" of the Background
            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", Nothing) = 0 AndAlso My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "TileWallpaper", Nothing) = 0 Then
                'RadioButton2.Checked = True
                BackPreview.BackgroundImageLayout = ImageLayout.Center
                PictureBox1.Visible = True
                PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
            ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", Nothing) = 0 AndAlso My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "TileWallpaper", Nothing) = 1 Then
                'RadioButton5.Checked = True
                BackPreview.BackgroundImageLayout = ImageLayout.Tile
                PictureBox1.Visible = False
            ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", Nothing) = 2 OrElse My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", Nothing) = 10 Then
                'RadioButton3.Checked = True
                BackPreview.BackgroundImageLayout = ImageLayout.Stretch
                PictureBox1.Visible = True
                PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
            ElseIf My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", Nothing) = 6 Then
                'RadioButton4.Checked = True
                BackPreview.BackgroundImageLayout = ImageLayout.Zoom
                PictureBox1.Visible = True
                PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            Else
                'RadioButton1.Checked = True
                BackPreview.BackgroundImageLayout = ImageLayout.None
                PictureBox1.SizeMode = PictureBoxSizeMode.Normal
            End If
        Else
            BackPreview.BackgroundImage = Image.FromFile("C:\Windows\Resources\Themes\Nothing.png")
            PictureBox1.Visible = False
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", 22)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "TileWallpaper", 0)
        BackUpdate()
        Desktop.BackUpdate()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", 0)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "TileWallpaper", 0)
        BackUpdate()
        Desktop.BackUpdate()
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", 2)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "TileWallpaper", 0)
        BackUpdate()
        Desktop.BackUpdate()
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", 6)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "TileWallpaper", 0)
        BackUpdate()
        Desktop.BackUpdate()
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "WallpaperStyle", 0)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "TileWallpaper", 1)
        BackUpdate()
        Desktop.BackUpdate()
    End Sub

    Private Sub ListView1_ItemActivate(sender As Object, e As EventArgs) Handles ListView1.ItemActivate
        ComboBox2.Text = ListView1.FocusedItem.ToolTipText
    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        If IO.File.Exists(ComboBox2.Text) = True Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", ComboBox2.Text)
            BackUpdate()
            Desktop.BackUpdate()
        End If
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub PictureBox1_LoadCompleted(sender As Object, e As AsyncCompletedEventArgs) Handles PictureBox1.LoadCompleted
        If PictureBox1.Image.Width > 224 Then
            PictureBox1.Width = PictureBox1.Image.Width
        Else
            PictureBox1.Width = 224
        End If

        If PictureBox1.Image.Height > 128 Then
            PictureBox1.Height = PictureBox1.Image.Height
        Else
            PictureBox1.Height = 128
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Shell("ms-settings:personalization")
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Process.Start("explorer", "::{26EE0668-A00A-44D7-9371-BEB064C98683}\1\::{ED834ED6-4B5A-4BFE-8F11-A626DCB6A921}")
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit", "LastKey", "Computer\Computer\HKEY_CURRENT_USER\Control Panel\Desktop")
        Process.Start("regedit")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 1 Then
            P1.Visible = False
            P2.Visible = True
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", "")
            BackUpdate()
            Desktop.BackUpdate()
        Else
            P1.Visible = True
            P2.Visible = False
            If IO.File.Exists(ComboBox2.Text) = True Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", ComboBox2.Text)
                BackUpdate()
                Desktop.BackUpdate()
            End If
        End If
    End Sub
    Private Declare Function SetSysColors Lib "user32.dll" (ByVal one As Integer, ByRef element As Integer, ByRef color As Integer) As Boolean
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim CD As New ColorDialog
        CD.FullOpen = True
        CD.AllowFullOpen = True
        If CD.ShowDialog(AppBar) = DialogResult.OK Then
            Dim BackgroundColor As Integer = ColorTranslator.ToWin32(CD.Color)
            SetSysColors(1, 1, BackgroundColor)
        End If
    End Sub
End Class
