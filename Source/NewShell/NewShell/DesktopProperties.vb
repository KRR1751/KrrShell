Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Diagnostics.Eventing
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization.Formatters
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
Imports Microsoft.Win32
Imports Microsoft.Win32.Registry
Imports NewShell.DesktopProperties

Public Class DesktopProperties

    Public Enum WallpaperStyle
        Stretched = 2
        Centered = 0
        Tiled = 1
        Fit = 6
        Fill = 10
        Span = 22
    End Enum

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        ApplyTheme()

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function ExitWindowsEx(uFlags As Integer, dwReason As Integer) As Boolean
    End Function

    Private Const EWX_LOGOFF As Integer = 0

    Private Sub ApplyAndAskLogOff()

        SaveThemeToRegistry()

        Dim result As DialogResult = MessageBox.Show(
        "For Windows Metrics and Fonts to take full effect, you must log off. Log off now?",
        "Apply Changes",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            ExitWindowsEx(EWX_LOGOFF, 0)
        End If
    End Sub

    Private Sub SaveThemeToRegistry()
        Using keyColors As RegistryKey = Registry.CurrentUser.OpenSubKey("Control Panel\Colors", True)
            If keyColors IsNot Nothing Then
                For Each kvp In ThemeColors
                    Dim c As Color = kvp.Value
                    Dim rgbString As String = $"{c.R} {c.G} {c.B}"
                    keyColors.SetValue(kvp.Key, rgbString)
                Next
            End If
        End Using

        Using keyMetrics As RegistryKey = Registry.CurrentUser.OpenSubKey("Control Panel\Desktop\WindowMetrics", True)
            If keyMetrics IsNot Nothing Then
                For Each kvp In ThemeMetrics
                    keyMetrics.SetValue(kvp.Key, (kvp.Value * -15).ToString())
                Next

                For Each kvp In ThemeFonts
                    keyMetrics.SetValue(kvp.Key, GetFontBinary(kvp.Value))
                Next
            End If
        End Using
    End Sub

    Private Function GetFontBinary(f As Font) As Byte()
        Dim lf As New LOGFONT()
        f.ToLogFont(lf)
        lf.lfHeight = -CInt(f.Size * (96 / 72))

        Dim size As Integer = Marshal.SizeOf(lf)
        Dim arr(size - 1) As Byte
        Dim ptr As IntPtr = Marshal.AllocHGlobal(size)
        Try
            Marshal.StructureToPtr(lf, ptr, False)
            Marshal.Copy(ptr, arr, 0, size)
        Finally
            Marshal.FreeHGlobal(ptr)
        End Try
        Return arr
    End Function

    Public Sub ApplyTheme()
        Select Case TabControl1.SelectedIndex
            Case 0
                SetWallpaper()

            Case 1
                Dim indices As New List(Of Integer)
                Dim values As New List(Of UInteger)

                Dim map As New Dictionary(Of Integer, String) From {
                {0, "Scrollbar"}, {1, "Background"}, {2, "ActiveTitle"},
                {3, "InactiveTitle"}, {4, "Menu"}, {5, "Window"},
                {6, "WindowFrame"}, {7, "MenuText"}, {8, "WindowText"},
                {9, "TitleText"}, {10, "ActiveBorder"}, {11, "InactiveBorder"},
                {12, "AppWorkspace"}, {13, "Highlight"}, {14, "HighlightText"},
                {15, "ButtonFace"}, {16, "ButtonShadow"}, {17, "GrayText"},
                {18, "ButtonText"}, {19, "InactiveTitleText"}, {20, "ButtonHilight"},
                {21, "ButtonDkShadow"}, {22, "ButtonLight"}, {23, "InfoText"},
                {24, "InfoWindow"}, {26, "HotTrackingColor"},
                {27, "GradientActiveTitle"}, {28, "GradientInactiveTitle"}
            }

                For Each kvp In map
                    If ThemeColors.ContainsKey(kvp.Value) Then
                        indices.Add(kvp.Key)
                        values.Add(CUInt(ColorTranslator.ToOle(ThemeColors(kvp.Value))))
                    End If
                Next

                If indices.Count > 0 Then
                    SetSysColors(indices.Count, indices.ToArray(), values.ToArray())
                End If

                If CheckBox1.Checked = True Then ApplyAndAskLogOff()
        End Select
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public regWallpaperHistory As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers\"
    Public regDefaultWallpaper As String = "Control Panel\Desktop"

    Private Sub DesktopProperties_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Desktop.SendToBack()
        Me.BringToFront()

        ComboBox2.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & regDefaultWallpaper, "Wallpaper", Nothing)

        ImageList1.Images.Clear()
        ListView1.Items.Clear()

        Dim i As Integer = 0

        Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(regWallpaperHistory)

            If key Is Nothing Then
                Return
            End If

            For Each v As String In key.GetValueNames()

                If IO.File.Exists(key.GetValue(v)) Then
                    Dim FI As New IO.FileInfo(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & regWallpaperHistory, v, Nothing))

                    Try
                        Dim wallpaperImg As Image = Image.FromFile(key.GetValue(v))

                        If wallpaperImg IsNot Nothing AndAlso IsValidImage(key.GetValue(v)) Then
                            ImageList1.Images.Add(wallpaperImg)
                        End If

                    Catch ex As Exception : End Try

                    Dim item As New ListViewItem With {
                        .Text = FI.Name,
                        .ToolTipText = FI.FullName,
                        .ImageIndex = i}

                    ListView1.Items.Add(item)

                    i += 1
                End If
            Next
        End Using

        BackUpdate()

        R_NUD.Value = SystemColors.Desktop.R
        G_NUD.Value = SystemColors.Desktop.G
        B_NUD.Value = SystemColors.Desktop.B

        ' Colors Tab

        LoadSchemesToComboBox()
        LoadAllFromRegistry()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim OFD As New OpenFileDialog

            OFD.AutoUpgradeEnabled = AppBar.UseExplorerFP
            OFD.Title = "Select a Background Image to appear on your Desktop!"
            OFD.Filter = "Picture formats (*.png;*.pneg;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff)|*.png;*.pneg;*.jpg;*.jpeg;*.bmp;*.gif"
            OFD.FileName = ""
            If OFD.ShowDialog = DialogResult.OK Then
                ComboBox2.Text = OFD.FileName
                BackUpdate()
            End If
        Catch ex As Exception

        End Try
    End Sub

    ' Wallpaper update here:

    Private Shared ReadOnly SPI_SETDESKWALLPAPER As Integer = &H14
    Private Shared ReadOnly SPIF_UPDATEINIFILE As Integer = &H1
    Private Shared ReadOnly SPIF_SENDWININICHANGE As Integer = &H2

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SystemParametersInfo(uAction As Integer, uParam As Integer, lpvParam As String, fuWinIni As Integer) As Integer
    End Function

    Public Function GetCurrentWallpaperPath() As String
        Try
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(regDefaultWallpaper, False)
                If key IsNot Nothing Then
                    Dim path As Object = key.GetValue("Wallpaper")
                    Return If(path IsNot Nothing, path.ToString(), "")
                End If
            End Using
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
        Return ""
    End Function

    Public Function GetCurrentWallpaperStyle() As WallpaperStyle
        Try
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(regDefaultWallpaper, False)
                If key IsNot Nothing Then
                    Dim tile As String = key.GetValue("TileWallpaper")?.ToString()
                    If tile = "1" Then Return WallpaperStyle.Tiled

                    Dim styleRaw As String = key.GetValue("WallpaperStyle")?.ToString()
                    Dim styleInt As Integer

                    If Integer.TryParse(styleRaw, styleInt) Then
                        If [Enum].IsDefined(GetType(WallpaperStyle), styleInt) Then
                            Return CType(styleInt, WallpaperStyle)
                        End If
                    End If
                End If
            End Using
        Catch ex As Exception
            Return WallpaperStyle.Fill
        End Try

        Return WallpaperStyle.Fill
    End Function

    Private Sub SetWallpaperStyle(style As WallpaperStyle)
        Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey(regDefaultWallpaper, True)

        Dim tileWall As String = "0"
        Dim wallStyle As String = style.ToString("D")

        If style = WallpaperStyle.Tiled Then
            tileWall = "1"
        End If

        key.SetValue("WallpaperStyle", wallStyle)
        key.SetValue("TileWallpaper", tileWall)
        key.Close()
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

    Public Sub BackUpdate()
        Dim curWallpaper As String = GetCurrentWallpaperPath()
        Dim curWallpaperStyle As WallpaperStyle = GetCurrentWallpaperStyle()

        If File.Exists(curWallpaper) Then

            If Not IsValidImage(curWallpaper) Then
                BackPreview.BackgroundImage = Nothing
                pbb_Preview.Visible = False
                Exit Sub
            End If

            BackPreview.BackgroundImage = Image.FromFile(curWallpaper)
            pbb_Preview.Image = Image.FromFile(curWallpaper)
            pbb_Preview.Visible = True

            Select Case curWallpaperStyle
                Case WallpaperStyle.Centered
                    BackPreview.BackgroundImageLayout = ImageLayout.Center

                    pbb_Preview.Visible = True
                    pbb_Preview.SizeMode = PictureBoxSizeMode.CenterImage

                    rbb_Tile.Checked = True

                Case WallpaperStyle.Tiled
                    BackPreview.BackgroundImageLayout = ImageLayout.Tile

                    pbb_Preview.Visible = False

                    rbb_Tile.Checked = True

                Case WallpaperStyle.Stretched, WallpaperStyle.Fill
                    BackPreview.BackgroundImageLayout = ImageLayout.Stretch

                    pbb_Preview.Visible = True
                    pbb_Preview.SizeMode = PictureBoxSizeMode.StretchImage

                    rbb_Stretch.Checked = True

                Case WallpaperStyle.Fit
                    BackPreview.BackgroundImageLayout = ImageLayout.Zoom

                    pbb_Preview.Visible = True
                    pbb_Preview.SizeMode = PictureBoxSizeMode.Zoom

                    rbb_Zoom.Checked = True

                Case Else
                    BackPreview.BackgroundImageLayout = ImageLayout.None
                    pbb_Preview.SizeMode = PictureBoxSizeMode.Normal

                    rbb_None.Checked = True

            End Select
        Else
            BackPreview.BackgroundImage = Nothing
            pbb_Preview.Visible = False
        End If
    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        If String.IsNullOrWhiteSpace(ComboBox2.Text) Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\" & regDefaultWallpaper, "Wallpaper", "")

            BackUpdate()
            Desktop.BackUpdate()
        Else
            If IO.File.Exists(ComboBox2.Text) = True Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\" & regDefaultWallpaper, "Wallpaper", ComboBox2.Text)

                BackUpdate()
                Desktop.BackUpdate()
            End If
        End If
    End Sub

    Private Sub pbb_Preview_LoadCompleted(sender As Object, e As AsyncCompletedEventArgs) Handles pbb_Preview.LoadCompleted, pbb_Preview.LoadCompleted
        If pbb_Preview.Image.Width > 224 Then
            pbb_Preview.Width = pbb_Preview.Image.Width
        Else
            pbb_Preview.Width = 224
        End If

        If pbb_Preview.Image.Height > 128 Then
            pbb_Preview.Height = pbb_Preview.Image.Height
        Else
            pbb_Preview.Height = 128
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        On Error Resume Next
        Process.Start("control.exe", "/name Microsoft.Personalization")
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim clsidToExecute As String = "::{26EE0668-A00A-44D7-9371-BEB064C98683}\1\::{ED834ED6-4B5A-4BFE-8F11-A626DCB6A921}"

        If AppBar.UseExplorerFM = True Then
            Process.Start("explorer", clsidToExecute)
        Else
            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell", "DisableCLSIDsWarnings", False) = False Then

                If MsgBox("Explorer is not your default File Manager, do you want execute this command via your custom explorer?", MsgBoxStyle.YesNo, "CLSID detected.") = MsgBoxResult.Yes Then Process.Start(AppBar.CustomFMPath, clsidToExecute)
            Else
                Process.Start(AppBar.CustomFMPath, clsidToExecute)
            End If
        End If
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        On Error Resume Next
        Process.Start("regedit.exe", "/s ""HKEY_CURRENT_USER\Control Panel\Desktop""")
    End Sub

    Public Sub SetWallpaper()
        On Error Resume Next

        If File.Exists(ComboBox2.Text) AndAlso IsValidImage(ComboBox2.Text) Then
            SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, ComboBox2.Text, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
        Else
            SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, "", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
        End If
    End Sub

    <DllImport("user32.dll")>
    Private Shared Function SetSysColors(nElements As Integer, lpaElements As Integer(), lpaRgbValues As UInteger()) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SystemParametersInfo(uiAction As UInteger, uiParam As UInteger, ByRef pvParam As NONCLIENTMETRICS, fWinIni As UInteger) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessageTimeout(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As String, flags As UInteger, timeout As UInteger, ByRef result As IntPtr) As IntPtr
    End Function

    Private Const WM_SETTINGCHANGE As UInteger = &H1A
    Private Const SPI_GETNONCLIENTMETRICS As UInteger = &H29
    Private Const SPI_SETNONCLIENTMETRICS As UInteger = &H2A

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim CD As New ColorDialog With {
            .Color = SystemColors.Desktop,
            .AnyColor = True,
            .FullOpen = True,
            .AllowFullOpen = True}

        If CD.ShowDialog(AppBar) = DialogResult.OK Then
            R_NUD.Value = CD.Color.R
            G_NUD.Value = CD.Color.G
            B_NUD.Value = CD.Color.B
        End If
    End Sub

    Private Sub ListView1_Click(sender As Object, e As EventArgs) Handles ListView1.Click
        If ListView1.FocusedItem IsNot Nothing Then

            Dim selectedImage As String = ListView1.FocusedItem.ToolTipText

            If File.Exists(selectedImage) AndAlso IsValidImage(selectedImage) Then ComboBox2.Text = selectedImage

        End If
    End Sub

    Private Sub R_NUD_ValueChanged(sender As Object, e As EventArgs) Handles R_NUD.ValueChanged, G_NUD.ValueChanged, B_NUD.ValueChanged
        On Error Resume Next
        Panel1.BackColor = Color.FromArgb(R_NUD.Value, G_NUD.Value, B_NUD.Value)
    End Sub

    Private Sub rbb_None_CheckedChanged(sender As Object, e As EventArgs) Handles rbb_None.CheckedChanged
        SetWallpaperStyle(WallpaperStyle.Span)
    End Sub

    Private Sub rbb_Stretch_CheckedChanged(sender As Object, e As EventArgs) Handles rbb_Stretch.CheckedChanged
        SetWallpaperStyle(WallpaperStyle.Stretched)
    End Sub

    Private Sub rbb_Tile_CheckedChanged(sender As Object, e As EventArgs) Handles rbb_Tile.CheckedChanged
        SetWallpaperStyle(WallpaperStyle.Tiled)
    End Sub

    Private Sub rbb_Center_CheckedChanged(sender As Object, e As EventArgs) Handles rbb_Center.CheckedChanged
        SetWallpaperStyle(WallpaperStyle.Centered)
    End Sub

    Private Sub rbb_Zoom_CheckedChanged(sender As Object, e As EventArgs) Handles rbb_Zoom.CheckedChanged
        SetWallpaperStyle(WallpaperStyle.Fit)
    End Sub

    Private Sub LinkLabel4_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        BackUpdate()
        Desktop.BackUpdate()
    End Sub

    ' System Colors here:

    <DllImport("shlwapi.dll", CharSet:=CharSet.Unicode, ExactSpelling:=True)>
    Private Shared Function SHLoadIndirectString(ByVal pszSource As String, ByVal pszOutBuf As StringBuilder, ByVal cchOutBuf As Integer, ByVal ppvReserved As IntPtr) As Integer
    End Function

    Private ActiveTitleColor As Color = Color.FromArgb(0, 0, 128)
    Private WindowBackColor As Color = Color.White
    Private ActiveTextColor As Color = Color.White

    Private Sub LoadSchemesToComboBox()
        ComboBoxSchemes.Items.Clear()
        Dim registryPath As String = "Control Panel\Appearance\Schemes"

        Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(registryPath)
            If key IsNot Nothing Then
                For Each valueName As String In key.GetValueNames()
                    Dim displayName As String = valueName

                    If valueName.StartsWith("@") Then
                        displayName = TranslateResourceString(valueName)
                    End If

                    ComboBoxSchemes.Items.Add(New SchemeItem(displayName, valueName))
                Next
            End If
        End Using

        If ComboBoxSchemes.Items.Count > 0 Then ComboBoxSchemes.SelectedIndex = 0
    End Sub

    Private Function TranslateResourceString(ByVal resource As String) As String
        If Not resource.StartsWith("@") Then Return resource

        If resource.Contains("-854") Then Return "Windows Classic"

        Dim sb As New StringBuilder(1024)
        Dim res As Integer = SHLoadIndirectString(resource, sb, sb.Capacity, IntPtr.Zero)

        Return If(res = 0, sb.ToString(), resource)
    End Function

    Private ThemeColors As New Dictionary(Of String, Color)
    Private ThemeMetrics As New Dictionary(Of String, Integer)
    Private ThemeFonts As New Dictionary(Of String, Font)

    Private Sub LoadAllFromRegistry()
        ' 1. Load Colors
        ComboBoxColors.Items.Clear()
        Using key As RegistryKey = Registry.CurrentUser.OpenSubKey("Control Panel\Colors")
            If key IsNot Nothing Then
                For Each valName As String In key.GetValueNames()
                    Dim rgb As String() = key.GetValue(valName).ToString().Split(" "c)
                    If rgb.Length = 3 Then
                        ThemeColors(valName) = Color.FromArgb(CInt(rgb(0)), CInt(rgb(1)), CInt(rgb(2)))
                        ComboBoxColors.Items.Add(valName)
                    End If
                Next
            End If
        End Using

        ' 2. Load Metrics and Fonts from WindowMetrics
        ComboBoxMetrics.Items.Clear()
        ComboBoxFonts.Items.Clear()

        Using key As RegistryKey = Registry.CurrentUser.OpenSubKey("Control Panel\Desktop\WindowMetrics")
            If key IsNot Nothing Then
                For Each valName As String In key.GetValueNames()
                    Dim rawValue = key.GetValue(valName)

                    ' If it's a font (Binary data)
                    If key.GetValueKind(valName) = RegistryValueKind.Binary Then
                        Try
                            Dim fontBytes As Byte() = DirectCast(rawValue, Byte())
                            ' LOGFONT conversion
                            Dim logFont As New LOGFONT()
                            Dim handle As GCHandle = GCHandle.Alloc(fontBytes, GCHandleType.Pinned)
                            logFont = Marshal.PtrToStructure(Of LOGFONT)(handle.AddrOfPinnedObject())
                            handle.Free()

                            Dim convertedFont As Font = Font.FromLogFont(logFont)
                            ThemeFonts(valName) = convertedFont
                            ComboBoxFonts.Items.Add(valName)
                        Catch
                            ' Skip values that are not valid LOGFONT
                        End Try

                        ' If it's a metric (String representing a number)
                    ElseIf key.GetValueKind(valName) = RegistryValueKind.String Then
                        Dim intVal As Integer
                        If Integer.TryParse(rawValue.ToString(), intVal) Then
                            If intVal < 0 Then intVal = Math.Abs(intVal \ 15) ' Twips to Pixels
                            ThemeMetrics(valName) = intVal
                            ComboBoxMetrics.Items.Add(valName)
                        End If
                    End If
                Next
            End If
        End Using
    End Sub

    ' Structure needed for font conversion
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure LOGFONT
        Public lfHeight As Integer
        Public lfWidth As Integer
        Public lfEscapement As Integer
        Public lfOrientation As Integer
        Public lfWeight As Integer
        Public lfItalic As Byte
        Public lfUnderline As Byte
        Public lfStrikeOut As Byte
        Public lfCharSet As Byte
        Public lfOutPrecision As Byte
        Public lfClipPrecision As Byte
        Public lfQuality As Byte
        Public lfPitchAndFamily As Byte
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)>
        Public lfFaceName As String
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure LOGFONTW
        Public lfHeight As Integer
        Public lfWidth As Integer
        Public lfEscapement As Integer
        Public lfOrientation As Integer
        Public lfWeight As Integer
        Public lfItalic As Byte
        Public lfUnderline As Byte
        Public lfStrikeOut As Byte
        Public lfCharSet As Byte
        Public lfOutPrecision As Byte
        Public lfClipPrecision As Byte
        Public lfQuality As Byte
        Public lfPitchAndFamily As Byte
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)>
        Public lfFaceName As String
    End Structure

    ' --- PAINTING WITH REGISTRY DATA ---

    Private Sub PanelPreview_Paint(sender As Object, e As PaintEventArgs) Handles PanelPreview.Paint
        Dim g As Graphics = e.Graphics
        g.SmoothingMode = Drawing2D.SmoothingMode.None ' Pro ten pravý pixelový vzhled

        ' Background (Desktop)
        g.Clear(ThemeColors("Background"))

        ' 1. Icons
        DrawDesktopIcon(g, 0, 0, "This PC", False, 15)
        DrawDesktopIcon(g, 0, 1, "Recycle Bin", True, 32)

        ' Metrics
        Dim capH As Integer = ThemeMetrics("CaptionHeight")
        Dim bordW As Integer = ThemeMetrics("BorderWidth")
        Dim tFont As Font = ThemeFonts("CaptionFont")
        Dim msgFont As Font = ThemeFonts("MessageFont")

        ' 2. Inactive Window
        DrawFullWindow(g, 100, 50, 300, 200, "Inactive Window", False, capH, bordW, tFont)

        ' 3. Active Window
        DrawFullWindow(g, 130, 80, 300, 200, "Active Window", True, capH, bordW, tFont)

        ' 4. Message Box
        DrawClassicMsgBox(g, 180, 130, 220, 100, capH, bordW, tFont, msgFont)

        ' 5. Tooltip
        DrawClassicTooltip(g, 380, 90, "Close")
    End Sub

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure NONCLIENTMETRICS
        Public cbSize As Integer
        Public iBorderWidth As Integer
        Public iScrollWidth As Integer
        Public iScrollHeight As Integer
        Public iCaptionWidth As Integer
        Public iCaptionHeight As Integer
        Public lfCaptionFont As LOGFONT
        Public iSmCaptionWidth As Integer
        Public iSmCaptionHeight As Integer
        Public lfSmCaptionFont As LOGFONT
        Public iMenuWidth As Integer
        Public iMenuHeight As Integer
        Public lfMenuFont As LOGFONT
        Public lfStatusFont As LOGFONT
        Public lfMessageFont As LOGFONT
        Public iPaddedBorderWidth As Integer
    End Structure

    Private Sub LoadSelectedScheme(ByVal schemeName As String)
        Using key As RegistryKey = Registry.CurrentUser.OpenSubKey("Control Panel\Appearance\Schemes")
            Dim rawData As Byte() = TryCast(key?.GetValue(schemeName), Byte())
            If rawData Is Nothing Then Return

            Using ms As New MemoryStream(rawData)
                Using reader As New BinaryReader(ms)
                    Try : Try
                            reader.BaseStream.Seek(8, origin:=SeekOrigin.Begin)

                            Dim currentPos As Long = reader.BaseStream.Position

                            ThemeMetrics("BorderWidth") = reader.ReadInt32()
                            ThemeMetrics("ScrollWidth") = reader.ReadInt32()
                            ThemeMetrics("ScrollHeight") = reader.ReadInt32()
                            ThemeMetrics("CaptionWidth") = reader.ReadInt32()
                            ThemeMetrics("CaptionHeight") = reader.ReadInt32()

                            ThemeMetrics("SmCaptionWidth") = Math.Max(12, ThemeMetrics("CaptionWidth") - 2)
                            ThemeMetrics("SmCaptionHeight") = Math.Max(12, ThemeMetrics("CaptionHeight") - 2)

                            ThemeMetrics("PaddedBorderWidth") = 0
                        Catch ex As Exception
                            Debug.WriteLine("Failed to load Metric: " & ex.Message)
                        End Try


                        reader.BaseStream.Seek(128, SeekOrigin.Begin)

                        Dim fontNames As String() = {"CaptionFont", "SmCaptionFont", "MenuFont", "StatusFont", "MessageFont"}
                        Dim currentFontIdx As Integer = 0

                        While reader.BaseStream.Position + 92 <= reader.BaseStream.Length AndAlso currentFontIdx < 5
                            Dim startPos As Long = reader.BaseStream.Position
                            Dim block As Byte() = reader.ReadBytes(92)

                            Dim name As String = System.Text.Encoding.Unicode.GetString(block, 28, 64).Split(Chr(0))(0)

                            If Not String.IsNullOrWhiteSpace(name) AndAlso (name.Contains("Tahoma") Or name.Contains("Segoe") Or name.Contains("Sans")) Then
                                Dim h As Integer = BitConverter.ToInt32(block, 0)
                                Dim w As Integer = BitConverter.ToInt32(block, 16)
                                Dim isItalic As Boolean = (block(20) > 0)

                                Dim fSize As Single = Math.Abs(h)
                                If fSize = 0 Or fSize > 50 Then fSize = 8
                                Dim fStyle As FontStyle = If(w >= 700, FontStyle.Bold, FontStyle.Regular)
                                If isItalic Then fStyle = fStyle Or FontStyle.Italic

                                ThemeFonts(fontNames(currentFontIdx)) = New Font(name, fSize, fStyle)
                                currentFontIdx += 1

                            Else
                                reader.BaseStream.Position = startPos + 4
                            End If
                        End While

                        Dim colorsOffset As Integer = 116

                        If rawData.Length >= colorsOffset Then
                            reader.BaseStream.Seek(rawData.Length - colorsOffset, SeekOrigin.Begin)
                        Else
                            Return
                        End If

                        Dim binaryLayout As String() = {
                        "Scrollbar", "Background", "ActiveTitle", "InactiveTitle",
                        "Menu", "Window", "WindowFrame", "MenuText",
                        "WindowText", "TitleText", "ActiveBorder", "InactiveBorder",
                        "AppWorkspace", "Highlight", "HighlightText", "ButtonFace",
                        "ButtonShadow", "GrayText", "ButtonText", "InactiveTitleText",
                        "ButtonHilight", "ButtonDkShadow", "ButtonLight", "InfoText",
                        "InfoWindow", "ButtonAlternateFace", "HotTrackingColor",
                        "GradientActiveTitle", "GradientInactiveTitle", "MenuHilight", "MenuBar"
                    }


                        For i As Integer = 0 To binaryLayout.Length - 1
                            If reader.BaseStream.Position + 4 <= reader.BaseStream.Length Then
                                Dim r As Byte = reader.ReadByte()
                                Dim g As Byte = reader.ReadByte()
                                Dim b As Byte = reader.ReadByte()
                                Dim dummy As Byte = reader.ReadByte()

                                If i < binaryLayout.Length Then
                                    Dim colName As String = binaryLayout(i)
                                    ThemeColors(colName) = Color.FromArgb(r, g, b)
                                End If
                            End If
                        Next

                    Catch ex As Exception
                        Debug.WriteLine("Error while loading Scheme: " & ex.Message)
                    End Try
                End Using
            End Using
        End Using
        PanelPreview.Invalidate()
    End Sub

    Public Class SchemeItem
        Public Property DisplayName As String
        Public Property RegistryName As String

        Public Sub New(display As String, registry As String)
            DisplayName = display
            RegistryName = registry
        End Sub

        Public Overrides Function ToString() As String
            Return DisplayName
        End Function
    End Class

    Private Sub ComboBoxSchemes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxSchemes.SelectedIndexChanged
        Dim selectedScheme = TryCast(ComboBoxSchemes.SelectedItem, SchemeItem)

        If selectedScheme IsNot Nothing Then
            LoadSelectedScheme(selectedScheme.RegistryName)

            PanelPreview.Invalidate()
            ComboBoxColors_SelectedIndexChanged(Nothing, Nothing)
        End If
    End Sub

    Private Function GetShellIcon(index As Integer, large As Boolean) As Icon
        Dim hImgLarge As IntPtr
        Dim hImgSmall As IntPtr

        ExtractIconEx("shell32.dll", index, hImgLarge, hImgSmall, 1)

        If large Then
            Return Icon.FromHandle(hImgLarge)
        Else
            Return Icon.FromHandle(hImgSmall)
        End If
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function ExtractIconEx(libName As String, iconIndex As Integer, ByRef largeIcon As IntPtr, ByRef smallIcon As IntPtr, nIcons As UInteger) As UInteger
    End Function

    Private Sub FillTitleGradient(g As Graphics, rect As Rectangle, c1 As Color, c2 As Color)
        If c1 = c2 Then
            Using b As New SolidBrush(c1)
                g.FillRectangle(b, rect)
            End Using
        Else
            Using b As New Drawing2D.LinearGradientBrush(rect, c1, c2, Drawing2D.LinearGradientMode.Horizontal)
                g.FillRectangle(b, rect)
            End Using
        End If
    End Sub

    Private Sub DrawDesktopIcon(g As Graphics, gridX As Integer, gridY As Integer, label As String, isSelected As Boolean, iconIndex As Integer)
        Dim iconSize As Integer = ThemeMetrics("Shell Icon Size") ' Standard 32
        Dim spacingH As Integer = ThemeMetrics("IconSpacing") ' Horizontální mezera
        Dim spacingV As Integer = ThemeMetrics("IconVerticalSpacing") ' Vertikální mezera
        Dim iconFont As Font

        Try
            iconFont = ThemeFonts("IconFont")
        Catch ex As Exception
            iconFont = SystemFonts.IconTitleFont
        End Try

        Dim posX As Integer = gridX * spacingH
        Dim posY As Integer = gridY * spacingV

        Using ico As Icon = GetShellIcon(iconIndex, True)
            Dim icoX As Integer = posX + (spacingH \ 2) - (iconSize \ 2)
            g.DrawIcon(ico, New Rectangle(icoX, posY + 5, iconSize, iconSize))
        End Using

        Dim textWidth As Integer = spacingH - 10
        Dim flags As TextFormatFlags = TextFormatFlags.HorizontalCenter Or TextFormatFlags.Top
        If ThemeMetrics("IconTitleWrap") = 1 Then flags = flags Or TextFormatFlags.WordBreak

        Dim textSize As Size = TextRenderer.MeasureText(label, iconFont, New Size(textWidth, 100), flags)
        Dim textRect As New Rectangle(posX + (spacingH \ 2) - (textSize.Width \ 2), posY + iconSize + 10, textSize.Width, textSize.Height)

        If isSelected Then
            g.FillRectangle(New SolidBrush(ThemeColors("Highlight")), textRect)
            TextRenderer.DrawText(g, label, iconFont, textRect, ThemeColors("HighlightText"), flags)
        Else
            TextRenderer.DrawText(g, label, iconFont, textRect, Color.White, flags)
        End If
    End Sub

    Private Sub DrawClassicTooltip(g As Graphics, x As Integer, y As Integer, text As String)
        Dim tipFont As Font = ThemeFonts("StatusFont") ' Tooltipy často sdílí font se stavovým řádkem
        Dim textSize As Size = TextRenderer.MeasureText(text, tipFont)
        Dim tipRect As New Rectangle(x, y, textSize.Width + 6, textSize.Height + 4)

        Using b As New SolidBrush(ThemeColors("InfoWindow"))
            g.FillRectangle(b, tipRect)
        End Using

        g.DrawRectangle(Pens.Black, tipRect)

        TextRenderer.DrawText(g, text, tipFont, New Point(tipRect.X + 3, tipRect.Y + 2), ThemeColors("InfoText"))
    End Sub

    Private Sub DrawFullWindow(g As Graphics, x As Integer, y As Integer, w As Integer, h As Integer, title As String, isActive As Boolean, capHeight As Integer, border As Integer, titleFont As Font)
        On Error Resume Next
        ' Colors
        Dim tColor As Color = If(isActive, ThemeColors("ActiveTitle"), ThemeColors("InactiveTitle"))
        Dim tText As Color = If(isActive, ThemeColors("TitleText"), ThemeColors("InactiveTitleText"))
        Dim bColor As Color = If(isActive, ThemeColors("ActiveBorder"), ThemeColors("InactiveBorder"))

        ' 1. Border
        ControlPaint.DrawBorder3D(g, New Rectangle(x, y, w, h), Border3DStyle.Raised)
        g.FillRectangle(New SolidBrush(bColor), x + 2, y + 2, w - 4, h - 4)

        ' 2. Title Bar
        Dim titleRect As New Rectangle(x + border + 2, y + border + 2, w - (border * 2) - 4, capHeight)
        FillTitleGradient(g, titleRect, tColor, If(isActive, ThemeColors("GradientActiveTitle"), ThemeColors("GradientInactiveTitle")))
        TextRenderer.DrawText(g, title, titleFont, New Rectangle(titleRect.X + 2, titleRect.Y, titleRect.Width - (capHeight * 3), titleRect.Height), tText, TextFormatFlags.VerticalCenter Or TextFormatFlags.Left)

        Dim menuHeight As Integer = ThemeMetrics("MenuHeight")
        If menuHeight <= 0 Then menuHeight = 19
        Dim menuRect As New Rectangle(titleRect.X, titleRect.Bottom, titleRect.Width, menuHeight + 1)
        g.FillRectangle(New SolidBrush(ThemeColors("Menu")), menuRect)

        Dim menuFont As Font = ThemeFonts("MenuFont")
        Dim itemX As Integer = menuRect.X + 2
        Dim selWidth As Integer = TextRenderer.MeasureText("File", menuFont).Width + 10
        Dim selRect As New Rectangle(itemX, menuRect.Y + 1, selWidth, menuRect.Height - 2)
        g.FillRectangle(New SolidBrush(ThemeColors("Highlight")), selRect)
        TextRenderer.DrawText(g, "File", menuFont, selRect, ThemeColors("HighlightText"), TextFormatFlags.VerticalCenter Or TextFormatFlags.HorizontalCenter)

        TextRenderer.DrawText(g, "Edit", menuFont, New Point(itemX + selWidth + 8, menuRect.Y + (menuRect.Height \ 2) - (menuFont.Height \ 2)), ThemeColors("MenuText"))
        TextRenderer.DrawText(g, "View", menuFont, New Point(itemX + selWidth + 45, menuRect.Y + (menuRect.Height \ 2) - (menuFont.Height \ 2)), ThemeColors("MenuText"))

        Dim scrollWidth As Integer = ThemeMetrics("ScrollWidth")
        Dim clientTop As Integer = menuRect.Bottom
        Dim clientAreaRect As New Rectangle(titleRect.X, clientTop, titleRect.Width, (y + h) - clientTop - border - 2)

        ' SCROLLBAR
        Dim scrollRect As New Rectangle(clientAreaRect.Right - scrollWidth, clientAreaRect.Y, scrollWidth, clientAreaRect.Height)
        Using hb As New Drawing2D.HatchBrush(Drawing2D.HatchStyle.Percent50, ThemeColors("ButtonHilight"), ThemeColors("ButtonFace"))
            g.FillRectangle(hb, scrollRect)
        End Using

        DrawCustomScrollButton(g, New Rectangle(scrollRect.X, scrollRect.Y, scrollWidth, scrollWidth), "up")
        DrawCustomScrollButton(g, New Rectangle(scrollRect.X, scrollRect.Bottom - scrollWidth, scrollWidth, scrollWidth), "down")

        Dim thumbRect As New Rectangle(scrollRect.X, scrollRect.Y + 30, scrollWidth, 40)
        DrawCustomButton(g, thumbRect, "", Me.Font)

        Dim whiteRect As New Rectangle(clientAreaRect.X, clientAreaRect.Y, clientAreaRect.Width - scrollWidth, clientAreaRect.Height)
        ControlPaint.DrawBorder3D(g, whiteRect, Border3DStyle.Sunken)

        Dim btnSize As Integer = capHeight - 2
        Dim btnY As Integer = titleRect.Y + 1
        Dim currentBtnX As Integer = titleRect.Right - btnSize - 2

        DrawCustomCaptionButton(g, New Rectangle(currentBtnX, btnY, btnSize + 1, btnSize), "close")
        currentBtnX -= (btnSize + 2)
        DrawCustomCaptionButton(g, New Rectangle(currentBtnX, btnY, btnSize + 1, btnSize), "max")
        currentBtnX -= (btnSize + 1)
        DrawCustomCaptionButton(g, New Rectangle(currentBtnX, btnY, btnSize + 1, btnSize), "min")

        Dim innerColor As Color = If(isActive, ThemeColors("Window"), ThemeColors("AppWorkspace"))
        g.FillRectangle(New SolidBrush(innerColor), whiteRect.X + 2, whiteRect.Y + 2, whiteRect.Width - 4, whiteRect.Height - 4)

        TextRenderer.DrawText(g, "Window Text Content", ThemeFonts("MessageFont"), New Point(whiteRect.X + 5, whiteRect.Y + 5), ThemeColors("WindowText"))
    End Sub

    Private Sub DrawClassicMsgBox(g As Graphics, x As Integer, y As Integer, w As Integer, h As Integer, capHeight As Integer, border As Integer, titleFont As Font, msgFont As Font)
        ControlPaint.DrawBorder3D(g, New Rectangle(x, y, w, h), Border3DStyle.Raised)
        g.FillRectangle(New SolidBrush(ThemeColors("ButtonFace")), x + 2, y + 2, w - 4, h - 4)

        ' Title
        Dim titleRect As New Rectangle(x + border + 2, y + border + 2, w - (border * 2) - 4, capHeight)
        FillTitleGradient(g, titleRect, ThemeColors("ActiveTitle"), ThemeColors("GradientActiveTitle"))

        TextRenderer.DrawText(g, "Message Box", titleFont, titleRect, ThemeColors("TitleText"), TextFormatFlags.VerticalCenter Or TextFormatFlags.Left)

        ' Close button
        Dim btnSize As Integer = capHeight - 2
        DrawCustomCaptionButton(g, New Rectangle(titleRect.Right - btnSize - 2, titleRect.Y + 1, btnSize + 1, btnSize), "close")

        Dim msgArea As New Rectangle(x + border + 10, titleRect.Bottom + 10, w - (border * 2) - 20, h - capHeight - 60)
        TextRenderer.DrawText(g, "Message Text", msgFont, msgArea, ThemeColors("WindowText"), TextFormatFlags.WordBreak Or TextFormatFlags.Top)

        ' OK Button
        Dim okRect As New Rectangle(x + (w \ 2) - 37, y + h - border - 35, 75, 23)
        Dim okFont As Font

        Try
            okFont = ThemeFonts("IconFont")
        Catch ex As Exception
            okFont = SystemFonts.IconTitleFont
        End Try
        DrawCustomButton(g, okRect, "OK", okFont)
    End Sub

    Private Sub DrawCustomScrollButton(g As Graphics, rect As Rectangle, direction As String)
        Dim sb As ScrollButton
        If direction = "up" Then
            sb = ScrollButton.Up
        Else
            sb = ScrollButton.Down
        End If

        ControlPaint.DrawScrollButton(g, rect, sb, ButtonState.Normal)
    End Sub

    Private Sub DrawCustomButton(g As Graphics, rect As Rectangle, text As String, btnFont As Font)
        ControlPaint.DrawBorder3D(g, rect, Border3DStyle.Raised)
        Using b As New SolidBrush(ThemeColors("ButtonFace"))
            g.FillRectangle(b, rect.X + 2, rect.Y + 2, rect.Width - 4, rect.Height - 4)
        End Using

        TextRenderer.DrawText(g, text, btnFont, rect, ThemeColors("ButtonText"), TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
    End Sub

    Private Sub DrawCustomCaptionButton(g As Graphics, rect As Rectangle, type As String)
        Dim face As Color = ThemeColors("ButtonFace")
        Dim light As Color = ThemeColors("ButtonHilight")
        Dim shadow As Color = ThemeColors("ButtonShadow")
        Dim dkShadow As Color = ThemeColors("ButtonDkShadow")
        Dim btnText As Color = ThemeColors("ButtonText")

        Using b As New SolidBrush(face)
            g.FillRectangle(b, rect)
        End Using

        Using pLight As New Pen(light), pShadow As New Pen(shadow), pDkShadow As New Pen(dkShadow)
            g.DrawLine(pLight, rect.X, rect.Y, rect.Right - 1, rect.Y)
            g.DrawLine(pLight, rect.X, rect.Y, rect.X, rect.Bottom - 1)
            g.DrawLine(pShadow, rect.X + 1, rect.Bottom - 2, rect.Right - 2, rect.Bottom - 2)
            g.DrawLine(pShadow, rect.Right - 2, rect.Y + 1, rect.Right - 2, rect.Bottom - 2)
            g.DrawLine(pDkShadow, rect.X, rect.Bottom - 1, rect.Right - 1, rect.Bottom - 1)
            g.DrawLine(pDkShadow, rect.Right - 1, rect.Y, rect.Right - 1, rect.Bottom - 1)
        End Using

        Using p As New Pen(btnText, 1)
            Dim center As New Point(rect.X + (rect.Width \ 2), rect.Y + (rect.Height \ 2))

            If type = "close" Then
                g.DrawLine(p, center.X - 3, center.Y - 3, center.X + 3, center.Y + 3)
                g.DrawLine(p, center.X - 2, center.Y - 3, center.X + 4, center.Y + 3) ' Tlustší křížek
                g.DrawLine(p, center.X + 3, center.Y - 3, center.X - 3, center.Y + 3)
                g.DrawLine(p, center.X + 4, center.Y - 3, center.X - 2, center.Y + 3)

            ElseIf type = "max" Then
                g.DrawRectangle(p, center.X - 4, center.Y - 4, 8, 8)
                g.DrawLine(p, center.X - 4, center.Y - 3, center.X + 4, center.Y - 3) ' Horní tlustá lišta

            ElseIf type = "min" Then
                g.DrawLine(p, center.X - 4, center.Y + 3, center.X + 1, center.Y + 3)
                g.DrawLine(p, center.X - 4, center.Y + 4, center.X + 1, center.Y + 4)
            End If
        End Using
    End Sub

    Private Function GetThemeColor(name As String, fallback As Color) As Color
        Return If(ThemeColors.ContainsKey(name), ThemeColors(name), fallback)
    End Function
    Private Function GetThemeMetric(name As String, fallback As Integer) As Integer
        Return If(ThemeMetrics.ContainsKey(name), ThemeMetrics(name), fallback)
    End Function
    Private Function GetThemeFont(name As String, fallback As Font) As Font
        Return If(ThemeFonts.ContainsKey(name), ThemeFonts(name), fallback)
    End Function

    Private Sub ComboBoxColors_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxColors.SelectedIndexChanged
        If ComboBoxColors.SelectedItem Is Nothing Then Exit Sub

        Dim selectedKey As String = ComboBoxColors.SelectedItem.ToString()
        If ThemeColors.ContainsKey(selectedKey) Then
            PanelCurrentColor.BackColor = ThemeColors(selectedKey)
        End If
    End Sub

    Private Sub ComboBoxMetrics_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxMetrics.SelectedIndexChanged
        If ComboBoxMetrics.SelectedItem Is Nothing Then Exit Sub

        Dim selectedKey As String = ComboBoxMetrics.SelectedItem.ToString()
        If ThemeMetrics.ContainsKey(selectedKey) Then
            Dim val As Integer = ThemeMetrics(selectedKey)
            If val >= NumericMetricValue.Minimum AndAlso val <= NumericMetricValue.Maximum Then
                NumericMetricValue.Value = val
            End If
        End If
    End Sub

    Private Sub ComboBoxFonts_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxFonts.SelectedIndexChanged
        ' We just store the selection, FontDialog will use it when opened
    End Sub

    ' --- EDITING LOGIC ---

    ' Change Color
    Private Sub ButtonEditColor_Click(sender As Object, e As EventArgs) Handles ButtonEditColor.Click
        If ComboBoxColors.SelectedItem Is Nothing Then Return

        Dim selectedKey As String = ComboBoxColors.SelectedItem.ToString()
        Using cd As New ColorDialog
            cd.Color = ThemeColors(selectedKey)
            If cd.ShowDialog() = DialogResult.OK Then
                ThemeColors(selectedKey) = cd.Color
                PanelCurrentColor.BackColor = cd.Color
                PanelPreview.Invalidate() ' Refresh the 98-style preview
            End If
        End Using
    End Sub

    ' Change Metric (Size)
    Private Sub NumericMetricValue_ValueChanged(sender As Object, e As EventArgs) Handles NumericMetricValue.ValueChanged
        If ComboBoxMetrics.SelectedItem Is Nothing Then Return

        Dim selectedKey As String = ComboBoxMetrics.SelectedItem.ToString()
        ThemeMetrics(selectedKey) = CInt(NumericMetricValue.Value)
        PanelPreview.Invalidate() ' Refresh the 98-style preview
    End Sub

    ' Change Font
    Private Sub ButtonEditFont_Click(sender As Object, e As EventArgs) Handles ButtonEditFont.Click
        If ComboBoxFonts.SelectedItem Is Nothing Then Return

        Dim selectedKey As String = ComboBoxFonts.SelectedItem.ToString()
        Using fd As New FontDialog
            fd.Font = ThemeFonts(selectedKey)
            If fd.ShowDialog() = DialogResult.OK Then
                ThemeFonts(selectedKey) = fd.Font
                PanelPreview.Invalidate() ' Refresh the 98-style preview
            End If
        End Using
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        Select Case TabControl1.SelectedIndex
            Case 0
                pbb_Preview.Visible = True
                PanelPreview.Visible = False
            Case 1
                PanelPreview.Visible = True
                pbb_Preview.Visible = False
        End Select
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ApplyTheme()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = False Then
            GroupBox2.Enabled = False
        Else
            If MessageBox.Show("Are you sure do you want to enable those settings? This is a warning because wrong values can break your themes completelly and turn this on only if you know what u are doing.", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                GroupBox2.Enabled = True
            Else
                CheckBox1.Checked = False
            End If
        End If
    End Sub
End Class
