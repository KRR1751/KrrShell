Imports System.IO
Imports System.Media
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class DateAndTime
    ' API Konstanty
    Private Const WM_NCPAINT As Integer = &H85
    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const WM_NCLBUTTONUP As Integer = &HA2
    Private Const WM_NCLBUTTONDBLCLK As Integer = &HA3
    Private Const WM_NCMOUSEMOVE As Integer = &HA0
    Private Const WM_NCACTIVATE As Integer = &H86
    Private Const WM_NCHITTEST As Integer = &H84
    Private Const WM_SETTEXT As Integer = &HC
    Private Const HT_CAPTION As Integer = 2
    Private Const DWMWA_NCRENDERING_POLICY As Integer = 2
    Private Const DWMNCRP_DISABLED As Integer = 1

    Private Const RDW_INVALIDATE As Integer = &H1
    Private Const RDW_FRAME As Integer = &H400
    Private Const RDW_UPDATENOW As Integer = &H100

    Private isButtonPressed As Boolean = False
    Private showCustomButton As Boolean = False ' Tuto proměnnou ovládá CheckBox

    <DllImport("user32.dll")>
    Private Shared Function RedrawWindow(hWnd As IntPtr, lprcUpdate As IntPtr, hrgnUpdate As IntPtr, flags As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindowDC(hWnd As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function ReleaseDC(hWnd As IntPtr, hDC As IntPtr) As Integer
    End Function

    <DllImport("dwmapi.dll")>
    Private Shared Function DwmSetWindowAttribute(hwnd As IntPtr, attr As Integer, ByRef attrValue As Integer, attrSize As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetSystemMetrics(nIndex As Integer) As Integer
    End Function

    Private ReadOnly Property CustomButtonRect As Rectangle
        Get
            Dim capHeight As Integer = GetSystemMetrics(4) ' SM_CYCAPTION
            Dim frameWidth As Integer = GetSystemMetrics(32) ' SM_CXFRAME

            Dim xOffset As Integer = If(Me.WindowState = FormWindowState.Maximized, 52, 52)
            Return New Rectangle(Me.Width - xOffset, 10, 20, capHeight - 5)
        End Get
    End Property

    Private ReadOnly Property CButtonRect As Rectangle
        Get
            Dim capHeight As Integer = GetSystemMetrics(4) ' SM_CYCAPTION
            Dim customWidth As Integer = 120

            Return New Rectangle((Me.Width / 2) - (customWidth / 2), 8, customWidth, capHeight - 2)
        End Get
    End Property

    <DllImport("uxtheme.dll", ExactSpelling:=True, CharSet:=CharSet.Unicode)>
    Private Shared Function SetWindowTheme(ByVal hWnd As IntPtr, ByVal pszSubAppName As String, ByVal pszSubIdList As String) As Integer
    End Function

    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        Dim policy As Integer = DWMNCRP_DISABLED
        DwmSetWindowAttribute(Me.Handle, DWMWA_NCRENDERING_POLICY, policy, Marshal.SizeOf(policy))

        'SetWindowTheme(MonthCalendar1.Handle, "", "")
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        Select Case m.Msg
            Case WM_NCPAINT, WM_NCACTIVATE, WM_SETTEXT
                MyBase.WndProc(m)
                DrawCustomButton()
                Return

            Case WM_NCLBUTTONDOWN
                If IsMouseInButton(m.LParam) Then
                    isButtonPressed = True
                    DrawCustomButton()
                    m.Result = IntPtr.Zero
                    Return
                End If

            Case WM_NCLBUTTONDBLCLK
                If IsMouseInButton(m.LParam) Then
                    isButtonPressed = True
                    DrawCustomButton()
                    m.Result = IntPtr.Zero
                    Return
                End If

            Case WM_NCMOUSEMOVE
                If isButtonPressed AndAlso Not IsMouseInButton(m.LParam) Then
                    isButtonPressed = False
                    DrawCustomButton()
                End If

            Case WM_NCLBUTTONUP
                If isButtonPressed Then
                    isButtonPressed = False
                    DrawCustomButton()

                    If Me.TopMost = True Then Me.TopMost = False Else Me.TopMost = True

                    RedrawWindow(Me.Handle, IntPtr.Zero, IntPtr.Zero, RDW_FRAME Or RDW_INVALIDATE Or RDW_UPDATENOW)

                    m.Result = IntPtr.Zero
                    Return
                End If
        End Select

        MyBase.WndProc(m)
    End Sub

    Private Function IsMouseInButton(lParam As IntPtr) As Boolean
        Dim screenPoint As New Point(lParam.ToInt32() And &HFFFF, (lParam.ToInt32() >> 16) And &HFFFF)
        Dim windowRect As New Rectangle(Me.Left, Me.Top, Me.Width, Me.Height)
        Dim relativePoint As New Point(screenPoint.X - windowRect.X, screenPoint.Y - windowRect.Y)

        Return CustomButtonRect.Contains(relativePoint)
    End Function

    Private Sub DrawCustomButton()
        Dim hdc As IntPtr = GetWindowDC(Me.Handle)
        If hdc <> IntPtr.Zero Then
            Using g As Graphics = Graphics.FromHdc(hdc)
                Dim state As ButtonState = If(isButtonPressed, ButtonState.Pushed, ButtonState.Normal)

                If Me.TopMost = True Then
                    ControlPaint.DrawScrollButton(g, CustomButtonRect, ScrollButton.Down, state)

                    ControlPaint.DrawBorder(g, CButtonRect, Color.LimeGreen, ButtonBorderStyle.Dashed)
                    ControlPaint.DrawStringDisabled(g, "Top Most Activated.", Me.Font, Me.ForeColor, CButtonRect, 5)
                Else
                    ControlPaint.DrawScrollButton(g, CustomButtonRect, ScrollButton.Min, state)

                    ControlPaint.DrawBorder(g, CButtonRect, Color.Red, ButtonBorderStyle.Dashed)
                    ControlPaint.DrawStringDisabled(g, "Top Most Dectivated.", Me.Font, Me.ForeColor, CButtonRect, 5)
                End If

                'ControlPaint.DrawCaptionButton(g, CustomButtonRect, CaptionButton.Maximize, state)
            End Using
            ReleaseDC(Me.Handle, hdc)
        End If
    End Sub

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As Integer, lParam As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function ReleaseCapture() As Boolean
    End Function

    Private Sub CustomCaption_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        If e.Button = MouseButtons.Left Then
            ReleaseCapture()
            SendMessage(Me.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0)
        End If
    End Sub

    Private Sub btnCustom_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub DateTime_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim bmp As Bitmap = My.Resources.Clock

        Dim hIcon As IntPtr = bmp.GetHicon()

        Dim windowIco As Icon = Icon.FromHandle(hIcon)

        Me.Icon = windowIco

        ' Position
        Me.Location = New Point(SystemInformation.WorkingArea.Width - Me.Width, SystemInformation.WorkingArea.Height - Me.Height)

        ' Time And Date (page)
        currentTime = DateTime.Now
        Dim timeZones = TimeZoneInfo.GetSystemTimeZones()

        TDLP.Items.Clear()

        TDLP.DataSource = timeZones
        TDLP.DisplayMember = "DisplayName"
        TDLP.ValueMember = "Id"

        Dim mainId As String = TimeZoneInfo.Local.Id

        TDLP.SelectedValue = mainId

        ' Timer (page)
        ComboBox1.Items.Clear()

        ComboBox1.DataSource = [Enum].GetNames(GetType(MessageBoxIcon))

        DefAlarmFilePath = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\DateTime\Timer", "DefaultSoundFile", DefAlarmFilePath)
        DefVolume = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\DateTime\Timer", "DefaultVolume", DefVolume)
        DefIsLoop = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\DateTime\Timer", "DefaultIsLoop", DefIsLoop)

        ' Alarms (page)
        Try
            For Each i In My.Computer.Registry.CurrentUser.OpenSubKey("Software\Shell\DateTime\Alarms", True).GetSubKeyNames
                AlarmList.Items.Add(i.Substring(i.LastIndexOf("\") + 1))
            Next
        Catch ex As Exception
            My.Computer.Registry.CurrentUser.CreateSubKey("Software\Shell\DateTime\Alarms", True)

            Console.WriteLine("Failed to locate Registry key for alarms, attempting to create in: HKEY_CURRENT_USER\Software\Shell\DateTime\Alarms")
        End Try

        ' Other
        'TabControl1.SelectTab(3)
    End Sub

    Private g As Graphics
    Private centerX As Integer
    Private centerY As Integer
    Private radius As Integer

    Private MainColor As Color = Color.Black
    Private SecondHandColor As Color = Color.Red
    Private ClockSize As Integer = 200

    Dim currentTime As DateTime

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Dim zoneId As String = TDLP.SelectedValue.ToString()
        Dim selectedZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById(zoneId)

        Dim timeInRegion As DateTime = TimeZoneInfo.ConvertTime(DateTime.Now, selectedZone)

        Dim formattedTime As String = timeInRegion.ToString("HH:mm:ss")

        Label1.Text = formattedTime
        Label2.Text = timeInRegion.ToString("dd.MM.yyyy")

        PictureBox1.Invalidate()
    End Sub

    Private Sub PictureBox1_Paint(sender As Object, e As PaintEventArgs) Handles PictureBox1.Paint, PictureBox2.Paint

        Dim g As Graphics = e.Graphics
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

        Dim centerX As Integer = PictureBox1.Width / 2
        Dim centerY As Integer = PictureBox1.Height / 2

        Dim maxRadius As Single = Math.Min(centerX, centerY) * 0.95R

        Dim zoneId As String = TDLP.SelectedValue.ToString()
        Dim selectedZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById(zoneId)

        Dim currentTime As DateTime = TimeZoneInfo.ConvertTime(DateTime.Now, selectedZone)

        g.DrawEllipse(New Pen(MainColor, 2), centerX - maxRadius, centerY - maxRadius, maxRadius * 2, maxRadius * 2)

        For i As Integer = 1 To 60
            Dim angleDeg As Double = i * 6 - 90
            Dim angleRad As Double = angleDeg * Math.PI / 180.0

            Dim startLen As Single, endLen As Single
            Dim thickness As Integer

            Dim penColor As Color = Color.DarkGray

            If i Mod 5 = 0 Then
                ' Hours
                startLen = maxRadius * 0.85R
                endLen = maxRadius * 0.95R
                thickness = 2
                penColor = MainColor
            Else
                ' Minutes
                startLen = maxRadius * 0.9R
                endLen = maxRadius * 0.95R
                thickness = 1
            End If

            ' Angles (X = r * cos, Y = r * sin)
            Dim startX As Single = centerX + startLen * CSng(Math.Cos(angleRad))
            Dim startY As Single = centerY + startLen * CSng(Math.Sin(angleRad))
            Dim endX As Single = centerX + endLen * CSng(Math.Cos(angleRad))
            Dim endY As Single = centerY + endLen * CSng(Math.Sin(angleRad))

            g.DrawLine(New Pen(penColor, thickness), startX, startY, endX, endY)
        Next

        ' Numbers
        Dim fontSize As Integer = Math.Max(10, CInt(maxRadius / 15))

        Using font As New Font("Arial", fontSize, FontStyle.Bold)
            Using brush As New SolidBrush(MainColor)

                Dim numberRadius As Single = maxRadius * 0.7R

                For i As Integer = 1 To 12
                    Dim hour As Integer = If(i = 12, 12, i)

                    Dim angleDeg As Double = i * 30 - 90
                    Dim angleRad As Double = angleDeg * Math.PI / 180.0

                    Dim textToDraw As String = hour.ToString()

                    Dim textX As Single = centerX + numberRadius * CSng(Math.Cos(angleRad))
                    Dim textY As Single = centerY + numberRadius * CSng(Math.Sin(angleRad))

                    Dim size As SizeF = g.MeasureString(textToDraw, font)

                    g.DrawString(textToDraw, font, brush, textX - size.Width / 2, textY - size.Height / 2)
                Next
            End Using
        End Using

        ' Drawing Arrows
        Dim hourAngle As Single = CSng((currentTime.Hour Mod 12 + currentTime.Minute / 60.0 + currentTime.Second / 3600.0) * 30.0)
        Dim minuteAngle As Single = CSng((currentTime.Minute + currentTime.Second / 60.0) * 6.0)
        Dim secondAngle As Single = CSng(currentTime.Second * 6.0)

        DrawHand(g, centerX, centerY, hourAngle, maxRadius * 0.5F, 6, New Pen(MainColor))

        DrawHand(g, centerX, centerY, minuteAngle, maxRadius * 0.75F, 3, New Pen(MainColor))

        DrawHand(g, centerX, centerY, secondAngle, maxRadius * 0.85F, 1, New Pen(SecondHandColor))

        g.FillEllipse(New SolidBrush(MainColor), centerX - 5, centerY - 5, 10, 10)
    End Sub

    Private Sub DrawHand(ByVal g As Graphics, ByVal centerX As Integer, ByVal centerY As Integer, ByVal angle As Single, ByVal length As Single, ByVal thickness As Integer, ByVal penColor As Pen)

        Dim radians As Single = CSng((angle - 90) * Math.PI / 180.0)

        Dim endX As Single = centerX + length * CSng(Math.Cos(radians))
        Dim endY As Single = centerY + length * CSng(Math.Sin(radians))

        Using handPen As New Pen(penColor.Color, thickness)
            handPen.StartCap = Drawing2D.LineCap.Round
            handPen.EndCap = Drawing2D.LineCap.Round

            g.DrawLine(handPen, centerX, centerY, endX, endY)
        End Using
    End Sub

    Private Sub TDLP_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TDLP.SelectedIndexChanged
        Try
            Dim zoneId As String = TDLP.SelectedValue.ToString()
            Dim selectedZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById(zoneId)

            Dim timeInRegion As DateTime = TimeZoneInfo.ConvertTime(DateTime.Now, selectedZone)

            currentTime = timeInRegion
            Label1.Text = timeInRegion.ToString("HH:mm:ss")
            Label2.Text = timeInRegion.ToString("dd.MM.yyyy")
        Catch ex As Exception

        End Try
    End Sub

    '------------STOP-WATCH--------------
    Private ReadOnly watch As New Stopwatch()
    Private offset As New TimeSpan(0)

    Private isCountdown As Boolean = False
    Private countdownSeconds As Integer = 0
    Private countdownStartTime As DateTime

    Private Sub tmrUpdate_Tick(sender As Object, e As EventArgs) Handles tmrSWUpdate.Tick
        If isCountdown Then
            UpdateCountdown()
        Else
            UpdateLabel()
        End If
    End Sub

    Private Sub UpdateCountdown()
        Dim elapsed As TimeSpan = DateTime.Now - countdownStartTime
        Dim remaining As Double = countdownSeconds - elapsed.TotalSeconds

        If remaining <= 0 Then
            StartStopwatch()
            UpdateLabel()
            btnSWPause.Enabled = True
        Else
            Label4.Text = Math.Ceiling(remaining).ToString("0")
            btnSWPause.Enabled = False
            btnSWPause.Text = "Pause"
        End If
    End Sub

    Private Sub btnSWStart_Click(sender As Object, e As EventArgs) Handles btnSWStart.Click
        If watch.IsRunning OrElse isCountdown Then
            watch.Reset()
            isCountdown = False
            offset = New TimeSpan(0)
            tmrSWUpdate.Stop()

            UpdateLabel()
            btnSWStart.Text = "Start"
            btnSWPause.Enabled = False
            btnSWPause.Text = "Pause"

            chkSWCountdown.Enabled = True
            If chkSWCountdown.Checked = True Then nudSWCountdown.Enabled = True
        Else
            If chkSWCountdown.Checked Then
                countdownSeconds = CInt(nudSWCountdown.Value)

                If countdownSeconds > 0 Then
                    isCountdown = True
                    countdownStartTime = DateTime.Now
                    tmrSWUpdate.Start()
                    UpdateCountdown()

                    btnSWPause.Enabled = False
                    btnSWPause.Text = "Pause"
                Else
                    StartStopwatch()
                    UpdateLabel()
                End If
            Else
                StartStopwatch()
                UpdateLabel()
            End If

            btnSWStart.Text = "Stop"
            btnSWPause.Enabled = True

            chkSWCountdown.Enabled = False
            nudSWCountdown.Enabled = False
        End If
    End Sub

    Private Sub StartStopwatch()
        isCountdown = False
        watch.Start()
        tmrSWUpdate.Start()
    End Sub

    Private Sub btnSWPause_Click(sender As Object, e As EventArgs) Handles btnSWPause.Click
        If watch.IsRunning Then
            watch.Stop()
            btnSWPause.Text = "Continue"
        Else
            watch.Start()
            btnSWPause.Text = "Pause"
        End If
    End Sub

    Private Sub tmrSWUpdate_Tick(sender As Object, e As EventArgs) Handles tmrSWUpdate.Tick
        If isCountdown Then
            UpdateCountdown()
        Else
            UpdateLabel()
        End If
    End Sub

    Private Sub UpdateLabel()
        Dim totalTime As TimeSpan = watch.Elapsed.Add(offset)

        Dim timeParts As New List(Of String)

        If chkSWDays.Checked Then timeParts.Add(Math.Floor(totalTime.TotalDays).ToString("0"))
        If chkSWHours.Checked Then timeParts.Add(totalTime.Hours.ToString("00"))
        If chkSWMinutes.Checked Then timeParts.Add(totalTime.Minutes.ToString("00"))

        timeParts.Add(totalTime.Seconds.ToString("00"))

        Dim mainTime As String = String.Join(":", timeParts)

        If chkSWMilliseconds.Checked Then
            Label4.Text = mainTime & "." & totalTime.Milliseconds.ToString("000")
        Else
            Label4.Text = mainTime
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim newInstance As New DateAndTime

        newInstance.Show()
        newInstance.TabControl1.SelectTab(TabControl1.SelectedIndex)
    End Sub

    Private Sub chkOptions_CheckedChanged(sender As Object, e As EventArgs) Handles chkSWMilliseconds.CheckedChanged, chkSWMinutes.CheckedChanged, chkSWHours.CheckedChanged, chkSWDays.CheckedChanged
        UpdateLabel()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        offset = offset.Add(TimeSpan.FromMilliseconds(nudSWMilliseconds.Value))
        offset = offset.Add(TimeSpan.FromSeconds(nudSWSeconds.Value))
        offset = offset.Add(TimeSpan.FromMinutes(nudSWMinutes.Value))
        offset = offset.Add(TimeSpan.FromHours(nudSWHours.Value))
        offset = offset.Add(TimeSpan.FromDays(nudSWDays.Value))

        UpdateLabel()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        offset = TimeSpan.FromMilliseconds(nudSWMilliseconds.Value)
        offset = TimeSpan.FromSeconds(nudSWSeconds.Value)
        offset = TimeSpan.FromMinutes(nudSWMinutes.Value)
        offset = TimeSpan.FromHours(nudSWHours.Value)
        offset = TimeSpan.FromDays(nudSWDays.Value)

        UpdateLabel()
    End Sub

    Private Sub chkSWCountdown_CheckedChanged(sender As Object, e As EventArgs) Handles chkSWCountdown.CheckedChanged
        nudSWCountdown.Enabled = chkSWCountdown.Checked
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        txtStoplist.Text += Label4.Text & Environment.NewLine
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If String.IsNullOrEmpty(txtStoplist.Text) Then Exit Sub

        If CheckBox7.Checked = True Then
            If MessageBox.Show("Do you really want to reset your Stoplist?", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                txtStoplist.Clear()
            End If
        Else
            txtStoplist.Clear()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim SFD As New SaveFileDialog With {
        .AutoUpgradeEnabled = AppBar.UseExplorerFP,
        .CreatePrompt = True,
        .SupportMultiDottedExtensions = True,
        .Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
        .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        .ValidateNames = True,
        .FileName = ""
        }

        If SFD.ShowDialog(Me) = DialogResult.OK Then
            Try
                Using sw As New StreamWriter(SFD.FileName, False)
                    sw.Write(txtStoplist.Text)
                    sw.Close()
                End Using
            Catch ex As Exception
                MsgBox("Failed to save a file. " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim OFD As New OpenFileDialog With {
        .AutoUpgradeEnabled = AppBar.UseExplorerFP,
        .SupportMultiDottedExtensions = True,
        .Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
        .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        .ValidateNames = True,
        .FileName = "",
        .CheckFileExists = True
        }

        If OFD.ShowDialog(Me) = DialogResult.OK Then
            Try
                Using sr As New StreamReader(OFD.FileName, False)
                    txtStoplist.Text = sr.ReadToEnd
                    sr.Close()
                End Using
            Catch ex As Exception
                MsgBox("Failed to open a file. " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub

    '--------------TIMER---------------
    Private TotalSeconds As Integer = 0
    Private IsPaused As Boolean = False

    Private Player As New SoundPlayer

    Public DefAlarmFilePath As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\Media\Alarm01.wav"
    Public DefVolume As Integer = 50
    Public DefIsLoop As Boolean = True

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnTStart.Click
        If TotalSeconds > 0 Then
            If Not IsPaused Then
                If tmrTCountdown.Enabled = True Then
                    tmrTCountdown.Stop()
                    IsPaused = False
                    TotalSeconds = 0
                    UpdateDisplay()

                    btnTStart.Text = "Start"
                    btnTPause.Text = "Pause"
                    btnTPause.Enabled = False
                Else
                    TotalSeconds = (nudTDays.Value * 86400) + (nudTHours.Value * 3600) + (nudTMinutes.Value * 60) + nudTSeconds.Value

                    tmrTCountdown.Start()
                    IsPaused = False
                    UpdateDisplay()

                    btnTStart.Text = "Stop"
                    btnTPause.Enabled = True
                End If
            End If
        Else
            MessageBox.Show("Please set a time greater than zero.")
        End If
    End Sub

    Private Function GetSecondsFromInputs() As Integer
        Dim ds As Integer = nudTDays.Value * 86400
        Dim hs As Integer = nudTHours.Value * 3600
        Dim ms As Integer = nudTMinutes.Value * 60
        Dim s As Integer = nudTSeconds.Value
        Return ds + hs + ms + s
    End Function

    Private Sub btnPause_Click(sender As Object, e As EventArgs) Handles btnTPause.Click
        If IsPaused = True Then
            tmrTCountdown.Start()
            IsPaused = False

            btnTPause.Text = "Pause"
        Else
            tmrTCountdown.Stop()
            IsPaused = True

            btnTPause.Text = "Continue"
        End If
    End Sub

    Private Sub CountdownTimer_Tick(sender As Object, e As EventArgs) Handles tmrTCountdown.Tick
        If TotalSeconds > 0 Then
            TotalSeconds -= 1
            UpdateDisplay()
        Else
            tmrTCountdown.Stop()

            If CheckBox2.Checked = True Then
                If CheckBox3.Checked = True Then PlayAlarm()

                If CheckBox4.Checked = True Then
                    Dim selectedIcon As MessageBoxIcon
                    selectedIcon = CType([Enum].Parse(GetType(MessageBoxIcon), ComboBox1.SelectedItem.ToString()), MessageBoxIcon)
                    MessageBox.Show("Time is up!", "Alarm", MessageBoxButtons.OK, selectedIcon, MessageBoxDefaultButton.Button1)
                End If
            End If

            Player.Stop()

            btnTStart.Text = "Start"
            btnTPause.Text = "Pause"
            btnTPause.Enabled = False
        End If
    End Sub

    Private Sub UpdateDisplay()
        Dim time As TimeSpan = TimeSpan.FromSeconds(TotalSeconds)
        lblTDisplay.Text = String.Format("{0}:{1:00}:{2:00}:{3:00}",
                                      Math.Floor(time.TotalDays),
                                      time.Hours,
                                      time.Minutes,
                                      time.Seconds)
    End Sub

    Private Sub PlayAlarm()
        Try
            Player.SoundLocation = DefAlarmFilePath
            'Player.settings.volume = DefVolume
            If DefIsLoop = True Then Player.PlayLooping() Else Player.Play()

        Catch ex As Exception
            Console.Beep()
        End Try
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TotalSeconds = GetSecondsFromInputs()
        UpdateDisplay()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TotalSeconds += GetSecondsFromInputs()
        UpdateDisplay()
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            ComboBox1.Enabled = True
            CheckBox3.Enabled = True
            If CheckBox3.Checked = True Then Button11.Enabled = True Else Button11.Enabled = False
        Else
            ComboBox1.Enabled = False
            CheckBox3.Enabled = False
            Button11.Enabled = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        Button11.Enabled = CheckBox3.Checked
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        GroupBox6.Enabled = CheckBox2.Checked
    End Sub

    Public Sub UpdateRegistry()
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\DateTime\Timer", "DefaultSoundFile", DefAlarmFilePath, Microsoft.Win32.RegistryValueKind.String)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\DateTime\Timer", "DefaultVolume", DefVolume, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\DateTime\Timer", "DefaultIsLoop", DefIsLoop, Microsoft.Win32.RegistryValueKind.DWord)
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        DateAndTimeProperties.ComboBox1.Text = DefAlarmFilePath
        DateAndTimeProperties.numVolume.Value = DefVolume
        DateAndTimeProperties.chkLoop.Checked = DefIsLoop

        If DateAndTimeProperties.ShowDialog(Me) = DialogResult.OK Then

            DefAlarmFilePath = DateAndTimeProperties.ComboBox1.Text
            DefVolume = DateAndTimeProperties.numVolume.Value
            DefIsLoop = DateAndTimeProperties.chkLoop.Checked

            UpdateRegistry()
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        txtMarkList.Text += lblTDisplay.Text & Environment.NewLine
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If String.IsNullOrEmpty(txtMarkList.Text) Then Exit Sub

        If CheckBox7.Checked = True Then
            If MessageBox.Show("Do you really want to reset your Marklist?", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                txtMarkList.Clear()
            End If
        Else
            txtMarkList.Clear()
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim SFD As New SaveFileDialog With {
        .AutoUpgradeEnabled = AppBar.UseExplorerFP,
        .CreatePrompt = True,
        .SupportMultiDottedExtensions = True,
        .Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
        .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        .ValidateNames = True,
        .FileName = ""
        }

        If SFD.ShowDialog(Me) = DialogResult.OK Then
            Try
                Using sw As New StreamWriter(SFD.FileName, False)
                    sw.Write(txtMarkList.Text)
                    sw.Close()
                End Using
            Catch ex As Exception
                MsgBox("Failed to save a file. " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim OFD As New OpenFileDialog With {
        .AutoUpgradeEnabled = AppBar.UseExplorerFP,
        .SupportMultiDottedExtensions = True,
        .Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
        .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        .ValidateNames = True,
        .FileName = "",
        .CheckFileExists = True
        }

        If OFD.ShowDialog(Me) = DialogResult.OK Then
            Try
                Using sr As New StreamReader(OFD.FileName, False)
                    txtMarkList.Text = sr.ReadToEnd
                    sr.Close()
                End Using
            Catch ex As Exception
                MsgBox("Failed to open a file. " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        On Error Resume Next
        Process.Start(New ProcessStartInfo(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "TimeDateProperties", Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\System32\timedate.cpl")) With {.UseShellExecute = True})
    End Sub
End Class