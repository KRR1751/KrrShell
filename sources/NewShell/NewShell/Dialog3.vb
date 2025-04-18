Imports System.Windows.Forms
Public Class Dialog3
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim value As CreateParams = MyBase.CreateParams
            'Don't allow the window to be activated.
            value.ExStyle = value.ExStyle Or NativeConstants.WS_EX_NOACTIVATE
            Return value
        End Get
    End Property

    Public Overloads Sub ShowUnactivated()
        NativeMethods.ShowWindow(Me.Handle, NativeConstants.SW_SHOWNOACTIVATE)
    End Sub

    Public Overloads Sub ShowUnactivated(ByVal owner As Form)
        Me.Owner = owner
        Me.ShowUnactivated()
    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        'If CheckBox1.Checked = True Then
        If m.Msg = NativeConstants.WM_MOUSEACTIVATE Then
            'Tell Windows not to activate the window on a mouse click.
            m.Result = CType(NativeConstants.MA_NOACTIVATE, IntPtr)
        Else
            MyBase.WndProc(m)
        End If
        'End If
    End Sub

    Friend NotInheritable Class NativeMethods
        <Runtime.InteropServices.DllImport("user32")> _
        Public Shared Function ShowWindow(ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
        End Function
    End Class

    Friend NotInheritable Class NativeConstants
        Public Const WS_EX_NOACTIVATE As Integer = &H8000000
        Public Const WM_MOUSEACTIVATE As Integer = &H21
        Public Const MA_NOACTIVATE As Integer = 3
        Public Const SW_SHOWNOACTIVATE As Integer = 4
    End Class
    'TIMEDATE---------------------------------------
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox1.Text = "📎"
            Me.TopMost = True
        Else
            CheckBox1.Text = "📌"
            Me.TopMost = False
        End If
    End Sub
    Dim m_windowdrag As System.Drawing.Point
    Private Sub Panel1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseDown
        m_windowdrag = New Point(-e.X, -e.Y)
    End Sub

    Private Sub Panel1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseMove
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Dim mlocation As Point
            mlocation = Control.MousePosition
            mlocation.Offset(m_windowdrag.X, m_windowdrag.Y)
            Me.Location = mlocation
        End If
    End Sub

    Private Sub Dialog3_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        If CheckBox1.Checked = False Then
            Me.Hide()
        End If
    End Sub

    Private Sub Dialog3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Location = New Point(My.Computer.Screen.WorkingArea.Width - Me.Width, My.Computer.Screen.WorkingArea.Height - Me.Height)
        TimeToCopy.Hide()
        TimeSToCopy.Hide()
        TimeSMToCopy.Hide()
        DateToCopy.Hide()
        DateToCopyO.Hide()
        SAddC.SelectedIndex = 2
        TAddC.SelectedIndex = 2
        DateTimePicker1.MinDate = Date.Today.AddDays(-14)

        DomainUpDown1.SelectedIndex = 0
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        If mls = 0 Then
            mlsf = "00"
        ElseIf mls = 1 Then
            mlsf = "01"
        ElseIf mls = 2 Then
            mlsf = "02"
        ElseIf mls = 3 Then
            mlsf = "03"
        ElseIf mls = 4 Then
            mlsf = "04"
        ElseIf mls = 5 Then
            mlsf = "05"
        ElseIf mls = 6 Then
            mlsf = "06"
        ElseIf mls = 7 Then
            mlsf = "07"
        ElseIf mls = 8 Then
            mlsf = "08"
        ElseIf mls = 9 Then
            mlsf = "09"
        Else
            mlsf = mls
        End If
        'Seconds
        If s = 0 Then
            sf = "00"
        ElseIf s = 1 Then
            sf = "01"
        ElseIf s = 2 Then
            sf = "02"
        ElseIf s = 3 Then
            sf = "03"
        ElseIf s = 4 Then
            sf = "04"
        ElseIf s = 5 Then
            sf = "05"
        ElseIf s = 6 Then
            sf = "06"
        ElseIf s = 7 Then
            sf = "07"
        ElseIf s = 8 Then
            sf = "08"
        ElseIf s = 9 Then
            sf = "09"
        Else
            sf = s
        End If
        'Minutes
        If m = 0 Then
            mf = "00"
        ElseIf m = 1 Then
            mf = "01"
        ElseIf m = 2 Then
            mf = "02"
        ElseIf m = 3 Then
            mf = "03"
        ElseIf m = 4 Then
            mf = "04"
        ElseIf m = 5 Then
            mf = "05"
        ElseIf m = 6 Then
            mf = "06"
        ElseIf m = 7 Then
            mf = "07"
        ElseIf m = 8 Then
            mf = "08"
        ElseIf m = 9 Then
            mf = "09"
        Else
            mf = m
        End If
        'Hours
        If h = 0 Then
            hf = "00"
        ElseIf h = 1 Then
            hf = "01"
        ElseIf h = 2 Then
            hf = "02"
        ElseIf h = 3 Then
            hf = "03"
        ElseIf h = 4 Then
            hf = "04"
        ElseIf h = 5 Then
            hf = "05"
        ElseIf h = 6 Then
            hf = "06"
        ElseIf h = 7 Then
            hf = "07"
        ElseIf h = 8 Then
            hf = "08"
        ElseIf h = 9 Then
            hf = "09"
        Else
            hf = h
        End If
    End Sub
    Dim fixmin As String
    Dim fixsec As String
    Dim fixday As String
    Dim fixmon As String
    Dim fixdayofweek As String
    Dim fixmont As String
    Private Sub Controller_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ControllerTAD.Tick
        Time.Text = AppBar.TimeLabel.Text
        'Minutees
        If DateTime.Now.Minute = 0 Then
            fixmin = "00"
        ElseIf DateTime.Now.Minute = 1 Then
            fixmin = "01"
        ElseIf DateTime.Now.Minute = 2 Then
            fixmin = "02"
        ElseIf DateTime.Now.Minute = 3 Then
            fixmin = "03"
        ElseIf DateTime.Now.Minute = 4 Then
            fixmin = "04"
        ElseIf DateTime.Now.Minute = 5 Then
            fixmin = "05"
        ElseIf DateTime.Now.Minute = 6 Then
            fixmin = "06"
        ElseIf DateTime.Now.Minute = 7 Then
            fixmin = "07"
        ElseIf DateTime.Now.Minute = 8 Then
            fixmin = "08"
        ElseIf DateTime.Now.Minute = 9 Then
            fixmin = "09"
        Else
            fixmin = DateTime.Now.Minute
        End If
        'Secoonds
        If DateTime.Now.Second = 0 Then
            fixsec = "00"
        ElseIf DateTime.Now.Second = 1 Then
            fixsec = "01"
        ElseIf DateTime.Now.Second = 2 Then
            fixsec = "02"
        ElseIf DateTime.Now.Second = 3 Then
            fixsec = "03"
        ElseIf DateTime.Now.Second = 4 Then
            fixsec = "04"
        ElseIf DateTime.Now.Second = 5 Then
            fixsec = "05"
        ElseIf DateTime.Now.Second = 6 Then
            fixsec = "06"
        ElseIf DateTime.Now.Second = 7 Then
            fixsec = "07"
        ElseIf DateTime.Now.Second = 8 Then
            fixsec = "08"
        ElseIf DateTime.Now.Second = 9 Then
            fixsec = "09"
        Else
            fixsec = DateTime.Now.Second
        End If
        'Day
        If DateTime.Now.Day = 0 Then
            fixday = "00"
        ElseIf DateTime.Now.Day = 1 Then
            fixday = "01"
        ElseIf DateTime.Now.Day = 2 Then
            fixday = "02"
        ElseIf DateTime.Now.Day = 3 Then
            fixday = "03"
        ElseIf DateTime.Now.Day = 4 Then
            fixday = "04"
        ElseIf DateTime.Now.Day = 5 Then
            fixday = "05"
        ElseIf DateTime.Now.Day = 6 Then
            fixday = "06"
        ElseIf DateTime.Now.Day = 7 Then
            fixday = "07"
        ElseIf DateTime.Now.Day = 8 Then
            fixday = "08"
        ElseIf DateTime.Now.Day = 9 Then
            fixday = "09"
        Else
            fixday = DateTime.Now.Day
        End If
        'Month
        If DateTime.Now.Month = 0 Then
            fixmon = "00"
        ElseIf DateTime.Now.Month = 1 Then
            fixmon = "01"
        ElseIf DateTime.Now.Month = 2 Then
            fixmon = "02"
        ElseIf DateTime.Now.Month = 3 Then
            fixmon = "03"
        ElseIf DateTime.Now.Month = 4 Then
            fixmon = "04"
        ElseIf DateTime.Now.Month = 5 Then
            fixmon = "05"
        ElseIf DateTime.Now.Month = 6 Then
            fixmon = "06"
        ElseIf DateTime.Now.Month = 7 Then
            fixmon = "07"
        ElseIf DateTime.Now.Month = 8 Then
            fixmon = "08"
        ElseIf DateTime.Now.Month = 9 Then
            fixmon = "09"
        Else
            fixmon = DateTime.Now.Month
        End If
        'DayOfWeek
        If DateTime.Now.DayOfWeek = 1 Then
            fixdayofweek = "Monday"
        ElseIf DateTime.Now.DayOfWeek = 2 Then
            fixdayofweek = "Tuesday"
        ElseIf DateTime.Now.DayOfWeek = 3 Then
            fixdayofweek = "Wednesday"
        ElseIf DateTime.Now.DayOfWeek = 4 Then
            fixdayofweek = "Thursday"
        ElseIf DateTime.Now.DayOfWeek = 5 Then
            fixdayofweek = "Friday"
        ElseIf DateTime.Now.DayOfWeek = 6 Then
            fixdayofweek = "Saturday"
        ElseIf DateTime.Now.DayOfWeek = 7 Then
            fixdayofweek = "Sunday"
        ElseIf DateTime.Now.DayOfWeek = 8 Then
            fixdayofweek = "DayOfErrors:)"
        ElseIf Not DateTime.Now.DayOfWeek >= 8 Then
            fixdayofweek = "Error On Us: <Date Error 404>"
        Else
            fixdayofweek = DateTime.Now.DayOfWeek
        End If
        'MonthNames!
        If DateTime.Now.Month = 0 Then
            fixmont = "Illegal Month Today(0)"
        ElseIf DateTime.Now.Month = 1 Then
            fixmont = "January"
        ElseIf DateTime.Now.Month = 2 Then
            fixmont = "February"
        ElseIf DateTime.Now.Month = 3 Then
            fixmont = "March"
        ElseIf DateTime.Now.Month = 4 Then
            fixmont = "April"
        ElseIf DateTime.Now.Month = 5 Then
            fixmont = "May"
        ElseIf DateTime.Now.Month = 6 Then
            fixmont = "June"
        ElseIf DateTime.Now.Month = 7 Then
            fixmont = "July"
        ElseIf DateTime.Now.Month = 8 Then
            fixmont = "August"
        ElseIf DateTime.Now.Month = 9 Then
            fixmont = "September"
        ElseIf DateTime.Now.Month = 10 Then
            fixmont = "October"
        ElseIf DateTime.Now.Month = 11 Then
            fixmont = "November"
        ElseIf DateTime.Now.Month = 12 Then
            fixmont = "December"
        Else
            fixmont = "Error <Month Is Higher or Less than Normal. 404>"
        End If
        Label1.Text = "Ticks: " & DateTime.Now.Ticks
        Label2.Text = "Time In Binary: "
        Label3.Text = DateTime.Now.ToBinary()
        TimeToCopy.Text = DateTime.Now.Hour & ":" & fixmin
        TimeSToCopy.Text = Time.Text
        TimeSMToCopy.Text = DateTime.Now.Hour & ":" & fixmin & ":" & fixsec & "." & DateTime.Now.Millisecond

        If MouseLocker.Enabled = False Then
            ML = Control.MousePosition
        End If
    End Sub
    Dim ML As New Point(0, 0)
    Private Sub WithSecondsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithSecondsToolStripMenuItem.Click
        TimeSToCopy.SelectAll()
        TimeSToCopy.Copy()
        ToolTip1.Active = True
        ToolTip1.Show("Successfully Copied!", Me, Time.Width / 2 + Time.Location.X, Time.Height / 2 + Time.Location.Y)
        ToolTipHider.Enabled = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TimeToCopy.SelectAll()
        TimeToCopy.Copy()
        ToolTip1.Active = True
        ToolTip1.Show("Successfully Copied!", Me, Time.Width / 2 + Time.Location.X, Time.Height / 2 + Time.Location.Y)
        ToolTipHider.Enabled = True
    End Sub

    Private Sub WithMilisecondsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WithMilisecondsToolStripMenuItem.Click
        TimeSMToCopy.SelectAll()
        TimeSMToCopy.Copy()
        ToolTip1.Active = True
        ToolTip1.Show("Successfully Copied!", Me, Time.Width / 2 + Time.Location.X, Time.Height / 2 + Time.Location.Y)
        ToolTipHider.Enabled = True
    End Sub

    Private Sub DateSettingsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateSettingsToolStripMenuItem.Click
        System.Diagnostics.Process.Start("C:\Windows\System32\timedate.cpl")
    End Sub

    Private Sub ToolTipHider_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolTipHider.Tick
        ToolTip1.Active = False
        ToolTipHider.Enabled = False
    End Sub

    Private Sub CopyDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyDateToolStripMenuItem.Click
        DateToCopy.Text = DateTime.Today
        DateToCopy.SelectAll()
        DateToCopy.Copy()
        DateCM.Hide()
        ToolTip1.Active = True
        ToolTip1.Show("Successfully Copied!", Me, MonthCalendar1.Width / 2 + MonthCalendar1.Location.X, 15 + MonthCalendar1.Location.Y)
        ToolTipHider.Enabled = True
    End Sub

    Private Sub DDMMYYYYToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DDMMYYYYToolStripMenuItem.Click
        DateToCopy.Text = fixday & "." & fixmon & "." & DateTime.Today.Year
        DateToCopy.SelectAll()
        DateToCopy.Copy()
        ToolTip1.Active = True
        ToolTip1.Show("Successfully Copied!", Me, MonthCalendar1.Width / 2 + MonthCalendar1.Location.X, 15 + MonthCalendar1.Location.Y)
        ToolTipHider.Enabled = True
    End Sub

    Private Sub DDMMYYYYToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DDMMYYYYToolStripMenuItem1.Click
        DateToCopy.Text = fixday & "/" & fixmon & "/" & DateTime.Today.Year
        DateToCopy.SelectAll()
        DateToCopy.Copy()
        ToolTip1.Active = True
        ToolTip1.Show("Successfully Copied!", Me, MonthCalendar1.Width / 2 + MonthCalendar1.Location.X, 15 + MonthCalendar1.Location.Y)
        ToolTipHider.Enabled = True
    End Sub

    Private Sub DDMMYYYYToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DDMMYYYYToolStripMenuItem2.Click
        DateToCopy.Text = fixday & "-" & fixmon & "-" & DateTime.Today.Year
        DateToCopy.SelectAll()
        DateToCopy.Copy()
        ToolTip1.Active = True
        ToolTip1.Show("Successfully Copied!", Me, MonthCalendar1.Width / 2 + MonthCalendar1.Location.X, 15 + MonthCalendar1.Location.Y)
        ToolTipHider.Enabled = True
    End Sub

    Private Sub YYYYMMDDToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles YYYYMMDDToolStripMenuItem.Click
        DateToCopy.Text = DateTime.Today.Year & "." & fixmon & "." & fixday
        DateToCopy.SelectAll()
        DateToCopy.Copy()
        ToolTip1.Active = True
        ToolTip1.Show("Successfully Copied!", Me, MonthCalendar1.Width / 2 + MonthCalendar1.Location.X, 15 + MonthCalendar1.Location.Y)
        ToolTipHider.Enabled = True
    End Sub

    Private Sub YYYYMMDDToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles YYYYMMDDToolStripMenuItem1.Click
        DateToCopy.Text = DateTime.Today.Year & "/" & fixmon & "/" & fixday
        DateToCopy.SelectAll()
        DateToCopy.Copy()
        ToolTip1.Active = True
        ToolTip1.Show("Successfully Copied!", Me, MonthCalendar1.Width / 2 + MonthCalendar1.Location.X, 15 + MonthCalendar1.Location.Y)
        ToolTipHider.Enabled = True
    End Sub

    Private Sub YYYYMMDDToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles YYYYMMDDToolStripMenuItem2.Click
        DateToCopy.Text = DateTime.Today.Year & "-" & fixmon & "-" & fixday
        DateToCopy.SelectAll()
        DateToCopy.Copy()
        ToolTip1.Active = True
        ToolTip1.Show("Successfully Copied!", Me, MonthCalendar1.Width / 2 + MonthCalendar1.Location.X, 15 + MonthCalendar1.Location.Y)
        ToolTipHider.Enabled = True
    End Sub

    Private Sub DayOfWeekToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DayOfWeekToolStripMenuItem.Click
        DateToCopy.Text = fixdayofweek
        DateToCopy.SelectAll()
        DateToCopy.Copy()
        ToolTip1.Active = True
        ToolTip1.Show("Successfully Copied!", Me, MonthCalendar1.Width / 2 + MonthCalendar1.Location.X, 15 + MonthCalendar1.Location.Y)
        ToolTipHider.Enabled = True
    End Sub

    Private Sub MonthToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MonthToolStripMenuItem.Click
        DateToCopy.Text = fixmont
        DateToCopy.SelectAll()
        DateToCopy.Copy()
        ToolTip1.Active = True
        ToolTip1.Show("Successfully Copied!", Me, MonthCalendar1.Width / 2 + MonthCalendar1.Location.X, 15 + MonthCalendar1.Location.Y)
        ToolTipHider.Enabled = True
    End Sub
    'Dim dateFromString As DateTime = DateTime.Parse(DateString, System.Globalization.CultureInfo.InvariantCulture)
    Dim TMZT As DateTime = New DateTime(Date.Now.Year, Date.Now.Month, Date.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, DateTimeKind.Local)
    Private Sub TMZ_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TMZ.Tick
        'TMZT = DateTime.Now
        'TMZT.AddHours(1)
        'TMZT.AddYears(35)
        'TMZT.ToString("dd.MM.yyyy hh:mm:ss")
        TMZT = DateTime.UtcNow.AddHours(1)
        If TDLP.SelectedIndex = 0 Then
            Label4.Text = TMZT.AddHours(-12)
        ElseIf TDLP.SelectedIndex = 1 Then
            Label4.Text = TMZT.AddHours(-11)
        ElseIf TDLP.SelectedIndex = 2 Then
            Label4.Text = TMZT.AddHours(-10)
        ElseIf TDLP.SelectedIndex = 3 Then
            Label4.Text = TMZT.AddHours(-9)
        ElseIf TDLP.SelectedIndex = 4 Then
            Label4.Text = TMZT.AddHours(-8)
        ElseIf TDLP.SelectedIndex = 5 Then
            Label4.Text = TMZT.AddHours(-7)
        ElseIf TDLP.SelectedIndex = 6 Then
            Label4.Text = TMZT.AddHours(-7)
        ElseIf TDLP.SelectedIndex = 7 Then
            Label4.Text = TMZT.AddHours(-6)
        ElseIf TDLP.SelectedIndex = 8 Then
            Label4.Text = TMZT.AddHours(-6)
        ElseIf TDLP.SelectedIndex = 9 Then
            Label4.Text = TMZT.AddHours(-6)
        ElseIf TDLP.SelectedIndex = 10 Then
            Label4.Text = TMZT.AddHours(-5)
        ElseIf TDLP.SelectedIndex = 11 Then
            Label4.Text = TMZT.AddHours(-5)
        ElseIf TDLP.SelectedIndex = 12 Then
            Label4.Text = TMZT.AddHours(-5)
        ElseIf TDLP.SelectedIndex = 13 Then
            Label4.Text = TMZT.AddHours(-4)
        ElseIf TDLP.SelectedIndex = 14 Then
            Label4.Text = TMZT.AddHours(-4)
        ElseIf TDLP.SelectedIndex = 15 Then
            Label4.Text = TMZT.AddHours(-4)
        ElseIf TDLP.SelectedIndex = 16 Then
            Label4.Text = TMZT.AddMinutes(-210)
        ElseIf TDLP.SelectedIndex = 17 Then
            Label4.Text = TMZT.AddHours(-3)
        ElseIf TDLP.SelectedIndex = 18 Then
            Label4.Text = TMZT.AddHours(-3)
        ElseIf TDLP.SelectedIndex = 19 Then
            Label4.Text = TMZT.AddHours(-2)
        ElseIf TDLP.SelectedIndex = 20 Then
            Label4.Text = TMZT.AddHours(-1)
        ElseIf TDLP.SelectedIndex = 21 Then
            Label4.Text = TMZT.AddHours(0)
        ElseIf TDLP.SelectedIndex = 22 Then
            Label4.Text = TMZT.AddHours(0)
        ElseIf TDLP.SelectedIndex = 23 Then
            Label4.Text = TMZT.AddHours(1)
        ElseIf TDLP.SelectedIndex = 24 Then
            Label4.Text = TMZT.AddHours(1)
        ElseIf TDLP.SelectedIndex = 25 Then
            Label4.Text = TMZT.AddHours(1)
        ElseIf TDLP.SelectedIndex = 26 Then
            Label4.Text = TMZT.AddHours(1)
        ElseIf TDLP.SelectedIndex = 27 Then
            Label4.Text = TMZT.AddHours(2)
        ElseIf TDLP.SelectedIndex = 28 Then
            Label4.Text = TMZT.AddHours(2)
        ElseIf TDLP.SelectedIndex = 29 Then
            Label4.Text = TMZT.AddHours(2)
        ElseIf TDLP.SelectedIndex = 30 Then
            Label4.Text = TMZT.AddHours(2)
        ElseIf TDLP.SelectedIndex = 31 Then
            Label4.Text = TMZT.AddHours(2)
        ElseIf TDLP.SelectedIndex = 32 Then
            Label4.Text = TMZT.AddHours(2)
        ElseIf TDLP.SelectedIndex = 33 Then
            Label4.Text = TMZT.AddHours(3)
        ElseIf TDLP.SelectedIndex = 34 Then
            Label4.Text = TMZT.AddHours(3)
        ElseIf TDLP.SelectedIndex = 35 Then
            Label4.Text = TMZT.AddHours(3)
        ElseIf TDLP.SelectedIndex = 36 Then
            Label4.Text = TMZT.AddMinutes(210)
        ElseIf TDLP.SelectedIndex = 37 Then
            Label4.Text = TMZT.AddHours(4)
        ElseIf TDLP.SelectedIndex = 38 Then
            Label4.Text = TMZT.AddHours(4)
        ElseIf TDLP.SelectedIndex = 39 Then
            Label4.Text = TMZT.AddMinutes(270)
        ElseIf TDLP.SelectedIndex = 40 Then
            Label4.Text = TMZT.AddHours(5)
        ElseIf TDLP.SelectedIndex = 41 Then
            Label4.Text = TMZT.AddHours(5)
        ElseIf TDLP.SelectedIndex = 42 Then
            Label4.Text = TMZT.AddMinutes(330)
        ElseIf TDLP.SelectedIndex = 43 Then
            Label4.Text = TMZT.AddHours(6)
        ElseIf TDLP.SelectedIndex = 44 Then
            Label4.Text = TMZT.AddHours(6)
        ElseIf TDLP.SelectedIndex = 45 Then
            Label4.Text = TMZT.AddHours(7)
        ElseIf TDLP.SelectedIndex = 46 Then
            Label4.Text = TMZT.AddHours(8)
        ElseIf TDLP.SelectedIndex = 47 Then
            Label4.Text = TMZT.AddHours(8)
        ElseIf TDLP.SelectedIndex = 48 Then
            Label4.Text = TMZT.AddHours(8)
        ElseIf TDLP.SelectedIndex = 49 Then
            Label4.Text = TMZT.AddHours(8)
        ElseIf TDLP.SelectedIndex = 50 Then
            Label4.Text = TMZT.AddHours(9)
        ElseIf TDLP.SelectedIndex = 51 Then
            Label4.Text = TMZT.AddHours(9)
        ElseIf TDLP.SelectedIndex = 52 Then
            Label4.Text = TMZT.AddHours(9)
        ElseIf TDLP.SelectedIndex = 53 Then
            Label4.Text = TMZT.AddMinutes(570)
        ElseIf TDLP.SelectedIndex = 54 Then
            Label4.Text = TMZT.AddMinutes(570)
        ElseIf TDLP.SelectedIndex = 55 Then
            Label4.Text = TMZT.AddHours(10)
        ElseIf TDLP.SelectedIndex = 56 Then
            Label4.Text = TMZT.AddHours(10)
        ElseIf TDLP.SelectedIndex = 57 Then
            Label4.Text = TMZT.AddHours(10)
        ElseIf TDLP.SelectedIndex = 58 Then
            Label4.Text = TMZT.AddHours(10)
        ElseIf TDLP.SelectedIndex = 59 Then
            Label4.Text = TMZT.AddHours(10)
        ElseIf TDLP.SelectedIndex = 60 Then
            Label4.Text = TMZT.AddHours(11)
        ElseIf TDLP.SelectedIndex = 61 Then
            Label4.Text = TMZT.AddHours(12)
        ElseIf TDLP.SelectedIndex = 62 Then
            Label4.Text = TMZT.AddHours(12)
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        CheckBox1.Checked = True
        System.Diagnostics.Process.Start("C:\Windows\System32\timedate.cpl")
    End Sub


    'Stoper

    Dim s As Integer = 0
    Dim mls As Integer = 0
    Dim m As Integer = 0
    Dim h As Integer = 0
    Dim d As Integer = 0

    Dim cd As Integer

    Dim mlsf As String
    Dim sf As String
    Dim mf As String
    Dim hf As String

    Private Sub SButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SStartB.Click
        If SStartB.Text = "Start" Then
            If CheckBox5.Checked = True Then
                SAddC.Enabled = False
                SAddN.Enabled = False
                SAddB.Enabled = False
                CheckBox5.Enabled = False
                cdN.Enabled = False
                GroupBox2.Enabled = False
                SPauseB.Enabled = False
                SMarkB.Enabled = False
                StopList.Enabled = False
                If CheckBox6.Checked = True Then
                    cs1.Enabled = False
                    cs2.Enabled = False
                    Nidk.Enabled = False
                    Label6.Text = cs1.Text & cdN.Value & cs2.Text
                Else
                    Label6.Text = cdN.Value
                End If
                cd = cdN.Value
                cdT.Enabled = True
                'ControllerS.Enabled = True
                SStartB.Text = "Stop"
            Else
                CheckBox5.Enabled = False
                GroupBox2.Enabled = False
                sT.Enabled = True
                mlsT.Enabled = True
                'ControllerS.Enabled = True
                SStartB.Text = "Stop"
            End If
        Else
            If CheckBox5.Checked = True Then
                SAddC.Enabled = True
                SAddN.Enabled = True
                SAddB.Enabled = True
                CheckBox5.Enabled = True
                cdN.Enabled = True
                GroupBox2.Enabled = True
                SPauseB.Enabled = True
                SMarkB.Enabled = True
                StopList.Enabled = True
                cdT.Enabled = False
                If CheckBox6.Checked = True Then

                    cs1.Enabled = True
                    cs2.Enabled = True
                    Nidk.Enabled = True
                    If CheckBox4.Checked = True Then
                        If CheckBox2.Checked = True Then
                            If CheckBox3.Checked = True Then
                                Label6.Text = d & ":" & hf & ":" & mf & ":" & sf & "." & mlsf
                            Else
                                Label6.Text = hf & ":" & mf & ":" & sf & "." & mlsf
                            End If
                        Else
                            Label6.Text = mf & ":" & sf & "." & mlsf
                        End If
                    Else
                        If CheckBox2.Checked = True Then
                            If CheckBox3.Checked = True Then
                                Label6.Text = d & ":" & hf & ":" & mf & ":" & sf
                            Else
                                Label6.Text = hf & ":" & mf & ":" & sf
                            End If
                        Else
                            Label6.Text = mf & ":" & sf
                        End If
                    End If
                Else
                    If CheckBox4.Checked = True Then
                        If CheckBox2.Checked = True Then
                            If CheckBox3.Checked = True Then
                                Label6.Text = d & ":" & hf & ":" & mf & ":" & sf & "." & mlsf
                            Else
                                Label6.Text = hf & ":" & mf & ":" & sf & "." & mlsf
                            End If
                        Else
                            Label6.Text = mf & ":" & sf & "." & mlsf
                        End If
                    Else
                        If CheckBox2.Checked = True Then
                            If CheckBox3.Checked = True Then
                                Label6.Text = d & ":" & hf & ":" & mf & ":" & sf
                            Else
                                Label6.Text = hf & ":" & mf & ":" & sf
                            End If
                        Else
                            Label6.Text = mf & ":" & sf
                        End If
                    End If
                End If
                'ControllerS.Enabled = True
                SStartB.Text = "Start"
            Else
                CheckBox5.Enabled = True
                sT.Enabled = False
                mlsT.Enabled = False
                SPauseB.Enabled = True
                SMarkB.Enabled = True
                StopList.Enabled = True
                'ControllerS.Enabled = False
                GroupBox2.Enabled = True
                SStartB.Text = "Start"
            End If
        End If
    End Sub

    Private Sub mlsT_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mlsT.Tick
        mls += 1
        If mls = 0 Then
            mlsf = "00"
        ElseIf mls = 1 Then
            mlsf = "01"
        ElseIf mls = 2 Then
            mlsf = "02"
        ElseIf mls = 3 Then
            mlsf = "03"
        ElseIf mls = 4 Then
            mlsf = "04"
        ElseIf mls = 5 Then
            mlsf = "05"
        ElseIf mls = 6 Then
            mlsf = "06"
        ElseIf mls = 7 Then
            mlsf = "07"
        ElseIf mls = 8 Then
            mlsf = "08"
        ElseIf mls = 9 Then
            mlsf = "09"
        Else
            mlsf = mls
        End If
    End Sub

    Private Sub ControllerS_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ControllerS.Tick
        If CheckBox5.Checked = False Then
            If CheckBox4.Checked = True Then
                If CheckBox2.Checked = True Then
                    If CheckBox3.Checked = True Then
                        Label6.Text = d & ":" & hf & ":" & mf & ":" & sf & "." & mlsf
                    Else
                        Label6.Text = hf & ":" & mf & ":" & sf & "." & mlsf
                    End If
                Else
                    Label6.Text = mf & ":" & sf & "." & mlsf
                End If
            Else
                If CheckBox2.Checked = True Then
                    If CheckBox3.Checked = True Then
                        Label6.Text = d & ":" & hf & ":" & mf & ":" & sf
                    Else
                        Label6.Text = hf & ":" & mf & ":" & sf
                    End If
                Else
                    Label6.Text = mf & ":" & sf
                End If
            End If
        End If
        If CheckBox2.Checked = False Then
            If SAddC.SelectedIndex = 2 Then
                SAddB.Enabled = False
            Else
                SAddB.Enabled = True
            End If
        End If
        If CheckBox3.Checked = False Then
            If SAddC.SelectedIndex = 3 Then
                SAddB.Enabled = False
            Else
                SAddB.Enabled = True
            End If
        End If
        If StopList.Items.Count = 0 Then
            Button4.Enabled = False
            Button5.Enabled = False
        Else
            Button4.Enabled = True
            Button5.Enabled = True
        End If
        'Seconds
        If s = 0 Then
            sf = "00"
        ElseIf s = 1 Then
            sf = "01"
        ElseIf s = 2 Then
            sf = "02"
        ElseIf s = 3 Then
            sf = "03"
        ElseIf s = 4 Then
            sf = "04"
        ElseIf s = 5 Then
            sf = "05"
        ElseIf s = 6 Then
            sf = "06"
        ElseIf s = 7 Then
            sf = "07"
        ElseIf s = 8 Then
            sf = "08"
        ElseIf s = 9 Then
            sf = "09"
        Else
            sf = s
        End If
        'Minutes
        If m = 0 Then
            mf = "00"
        ElseIf m = 1 Then
            mf = "01"
        ElseIf m = 2 Then
            mf = "02"
        ElseIf m = 3 Then
            mf = "03"
        ElseIf m = 4 Then
            mf = "04"
        ElseIf m = 5 Then
            mf = "05"
        ElseIf m = 6 Then
            mf = "06"
        ElseIf m = 7 Then
            mf = "07"
        ElseIf m = 8 Then
            mf = "08"
        ElseIf m = 9 Then
            mf = "09"
        Else
            mf = m
        End If
        'Hours
        If h = 0 Then
            hf = "00"
        ElseIf h = 1 Then
            hf = "01"
        ElseIf h = 2 Then
            hf = "02"
        ElseIf h = 3 Then
            hf = "03"
        ElseIf h = 4 Then
            hf = "04"
        ElseIf h = 5 Then
            hf = "05"
        ElseIf h = 6 Then
            hf = "06"
        ElseIf h = 7 Then
            hf = "07"
        ElseIf h = 8 Then
            hf = "08"
        ElseIf h = 9 Then
            hf = "09"
        Else
            hf = h
        End If
    End Sub

    Private Sub sT_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sT.Tick
        If s >= 59 Then
            If m >= 59 Then
                If CheckBox2.Checked = True Then
                    If h >= 23 Then
                        If CheckBox3.Checked = True Then
                            d += 1
                            h = 0
                        Else
                            h += 1
                        End If
                    Else
                        h += 1
                    End If
                    m = 0
                Else
                    m += 1
                End If
            Else
                m += 1
            End If
            s = 0
        Else
            s += 1
        End If
        mls = 0
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            CheckBox3.Enabled = True
            mn.Maximum = 59
            hn.Enabled = True
        Else
            CheckBox3.Enabled = False
            CheckBox3.Checked = False
            hn.Enabled = False
            mn.Maximum = 32767
            dn.Enabled = False
        End If
        If CheckBox4.Checked = True Then
            If CheckBox2.Checked = True Then
                If CheckBox3.Checked = True Then
                    Label6.Text = d & ":" & hf & ":" & mf & ":" & sf & "." & mlsf
                Else
                    Label6.Text = hf & ":" & mf & ":" & sf & "." & mlsf
                End If
            Else
                Label6.Text = mf & ":" & sf & "." & mlsf
            End If
        Else
            If CheckBox2.Checked = True Then
                If CheckBox3.Checked = True Then
                    Label6.Text = d & ":" & hf & ":" & mf & ":" & sf
                Else
                    Label6.Text = hf & ":" & mf & ":" & sf
                End If
            Else
                Label6.Text = mf & ":" & sf
            End If
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            mln.Enabled = True
        Else
            mln.Enabled = False
        End If
        If CheckBox4.Checked = True Then
            If CheckBox2.Checked = True Then
                If CheckBox3.Checked = True Then
                    Label6.Text = d & ":" & hf & ":" & mf & ":" & sf & "." & mlsf
                Else
                    Label6.Text = hf & ":" & mf & ":" & sf & "." & mlsf
                End If
            Else
                Label6.Text = mf & ":" & sf & "." & mlsf
            End If
        Else
            If CheckBox2.Checked = True Then
                If CheckBox3.Checked = True Then
                    Label6.Text = d & ":" & hf & ":" & mf & ":" & sf
                Else
                    Label6.Text = hf & ":" & mf & ":" & sf
                End If
            Else
                Label6.Text = mf & ":" & sf
            End If
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            hn.Maximum = 23
            dn.Enabled = True
        Else
            dn.Enabled = False
            hn.Maximum = 32767
        End If
        If CheckBox4.Checked = True Then
            If CheckBox2.Checked = True Then
                If CheckBox3.Checked = True Then
                    Label6.Text = d & ":" & hf & ":" & mf & ":" & sf & "." & mlsf
                Else
                    Label6.Text = hf & ":" & mf & ":" & sf & "." & mlsf
                End If
            Else
                Label6.Text = mf & ":" & sf & "." & mlsf
            End If
        Else
            If CheckBox2.Checked = True Then
                If CheckBox3.Checked = True Then
                    Label6.Text = d & ":" & hf & ":" & mf & ":" & sf
                Else
                    Label6.Text = hf & ":" & mf & ":" & sf
                End If
            Else
                Label6.Text = mf & ":" & sf
            End If
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        mls = mln.Value
        s = sn.Value
        m = mn.Value
        h = hn.Value
        d = dn.Value
        ToolTip1.Show("Time successfully pasted!", Me, Control.MousePosition.X - Me.Location.X, Control.MousePosition.Y - Me.Location.Y, 4000)
    End Sub

    Private Sub SPauseB_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles SPauseB.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Right Then
            If SStartB.Text = "Stop" Then
                CheckBox5.Enabled = True
                sT.Enabled = False
                mlsT.Enabled = False
                'ControllerS.Enabled = False
                GroupBox2.Enabled = True
                SStartB.Text = "Start"
                mlsf = "00"
                mls = mln.Value
                s = sn.Value
                m = mn.Value
                h = hn.Value
                d = dn.Value
            End If
        ElseIf e.Button = Windows.Forms.MouseButtons.Left Then
            CheckBox5.Enabled = True
            sT.Enabled = False
            mlsT.Enabled = False
            'ControllerS.Enabled = False
            GroupBox2.Enabled = True
            SStartB.Text = "Start"
            mlsf = "00"
            mls = 0
            s = 0
            m = 0
            h = 0
            d = 0
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        mln.Value = 0
        sn.Value = 0
        mn.Value = 0
        hn.Value = 0
        dn.Value = 0
        ToolTip1.Show("Paste boxes successfully reseted!", Me, Control.MousePosition.X - Me.Location.X, Control.MousePosition.Y - Me.Location.Y, 4000)
    End Sub

    Private Sub Dialog3_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LostFocus
        If CheckBox1.Checked = False Then
            Me.Hide()
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            cdN.Enabled = True
            CheckBox6.Enabled = True
        Else
            CheckBox6.Enabled = False
            CheckBox6.Checked = False
            cdN.Enabled = False
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            cs1.Enabled = True
            cs2.Enabled = True
            Nidk.Enabled = True
        Else
            cs1.Enabled = False
            cs2.Enabled = False
            Nidk.Enabled = False
        End If
    End Sub

    Private Sub cdT_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cdT.Tick
        If Not cd <= 1 Then
            cd -= 1
            If CheckBox6.Checked = True Then
                Label6.Text = cs1.Text & cd & cs2.Text
            Else
                Label6.Text = cd
            End If
        Else
            SAddC.Enabled = True
            SAddN.Enabled = True
            SAddB.Enabled = True
            CheckBox5.Checked = False
            StopList.Enabled = True
            SPauseB.Enabled = True
            SMarkB.Enabled = True
            GroupBox2.Enabled = False
            sT.Enabled = True
            mlsT.Enabled = True
            'ControllerS.Enabled = True
            SStartB.Text = "Stop"
            cdT.Enabled = False
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If CheckBox7.Checked = True Then
            CheckBox1.Checked = True
            If MessageBox.Show("Are you sure, do you want reset all the items on Stop List?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                remSD.Clear()
                remSH.Clear()
                remSM.Clear()
                remSS.Clear()
                remSMS.Clear()
                StopList.Items.Clear()
            End If
        Else
            remSD.Clear()
            remSH.Clear()
            remSM.Clear()
            remSS.Clear()
            remSMS.Clear()
            StopList.Items.Clear()
        End If
    End Sub
    Dim remSMS As New Collection
    Dim remSS As New Collection
    Dim remSM As New Collection
    Dim remSH As New Collection
    Dim remSD As New Collection
    Private Sub SMarkB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SMarkB.Click
        If CheckBox4.Checked = False Then
            If CheckBox2.Checked = True Then
                StopList.Items.Add(Label6.Text)
                remSS.Add(sf)
                remSM.Add(mf)
                remSH.Add(hf)
            ElseIf CheckBox2.Checked = True AndAlso CheckBox3.Checked = True Then
                StopList.Items.Add(Label6.Text)
                remSS.Add(sf)
                remSM.Add(mf)
                remSH.Add(hf)
                remSD.Add(d)
            Else
                StopList.Items.Add(Label6.Text)
                remSS.Add(sf)
                remSM.Add(mf)
            End If
        Else
            If CheckBox2.Checked = True Then
                StopList.Items.Add(Label6.Text)
                remSMS.Add(mls)
                remSS.Add(sf)
                remSM.Add(mf)
                remSH.Add(hf)
            ElseIf CheckBox2.Checked = True AndAlso CheckBox3.Checked = True Then
                StopList.Items.Add(Label6.Text)
                remSMS.Add(mls)
                remSS.Add(sf)
                remSM.Add(mf)
                remSH.Add(hf)
                remSD.Add(d)
            Else
                StopList.Items.Add(Label6.Text)
                remSMS.Add(mls)
                remSS.Add(sf)
                remSM.Add(mf)
            End If
        End If
    End Sub

    Private Sub SLcm_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SLcm.Opening
        If StopList.SelectedItem = "" Then
            ricmb.Enabled = False
            cicmb.Enabled = False
            cacmb.Enabled = False
        Else
            ricmb.Enabled = True
            If StopList.SelectedItems.Count > 1 Then
                cicmb.Enabled = False
                cacmb.Enabled = False
            ElseIf StopList.SelectedItems.Count = 1 Then
                cicmb.Enabled = True
                cacmb.Enabled = True
            End If
        End If
    End Sub

    Private Sub SCopyB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SCopyB.Click
        TimeToCopy.Text = Label6.Text
        TimeToCopy.SelectAll()
        TimeToCopy.Copy()
        ToolTip1.Show("Successfully Copied!", Me, Control.MousePosition.X - Me.Location.X, Control.MousePosition.Y - Me.Location.Y, 4000)
    End Sub

    Private Sub ricmb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ricmb.Click
        Dim knt As Integer = StopList.SelectedIndices.Count
        Dim i As Integer
        For i = 0 To knt - 1
            StopList.Items.RemoveAt(StopList.SelectedIndex)
        Next
        'StopList.(StopList.SelectedItems)
    End Sub

    Private Sub StopList_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles StopList.KeyPress
        If Not StopList.SelectedItem = "" Then
            If e.KeyChar = "d"c Then
                StopList.Items.RemoveAt(StopList.SelectedIndex)
                StopList.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        On Error GoTo err
        Dim sfd As New SaveFileDialog
        sfd.Title = "Save File Dialog"
        sfd.FileName = "MyRecords.txt"
        sfd.InitialDirectory = "C:\Users\Kristian\Documents\"
        sfd.Filter = "TXT File (*.txt)|*.txt|All file types|*.*"
        CheckBox1.Checked = True

        StopListToSave.Text = ""
        StopList.SelectedIndex = 0
        StopListToSave.SelectedText = StopList.SelectedItem & Environment.NewLine
repeat:
        If StopList.SelectedIndex = StopList.Items.Count - 1 Then
            StopListToSave.SelectedText = StopList.SelectedItem
            If sfd.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim writer As New System.IO.StreamWriter(sfd.FileName, True, System.Text.Encoding.Default)
                writer.Write(StopListToSave.Text)
                writer.Close()
                ToolTip1.Show("Successfully Saved to: " & sfd.FileName & "!", Me, Control.MousePosition.X - Me.Location.X, Control.MousePosition.Y - Me.Location.Y, 4000)
            End If
        Else
            StopList.SelectedIndex += 1
            StopListToSave.SelectedText = StopList.SelectedItem & Environment.NewLine
            GoTo repeat
        End If
        If Me.Tag = "rgjfkjztrfzt" Then
err:
            MessageBox.Show("Error while converting list to TXT, because you maybe had that list empty or it contains some characters that he can't understand." & Environment.NewLine & "Please check the list again, if is something wrong. Thanks!", "Error while Converting", MessageBoxButtons.OK)
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SAddB.Click
        If SAddC.SelectedIndex = 0 Then
            s += SAddN.Value
        ElseIf SAddC.SelectedIndex = 1 Then
            m += SAddN.Value
        ElseIf SAddC.SelectedIndex = 2 Then
            h += SAddN.Value
        ElseIf SAddC.SelectedIndex = 3 Then
            d += SAddN.Value
        End If
    End Sub

    Private Sub ClearToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearToolStripMenuItem.Click
        If CheckBox7.Checked = True Then
            CheckBox1.Checked = True
            If MessageBox.Show("Are you sure, do you want reset all the items on Stop List?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                StopList.Items.Clear()
            End If
        Else
            StopList.Items.Clear()
        End If
    End Sub

    Private Sub cacmb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cacmb.Click
        If CheckBox4.Checked = False Then
            If CheckBox2.Checked = True Then
                If CheckBox3.Checked = True Then
                    s = remSS.Item(StopList.SelectedIndex + 1)
                    m = remSM.Item(StopList.SelectedIndex + 1)
                    h = remSH.Item(StopList.SelectedIndex + 1)
                    d = remSD.Item(StopList.SelectedIndex + 1)
                Else
                    s = remSS.Item(StopList.SelectedIndex + 1)
                    m = remSM.Item(StopList.SelectedIndex + 1)
                    h = remSH.Item(StopList.SelectedIndex + 1)
                End If
            Else
                s = remSS.Item(StopList.SelectedIndex + 1)
                m = remSM.Item(StopList.SelectedIndex + 1)
            End If
        Else
            If CheckBox2.Checked = True Then
                If CheckBox3.Checked = True Then
                    mls = remSMS.Item(StopList.SelectedIndex + 1)
                    s = remSS.Item(StopList.SelectedIndex + 1)
                    m = remSM.Item(StopList.SelectedIndex + 1)
                    h = remSH.Item(StopList.SelectedIndex + 1)
                    d = remSD.Item(StopList.SelectedIndex + 1)
                Else
                    mls = remSMS.Item(StopList.SelectedIndex + 1)
                    s = remSS.Item(StopList.SelectedIndex + 1)
                    m = remSM.Item(StopList.SelectedIndex + 1)
                    h = remSH.Item(StopList.SelectedIndex + 1)
                End If
            Else
                mls = remSMS.Item(StopList.SelectedIndex + 1)
                s = remSS.Item(StopList.SelectedIndex + 1)
                m = remSM.Item(StopList.SelectedIndex + 1)
            End If
        End If
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cicmb.Click
        TimeToCopy.Text = StopList.SelectedItem
        TimeToCopy.SelectAll()
        TimeToCopy.Copy()
        ToolTip1.Show("Successfully copied!", Me, Control.MousePosition.X - Me.Location.X, Control.MousePosition.Y - Me.Location.Y, 4000)
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TimeToCopy.Text = Label13.Text
        TimeToCopy.SelectAll()
        TimeToCopy.Copy()
        ToolTip1.Show("Successfully copied!", Me, Control.MousePosition.X - Me.Location.X, Control.MousePosition.Y - Me.Location.Y, 4000)
    End Sub

    'TIMERRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRR
    Public Ts As Integer
    Public Tm As Integer
    Public Th As Integer
    Public Td As Integer

    Dim Tsf As String
    Dim Tmf As String
    Dim Thf As String
    Dim Tdf As String
    Private Sub ControllerTT_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ControllerTT.Tick
        If CheckBox13.Checked = True Then
            If CheckBox12.Checked = True Then
                Label13.Text = Td & ":" & Thf & ":" & Tmf & ":" & Tsf
            Else
                Label13.Text = Thf & ":" & Tmf & ":" & Tsf
            End If
        Else
            Label13.Text = Tmf & ":" & Tsf
        End If

        If CheckBox13.Checked = False Then
            If TAddC.SelectedIndex = 2 Then
                TAddB.Enabled = False
            Else
                TAddB.Enabled = True
            End If
        End If
        If CheckBox12.Checked = False Then
            If TAddC.SelectedIndex = 3 Then
                TAddB.Enabled = False
            Else
                TAddB.Enabled = True
            End If
        End If
        If MarkList.Items.Count = 0 Then
            GroupBox6.Enabled = False
        Else
            GroupBox6.Enabled = True
        End If
        'Seconds
        If Ts = 0 Then
            Tsf = "00"
        ElseIf Ts = 1 Then
            Tsf = "01"
        ElseIf Ts = 2 Then
            Tsf = "02"
        ElseIf Ts = 3 Then
            Tsf = "03"
        ElseIf Ts = 4 Then
            Tsf = "04"
        ElseIf Ts = 5 Then
            Tsf = "05"
        ElseIf Ts = 6 Then
            Tsf = "06"
        ElseIf Ts = 7 Then
            Tsf = "07"
        ElseIf Ts = 8 Then
            Tsf = "08"
        ElseIf Ts = 9 Then
            Tsf = "09"
        Else
            Tsf = Ts
        End If
        'Minutes
        If Tm = 0 Then
            Tmf = "00"
        ElseIf Tm = 1 Then
            Tmf = "01"
        ElseIf Tm = 2 Then
            Tmf = "02"
        ElseIf Tm = 3 Then
            Tmf = "03"
        ElseIf Tm = 4 Then
            Tmf = "04"
        ElseIf Tm = 5 Then
            Tmf = "05"
        ElseIf Tm = 6 Then
            Tmf = "06"
        ElseIf Tm = 7 Then
            Tmf = "07"
        ElseIf Tm = 8 Then
            Tmf = "08"
        ElseIf Tm = 9 Then
            Tmf = "09"
        Else
            Tmf = Tm
        End If
        'Hours
        If Th = 0 Then
            Thf = "00"
        ElseIf Th = 1 Then
            Thf = "01"
        ElseIf Th = 2 Then
            Thf = "02"
        ElseIf Th = 3 Then
            Thf = "03"
        ElseIf Th = 4 Then
            Thf = "04"
        ElseIf Th = 5 Then
            Thf = "05"
        ElseIf Th = 6 Then
            Thf = "06"
        ElseIf Th = 7 Then
            Thf = "07"
        ElseIf Th = 8 Then
            Thf = "08"
        ElseIf Th = 9 Then
            Thf = "09"
        Else
            Thf = Th
        End If
        If MarkList.Items.Count = 0 Then
            GroupBox6.Enabled = False
        Else
            GroupBox6.Enabled = True
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        MarkList.Items.Clear()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        If CheckBox13.Checked = True Then
            If CheckBox12.Checked = True Then
                Ts = tsn.Value
                Tm = tmn.Value
                Th = thn.Value
                Td = tdn.Value
            Else
                Ts = tsn.Value
                Tm = tmn.Value
                Th = thn.Value
            End If
        Else
            Ts = tsn.Value
            Tm = tmn.Value
        End If
    End Sub

    Private Sub BoomAction()
        If CheckBox16.Checked = False Then
            TremHM = ThmN.Value
        End If
        If DomainUpDown1.SelectedIndex = 0 Then
            If ComboBox1.SelectedIndex = 0 Then
                If CheckBox16.Checked = True Then
                    Media.SystemSounds.Asterisk.Play()
                Else
repeat11:
                    If TremHM < 1 Then
                        TPauseB.Enabled = False
                        TPauseB.Text = "Pause"
                        TsT.Enabled = False
                        GroupBox4.Enabled = True
                        If CheckBox15.Checked = True Then
                            GroupBox7.Enabled = True
                        End If
                        CheckBox15.Enabled = True
                        TStartB.Text = "Start"
                        alarm.Stop()
                    Else
                        Media.SystemSounds.Asterisk.Play()
                        TremHM -= 1
                        GoTo repeat11
                    End If
                End If
            ElseIf ComboBox1.SelectedIndex = 1 Then
                If CheckBox16.Checked = True Then
                    Media.SystemSounds.Beep.Play()
                Else
repeat12:
                    If TremHM < 1 Then
                        TPauseB.Enabled = False
                        TPauseB.Text = "Pause"
                        TsT.Enabled = False
                        GroupBox4.Enabled = True
                        If CheckBox15.Checked = True Then
                            GroupBox7.Enabled = True
                        End If
                        CheckBox15.Enabled = True
                        TStartB.Text = "Start"
                        alarm.Stop()
                    Else
                        Media.SystemSounds.Beep.Play()
                        TremHM -= 1
                        GoTo repeat12
                    End If
                End If
            ElseIf ComboBox1.SelectedIndex = 2 Then
                If CheckBox16.Checked = True Then
                    Media.SystemSounds.Exclamation.Play()
                Else
repeat13:
                    If TremHM < 1 Then
                        TPauseB.Enabled = False
                        TPauseB.Text = "Pause"
                        TsT.Enabled = False
                        GroupBox4.Enabled = True
                        If CheckBox15.Checked = True Then
                            GroupBox7.Enabled = True
                        End If
                        CheckBox15.Enabled = True
                        TStartB.Text = "Start"
                        alarm.Stop()
                    Else
                        Media.SystemSounds.Exclamation.Play()
                        TremHM -= 1
                        GoTo repeat13
                    End If
                End If
            ElseIf ComboBox1.SelectedIndex = 3 Then
                If CheckBox16.Checked = True Then
                    Media.SystemSounds.Hand.Play()
                Else
repeat14:
                    If TremHM < 1 Then
                        TPauseB.Enabled = False
                        TPauseB.Text = "Pause"
                        TsT.Enabled = False
                        GroupBox4.Enabled = True
                        If CheckBox15.Checked = True Then
                            GroupBox7.Enabled = True
                        End If
                        CheckBox15.Enabled = True
                        TStartB.Text = "Start"
                        alarm.Stop()
                    Else
                        Media.SystemSounds.Hand.Play()
                        TremHM -= 1
                        GoTo repeat14
                    End If
                End If
            ElseIf ComboBox1.SelectedIndex = 4 Then
                If CheckBox16.Checked = True Then
                    Media.SystemSounds.Question.Play()
                Else
repeat15:
                    If TremHM < 1 Then
                        TPauseB.Enabled = False
                        TPauseB.Text = "Pause"
                        TsT.Enabled = False
                        GroupBox4.Enabled = True
                        If CheckBox15.Checked = True Then
                            GroupBox7.Enabled = True
                        End If
                        CheckBox15.Enabled = True
                        TStartB.Text = "Start"
                        alarm.Stop()
                    Else
                        Media.SystemSounds.Question.Play()
                        TremHM -= 1
                        GoTo repeat15
                    End If
                End If
            End If
        End If
        If DomainUpDown1.SelectedIndex = 1 Then
            alpath = "C:\Windows\Media\Alarms\" & ComboBox3.SelectedItem
            alarm = New Media.SoundPlayer(alpath)
            If CheckBox16.Checked = True Then
                alarm.PlayLooping()
            Else
repeat21:
                If TremHM < 1 Then
                    TPauseB.Enabled = False
                    TPauseB.Text = "Pause"
                    TsT.Enabled = False
                    GroupBox4.Enabled = True
                    If CheckBox15.Checked = True Then
                        GroupBox7.Enabled = True
                    End If
                    CheckBox15.Enabled = True
                    TStartB.Text = "Start"
                    alarm.Stop()
                Else
                    alarm.PlaySync()
                    TremHM -= 1
                    GoTo repeat21
                End If
            End If
        End If
        If DomainUpDown1.SelectedIndex = 2 Then
            alpath = ComboBox4.Text
            If My.Computer.FileSystem.FileExists(alpath) = True Then
                alarm = New Media.SoundPlayer(alpath)
                If CheckBox16.Checked = True Then
                    alarm.PlayLooping()
                Else
repeat31:
                    If TremHM < 1 Then
                        TPauseB.Enabled = False
                        TPauseB.Text = "Pause"
                        TsT.Enabled = False
                        GroupBox4.Enabled = True
                        If CheckBox15.Checked = True Then
                            GroupBox7.Enabled = True
                        End If
                        CheckBox15.Enabled = True
                        TStartB.Text = "Start"
                        alarm.Stop()
                    Else
                        alarm.PlaySync()
                        TremHM -= 1
                        GoTo repeat31
                    End If
                End If
            Else
                TPauseB.Enabled = False
                TPauseB.Text = "Pause"
                TsT.Enabled = False
                GroupBox4.Enabled = True
                If CheckBox15.Checked = True Then
                    GroupBox7.Enabled = True
                End If
                CheckBox15.Enabled = True
                TStartB.Text = "Start"
                alarm.Stop()
            End If
        End If
    End Sub

    Private Sub BoomACT()
        If CheckBox14.Checked = True Then
            If ComboBox2.SelectedIndex = 0 Then
                If TActL.Text = "Info" Then
                    MessageBox.Show("Time's up!", "Microsoft Windows", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                ElseIf TActL.Text = "Question" Then
                    MessageBox.Show("Time's up!", "Microsoft Windows", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
                ElseIf TActL.Text = "Warning" Then
                    MessageBox.Show("Time's up!", "Microsoft Windows", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                ElseIf TActL.Text = "Error" Then
                    MessageBox.Show("Time's up!", "Microsoft Windows", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                End If
            ElseIf ComboBox2.SelectedIndex = 1 Then
                If TActL.Text = "RightDown" Then
                    NW.MODE = 2
                    NW.TL.Text = "TIMER REACH THE END!"
                    NW.IP.Text = "Click the ''Stop'' button to eliminate those attentions and stop the timer."
                    NW.B2.Text = "Stop"
                    NW.B3.Text = "Show Timer"
                    NW.Location = New Point(My.Computer.Screen.WorkingArea.Width - NW.Width, My.Computer.Screen.WorkingArea.Height - NW.Height)
                    Select Case NW.ShowDialog
                        Case Windows.Forms.DialogResult.Yes
                            TPauseB.Enabled = False
                            TPauseB.Text = "Pause"
                            TsT.Enabled = False
                            Td = 0
                            Th = 0
                            Tm = 0
                            Ts = 0
                            GroupBox4.Enabled = True
                            If CheckBox15.Checked = True Then
                                GroupBox7.Enabled = True
                            End If
                            CheckBox15.Enabled = True
                            alarm.Stop()
                            TStartB.Text = "Start"
                        Case Windows.Forms.DialogResult.No
                            Me.Show()
                            CheckBox1.Checked = True
                            TabControl1.SelectedIndex = 2
                    End Select
                ElseIf TActL.Text = "Right" Then
                    NW.MODE = 2
                    NW.TL.Text = "TIMER REACH THE END!"
                    NW.IP.Text = "Click the ''Stop'' button to eliminate those attentions and stop the timer."
                    NW.B2.Text = "Stop"
                    NW.B3.Text = "Show Timer"
                    NW.Location = New Point(My.Computer.Screen.WorkingArea.Width - NW.Width, My.Computer.Screen.WorkingArea.Height / 2 - NW.Height / 2)
                    Select Case NW.ShowDialog
                        Case Windows.Forms.DialogResult.Yes
                            TPauseB.Enabled = False
                            TPauseB.Text = "Pause"
                            TsT.Enabled = False
                            Td = 0
                            Th = 0
                            Tm = 0
                            Ts = 0
                            GroupBox4.Enabled = True
                            If CheckBox15.Checked = True Then
                                GroupBox7.Enabled = True
                            End If
                            CheckBox15.Enabled = True
                            alarm.Stop()
                            TStartB.Text = "Start"
                        Case Windows.Forms.DialogResult.No
                            Me.Show()
                            CheckBox1.Checked = True
                            TabControl1.SelectedIndex = 2
                    End Select
                ElseIf TActL.Text = "RightUp" Then
                    NW.MODE = 2
                    NW.TL.Text = "TIMER REACH THE END!"
                    NW.IP.Text = "Click the ''Stop'' button to eliminate those attentions and stop the timer."
                    NW.B2.Text = "Stop"
                    NW.B3.Text = "Show Timer"
                    NW.Location = New Point(My.Computer.Screen.WorkingArea.Width - NW.Width, 0)
                    Select Case NW.ShowDialog
                        Case Windows.Forms.DialogResult.Yes
                            TPauseB.Enabled = False
                            TPauseB.Text = "Pause"
                            TsT.Enabled = False
                            Td = 0
                            Th = 0
                            Tm = 0
                            Ts = 0
                            GroupBox4.Enabled = True
                            If CheckBox15.Checked = True Then
                                GroupBox7.Enabled = True
                            End If
                            CheckBox15.Enabled = True
                            alarm.Stop()
                            TStartB.Text = "Start"
                        Case Windows.Forms.DialogResult.No
                            Me.Show()
                            CheckBox1.Checked = True
                            TabControl1.SelectedIndex = 2
                    End Select
                ElseIf TActL.Text = "Up" Then
                    NW.MODE = 2
                    NW.TL.Text = "TIMER REACH THE END!"
                    NW.IP.Text = "Click the ''Stop'' button to eliminate those attentions and stop the timer."
                    NW.B2.Text = "Stop"
                    NW.B3.Text = "Show Timer"
                    NW.Location = New Point(My.Computer.Screen.WorkingArea.Width / 2 - NW.Width / 2, 0)
                    Select Case NW.ShowDialog
                        Case Windows.Forms.DialogResult.Yes
                            TPauseB.Enabled = False
                            TPauseB.Text = "Pause"
                            TsT.Enabled = False
                            Td = 0
                            Th = 0
                            Tm = 0
                            Ts = 0
                            GroupBox4.Enabled = True
                            If CheckBox15.Checked = True Then
                                GroupBox7.Enabled = True
                            End If
                            CheckBox15.Enabled = True
                            alarm.Stop()
                            TStartB.Text = "Start"
                        Case Windows.Forms.DialogResult.No
                            Me.Show()
                            CheckBox1.Checked = True
                            TabControl1.SelectedIndex = 2
                    End Select
                ElseIf TActL.Text = "LeftUp" Then
                    NW.MODE = 2
                    NW.TL.Text = "TIMER REACH THE END!"
                    NW.IP.Text = "Click the ''Stop'' button to eliminate those attentions and stop the timer."
                    NW.B2.Text = "Stop"
                    NW.B3.Text = "Show Timer"
                    NW.Location = New Point(0, 0)
                    Select Case NW.ShowDialog
                        Case Windows.Forms.DialogResult.Yes
                            TPauseB.Enabled = False
                            TPauseB.Text = "Pause"
                            TsT.Enabled = False
                            Td = 0
                            Th = 0
                            Tm = 0
                            Ts = 0
                            GroupBox4.Enabled = True
                            If CheckBox15.Checked = True Then
                                GroupBox7.Enabled = True
                            End If
                            CheckBox15.Enabled = True
                            alarm.Stop()
                            TStartB.Text = "Start"
                        Case Windows.Forms.DialogResult.No
                            Me.Show()
                            CheckBox1.Checked = True
                            TabControl1.SelectedIndex = 2
                    End Select
                ElseIf TActL.Text = "Left" Then
                    NW.MODE = 2
                    NW.TL.Text = "TIMER REACH THE END!"
                    NW.IP.Text = "Click the ''Stop'' button to eliminate those attentions and stop the timer."
                    NW.B2.Text = "Stop"
                    NW.B3.Text = "Show Timer"
                    NW.Location = New Point(0, My.Computer.Screen.WorkingArea.Height / 2 - NW.Height / 2)
                    Select Case NW.ShowDialog
                        Case Windows.Forms.DialogResult.Yes
                            TPauseB.Enabled = False
                            TPauseB.Text = "Pause"
                            TsT.Enabled = False
                            Td = 0
                            Th = 0
                            Tm = 0
                            Ts = 0
                            GroupBox4.Enabled = True
                            If CheckBox15.Checked = True Then
                                GroupBox7.Enabled = True
                            End If
                            CheckBox15.Enabled = True
                            alarm.Stop()
                            TStartB.Text = "Start"
                        Case Windows.Forms.DialogResult.No
                            Me.Show()
                            CheckBox1.Checked = True
                            TabControl1.SelectedIndex = 2
                    End Select
                ElseIf TActL.Text = "LeftDown" Then
                    NW.MODE = 2
                    NW.TL.Text = "TIMER REACH THE END!"
                    NW.IP.Text = "Click the ''Stop'' button to eliminate those attentions and stop the timer."
                    NW.B2.Text = "Stop"
                    NW.B3.Text = "Show Timer"
                    NW.Location = New Point(0, My.Computer.Screen.WorkingArea.Height - NW.Height)
                    Select Case NW.ShowDialog
                        Case Windows.Forms.DialogResult.Yes
                            TPauseB.Enabled = False
                            TPauseB.Text = "Pause"
                            TsT.Enabled = False
                            Td = 0
                            Th = 0
                            Tm = 0
                            Ts = 0
                            GroupBox4.Enabled = True
                            If CheckBox15.Checked = True Then
                                GroupBox7.Enabled = True
                            End If
                            CheckBox15.Enabled = True
                            alarm.Stop()
                            TStartB.Text = "Start"
                        Case Windows.Forms.DialogResult.No
                            Me.Show()
                            CheckBox1.Checked = True
                            TabControl1.SelectedIndex = 2
                    End Select
                ElseIf TActL.Text = "Down" Then
                    NW.MODE = 2
                    NW.TL.Text = "TIMER REACH THE END!"
                    NW.IP.Text = "Click the ''Stop'' button to eliminate those attentions and stop the timer."
                    NW.B2.Text = "Stop"
                    NW.B3.Text = "Show Timer"
                    NW.Location = New Point(My.Computer.Screen.WorkingArea.Width / 2 - NW.Width / 2, My.Computer.Screen.WorkingArea.Height - NW.Height)
                    Select Case NW.ShowDialog
                        Case Windows.Forms.DialogResult.Yes
                            TPauseB.Enabled = False
                            TPauseB.Text = "Pause"
                            TsT.Enabled = False
                            Td = 0
                            Th = 0
                            Tm = 0
                            Ts = 0
                            GroupBox4.Enabled = True
                            If CheckBox15.Checked = True Then
                                GroupBox7.Enabled = True
                            End If
                            CheckBox15.Enabled = True
                            alarm.Stop()
                            TStartB.Text = "Start"
                        Case Windows.Forms.DialogResult.No
                            Me.Show()
                            CheckBox1.Checked = True
                            TabControl1.SelectedIndex = 2
                    End Select
                ElseIf TActL.Text = "Center" Then
                    NW.MODE = 2
                    NW.TL.Text = "TIMER REACH THE END!"
                    NW.IP.Text = "Click the ''Stop'' button to eliminate those attentions and stop the timer."
                    NW.B2.Text = "Stop"
                    NW.B3.Text = "Show Timer"
                    NW.Location = New Point(My.Computer.Screen.WorkingArea.Width / 2 - NW.Width / 2, My.Computer.Screen.WorkingArea.Height / 2 - NW.Height / 2)
                    Select Case NW.ShowDialog
                        Case Windows.Forms.DialogResult.Yes
                            TPauseB.Enabled = False
                            TPauseB.Text = "Pause"
                            TsT.Enabled = False
                            Td = 0
                            Th = 0
                            Tm = 0
                            Ts = 0
                            GroupBox4.Enabled = True
                            If CheckBox15.Checked = True Then
                                GroupBox7.Enabled = True
                            End If
                            CheckBox15.Enabled = True
                            alarm.Stop()
                            TStartB.Text = "Start"
                        Case Windows.Forms.DialogResult.No
                            Me.Show()
                            CheckBox1.Checked = True
                            TabControl1.SelectedIndex = 2
                    End Select
                End If
            ElseIf ComboBox2.SelectedIndex = 2 Then
                ScreenLocker.BackColor = Color.LightBlue
                ScreenLocker.BackgroundImage = Button1.BackgroundImage
                ScreenLocker.ShowDialog()
            ElseIf ComboBox2.SelectedIndex = 3 Then
                ScreenLocker.BackColor = TActL.BackColor
                ScreenLocker.BackgroundImage = Button1.BackgroundImage
                If tacthccmb.Checked = True Then
                    Cursor.Hide()
                End If
                ScreenLocker.ShowDialog()
            ElseIf ComboBox2.SelectedIndex = 4 Then
                On Error Resume Next
                ScreenLocker.BackColor = Color.LightBlue
                If My.Computer.FileSystem.FileExists(TActL.Text) = True Then
                    ScreenLocker.BackgroundImage = New Bitmap(TActL.Text)
                    If tacthccmb.Checked = True Then
                        Cursor.Hide()
                    End If
                    ScreenLocker.ShowDialog()
                End If
            ElseIf ComboBox2.SelectedIndex = 5 Then
                MouseLocker.Enabled = True
                MLf.Show()
            ElseIf ComboBox2.SelectedIndex = 6 Then
                On Error Resume Next
                If My.Computer.FileSystem.FileExists(TActL.Text) = True Then
                    System.Diagnostics.Process.Start(TActL.Text)
                End If
            ElseIf ComboBox2.SelectedIndex = 7 Then
                On Error Resume Next
                System.Diagnostics.Process.Start("C:\Windows\shutdown_and_logoff_scripts\" & TActL.Text & ".lnk", AppWinStyle.Hide)
            ElseIf ComboBox2.SelectedIndex = 8 Then
                On Error Resume Next
                System.Diagnostics.Process.Start("C:\Windows\shutdown_and_logoff_scripts\logon\" & TActL.Text & ".lnk", AppWinStyle.Hide)
            End If
        End If
    End Sub
    Dim alpath As String
    Public alarm As Media.SoundPlayer
    Dim TremHM As Integer
    Private Sub TsT_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TsT.Tick
        On Error Resume Next
        Ts -= 1
        If Ts <= -1 Then
            Tm -= 1
            Ts = 59
            If Tm <= -1 Then
                Th -= 1
                Tm = 59
                If CheckBox13.Checked = True Then
                    If Th <= -1 Then
                        Td -= 1
                        Th = 23
                        If CheckBox12.Checked = True Then
                            If Td <= -1 Then
                                TsT.Enabled = False
                                Td = 0
                                Th = 0
                                Tm = 0
                                Ts = 0
                                If CheckBox15.Checked = True Then
                                    BoomAction()
                                    If CheckBox16.Checked = True Then
                                        BoomACT()
                                    End If
                                Else
                                    If CheckBox16.Checked = True Then
                                        BoomACT()
                                    End If
                                    TPauseB.Enabled = False
                                    TPauseB.Text = "Pause"
                                    TsT.Enabled = False
                                    GroupBox4.Enabled = True
                                    If CheckBox15.Checked = True Then
                                        GroupBox7.Enabled = True
                                    End If
                                    CheckBox15.Enabled = True
                                    TStartB.Text = "Start"
                                    alarm.Stop()
                                End If
                            End If
                        Else
                            TsT.Enabled = False
                            Td = 0
                            Th = 0
                            Tm = 0
                            Ts = 0
                            If CheckBox15.Checked = True Then
                                BoomAction()
                                If CheckBox16.Checked = True Then
                                    BoomACT()
                                End If
                            Else
                                If CheckBox16.Checked = True Then
                                    BoomACT()
                                End If
                                TPauseB.Enabled = False
                                TPauseB.Text = "Pause"
                                TsT.Enabled = False
                                GroupBox4.Enabled = True
                                If CheckBox15.Checked = True Then
                                    GroupBox7.Enabled = True
                                End If
                                CheckBox15.Enabled = True
                                TStartB.Text = "Start"
                                alarm.Stop()
                            End If
                        End If
                    End If
                Else
                    TsT.Enabled = False
                    Td = 0
                    Th = 0
                    Tm = 0
                    Ts = 0
                    If CheckBox15.Checked = True Then
                        BoomAction()
                        If CheckBox16.Checked = True Then
                            BoomACT()
                        End If
                    Else
                        If CheckBox16.Checked = True Then
                            BoomACT()
                        End If
                        TPauseB.Enabled = False
                        TPauseB.Text = "Pause"
                        TsT.Enabled = False
                        GroupBox4.Enabled = True
                        If CheckBox15.Checked = True Then
                            GroupBox7.Enabled = True
                        End If
                        CheckBox15.Enabled = True
                        TStartB.Text = "Start"
                        alarm.Stop()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TStartB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TStartB.Click
        On Error Resume Next
        If TStartB.Text = "Start" Then
            GroupBox4.Enabled = False
            GroupBox7.Enabled = False
            CheckBox15.Enabled = False
            TsT.Enabled = True
            TPauseB.Enabled = True
            TStartB.Text = "Stop"
        Else
            TPauseB.Enabled = False
            TPauseB.Text = "Pause"
            TsT.Enabled = False
            Td = 0
            Th = 0
            Tm = 0
            Ts = 0
            GroupBox4.Enabled = True
            If CheckBox15.Checked = True Then
                GroupBox7.Enabled = True
            End If
            CheckBox15.Enabled = True
            alarm.Stop()
            TStartB.Text = "Start"
        End If
    End Sub

    Private Sub CheckBox13_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox13.CheckedChanged
        If CheckBox13.Checked = True Then
            CheckBox12.Enabled = True
            tmn.Maximum = 59
            thn.Enabled = True
        Else
            CheckBox12.Enabled = False
            CheckBox12.Checked = False
            thn.Enabled = False
            tmn.Maximum = 32767
            tdn.Enabled = False
        End If
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        tsn.Value = 0
        tmn.Value = 0
        thn.Value = 0
        tdn.Value = 0
        ToolTip1.Show("Paste boxes successfully reseted!", Me, Control.MousePosition.X - Me.Location.X, Control.MousePosition.Y - Me.Location.Y, 4000)
    End Sub

    Private Sub TAddB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TAddB.Click
        If TAddC.SelectedIndex = 0 Then
            Ts += TAddN.Value
        ElseIf TAddC.SelectedIndex = 1 Then
            Tm += TAddN.Value
        ElseIf TAddC.SelectedIndex = 2 Then
            Th += TAddN.Value
        ElseIf TAddC.SelectedIndex = 3 Then
            Td += TAddN.Value
        End If
    End Sub

    Private Sub CheckBox12_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox12.CheckedChanged
        If CheckBox12.Checked = True Then
            thn.Maximum = 23
            tdn.Enabled = True
        Else
            tdn.Enabled = False
            thn.Maximum = 32767
        End If
    End Sub

    Private Sub CheckBox15_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox15.CheckedChanged
        If CheckBox15.Checked = True Then
            GroupBox7.Enabled = True
        Else
            GroupBox7.Enabled = False
        End If
    End Sub
    Dim remTMS As New Collection
    Dim remTS As New Collection
    Dim remTM As New Collection
    Dim remTH As New Collection
    Dim remTD As New Collection
    Private Sub TMarkB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TMarkB.Click
        If CheckBox13.Checked = True Then
            If CheckBox12.Checked = True Then
                MarkList.Items.Add(Label13.Text)
                remTS.Add(Tsf)
                remTM.Add(Tmf)
                remTH.Add(Thf)
                remTD.Add(Td)
            Else
                MarkList.Items.Add(Label13.Text)
                remTS.Add(Tsf)
                remTM.Add(Tmf)
                remTH.Add(Thf)
            End If
        Else
            MarkList.Items.Add(Label13.Text)
            remTS.Add(Tsf)
            remTM.Add(Tmf)
        End If
    End Sub

    Private Sub DomainUpDown1_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DomainUpDown1.SelectedItemChanged
        If DomainUpDown1.SelectedIndex = 0 Then
            ComboBox3.Visible = False
            ComboBox4.Visible = False
            Button16.Visible = False
            ComboBox1.Visible = True
        ElseIf DomainUpDown1.SelectedIndex = 1 Then
            ComboBox1.Visible = False
            ComboBox4.Visible = False
            Button16.Visible = False
            ComboBox3.Visible = True
        ElseIf DomainUpDown1.SelectedIndex = 2 Then
            ComboBox1.Visible = False
            ComboBox3.Visible = False
            Button16.Visible = True
            ComboBox4.Visible = True
        End If
    End Sub

    Private Sub CheckBox16_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox16.CheckedChanged
        If CheckBox16.Checked = True Then
            Label27.Enabled = False
            ThmN.Enabled = False
        Else
            Label27.Enabled = True
            ThmN.Enabled = True
        End If
    End Sub
    Dim remSEMI3 As Integer = 0
    Private Sub ComboBox3_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.DropDown
        On Error Resume Next
        ComboBox3.Items.Clear()
        Dim i As String
        Dim CTSPath As String = "C:\Windows\Media\Alarms"
        If My.Computer.FileSystem.DirectoryExists(CTSPath) Then
            For Each i In My.Computer.FileSystem.GetFiles(CTSPath)
                ComboBox3.Items.Add(i.Substring(i.LastIndexOf("\") + 1))
            Next
        End If
        ComboBox3.SelectedIndex = remSEMI3
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Dim ofd As New OpenFileDialog
        ofd.CheckFileExists = True
        ofd.CheckPathExists = True
        ofd.ShowReadOnly = True
        ofd.Title = "Open File Dialog for Alarm choosing"
        ofd.Filter = "Supported media (*.wav)|*.wav"
        ofd.CustomPlaces.Clear()
        ofd.CustomPlaces.Add("C:\Windows\Media\Alarms\")
        ofd.InitialDirectory = "C:\Windows\Media\Alarms\"
        CheckBox1.Checked = True
        If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
            ComboBox4.Text = ofd.FileName
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        remSEMI3 = ComboBox3.SelectedIndex
    End Sub

    Private Sub TPauseB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TPauseB.Click
        If TPauseB.Text = "Pause" Then
            TsT.Enabled = False
            TPauseB.Text = "UnPause"
        Else
            TsT.Enabled = True
            TPauseB.Text = "Pause"
        End If
    End Sub

    Private Sub MLcm_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MLcm.Opening
        If MarkList.SelectedItem = "" Then
            tricmb.Enabled = False
            tcicmb.Enabled = False
            tcacmb.Enabled = False
        Else
            tricmb.Enabled = True
            If MarkList.SelectedItems.Count > 1 Then
                tcicmb.Enabled = False
                tcacmb.Enabled = False
            ElseIf MarkList.SelectedItems.Count = 1 Then
                tcicmb.Enabled = True
                tcacmb.Enabled = True
            End If
        End If
    End Sub

    Private Sub ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click
        MarkList.Items.Clear()
    End Sub

    Private Sub tricmb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tricmb.Click
        Dim knt As Integer = MarkList.SelectedIndices.Count
        Dim i As Integer
        For i = 0 To knt - 1
            MarkList.Items.RemoveAt(MarkList.SelectedIndex)
        Next
    End Sub

    Private Sub tcacmb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tcacmb.Click
        If CheckBox13.Checked = True Then
            If CheckBox12.Checked = True Then
                Ts = remTS.Item(MarkList.SelectedIndex + 1)
                Tm = remTM.Item(MarkList.SelectedIndex + 1)
                Th = remTH.Item(MarkList.SelectedIndex + 1)
                Td = remTD.Item(MarkList.SelectedIndex + 1)
            Else
                Ts = remTS.Item(MarkList.SelectedIndex + 1)
                Tm = remTM.Item(MarkList.SelectedIndex + 1)
                Th = remTH.Item(MarkList.SelectedIndex + 1)
            End If
        Else
            Ts = remTS.Item(MarkList.SelectedIndex + 1)
            Tm = remTM.Item(MarkList.SelectedIndex + 1)
        End If
    End Sub

    Private Sub tcicmb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tcicmb.Click
        TimeToCopy.Text = MarkList.SelectedItem
        TimeToCopy.SelectAll()
        TimeToCopy.Copy()
        ToolTip1.Show("Successfully copied!", Me, Control.MousePosition.X - Me.Location.X, Control.MousePosition.Y - Me.Location.Y, 4000)
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        On Error GoTo err
        Dim sfd As New SaveFileDialog
        sfd.Title = "Save File Dialog"
        sfd.FileName = "MyRecords.txt"
        sfd.InitialDirectory = "C:\Users\Kristian\Documents\"
        sfd.Filter = "TXT File (*.txt)|*.txt|All file types|*.*"
        CheckBox1.Checked = True

        MarkListToSave.Text = ""
        MarkList.SelectedIndex = 0
        MarkListToSave.SelectedText = MarkList.SelectedItem & Environment.NewLine
repeat:
        If MarkList.SelectedIndex = MarkList.Items.Count - 1 Then
            MarkListToSave.SelectedText = MarkList.SelectedItem
            If sfd.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim writer As New System.IO.StreamWriter(sfd.FileName, True, System.Text.Encoding.Default)
                writer.Write(MarkListToSave.Text)
                writer.Close()
                ToolTip1.Show("Successfully Saved to: " & sfd.FileName & "!", Me, Control.MousePosition.X - Me.Location.X, Control.MousePosition.Y - Me.Location.Y, 4000)
            End If
        Else
            MarkList.SelectedIndex += 1
            MarkListToSave.SelectedText = MarkList.SelectedItem & Environment.NewLine
            GoTo repeat
        End If
        If Me.Tag = "rgjfkjztrfzt" Then
err:
            MessageBox.Show("Error while converting list to TXT, because you maybe had that list empty or it contains some characters that he can't understand." & Environment.NewLine & "Please check the list again, if is something wrong. Thanks!", "Error while Converting", MessageBoxButtons.OK)
        End If
    End Sub

    Private Sub CheckBox14_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox14.CheckedChanged
        ComboBox2.Enabled = CheckBox14.Checked
        If ComboBox2.SelectedIndex = 2 Then
            TActL.Enabled = False
        Else
            TActL.Enabled = CheckBox14.Checked
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex = 0 Then
            TActL.Enabled = True
            TActL.Text = "Info"
            TActL.BackColor = Color.LightGray
        ElseIf ComboBox2.SelectedIndex = 1 Then
            TActL.Enabled = True
            TActL.Text = "RightDown"
            TActL.BackColor = Color.LightGray
        ElseIf ComboBox2.SelectedIndex = 2 Then
            TActL.Enabled = False
            TActL.Text = "'None'"
            TActL.BackColor = Color.LightGray
        ElseIf ComboBox2.SelectedIndex = 3 Then
            TActL.Enabled = True
            TActL.Text = ""
            TActL.BackColor = Color.Black
        ElseIf ComboBox2.SelectedIndex = 4 Then
            TActL.Enabled = True
            TActL.Text = "File ''None''"
            TActL.BackColor = Color.LightGray
        ElseIf ComboBox2.SelectedIndex = 5 Then
            TActL.Enabled = False
            TActL.Text = "'None'"
            TActL.BackColor = Color.LightGray
        ElseIf ComboBox2.SelectedIndex = 6 Then
            TActL.Enabled = True
            TActL.Text = "File ''None''"
            TActL.BackColor = Color.LightGray
        ElseIf ComboBox2.SelectedIndex = 7 Then
            TActL.Enabled = True
            TActL.Text = "Shutdown"
            TActL.BackColor = Color.LightGray
        ElseIf ComboBox2.SelectedIndex = 8 Then
            TActL.Enabled = True
            TActL.Text = "LogOff"
            TActL.BackColor = Color.LightGray
        End If
    End Sub

    Private Sub tactcm_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles tactcm.Opening
        If ComboBox2.SelectedIndex = 2 Or 3 Or 4 Then
            tacthccmb.Enabled = True
        Else
            tacthccmb.Enabled = False
            tacthccmb.Checked = False
        End If
    End Sub

    Private Sub TActL_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TActL.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Left Then
            If ComboBox2.SelectedIndex = 0 Then
                If TActL.Text = "Info" Then
                    TActL.Text = "Question"
                ElseIf TActL.Text = "Question" Then
                    TActL.Text = "Warning"
                ElseIf TActL.Text = "Warning" Then
                    TActL.Text = "Error"
                ElseIf TActL.Text = "Error" Then
                    TActL.Text = "Info"
                End If
            ElseIf ComboBox2.SelectedIndex = 1 Then
                If TActL.Text = "RightDown" Then
                    TActL.Text = "Right"
                ElseIf TActL.Text = "Right" Then
                    TActL.Text = "RightUp"
                ElseIf TActL.Text = "RightUp" Then
                    TActL.Text = "Up"
                ElseIf TActL.Text = "Up" Then
                    TActL.Text = "LeftUp"
                ElseIf TActL.Text = "LeftUp" Then
                    TActL.Text = "Left"
                ElseIf TActL.Text = "Left" Then
                    TActL.Text = "LeftDown"
                ElseIf TActL.Text = "LeftDown" Then
                    TActL.Text = "Down"
                ElseIf TActL.Text = "Down" Then
                    TActL.Text = "Center"
                ElseIf TActL.Text = "Center" Then
                    TActL.Text = "RightDown"
                End If
            ElseIf ComboBox2.SelectedIndex = 3 Then
                Dim cd As New ColorDialog
                cd.AnyColor = True
                cd.FullOpen = True
                cd.AllowFullOpen = True
                cd.Color = ScreenLocker.BackColor
                CheckBox1.Checked = True
                If cd.ShowDialog = Windows.Forms.DialogResult.OK Then
                    TActL.BackColor = cd.Color
                End If
            ElseIf ComboBox2.SelectedIndex = 4 Then
                Dim ofd As New OpenFileDialog
                ofd.CheckFileExists = True
                ofd.CheckPathExists = True
                ofd.ShowReadOnly = True
                ofd.Title = "Open File Dialog for Picture choosing"
                ofd.Filter = "All picture formats that can be opened (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp"
                ofd.InitialDirectory = "%userprofile%\Pictures"
                CheckBox1.Checked = True
                If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
                    TActL.Text = ofd.FileName
                End If
            ElseIf ComboBox2.SelectedIndex = 6 Then
                Dim ofd As New OpenFileDialog
                ofd.CheckFileExists = True
                ofd.CheckPathExists = True
                ofd.ShowReadOnly = True
                ofd.Title = "Open File Dialog for Program to Start choosing"
                ofd.Filter = "All programs that can be runned (*.exe;*.bat;*.cmd;*.scr)|*.exe;*.bat;*.cmd;*.scr|All files (*.*)|*.*"
                ofd.InitialDirectory = "%userprofile%\Documents"
                CheckBox1.Checked = True
                If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
                    TActL.Text = ofd.FileName
                End If
            ElseIf ComboBox2.SelectedIndex = 7 Then
                If TActL.Text = "Shutdown" Then
                    TActL.Text = "Restart"
                ElseIf TActL.Text = "Restart" Then
                    TActL.Text = "Sleep"
                ElseIf TActL.Text = "Sleep" Then
                    TActL.Text = "Hibernate"
                ElseIf TActL.Text = "Hibernate" Then
                    TActL.Text = "Shutdown"
                End If
            ElseIf ComboBox2.SelectedIndex = 8 Then
                If TActL.Text = "LogOff" Then
                    TActL.Text = "Switch_User"
                ElseIf TActL.Text = "Switch_User" Then
                    TActL.Text = "Lock_Screen"
                ElseIf TActL.Text = "Lock" Then
                    TActL.Text = "LogOff"
                End If
            End If
        ElseIf e.Button = Windows.Forms.MouseButtons.Middle Then
            If ComboBox2.SelectedIndex = 1 Then
                CheckBox1.Checked = True
                NotifyWindowD.ShowDialog()
            End If
        End If
    End Sub

    Private Sub MouseLocker_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MouseLocker.Tick
        Cursor.Position = ML
    End Sub


    REM ALAAARMMMM


    Private Sub CheckBox17_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox17.CheckedChanged
        If Button17.Text = "Save" OrElse CheckBox25.Checked = True Then
            GroupBox9.Enabled = CheckBox17.Checked
        End If
        If CheckBox17.Checked = True Then
            DateTimePicker1.Visible = False
            DateTimePicker1.Enabled = False
            Label40.Visible = True
            Label40.Text = "''By repeating cannot be Date modified.''"
            If Not AlarmList.SelectedIndex = -1 Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Date", "Nothing", Microsoft.Win32.RegistryValueKind.String)
            End If
        Else
            DateTimePicker1.Visible = True
            DateTimePicker1.Enabled = True
            Label40.Visible = False
            Label40.Text = DateTimePicker1.Text
            If Not AlarmList.SelectedIndex = -1 Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Date", DateTimePicker1.Text, Microsoft.Win32.RegistryValueKind.String)
            End If
        End If
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        If CheckBox25.Checked = True Then
            AppBar.AlarmLoadDefault()
        Else
            AlarmLoadDefault()
        End If
        CheckBox1.Checked = True
        If ComboBox5.SelectedIndex = 0 Then
            APD.MODE = 1
            APD.ShowDialog()
        ElseIf ComboBox5.SelectedIndex = 1 Then
            APD.MODE = 2
            APD.ShowDialog()
        ElseIf ComboBox5.SelectedIndex = 2 Then
            APD.MODE = 3
            If CheckBox25.Checked = True Then
                APD.ShowDialog(Me)
            Else
                If APD.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\", "Path", APD.remCB21, Microsoft.Win32.RegistryValueKind.String)
                    If APD.remRB21 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\", "Mode", "1", Microsoft.Win32.RegistryValueKind.DWord)

                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "Arguments", APD.remTXT211, Microsoft.Win32.RegistryValueKind.String)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "AppWinStyle", APD.remCB211, Microsoft.Win32.RegistryValueKind.DWord)
                        If APD.remCHB211 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "CreateNoWindow", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "CreateNoWindow", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If APD.remCHB212 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "UseShellExecute", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "UseShellExecute", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If APD.remCHB213 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RunAs", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "Username", APD.remTXT212, Microsoft.Win32.RegistryValueKind.String)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RunAs", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If APD.remCHB214 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "LoadUserProfile", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "LoadUserProfile", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If APD.remCHB215 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "EnableCustomWorkingDir", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "CustomWorkDir", APD.remTXT212, Microsoft.Win32.RegistryValueKind.String)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "EnableCustomWorkingDir", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If APD.remCHB216 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "EnableDomain", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "Domain", APD.remTXT215, Microsoft.Win32.RegistryValueKind.String)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "EnableDomain", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If APD.remCHB217 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "ErrorDialog", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "ErrorDialog", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If APD.remCHB218 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "ErrorDialogParentHandle", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "ErrorDialogParentHandle", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If APD.remCHB219 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSError", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSError", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If APD.remCHB21a = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSInput", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSInput", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If APD.remCHB21b = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSOutput", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSOutput", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\", "Mode", "0", Microsoft.Win32.RegistryValueKind.DWord)

                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\2\", "AppWinStyle", APD.remCB221, Microsoft.Win32.RegistryValueKind.DWord)
                        If APD.remCHB221 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\2\", "Wait", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\2\", "Timeout", APD.remNUP221, Microsoft.Win32.RegistryValueKind.String)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\2\", "Wait", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\2\", "Timeout", "-1", Microsoft.Win32.RegistryValueKind.String)
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ControllerA_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ControllerA.Tick

    End Sub
    Dim remLAIF As Integer = -1
    Private Sub CheckBox25_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox25.CheckedChanged
        If CheckBox25.Checked = True Then
            CheckBox17.Checked = False
            TextBox1.Text = ""
            TextBox1.Enabled = True
            remLAIF = AlarmList.SelectedIndex
            AlarmList.SelectedIndex = -1
            AlarmList.Enabled = False
            GroupBox8.Visible = True
            GroupBox8.Text = "New Alarm"
            DateTimePicker1.Enabled = True
            DateTimePicker1.Visible = True
            ComboBox5.Enabled = True
            Button19.Enabled = True
            Button17.Text = "Create"
            Label40.Visible = False
            Button20.Visible = False
            A_H.Enabled = True
            A_M.Enabled = True
            A_S.Enabled = True
            CheckBox17.Enabled = True
            If CheckBox17.Checked = True Then
                GroupBox9.Enabled = True
            End If
        Else
            AlarmList.Enabled = True
            If remLAIF = -1 Then
                GroupBox8.Visible = False
            Else
                AlarmList.SelectedIndex = remLAIF
            End If
        End If
    End Sub

    Private Sub AlarmList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AlarmList.SelectedIndexChanged
        Dim readV = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Description", Nothing)
        If CheckBox25.Checked = True Then
            CheckBox1.Checked = True
            If Not AlarmList.SelectedIndex = -1 Then
                Select Case MessageBox.Show("Are you sure do you want to leave?", "Unsaved work detected", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
                    Case Windows.Forms.DialogResult.No
                        remLAIF = AlarmList.SelectedIndex
                        AlarmList.SelectedIndex = -1
                    Case Windows.Forms.DialogResult.Yes
                        CheckBox25.Checked = False
                        GroupBox8.Visible = True
                        GroupBox8.Text = "Alarm Properties"
                        Button17.Text = "TEST"
                        DateTimePicker1.Visible = False
                        DateTimePicker1.Enabled = False
                        Label40.Visible = True
                        Button20.Visible = True
                        A_H.Enabled = False
                        A_M.Enabled = False
                        A_S.Enabled = False
                        CheckBox17.Enabled = False
                        GroupBox9.Enabled = False
                        mi = False
                        TextBox1.Enabled = False
                        Dim readValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Date", Nothing)
                        Dim RV = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\", "MODE", Nothing)
                        Dim readH = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Hour", Nothing)
                        Dim readM = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Minute", Nothing)
                        Dim readS = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Second", Nothing)
                        ComboBox5.SelectedIndex = RV
                        Label40.Text = readValue
                        TextBox1.Text = readV
                        A_H.Value = readH
                        A_M.Value = readM
                        A_S.Value = readS
                        If My.Computer.Registry.CurrentUser.OpenSubKey("ALARMS\" & AlarmList.SelectedItem & "\Repeat\", True) Is Nothing Then
                            CheckBox17.Checked = False
                            M_CHB.Checked = False
                            T_CHB.Checked = False
                            W_CHB.Checked = False
                            Th_CHB.Checked = False
                            F_CHB.Checked = False
                            SA_CHB.Checked = False
                            SU_CHB.Checked = False
                        Else
                            CheckBox17.Checked = True
                            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Monday", Nothing) Is Nothing Then
                                M_CHB.Checked = False
                            Else
                                M_CHB.Checked = True
                            End If
                            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Tuesday", Nothing) Is Nothing Then
                                T_CHB.Checked = False
                            Else
                                T_CHB.Checked = True
                            End If
                            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Wednesday", Nothing) Is Nothing Then
                                W_CHB.Checked = False
                            Else
                                W_CHB.Checked = True
                            End If
                            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Thursday", Nothing) Is Nothing Then
                                Th_CHB.Checked = False
                            Else
                                Th_CHB.Checked = True
                            End If
                            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Friday", Nothing) Is Nothing Then
                                F_CHB.Checked = False
                            Else
                                F_CHB.Checked = True
                            End If
                            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Saturday", Nothing) Is Nothing Then
                                SA_CHB.Checked = False
                            Else
                                SA_CHB.Checked = True
                            End If
                            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Sunday", Nothing) Is Nothing Then
                                SU_CHB.Checked = False
                            Else
                                SU_CHB.Checked = True
                            End If
                        End If
                End Select
            End If
        Else
            GroupBox8.Visible = True
            GroupBox8.Text = "Alarm Properties"
            DateTimePicker1.Visible = False
            DateTimePicker1.Enabled = False
            Button17.Text = "TEST"
            Label40.Visible = True
            Button20.Visible = True
            A_H.Enabled = False
            A_M.Enabled = False
            A_S.Enabled = False
            CheckBox17.Enabled = False
            GroupBox9.Enabled = False
            mi = False
            TextBox1.Enabled = False
            Dim readValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Date", Nothing)
            Dim RV = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\", "MODE", Nothing)
            Dim readH = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Hour", Nothing)
            Dim readM = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Minute", Nothing)
            Dim readS = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Second", Nothing)
            TextBox1.Text = readV
            Label40.Text = readValue
            ComboBox5.SelectedIndex = RV
            A_H.Value = readH
            A_M.Value = readM
            A_S.Value = readS
            If My.Computer.Registry.CurrentUser.OpenSubKey("ALARMS\" & AlarmList.SelectedItem & "\Repeat\") Is Nothing Then
                CheckBox17.Checked = False
                M_CHB.Checked = False
                T_CHB.Checked = False
                W_CHB.Checked = False
                Th_CHB.Checked = False
                F_CHB.Checked = False
                SA_CHB.Checked = False
                SU_CHB.Checked = False
            Else
                CheckBox17.Checked = True
                If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Monday", Nothing) Is Nothing Then
                    M_CHB.Checked = False
                Else
                    M_CHB.Checked = True
                End If
                If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Tuesday", Nothing) Is Nothing Then
                    T_CHB.Checked = False
                Else
                    T_CHB.Checked = True
                End If
                If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Wednesday", Nothing) Is Nothing Then
                    W_CHB.Checked = False
                Else
                    W_CHB.Checked = True
                End If
                If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Thursday", Nothing) Is Nothing Then
                    Th_CHB.Checked = False
                Else
                    Th_CHB.Checked = True
                End If
                If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Friday", Nothing) Is Nothing Then
                    F_CHB.Checked = False
                Else
                    F_CHB.Checked = True
                End If
                If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Saturday", Nothing) Is Nothing Then
                    SA_CHB.Checked = False
                Else
                    SA_CHB.Checked = True
                End If
                If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Sunday", Nothing) Is Nothing Then
                    SU_CHB.Checked = False
                Else
                    SU_CHB.Checked = True
                End If
            End If
        End If
        GroupBox8.Text = "Alarm Options"
        DateTimePicker1.Visible = False
        DateTimePicker1.Enabled = False
        Label40.Visible = True
        CheckBox17.Enabled = False
        If AlarmList.SelectedIndex = -1 Then
            Button18.Enabled = False
        Else
            Button18.Enabled = True
        End If
    End Sub
    Dim WP As String
    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click, ToolStripMenuItem3.Click
        On Error Resume Next
        CheckBox1.Checked = True
        If MessageBox.Show("Are you sure do you want to remove this alarm?", "Microsoft Windows", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
            My.Computer.Registry.CurrentUser.DeleteSubKeyTree("ALARMS\" & AlarmList.SelectedItem)
            AlarmList.Items.RemoveAt(AlarmList.SelectedIndex)
            If AlarmList.SelectedIndex = -1 Then
                GroupBox8.Visible = False
            End If
        End If
    End Sub

    Public Sub AlarmLoadDefault()
        On Error Resume Next
        REM 1
        Dim RV11 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomB1", Nothing)
        Dim RV11_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "CustomB1", Nothing)
        Dim RV12 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableB2", Nothing)
        Dim RV12_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomB2", Nothing)
        Dim RV12__ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "CustomB2", Nothing)
        Dim RV13 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomPicture", Nothing)
        Dim RV13_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "CustomPicturePath", Nothing)
        Dim RV14 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomTitle", Nothing)
        Dim RV14_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "CustomTitle", Nothing)
        Dim RV15 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "Position", Nothing)
        Dim RV15X = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "X", Nothing)
        Dim RV15Y = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "Y", Nothing)
        Dim RV16 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableRinging", Nothing)
        Dim RV17 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingMode", Nothing)
        Dim RV18 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingALIST", Nothing)
        Dim RV19 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingPath", Nothing)
        Dim RV1a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingSYS", Nothing)
        Dim RV1b = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingRepeat", Nothing)
        Dim RV1c = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingTimesRepeat", Nothing)
        Dim RV1d = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingVol", Nothing)

        APD.NumericUpDown1.Maximum = SystemInformation.PrimaryMonitorSize.Width
        APD.NumericUpDown2.Maximum = SystemInformation.PrimaryMonitorSize.Height
        If RV11 = "1" Then
            APD.remDUD4 = 1
            APD.DomainUpDown4.SelectedIndex = 1
            If Not RV11_ = String.Empty Then
                APD.remTXT3 = RV11_
                APD.TextBox3.Text = RV11_
            End If
        Else
            APD.remDUD4 = 0
            APD.DomainUpDown4.SelectedIndex = 0
            APD.remTXT3 = ""
            APD.TextBox3.Text = ""
        End If
        If RV12 = "1" Then
            APD.remCHB3 = True
            APD.CheckBox2.Checked = True
            If RV12_ = "1" Then
                APD.remDUD5 = 1
                APD.DomainUpDown5.SelectedIndex = 1
                If Not RV12__ = String.Empty Then
                    APD.remTXT4 = RV12__
                    APD.TextBox2.Text = RV12__
                End If
            Else
                APD.remDUD5 = 0
                APD.DomainUpDown5.SelectedIndex = 0
                APD.remTXT4 = ""
                APD.TextBox2.Text = ""
            End If
        Else
            APD.remCHB3 = False
            APD.CheckBox2.Checked = False
        End If
        If RV13 = "2" Then
            APD.remDUD3 = 2
            APD.DomainUpDown2.SelectedIndex = 2
            If Not RV13_ = String.Empty Then
                APD.remTXT2 = RV13_
                APD.TextBox4.Text = RV13_
            End If
        ElseIf RV13 = "1" Then
            APD.remDUD3 = 1
            APD.DomainUpDown2.SelectedIndex = 1
            APD.remTXT2 = ""
            APD.TextBox4.Text = ""
        Else
            APD.remDUD3 = 0
            APD.DomainUpDown2.SelectedIndex = 0
            APD.remTXT2 = ""
            APD.TextBox4.Text = ""
        End If
        If RV14 = "1" Then
            APD.remDUD2 = 1
            APD.DomainUpDown3.SelectedIndex = 1
            If Not RV14_ = String.Empty Then
                APD.remTXT1 = RV14_
                APD.TextBox1.Text = RV14_
            End If
        Else
            APD.remDUD2 = 0
            APD.DomainUpDown3.SelectedIndex = 0
            APD.remTXT1 = ""
            APD.TextBox1.Text = ""
        End If
        If RV15 >= 9 Then
            APD.remCB1 = 9
            APD.ComboBox1.SelectedIndex = 9
            APD.remNUPX = RV15X
            APD.remNUPY = RV15Y
            APD.NumericUpDown1.Value = RV15X
            APD.NumericUpDown2.Value = RV15Y
        ElseIf RV15 < 9 Then
            APD.remCB1 = RV15
            APD.ComboBox1.SelectedIndex = RV15
            APD.remNUPX = RV15X
            APD.remNUPY = RV15Y
            APD.NumericUpDown1.Value = RV15X
            APD.NumericUpDown2.Value = RV15Y
        End If
        If RV16 = "1" Then
            APD.remCHB1 = True
            APD.CheckBox1.Checked = True
        Else
            APD.remCHB1 = True
            APD.CheckBox1.Checked = True
        End If
        If RV1b = "1" Then
            APD.remCHB2 = True
            APD.CheckBox16.Checked = True
        Else
            APD.remCHB2 = False
            APD.CheckBox16.Checked = False
        End If
        If RV1c >= 1 Then
            APD.remNUP2 = RV1c
            APD.ThmN.Value = RV1c
        End If
        If RV1d >= 0 AndAlso RV1d <= 100 Then
            APD.remNUP1 = RV1d
            APD.NumericUpDown12.Value = RV1d
        End If
        If RV17 = "0" Then
            APD.remDUD1 = 0
            APD.DomainUpDown1.SelectedIndex = 0
            APD.remCB2 = RV1a
            APD.ComboBox2.SelectedIndex = RV1a
        ElseIf RV17 = "1" Then
            Dim i As String
            Dim CTSPath As String = "C:\Windows\Media\Alarms"
            If My.Computer.FileSystem.DirectoryExists(CTSPath) Then
                For Each i In My.Computer.FileSystem.GetFiles(CTSPath)
                    APD.ComboBox3.Items.Add(i.Substring(i.LastIndexOf("\") + 1))
                Next
            End If
            APD.remDUD1 = 1
            APD.DomainUpDown1.SelectedIndex = 1
            APD.remCB3 = RV18
            APD.ComboBox3.SelectedIndex = RV18
        Else
            APD.remDUD1 = 2
            APD.DomainUpDown1.SelectedIndex = 2
            APD.remCB4 = RV19
            APD.ComboBox4.Text = RV19
        End If
        REM 2
        Dim RV21 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingMode", Nothing)
        Dim RV22 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingALIST", Nothing)
        Dim RV23 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingPath", Nothing)
        Dim RV24 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingSYS", Nothing)
        Dim RV25 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingRepeat", Nothing)
        Dim RV26 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingTimesRepeat", Nothing)
        Dim RV27 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingVol", Nothing)

        If RV25 = "1" Then
            APD.remCHB11 = True
            APD.CheckBox3.Checked = True
        Else
            APD.remCHB11 = False
            APD.CheckBox3.Checked = False
        End If
        If RV26 >= 1 Then
            APD.remNUP12 = RV26
            APD.NumericUpDown3.Value = RV26
        End If
        If RV27 >= 0 AndAlso RV27 <= 100 Then
            APD.remNUP11 = RV27
            APD.NumericUpDown4.Value = RV27
        End If
        If RV21 = "0" Then
            APD.remDUD11 = 0
            APD.DomainUpDown6.SelectedIndex = 0
            APD.remCB11 = RV24
            APD.ComboBox6.SelectedIndex = RV24
        ElseIf RV21 = "1" Then
            Dim i As String
            Dim CTSPath As String = "C:\Windows\Media\Alarms"
            If My.Computer.FileSystem.DirectoryExists(CTSPath) Then
                For Each i In My.Computer.FileSystem.GetFiles(CTSPath)
                    APD.ComboBox7.Items.Add(i.Substring(i.LastIndexOf("\") + 1))
                Next
            End If
            APD.remDUD11 = 1
            APD.DomainUpDown6.SelectedIndex = 1
            APD.remCB12 = RV22
            APD.ComboBox7.SelectedIndex = RV22
        Else
            APD.remDUD11 = 2
            APD.DomainUpDown6.SelectedIndex = 2
            APD.remCB13 = RV23
            APD.ComboBox5.Text = RV23
        End If
        REM 3
        Dim RV31 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\", "MODE", Nothing)
        Dim RV32 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\", "Path", Nothing)

        If RV31 = "0" Then
            APD.remRB21 = False
            APD.RadioButton3.Checked = False
            APD.RadioButton2.Checked = True
        Else
            APD.remRB21 = True
            APD.RadioButton3.Checked = True
            APD.RadioButton2.Checked = False
        End If
        If IO.File.Exists(RV32) = True Then
            APD.remCB21 = RV32
            APD.ComboBox8.Text = RV32
        End If
        REM 3.1
        Dim RV311 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "AppWinStyle", Nothing)
        Dim RV312 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "Arguments", Nothing)
        Dim RV313 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "CreateNoWindow", Nothing)
        Dim RV314 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "EnableCustomWorkingDir", Nothing)
        Dim RV315 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "CustomWorkDir", Nothing)
        Dim RV316 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "EnableDomain", Nothing)
        Dim RV316_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "Domain", Nothing)
        Dim RV317 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "RunAs", Nothing)
        Dim RV318 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "Username", Nothing)
        Dim RV319 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "LoadUserProfile", Nothing)
        Dim RV31a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "UseShellExecute", Nothing)
        Dim RV31b = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "ErrorDialog", Nothing)
        Dim RV31c = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "ErrorDialogParentHandle", Nothing)
        Dim RV31d = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "RSError", Nothing)
        Dim RV31e = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "RSInput", Nothing)
        Dim RV31f = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "RSOutput", Nothing)

        If RV311 < 4 Then
            APD.remCB211 = RV311
            APD.ComboBox10.SelectedIndex = RV311
        End If
        APD.remTXT211 = RV312
        APD.TextBox5.Text = RV312
        If RV313 = "1" Then
            APD.remCHB211 = True
            APD.CheckBox7.Checked = True
        Else
            APD.remCHB211 = False
            APD.CheckBox7.Checked = False
        End If
        If RV31a = "1" Then
            APD.remCHB212 = True
            APD.CheckBox8.Checked = True
        Else
            APD.remCHB212 = False
            APD.CheckBox8.Checked = False
        End If
        If RV319 = "1" Then
            APD.remCHB214 = True
            APD.CheckBox9.Checked = True
        Else
            APD.remCHB214 = False
            APD.CheckBox9.Checked = False
        End If
        If RV31b = "1" Then
            APD.remCHB217 = True
            APD.CheckBox12.Checked = True
        Else
            APD.remCHB217 = False
            APD.CheckBox12.Checked = False
        End If
        If RV31c = "1" Then
            APD.remCHB218 = True
            APD.CheckBox13.Checked = True
        Else
            APD.remCHB218 = False
            APD.CheckBox13.Checked = False
        End If
        If RV31d = "1" Then
            APD.remCHB219 = True
            APD.CheckBox14.Checked = True
        Else
            APD.remCHB219 = False
            APD.CheckBox14.Checked = False
        End If
        If RV31e = "1" Then
            APD.remCHB21a = True
            APD.CheckBox15.Checked = True
        Else
            APD.remCHB21a = False
            APD.CheckBox15.Checked = False
        End If
        If RV31f = "1" Then
            APD.remCHB21b = True
            APD.CheckBox17.Checked = True
        Else
            APD.remCHB21b = False
            APD.CheckBox17.Checked = False
        End If
        APD.remTXT214 = RV315
        APD.TextBox7.Text = RV315
        If RV314 = "1" Then
            APD.remCHB215 = True
            APD.CheckBox10.Checked = True
        Else
            APD.remCHB215 = False
            APD.CheckBox10.Checked = False
        End If
        APD.remTXT215 = RV316_
        APD.TextBox8.Text = RV316_
        If RV316 = "1" Then
            APD.remCHB216 = True
            APD.CheckBox11.Checked = True
        Else
            APD.remCHB216 = False
            APD.CheckBox11.Checked = False
        End If
        APD.remTXT212 = RV318
        APD.TextBox6.Text = RV318
        If RV317 = "1" Then
            APD.remCHB213 = True
            APD.CheckBox4.Checked = True
        Else
            APD.remCHB213 = False
            APD.CheckBox4.Checked = False
        End If
        REM 3.2
        Dim RV321 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\2\", "AppWinStyle", Nothing)
        Dim RV322 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\2\", "Wait", Nothing)
        Dim RV323 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\2\", "Timeout", Nothing)

        If RV321 < 6 Then
            APD.remCB221 = RV321
            APD.ComboBox9.SelectedIndex = RV321
        End If
        If RV322 = "1" Then
            APD.remCHB221 = True
            APD.CheckBox6.Checked = True
        Else
            APD.remCHB221 = False
            APD.CheckBox6.Checked = False
        End If
        If RV323 >= -1 Then
            APD.remNUP221 = RV323
            APD.NumericUpDown5.Value = RV323
        Else
            APD.remNUP221 = -1
            APD.NumericUpDown5.Value = -1
        End If
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        If Button17.Text = "Create" Then
            If My.Computer.Registry.CurrentUser.OpenSubKey("ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS, True) Is Nothing Then
0:
                AlarmList.Items.Add(DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS)
                My.Computer.Registry.CurrentUser.CreateSubKey("ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS, "Hour", AFixH, Microsoft.Win32.RegistryValueKind.String)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS, "Minute", AFixM, Microsoft.Win32.RegistryValueKind.String)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS, "Second", AFixS, Microsoft.Win32.RegistryValueKind.String)
                If CheckBox17.Checked = True Then
                    My.Computer.Registry.CurrentUser.CreateSubKey("ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Repeat\", Microsoft.Win32.RegistryKeyPermissionCheck.Default)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS, "Date", "Nothing", Microsoft.Win32.RegistryValueKind.String)
                    If M_CHB.Checked = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Repeat\", "Monday", "", Microsoft.Win32.RegistryValueKind.String)
                        'CreateObject("WScript.Shell").RegDelete("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Repeat\Monday")
                    End If
                    If T_CHB.Checked = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Repeat\", "Tuesday", "", Microsoft.Win32.RegistryValueKind.String)
                    End If
                    If W_CHB.Checked = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Repeat\", "Wednesday", "", Microsoft.Win32.RegistryValueKind.String)
                    End If
                    If Th_CHB.Checked = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Repeat\", "Thursday", "", Microsoft.Win32.RegistryValueKind.String)
                    End If
                    If F_CHB.Checked = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Repeat\", "Friday", "", Microsoft.Win32.RegistryValueKind.String)
                    End If
                    If SA_CHB.Checked = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Repeat\", "Saturday", "", Microsoft.Win32.RegistryValueKind.String)
                    End If
                    If SU_CHB.Checked = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Repeat\", "Sunday", "", Microsoft.Win32.RegistryValueKind.String)
                    End If
                Else
                    If My.Computer.Registry.CurrentUser.OpenSubKey("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Repeat\", True) IsNot Nothing Then My.Computer.Registry.CurrentUser.DeleteSubKeyTree("ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Repeat\")
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS, "Date", DateTimePicker1.Text, Microsoft.Win32.RegistryValueKind.String)
                End If
                'if textbox1.Text
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS, "Description", TextBox1.Text, Microsoft.Win32.RegistryValueKind.String)
                My.Computer.Registry.CurrentUser.CreateSubKey("ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\", Microsoft.Win32.RegistryKeyPermissionCheck.Default)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\", "MODE", ComboBox5.SelectedIndex, Microsoft.Win32.RegistryValueKind.DWord)
                If APD.SaveBTN.Checked = True Then
                    If ComboBox5.SelectedIndex = 0 Then
                        If APD.remCB1 = 9 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "X", APD.remNUPX, Microsoft.Win32.RegistryValueKind.String)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "Y", APD.remNUPY, Microsoft.Win32.RegistryValueKind.String)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "Position", APD.remCB1, Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "X", APD.remNUPX, Microsoft.Win32.RegistryValueKind.String)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "Y", APD.remNUPY, Microsoft.Win32.RegistryValueKind.String)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "Position", APD.remCB1, Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "Position", APD.remCB1, Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "Position", APD.remCB1, Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If APD.remCHB1 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableRinging", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingMode", APD.remDUD1, Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableRinging", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "RingMode", APD.remDUD1, Microsoft.Win32.RegistryValueKind.DWord)
                            If APD.remDUD1 = 0 Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingSYS", APD.remCB2, Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "RingSYS", APD.remCB2, Microsoft.Win32.RegistryValueKind.DWord)
                            ElseIf APD.remDUD1 = 1 Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingALIST", APD.remCB3, Microsoft.Win32.RegistryValueKind.String)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "RingALIST", APD.remCB3, Microsoft.Win32.RegistryValueKind.String)
                            ElseIf APD.remDUD1 = 2 Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingPath", APD.remCB4, Microsoft.Win32.RegistryValueKind.String)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "RingPath", APD.remCB4, Microsoft.Win32.RegistryValueKind.String)
                            End If
                            If APD.remCHB2 = False Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingRepeat", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingTimesRepeat", APD.remNUP2, Microsoft.Win32.RegistryValueKind.String)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "RingRepeat", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "RingTimesRepeat", APD.remNUP12, Microsoft.Win32.RegistryValueKind.String)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingRepeat", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "RingVol", APD.remNUP11, Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "RingVol", APD.remNUP11, Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableRinging", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableRinging", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If APD.remCHB3 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableB2", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableB2", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            If APD.remDUD5 = 0 Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomB2", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomB2", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomB2", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomB2", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "CustomB2", APD.remTXT4, Microsoft.Win32.RegistryValueKind.String)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "CustomB2", APD.remTXT4, Microsoft.Win32.RegistryValueKind.String)
                            End If
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableB2", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableB2", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If APD.remDUD2 = 0 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomTitle", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomTitle", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomTitle", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomTitle", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "CustomTitle", APD.remTXT1, Microsoft.Win32.RegistryValueKind.String)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "CustomTitle", APD.remTXT1, Microsoft.Win32.RegistryValueKind.String)
                        End If
                        If APD.remDUD3 = 0 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomPicture", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomPicture", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        ElseIf APD.remDUD3 = 1 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomPicture", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomPicture", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        ElseIf APD.remDUD3 = 2 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomPicture", "2", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomPicture", "2", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "CustomPicturePath", APD.remTXT2, Microsoft.Win32.RegistryValueKind.String)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "CustomPicturePath", APD.remTXT2, Microsoft.Win32.RegistryValueKind.String)
                        End If
                        If APD.remDUD4 = 0 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomB1", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomB1", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "EnableCustomB1", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomB1", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\1\", "CustomB1", APD.remTXT3, Microsoft.Win32.RegistryValueKind.String)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "CustomB1", APD.remTXT3, Microsoft.Win32.RegistryValueKind.String)
                        End If
                    ElseIf ComboBox5.SelectedIndex = 1 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingMode", APD.remDUD11, Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\2\", "RingMode", APD.remDUD11, Microsoft.Win32.RegistryValueKind.DWord)
                        If APD.remDUD11 = 0 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingSYS", APD.remCB2, Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\2\", "RingSYS", APD.remCB11, Microsoft.Win32.RegistryValueKind.DWord)
                        ElseIf APD.remDUD11 = 1 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingALIST", APD.remCB3, Microsoft.Win32.RegistryValueKind.String)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\2\", "RingALIST", APD.remCB12, Microsoft.Win32.RegistryValueKind.String)
                        ElseIf APD.remDUD11 = 2 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingPath", APD.remCB4, Microsoft.Win32.RegistryValueKind.String)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\2\", "RingPath", APD.remCB13, Microsoft.Win32.RegistryValueKind.String)
                        End If
                        If APD.remCHB11 = False Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingRepeat", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingTimesRepeat", APD.remNUP12, Microsoft.Win32.RegistryValueKind.String)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\2\", "RingRepeat", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\2\", "RingTimesRepeat", APD.remNUP12, Microsoft.Win32.RegistryValueKind.String)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingRepeat", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\2\", "RingRepeat", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\2\", "RingVol", APD.remNUP11, Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\2\", "RingVol", APD.remNUP11, Microsoft.Win32.RegistryValueKind.DWord)
                    ElseIf ComboBox5.SelectedIndex = 2 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\", "Path", APD.remCB21, Microsoft.Win32.RegistryValueKind.String)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\", "Path", APD.remCB21, Microsoft.Win32.RegistryValueKind.String)
                        If APD.remRB21 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\", "Mode", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\", "Mode", "1", Microsoft.Win32.RegistryValueKind.DWord)

                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "Arguments", APD.remTXT211, Microsoft.Win32.RegistryValueKind.String)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "AppWinStyle", APD.remCB211, Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "Arguments", APD.remTXT211, Microsoft.Win32.RegistryValueKind.String)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "AppWinStyle", APD.remCB211, Microsoft.Win32.RegistryValueKind.DWord)
                            If APD.remCHB211 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "CreateNoWindow", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "CreateNoWindow", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "CreateNoWindow", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "CreateNoWindow", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB212 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "UseShellExecute", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "UseShellExecute", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "UseShellExecute", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "UseShellExecute", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB213 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RunAs", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "Username", APD.remTXT212, Microsoft.Win32.RegistryValueKind.String)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RunAs", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "Username", APD.remTXT212, Microsoft.Win32.RegistryValueKind.String)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RunAs", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RunAs", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB214 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "LoadUserProfile", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "LoadUserProfile", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "LoadUserProfile", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "LoadUserProfile", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB215 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "EnableCustomWorkingDir", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "CustomWorkDir", APD.remTXT212, Microsoft.Win32.RegistryValueKind.String)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "EnableCustomWorkingDir", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "CustomWorkDir", APD.remTXT212, Microsoft.Win32.RegistryValueKind.String)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "EnableCustomWorkingDir", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "EnableCustomWorkingDir", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB216 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "EnableDomain", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "Domain", APD.remTXT215, Microsoft.Win32.RegistryValueKind.String)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "EnableDomain", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "Domain", APD.remTXT215, Microsoft.Win32.RegistryValueKind.String)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "EnableDomain", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "EnableDomain", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB217 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "ErrorDialog", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "ErrorDialog", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "ErrorDialog", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "ErrorDialog", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB218 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "ErrorDialogParentHandle", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "ErrorDialogParentHandle", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "ErrorDialogParentHandle", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "ErrorDialogParentHandle", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB219 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RSError", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSError", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RSError", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSError", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB21a = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RSInput", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSInput", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RSInput", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSInput", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB21b = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RSOutput", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSOutput", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\1\", "RSOutput", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSOutput", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\", "Mode", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\", "Mode", "0", Microsoft.Win32.RegistryValueKind.DWord)

                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\2\", "AppWinStyle", APD.remCB221, Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\2\", "AppWinStyle", APD.remCB221, Microsoft.Win32.RegistryValueKind.DWord)
                            If APD.remCHB221 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\2\", "Wait", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\2\", "Timeout", APD.remNUP221, Microsoft.Win32.RegistryValueKind.String)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\2\", "Wait", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\2\", "Timeout", APD.remNUP221, Microsoft.Win32.RegistryValueKind.String)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\2\", "Wait", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\NW\Alarm\3\2\", "Timeout", "-1", Microsoft.Win32.RegistryValueKind.String)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\2\", "Wait", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\2\", "Timeout", "-1", Microsoft.Win32.RegistryValueKind.String)
                            End If
                        End If
                    End If
                Else
                    If ComboBox5.SelectedIndex = 0 Then
                        If APD.remCB1 = 9 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "X", APD.remNUPX, Microsoft.Win32.RegistryValueKind.String)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "Y", APD.remNUPY, Microsoft.Win32.RegistryValueKind.String)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "Position", APD.remCB1, Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "Position", APD.remCB1, Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If APD.remCHB1 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableRinging", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "RingMode", APD.remDUD1, Microsoft.Win32.RegistryValueKind.DWord)
                            If APD.remDUD1 = 0 Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "RingSYS", APD.remCB2, Microsoft.Win32.RegistryValueKind.DWord)
                            ElseIf APD.remDUD1 = 1 Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "RingALIST", APD.remCB3, Microsoft.Win32.RegistryValueKind.String)
                            ElseIf APD.remDUD1 = 2 Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "RingPath", APD.remCB4, Microsoft.Win32.RegistryValueKind.String)
                            End If
                            If APD.remCHB2 = False Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "RingRepeat", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "RingTimesRepeat", APD.remNUP2, Microsoft.Win32.RegistryValueKind.String)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "RingRepeat", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "RingVol", APD.remNUP1, Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableRinging", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If APD.remCHB3 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableB2", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            If APD.remDUD5 = 0 Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomB2", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomB2", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "CustomB2", APD.remTXT4, Microsoft.Win32.RegistryValueKind.String)
                            End If
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableB2", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        If APD.remDUD2 = 0 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomTitle", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomTitle", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "CustomTitle", APD.remTXT1, Microsoft.Win32.RegistryValueKind.String)
                        End If
                        If APD.remDUD3 = 0 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomPicture", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        ElseIf APD.remDUD3 = 1 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomPicture", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        ElseIf APD.remDUD3 = 2 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomPicture", "2", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "CustomPicturePath", APD.remTXT2, Microsoft.Win32.RegistryValueKind.String)
                        End If
                        If APD.remDUD4 = 0 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomB1", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "EnableCustomB1", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\1\", "CustomB1", APD.remTXT3, Microsoft.Win32.RegistryValueKind.String)
                        End If
                    ElseIf ComboBox5.SelectedIndex = 1 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\2\", "RingMode", APD.remDUD11, Microsoft.Win32.RegistryValueKind.DWord)
                        If APD.remDUD11 = 0 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\2\", "RingSYS", APD.remCB11, Microsoft.Win32.RegistryValueKind.DWord)
                        ElseIf APD.remDUD11 = 1 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\2\", "RingALIST", APD.remCB12, Microsoft.Win32.RegistryValueKind.String)
                        ElseIf APD.remDUD11 = 2 Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\2\", "RingPath", APD.remCB13, Microsoft.Win32.RegistryValueKind.String)
                        End If
                        If APD.remCHB11 = False Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\2\", "RingRepeat", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\2\", "RingTimesRepeat", APD.remNUP12, Microsoft.Win32.RegistryValueKind.String)
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\2\", "RingRepeat", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        End If
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\2\", "RingVol", APD.remNUP11, Microsoft.Win32.RegistryValueKind.DWord)
                    ElseIf ComboBox5.SelectedIndex = 2 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\", "Path", APD.remCB21, Microsoft.Win32.RegistryValueKind.String)
                        If APD.remRB21 = True Then
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\", "Mode", "1", Microsoft.Win32.RegistryValueKind.DWord)

                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "Arguments", APD.remTXT211, Microsoft.Win32.RegistryValueKind.String)
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "AppWinStyle", APD.remCB211, Microsoft.Win32.RegistryValueKind.DWord)
                            If APD.remCHB211 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "CreateNoWindow", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "CreateNoWindow", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB212 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "UseShellExecute", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "UseShellExecute", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB213 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RunAs", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "Username", APD.remTXT212, Microsoft.Win32.RegistryValueKind.String)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RunAs", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB214 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "LoadUserProfile", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "LoadUserProfile", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB215 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "EnableCustomWorkingDir", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "CustomWorkDir", APD.remTXT212, Microsoft.Win32.RegistryValueKind.String)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "EnableCustomWorkingDir", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB216 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "EnableDomain", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "Domain", APD.remTXT215, Microsoft.Win32.RegistryValueKind.String)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "EnableDomain", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB217 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "ErrorDialog", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "ErrorDialog", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB218 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "ErrorDialogParentHandle", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "ErrorDialogParentHandle", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB219 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSError", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSError", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB21a = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSInput", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSInput", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                            If APD.remCHB21b = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSOutput", "1", Microsoft.Win32.RegistryValueKind.DWord)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\1\", "RSOutput", "0", Microsoft.Win32.RegistryValueKind.DWord)
                            End If
                        Else
                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\", "Mode", "0", Microsoft.Win32.RegistryValueKind.DWord)

                            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\2\", "AppWinStyle", APD.remCB221, Microsoft.Win32.RegistryValueKind.DWord)
                            If APD.remCHB221 = True Then
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\2\", "Wait", "1", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\2\", "Timeout", APD.remNUP221, Microsoft.Win32.RegistryValueKind.String)
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\2\", "Wait", "0", Microsoft.Win32.RegistryValueKind.DWord)
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "\Mode\3\2\", "Timeout", "-1", Microsoft.Win32.RegistryValueKind.String)
                            End If
                        End If
                    End If
                End If
                CheckBox25.Checked = False
            Else
                If MessageBox.Show("Oh no, alarm by name: '" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS & "' already exists." & Environment.NewLine & "Do you want replace the alarm with that current one?", "Confirm Overwriting", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                    My.Computer.Registry.CurrentUser.DeleteSubKeyTree("ALARMS\" & DateTimePicker1.Text & "_" & AFixH & ":" & AFixM & ":" & AFixS)
                    GoTo 0
                End If
            End If
        ElseIf Button17.Text = "Save" Then
            My.Computer.Registry.CurrentUser.CreateSubKey("ALARMS\" & AlarmList.SelectedItem)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Hour", AFixH, Microsoft.Win32.RegistryValueKind.String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Minute", AFixM, Microsoft.Win32.RegistryValueKind.String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Second", AFixS, Microsoft.Win32.RegistryValueKind.String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Description", TextBox1.Text, Microsoft.Win32.RegistryValueKind.String)
            If CheckBox17.Checked = True Then
                My.Computer.Registry.CurrentUser.CreateSubKey("ALARMS\" & AlarmList.SelectedItem & "\Repeat\", Microsoft.Win32.RegistryKeyPermissionCheck.Default)
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Date", "Nothing", Microsoft.Win32.RegistryValueKind.String)
                On Error Resume Next
                If M_CHB.Checked = True Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Monday", "", Microsoft.Win32.RegistryValueKind.String)
                    'CreateObject("WScript.Shell").RegDelete("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\Monday")
                Else
                    If My.Computer.Registry.CurrentUser.GetValue("ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Monday") IsNot Nothing Then
                        CreateObject("WScript.Shell").RegDelete("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\Monday")
                    End If
                End If
                If T_CHB.Checked = True Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Tuesday", "", Microsoft.Win32.RegistryValueKind.String)
                Else
                    If My.Computer.Registry.CurrentUser.GetValue("ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Tuesday") IsNot Nothing Then
                        CreateObject("WScript.Shell").RegDelete("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\Tuesday")
                    End If
                    End If
                    If W_CHB.Checked = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Wednesday", "", Microsoft.Win32.RegistryValueKind.String)
                Else
                    If My.Computer.Registry.CurrentUser.GetValue("ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Wednesday") IsNot Nothing Then
                        CreateObject("WScript.Shell").RegDelete("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\Wednesday")
                    End If
                End If
                If Th_CHB.Checked = True Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Thursday", "", Microsoft.Win32.RegistryValueKind.String)
                Else
                    If My.Computer.Registry.CurrentUser.GetValue("ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Thursday") IsNot Nothing Then
                        CreateObject("WScript.Shell").RegDelete("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\Thursday")
                    End If
                End If
                If F_CHB.Checked = True Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Friday", "", Microsoft.Win32.RegistryValueKind.String)
                Else
                    If My.Computer.Registry.CurrentUser.GetValue("ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Friday") IsNot Nothing Then
                        CreateObject("WScript.Shell").RegDelete("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\Friday")
                    End If
                End If
                If SA_CHB.Checked = True Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Saturday", "", Microsoft.Win32.RegistryValueKind.String)
                Else
                    If My.Computer.Registry.CurrentUser.GetValue("ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Saturday") IsNot Nothing Then
                        CreateObject("WScript.Shell").RegDelete("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\Saturday")
                    End If
                End If
                If SU_CHB.Checked = True Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Sunday", "", Microsoft.Win32.RegistryValueKind.String)
                Else
                    If My.Computer.Registry.CurrentUser.GetValue("ALARMS\" & AlarmList.SelectedItem & "\Repeat\", "Sunday") IsNot Nothing Then
                        CreateObject("WScript.Shell").RegDelete("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\Sunday")
                    End If
                End If
            Else
                If My.Computer.Registry.CurrentUser.OpenSubKey("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Repeat\", True) IsNot Nothing Then My.Computer.Registry.CurrentUser.DeleteSubKeyTree("ALARMS\" & AlarmList.SelectedItem & "\Repeat\")
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem, "Date", DateTimePicker1.Text, Microsoft.Win32.RegistryValueKind.String)
            End If
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\", "MODE", ComboBox5.SelectedIndex, Microsoft.Win32.RegistryValueKind.String)
            If ComboBox5.SelectedIndex = 0 Then
                If APD.remCB1 = 9 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "X", APD.remNUPX, Microsoft.Win32.RegistryValueKind.String)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "Y", APD.remNUPY, Microsoft.Win32.RegistryValueKind.String)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "Position", APD.remCB1, Microsoft.Win32.RegistryValueKind.DWord)
                Else
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "Position", APD.remCB1, Microsoft.Win32.RegistryValueKind.DWord)
                End If
                If APD.remCHB1 = True Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableRinging", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingMode", APD.remDUD1, Microsoft.Win32.RegistryValueKind.DWord)
                    If APD.remDUD1 = 0 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingSYS", APD.remCB2, Microsoft.Win32.RegistryValueKind.DWord)
                    ElseIf APD.remDUD1 = 1 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingALIST", APD.remCB3, Microsoft.Win32.RegistryValueKind.String)
                    ElseIf APD.remDUD1 = 2 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingPath", APD.remCB4, Microsoft.Win32.RegistryValueKind.String)
                    End If
                    If APD.remCHB2 = False Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingRepeat", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingTimesRepeat", APD.remNUP2, Microsoft.Win32.RegistryValueKind.String)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingRepeat", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingVol", APD.remNUP1, Microsoft.Win32.RegistryValueKind.DWord)
                Else
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableRinging", "0", Microsoft.Win32.RegistryValueKind.DWord)
                End If
                If APD.remCHB3 = True Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableB2", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    If APD.remDUD5 = 0 Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomB2", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomB2", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "CustomB2", APD.remTXT4, Microsoft.Win32.RegistryValueKind.String)
                    End If
                Else
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableB2", "0", Microsoft.Win32.RegistryValueKind.DWord)
                End If
                If APD.remDUD2 = 0 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomTitle", "0", Microsoft.Win32.RegistryValueKind.DWord)
                Else
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomTitle", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "CustomTitle", APD.remTXT1, Microsoft.Win32.RegistryValueKind.String)
                End If
                If APD.remDUD3 = 0 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomPicture", "0", Microsoft.Win32.RegistryValueKind.DWord)
                ElseIf APD.remDUD3 = 1 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomPicture", "1", Microsoft.Win32.RegistryValueKind.DWord)
                ElseIf APD.remDUD3 = 2 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomPicture", "2", Microsoft.Win32.RegistryValueKind.DWord)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "CustomPicturePath", APD.remTXT2, Microsoft.Win32.RegistryValueKind.String)
                End If
                If APD.remDUD4 = 0 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomB1", "0", Microsoft.Win32.RegistryValueKind.DWord)
                Else
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomB1", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "CustomB1", APD.remTXT3, Microsoft.Win32.RegistryValueKind.String)
                End If
            ElseIf ComboBox5.SelectedIndex = 1 Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingMode", APD.remDUD11, Microsoft.Win32.RegistryValueKind.DWord)
                If APD.remDUD11 = 0 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingSYS", APD.remCB11, Microsoft.Win32.RegistryValueKind.DWord)
                ElseIf APD.remDUD11 = 1 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingALIST", APD.remCB12, Microsoft.Win32.RegistryValueKind.String)
                ElseIf APD.remDUD11 = 2 Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingPath", APD.remCB13, Microsoft.Win32.RegistryValueKind.String)
                End If
                If APD.remCHB11 = False Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingRepeat", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingTimesRepeat", APD.remNUP12, Microsoft.Win32.RegistryValueKind.String)
                Else
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingRepeat", "1", Microsoft.Win32.RegistryValueKind.DWord)
                End If
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingVol", APD.remNUP11, Microsoft.Win32.RegistryValueKind.DWord)
            ElseIf ComboBox5.SelectedIndex = 2 Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\", "Path", APD.remCB21, Microsoft.Win32.RegistryValueKind.String)
                If APD.remRB21 = True Then
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\", "Mode", "1", Microsoft.Win32.RegistryValueKind.DWord)

                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "Arguments", APD.remTXT211, Microsoft.Win32.RegistryValueKind.String)
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "AppWinStyle", APD.remCB211, Microsoft.Win32.RegistryValueKind.DWord)
                    If APD.remCHB211 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "CreateNoWindow", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "CreateNoWindow", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If APD.remCHB212 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "UseShellExecute", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "UseShellExecute", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If APD.remCHB213 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "RunAs", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "Username", APD.remTXT212, Microsoft.Win32.RegistryValueKind.String)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "RunAs", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If APD.remCHB214 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "LoadUserProfile", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "LoadUserProfile", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If APD.remCHB215 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "EnableCustomWorkingDir", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "CustomWorkDir", APD.remTXT212, Microsoft.Win32.RegistryValueKind.String)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "EnableCustomWorkingDir", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If APD.remCHB216 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "EnableDomain", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "Domain", APD.remTXT215, Microsoft.Win32.RegistryValueKind.String)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "EnableDomain", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If APD.remCHB217 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "ErrorDialog", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "ErrorDialog", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If APD.remCHB218 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "ErrorDialogParentHandle", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "ErrorDialogParentHandle", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If APD.remCHB219 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "RSError", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "RSError", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If APD.remCHB21a = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "RSInput", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "RSInput", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                    If APD.remCHB21b = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "RSOutput", "1", Microsoft.Win32.RegistryValueKind.DWord)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "RSOutput", "0", Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                Else
                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\", "Mode", "0", Microsoft.Win32.RegistryValueKind.DWord)

                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\2\", "AppWinStyle", APD.remCB221, Microsoft.Win32.RegistryValueKind.DWord)
                    If APD.remCHB221 = True Then
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\2\", "Wait", "1", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\2\", "Timeout", APD.remNUP221, Microsoft.Win32.RegistryValueKind.String)
                    Else
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\2\", "Wait", "0", Microsoft.Win32.RegistryValueKind.DWord)
                        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\2\", "Timeout", "-1", Microsoft.Win32.RegistryValueKind.String)
                    End If
                End If
            End If
            AlarmList.Enabled = True
            CheckBox25.Enabled = True
            Button18.Enabled = False
            Label40.Visible = True
            CheckBox17.Enabled = False
            GroupBox9.Enabled = False
            mi = False
            A_H.Enabled = False
            A_M.Enabled = False
            A_S.Enabled = False
            TextBox1.Enabled = False
            DateTimePicker1.Visible = False
            DateTimePicker1.Enabled = False
            Button20.Enabled = True
            Button18.Enabled = True
            Button17.Text = "TEST"
        ElseIf Button17.Text = "TEST" Then
            CheckBox1.Checked = True
            ACTION()
        ElseIf Button17.Text = "Stop" Then
            On Error Resume Next
            ComboBox5.Enabled = True
            Button19.Enabled = True
            Button20.Enabled = True
            AlarmList.Enabled = True
            Button18.Enabled = True
            CheckBox15.Enabled = True
            alarm.Stop()
            AppBar.alarm.Stop()
            Button17.Text = "TEST"
            AppBar.ALARMSTOPToolStripMenuItem.Visible = False
            AppBar.ToolStripSeparator20.Visible = False
        End If
    End Sub
    Public Sub ACTION()
        On Error Resume Next
        Dim RV1 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\", "Date", Nothing)
        Dim RV2 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\", "Description", Nothing)
        Dim RV3 As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\", "Hour", Nothing)
        Dim RV4 As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\", "Minute", Nothing)
        Dim RV5 As Integer = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\", "Second", Nothing)
        If ComboBox5.SelectedIndex = 0 Then
            Dim RV11 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomB1", Nothing)
            Dim RV11_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "CustomB1", Nothing)
            Dim RV12 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableB2", Nothing)
            Dim RV12_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomB2", Nothing)
            Dim RV12__ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "CustomB2", Nothing)
            Dim RV13 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomPicture", Nothing)
            Dim RV13_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "CustomPicturePath", Nothing)
            Dim RV14 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableCustomTitle", Nothing)
            Dim RV14_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "CustomTitle", Nothing)
            Dim RV15 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "Position", Nothing)
            Dim RV15X = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "X", Nothing)
            Dim RV15Y = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "Y", Nothing)
            Dim RV16 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "EnableRinging", Nothing)
            Dim RV17 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingMode", Nothing)
            Dim RV18 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingALIST", Nothing)
            Dim RV19 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingPath", Nothing)
            Dim RV1a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingSYS", Nothing)
            Dim RV1b = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingRepeat", Nothing)
            Dim RV1c = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingTimesRepeat", Nothing)
            Dim RV1d = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\1\", "RingVol", Nothing)

            NW.IP.Text = RV2
            If RV13 = "0" Then
                NW.PB.Image = My.Resources.Clock
            ElseIf RV13 = "1" Then
                NW.PB.Image = My.Resources.Information
            Else
                NW.PB.ImageLocation = RV13_
            End If
            If RV14 = "1" Then
                NW.TL.Text = RV14_
            Else
                NW.TL.Text = "Alarm rich the end."
            End If
            If RV11 = "1" Then
                NW.B1.Text = RV11_
            Else
                NW.B1.Text = "Stop"
            End If
            If RV12 = "1" Then
                If RV12_ = "1" Then
                    NW.B2.Text = RV12__
                Else
                    NW.B2.Text = "I'll put it off later."
                    NW.B3.Text = "Stop"
                End If
                NW.MODE = 2
            Else
                NW.MODE = 1
            End If
            If RV15 = "0" Then
                NW.Location = New Point(My.Computer.Screen.WorkingArea.Width - NW.Width, My.Computer.Screen.WorkingArea.Height - NW.Height)
            ElseIf RV15 = "1" Then
                NW.Location = New Point(My.Computer.Screen.WorkingArea.Width - NW.Width, My.Computer.Screen.WorkingArea.Height / 2 - NW.Height / 2)
            ElseIf RV15 = "2" Then
                NW.Location = New Point(My.Computer.Screen.WorkingArea.Width - NW.Width, 0)
            ElseIf RV15 = "3" Then
                NW.Location = New Point(My.Computer.Screen.WorkingArea.Width / 2 - NW.Width / 2, 0)
            ElseIf RV15 = "4" Then
                NW.Location = New Point(0, 0)
            ElseIf RV15 = "5" Then
                NW.Location = New Point(0, My.Computer.Screen.WorkingArea.Height / 2 - NW.Height / 2)
            ElseIf RV15 = "6" Then
                NW.Location = New Point(0, My.Computer.Screen.WorkingArea.Height - NW.Height)
            ElseIf RV15 = "7" Then
                NW.Location = New Point(My.Computer.Screen.WorkingArea.Width / 2 - NW.Width / 2, My.Computer.Screen.WorkingArea.Height - NW.Height)
            ElseIf RV15 = "8" Then
                NW.Location = New Point(My.Computer.Screen.WorkingArea.Width / 2 - NW.Width / 2, My.Computer.Screen.WorkingArea.Height / 2 - NW.Height / 2)
            ElseIf RV15 = "9" Then
                NW.Location = New Point(RV15X, RV15Y)
            End If

            Select Case NW.ShowDialog
                Case Windows.Forms.DialogResult.OK
                    alarm.Stop()
                    NW.Close()
                Case Windows.Forms.DialogResult.No
                    alarm.Stop()
                    NW.Close()
                Case Windows.Forms.DialogResult.Yes
1:
                    Dim mymess As String
                    Dim myanswer As String
                    mymess = Microsoft.VisualBasic.InputBox("Type in how many minutes this alarm will pop-up again:" & Environment.NewLine & "''Type a number in minutes, so 1 hour is 60 mins etc.''", "When we will attention you again?", "")
                    myanswer = mymess
                    If Integer.TryParse(myanswer, 0) Then
                        If myanswer < 0 Then
                            MessageBox.Show("U can't type numbers lower than 0.. Please try it again.")
                            GoTo 1
                        Else
                            Dim FM As Integer = RV4 + myanswer
                            If FM > 59 Then
                                Dim i As Integer = 0
                                i = FM / 59
                                For index As Integer = 1 To i
                                    FM -= 59
                                Next
                                Dim FH As Integer = RV3 + i
                                If FH > 23 Then
                                    ADTA.Value = RV2
                                    Dim i2 As Integer = 0
                                    i2 = FH / 23
                                    For index As Integer = 1 To i2
                                        FH -= 23
                                    Next
                                    ADTA.Value = ADTA.Value.AddDays(i2)
                                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\", "Hour", FH, Microsoft.Win32.RegistryValueKind.String)
                                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\", "Date", ADTA.Value.ToString("dd. MM. yyyy"), Microsoft.Win32.RegistryValueKind.String)
                                Else
                                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\", "Hour", FH, Microsoft.Win32.RegistryValueKind.String)
                                    My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\", "Minute", FM, Microsoft.Win32.RegistryValueKind.String)
                                End If
                            Else
                                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\", "Minute", FM, Microsoft.Win32.RegistryValueKind.String)
                            End If
                        End If
                    Else
                        GoTo 1
                    End If
            End Select

        ElseIf ComboBox5.SelectedIndex = 1 Then
            Dim RV21 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingMode", Nothing)
            Dim RV22 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingALIST", Nothing)
            Dim RV23 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingPath", Nothing)
            Dim RV24 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingSYS", Nothing)
            Dim RV25 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingRepeat", Nothing)
            Dim RV26 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingTimesRepeat", Nothing)
            Dim RV27 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\2\", "RingVol", Nothing)

            If RV21 = "0" Then
                If RV24 = "0" Then
                    Media.SystemSounds.Asterisk.Play()
                ElseIf RV24 = "1" Then
                    Media.SystemSounds.Beep.Play()
                ElseIf RV24 = "2" Then
                    Media.SystemSounds.Exclamation.Play()
                ElseIf RV24 = "3" Then
                    Media.SystemSounds.Hand.Play()
                ElseIf RV24 = "4" Then
                    Media.SystemSounds.Question.Play()
                End If
            ElseIf RV21 = "1" Then
                AlCol.Items.Clear()
                Dim i As String
                If My.Computer.FileSystem.DirectoryExists("C:\Windows\Media\Alarms\") Then
                    For Each i In My.Computer.FileSystem.GetFiles("C:\Windows\Media\Alarms\")
                        AlCol.Items.Add(i.Substring(i.LastIndexOf("\") + 1))
                    Next
                End If
                Dim alrpath As String = "C:\Windows\Media\Alarms\" & AlCol.Items.Item(RV22)
                alarm = New Media.SoundPlayer(alrpath)
                If RV25 = "1" Then
                    alarm.PlayLooping()
                Else
                    'if RingTimesRepeat
                    alarm.Play()
                End If
            ElseIf RV21 = "2" Then
                If IO.File.Exists(RV23) = True Then
                    Dim alrpath As String = RV23
                    alarm = New Media.SoundPlayer(alrpath)
                    If RV25 = "1" Then
                        alarm.PlayLooping()
                    Else
                        'if RingTimesRepeat
                        alarm.Play()
                    End If
                End If
            End If
            Button17.Text = "Stop"
            ComboBox5.Enabled = False
            Button19.Enabled = False
            Button20.Enabled = False
            AlarmList.Enabled = False
            Button18.Enabled = False
            CheckBox15.Enabled = False
            mi = False
            GroupBox9.Enabled = False
            A_H.Enabled = False
            A_M.Enabled = False
            A_S.Enabled = False
        ElseIf ComboBox5.SelectedIndex = 2 Then
            Dim RV31 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\", "MODE", Nothing)
            Dim RV32 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\", "Path", Nothing)

            If IO.File.Exists(RV32) = True Then
                WP = RV32
            End If
            If RV31 = "0" Then
                Dim RV321 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\2\", "AppWinStyle", Nothing)
                Dim RV322 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\2\", "Wait", Nothing)
                Dim RV323 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\2\", "Timeout", Nothing)

                If RV321 = "0" Then
                    If RV322 = "1" Then
                        If RV323 >= -1 Then
                            Shell(WP, AppWinStyle.NormalFocus, True, RV323)
                        Else
                            Shell(WP, AppWinStyle.NormalFocus, True, -1)
                        End If
                    Else
                        If RV323 >= -1 Then
                            Shell(WP, AppWinStyle.NormalFocus, False, RV323)
                        Else
                            Shell(WP, AppWinStyle.NormalFocus, False, -1)
                        End If
                    End If
                ElseIf RV321 = "1" Then
                    If RV322 = "1" Then
                        If RV323 >= -1 Then
                            Shell(WP, AppWinStyle.MinimizedFocus, True, RV323)
                        Else
                            Shell(WP, AppWinStyle.MinimizedFocus, True, -1)
                        End If
                    Else
                        If RV323 >= -1 Then
                            Shell(WP, AppWinStyle.MinimizedFocus, False, RV323)
                        Else
                            Shell(WP, AppWinStyle.MinimizedFocus, False, -1)
                        End If
                    End If
                ElseIf RV321 = "2" Then
                    If RV322 = "1" Then
                        If RV323 >= -1 Then
                            Shell(WP, AppWinStyle.MaximizedFocus, True, RV323)
                        Else
                            Shell(WP, AppWinStyle.MaximizedFocus, True, -1)
                        End If
                    Else
                        If RV323 >= -1 Then
                            Shell(WP, AppWinStyle.MaximizedFocus, False, RV323)
                        Else
                            Shell(WP, AppWinStyle.MaximizedFocus, False, -1)
                        End If
                    End If
                ElseIf RV321 = "3" Then
                    If RV322 = "1" Then
                        If RV323 >= -1 Then
                            Shell(WP, AppWinStyle.NormalNoFocus, True, RV323)
                        Else
                            Shell(WP, AppWinStyle.NormalNoFocus, True, -1)
                        End If
                    Else
                        If RV323 >= -1 Then
                            Shell(WP, AppWinStyle.NormalNoFocus, False, RV323)
                        Else
                            Shell(WP, AppWinStyle.NormalNoFocus, False, -1)
                        End If
                    End If
                ElseIf RV321 = "4" Then
                    If RV322 = "1" Then
                        If RV323 >= -1 Then
                            Shell(WP, AppWinStyle.MinimizedNoFocus, True, RV323)
                        Else
                            Shell(WP, AppWinStyle.MinimizedNoFocus, True, -1)
                        End If
                    Else
                        If RV323 >= -1 Then
                            Shell(WP, AppWinStyle.MinimizedNoFocus, False, RV323)
                        Else
                            Shell(WP, AppWinStyle.MinimizedNoFocus, False, -1)
                        End If
                    End If
                ElseIf RV321 = "5" Then
                    If RV322 = "1" Then
                        If RV323 >= -1 Then
                            Shell(WP, AppWinStyle.Hide, True, RV323)
                        Else
                            Shell(WP, AppWinStyle.Hide, True, -1)
                        End If
                    Else
                        If RV323 >= -1 Then
                            Shell(WP, AppWinStyle.Hide, False, RV323)
                        Else
                            Shell(WP, AppWinStyle.Hide, False, -1)
                        End If
                    End If
                End If
            Else
                Dim RV311 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "AppWinStyle", Nothing)
                Dim RV312 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "Arguments", Nothing)
                Dim RV313 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "CreateNoWindow", Nothing)
                Dim RV314 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "EnableCustomWorkingDir", Nothing)
                Dim RV315 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "CustomWorkDir", Nothing)
                Dim RV316 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "EnableDomain", Nothing)
                Dim RV316_ = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "Domain", Nothing)
                Dim RV317 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "RunAs", Nothing)
                Dim RV318 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "Username", Nothing)
                Dim RV319 = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "LoadUserProfile", Nothing)
                Dim RV31a = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "UseShellExecute", Nothing)
                Dim RV31b = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "ErrorDialog", Nothing)
                Dim RV31c = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "ErrorDialogParentHandle", Nothing)
                Dim RV31d = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "RSError", Nothing)
                Dim RV31e = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "RSInput", Nothing)
                Dim RV31f = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\Mode\3\1\", "RSOutput", Nothing)

                Dim start As New ProcessStartInfo(WP)
                start.Arguments = RV312
                start.CreateNoWindow = RV313
                start.UseShellExecute = RV31a
                start.LoadUserProfile = RV319
                start.ErrorDialog = RV31b
                start.ErrorDialogParentHandle = RV31c
                start.RedirectStandardError = RV31d
                start.RedirectStandardInput = RV31e
                start.RedirectStandardOutput = RV31f
                If RV314 = "1" Then
                    If RV315 = String.Empty Then
                    Else
                        start.Domain = RV315
                    End If
                End If
                If RV316 = "1" Then
                    If RV316_ = String.Empty Then
                    Else
                        start.Domain = RV316_
                    End If
                End If
                If RV317 = "1" Then
                    On Error GoTo 2
                    start.UserName = RV318
                    start.Password = ToSecureString(APD.MaskedTextBox1.Text)
                End If
2:
                If RV311 = "0" Then
                    start.WindowStyle = ProcessWindowStyle.Normal
                ElseIf RV311 = "1" Then
                    start.WindowStyle = ProcessWindowStyle.Minimized
                ElseIf RV311 = "2" Then
                    start.WindowStyle = ProcessWindowStyle.Maximized
                ElseIf RV311 = "3" Then
                    start.WindowStyle = ProcessWindowStyle.Hidden
                End If
                Dim proc As New Process()
                proc.StartInfo = start
                proc.Start()
            End If
        End If
    End Sub
    Public Shared Function ToSecureString(ByVal source As String) As Security.SecureString
        Dim result = New Security.SecureString()
        For Each c As Char In source
            result.AppendChar(c)
        Next
        Return result
    End Function
    Dim AFixH As String = "00"
    Dim AFixM As String = "00"
    Dim AFixS As String = "00"
    Private Sub NumericUpDown8_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles A_H.ValueChanged
        'Hours
        If A_H.Value = 0 Then
            AFixH = "00"
        ElseIf A_H.Value = 1 Then
            AFixH = "01"
        ElseIf A_H.Value = 2 Then
            AFixH = "02"
        ElseIf A_H.Value = 3 Then
            AFixH = "03"
        ElseIf A_H.Value = 4 Then
            AFixH = "04"
        ElseIf A_H.Value = 5 Then
            AFixH = "05"
        ElseIf A_H.Value = 6 Then
            AFixH = "06"
        ElseIf A_H.Value = 7 Then
            AFixH = "07"
        ElseIf A_H.Value = 8 Then
            AFixH = "08"
        ElseIf A_H.Value = 9 Then
            AFixH = "09"
        Else
            AFixH = A_H.Value
        End If
    End Sub

    Private Sub NumericUpDown7_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles A_M.ValueChanged
        'Minutes
        If A_M.Value = 0 Then
            AFixM = "00"
        ElseIf A_M.Value = 1 Then
            AFixM = "01"
        ElseIf A_M.Value = 2 Then
            AFixM = "02"
        ElseIf A_M.Value = 3 Then
            AFixM = "03"
        ElseIf A_M.Value = 4 Then
            AFixM = "04"
        ElseIf A_M.Value = 5 Then
            AFixM = "05"
        ElseIf A_M.Value = 6 Then
            AFixM = "06"
        ElseIf A_M.Value = 7 Then
            AFixM = "07"
        ElseIf A_M.Value = 8 Then
            AFixM = "08"
        ElseIf A_M.Value = 9 Then
            AFixM = "09"
        Else
            AFixM = A_M.Value
        End If
    End Sub

    Private Sub NumericUpDown6_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles A_S.ValueChanged
        'Seconds
        If A_S.Value = 0 Then
            AFixS = "00"
        ElseIf A_S.Value = 1 Then
            AFixS = "01"
        ElseIf A_S.Value = 2 Then
            AFixS = "02"
        ElseIf A_S.Value = 3 Then
            AFixS = "03"
        ElseIf A_S.Value = 4 Then
            AFixS = "04"
        ElseIf A_S.Value = 5 Then
            AFixS = "05"
        ElseIf A_S.Value = 6 Then
            AFixS = "06"
        ElseIf A_S.Value = 7 Then
            AFixS = "07"
        ElseIf A_S.Value = 8 Then
            AFixS = "08"
        ElseIf A_S.Value = 9 Then
            AFixS = "09"
        Else
            AFixS = A_S.Value
        End If
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click, ETI.Click
        Button17.Text = "Save"
        AlarmList.Enabled = False
        CheckBox25.Enabled = False
        Button18.Enabled = False
        CheckBox17.Enabled = True
        GroupBox9.Enabled = True
        mi = True
        A_H.Enabled = True
        A_M.Enabled = True
        A_S.Enabled = True
        TextBox1.Enabled = True
        If CheckBox17.Checked = True Then
            Label40.Visible = True
            DateTimePicker1.Visible = False
            DateTimePicker1.Enabled = False
        Else
            Label40.Visible = False
            DateTimePicker1.Visible = True
            DateTimePicker1.Enabled = True
        End If
        Button20.Enabled = False
    End Sub
    Dim mi As Boolean = False
    Private Sub Label40_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label40.TextChanged
        If Label40.Text = "Nothing" Then
            Label40.Text = "''By repeating cannot be Date modified.''"
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        TimeToCopy.Text = AlarmList.SelectedItem
        TimeToCopy.SelectAll()
        TimeToCopy.Copy()
        ToolTip1.Show("Successfully copied!", Me, Control.MousePosition.X - Me.Location.X, Control.MousePosition.Y - Me.Location.Y, 4000)
    End Sub

    Private Sub ALcm_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ALcm.Opening
        If AlarmList.SelectedIndex = -1 Then
            NATI.Visible = True
            ToolStripMenuItem1.Visible = False
            ToolStripMenuItem2.Visible = False
            ToolStripMenuItem3.Visible = False
            ETI.Visible = False
            ToolStripSeparator7.Visible = False
            ToolStripSeparator8.Visible = False
            ToolStripSeparator6.Visible = False
        Else
            NATI.Visible = False
            ToolStripMenuItem1.Visible = True
            ToolStripMenuItem2.Visible = True
            ToolStripMenuItem3.Visible = True
            ETI.Visible = True
            ToolStripSeparator7.Visible = False
            ToolStripSeparator8.Visible = True
            ToolStripSeparator6.Visible = True
        End If
    End Sub

    Private Sub ToolStripMenuItem2_DropDownOpening(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.DropDownOpening
        If My.Computer.Registry.CurrentUser.OpenSubKey("ALARMS\" & AlarmList.SelectedItem & "\", False) Is Nothing Then
            NRTI.Enabled = False
            NRTI.Text = "(Cannot find the name.)"
        Else
            NRTI.Enabled = True
            NRTI.Text = "HKEY_CURRENT_USER\ALARMS\" & AlarmList.SelectedItem & "\"
        End If
    End Sub

    Private Sub NRTI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NRTI.Click
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit\", "LastKey", "Computer\" & NRTI.Text, Microsoft.Win32.RegistryValueKind.String)
        CheckBox1.Checked = True
        System.Diagnostics.Process.Start("C:\Windows\regedit.exe")
    End Sub

    Private Sub NATI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NATI.Click
        CheckBox25.Checked = True
    End Sub
End Class
