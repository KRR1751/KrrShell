Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports NAudio.CoreAudioApi

Public Class VolumeControl

    <DllImport("dwmapi.dll")>
    Public Shared Function DwmSetWindowAttribute(hwnd As IntPtr, attr As Integer, ByRef attrValue As Integer, attrSize As Integer) As Integer
    End Function

    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)

        Dim value As Integer = 1
        DwmSetWindowAttribute(Me.Handle, 2, value, Marshal.SizeOf(value))
    End Sub

    Public Sub New()
        InitializeComponent()

        Me.FormBorderStyle = FormBorderStyle.None
    End Sub

    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.Style = cp.Style Or &H80000   ' WS_SYSMENU '
            cp.Style = cp.Style Or &H20000   ' WS_MINIMIZEBOX '

            cp.Style = cp.Style Or &H800000  ' WS_BORDER '
            Return cp
        End Get
    End Property

    ' VOLUME MAGIC HERE '

    Private DeviceEnumerator As New MMDeviceEnumerator()
    Private DefaultDevice As MMDevice = DeviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia)

    Public Function IsSysMuted() As Boolean
        Return DefaultDevice.AudioEndpointVolume.Mute
    End Function

    Public Function GetCurrentVolume() As Single
        Return DefaultDevice.AudioEndpointVolume.MasterVolumeLevelScalar
    End Function

    Public Sub SetVolume(volumeLevel As Single)
        If volumeLevel < 0.0F Then volumeLevel = 0.0F
        If volumeLevel > 1.0F Then volumeLevel = 1.0F

        DefaultDevice.AudioEndpointVolume.MasterVolumeLevelScalar = volumeLevel
    End Sub

    Public Sub ToggleMute(onoff As Boolean)
        Select Case onoff
            Case True
                DefaultDevice.AudioEndpointVolume.Mute = True
            Case False
                DefaultDevice.AudioEndpointVolume.Mute = False
        End Select
    End Sub

    Public IsMuted As Boolean = False

    ' OTHER STUFF '

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        SetVolume(TrackBar1.Value / 100)

        Label2.Text = TrackBar1.Value & "%"
        UpdateIcons()
    End Sub
    Dim liveBar = Application.OpenForms.OfType(Of AppBar)().FirstOrDefault()

    Public Sub UpdateIcons()

        If Not IsSysMuted() Then

            Dim curVol As Integer = GetCurrentVolume() * 100

            If liveBar IsNot Nothing Then
                liveBar.Invoke(Sub()
                                   If curVol = 0 Then
                                       liveBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeZero

                                   ElseIf curVol < 50 AndAlso Not curVol = 0 Then
                                       liveBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeBelow50

                                   ElseIf curVol >= 50 AndAlso Not curVol = 100 Then
                                       liveBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeAbove50.ToBitmap

                                   ElseIf curVol = 100 Then
                                       liveBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeFull

                                   End If
                               End Sub)
            End If
        Else
            If liveBar IsNot Nothing Then
                liveBar.Invoke(Sub()
                                   liveBar.ToolStripButton1.BackgroundImage = My.Resources.VolumeMute
                               End Sub)
            End If
        End If
    End Sub

    Private Sub VolumeControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TrackBar1.Value = GetCurrentVolume() * 100

        IsMuted = DefaultDevice.AudioEndpointVolume.Mute
        CheckBox1.Checked = IsMuted

        Label2.Text = TrackBar1.Value & "%"

        ' Showing the form
        Me.Location = New Point(MousePosition.X - Me.Width / 2, SystemInformation.WorkingArea.Height - Me.Height)
        Me.BringToFront()
    End Sub

    Private Sub VolumeControl_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Volume", "LastVol", TrackBar1.Value, Microsoft.Win32.RegistryValueKind.DWord)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Volume", "IsMuted", CheckBox1.Checked, Microsoft.Win32.RegistryValueKind.DWord)

        Me.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        ToggleMute(CheckBox1.Checked)
        IsMuted = CheckBox1.Checked

        UpdateIcons()
    End Sub

    Private Sub VolumeControl_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If liveBar IsNot Nothing Then
            Try
                liveBar.ToolStripButton1.Checked = False
            Catch ex As Exception
                AppBar.ToolStripButton1.Checked = False
            End Try
        End If
    End Sub
End Class