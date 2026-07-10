Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text

Public Class PowerControl

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

    ' POWER MAGIC HERE '

    Private Sub PowerControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim liveBar = Application.OpenForms.OfType(Of AppBar)().FirstOrDefault()

        PowerSchemeManager.GenerateRadioButtons(pnl_PWRP)

        pb_Icon.Image = liveBar.ToolStripButton4.BackgroundImage
        lbl_Info.Text = liveBar.ToolStripButton4.Text
        ' Showing the form

        Select Case AppBar.currentDockPos
            Case AppBar.DockPosition.Top
                Me.Location = New Point(MousePosition.X - Me.Width / 2, AppBar.Height)
            Case AppBar.DockPosition.Left
                Me.Location = New Point(AppBar.Width, MousePosition.Y - Me.Height / 2)
            Case AppBar.DockPosition.Right
                Me.Location = New Point(SystemInformation.PrimaryMonitorSize.Width - Me.Width, MousePosition.Y - Me.Height / 2)
            Case AppBar.DockPosition.Bottom
                Me.Location = New Point(MousePosition.X - Me.Width / 2, SystemInformation.WorkingArea.Height - Me.Height)
        End Select

        Me.BringToFront()
    End Sub

    Private Sub PowerControl_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Dim liveBar = Application.OpenForms.OfType(Of AppBar)().FirstOrDefault()

        If liveBar IsNot Nothing Then
            Try
                liveBar.ToolStripButton4.Checked = False
            Catch ex As Exception
                AppBar.ToolStripButton4.Checked = False
            End Try
        End If
    End Sub

    Public Sub PowerControl_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown, llbl_1.KeyDown, llbl_2.KeyDown, pnl_PWRP.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub PowerControl_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate
        Me.Close()
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles llbl_2.LinkClicked
        On Error Resume Next

        Dim decodedText As String = AppBar.EncodeFilePath(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\CustomPaths", "Settings", "control"))
        Process.Start(decodedText, "/name Microsoft.PowerOptions")
    End Sub
End Class

Public Class PowerSchemeManager

    <DllImport("powrprof.dll", SetLastError:=True)>
    Private Shared Function PowerGetActiveScheme(
        ByVal UserRootPowerKey As IntPtr,
        ByRef ActivePolicyGuid As IntPtr) As UInteger
    End Function

    <DllImport("powrprof.dll", SetLastError:=True)>
    Private Shared Function PowerSetActiveScheme(
        ByVal UserRootPowerKey As IntPtr,
        ByRef SchemeGuid As Guid) As UInteger
    End Function

    <DllImport("powrprof.dll", SetLastError:=True)>
    Private Shared Function PowerEnumerate(
        ByVal RootPowerKey As IntPtr,
        ByVal SchemeGuid As IntPtr,
        ByVal SubGroupOfPowerSettingsGuid As IntPtr,
        ByVal AccessFlags As UInteger,
        ByVal Index As UInteger,
        ByVal Buffer As IntPtr,
        ByRef BufferSize As UInteger) As UInteger
    End Function

    <DllImport("powrprof.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Private Shared Function PowerReadFriendlyName(
        ByVal RootPowerKey As IntPtr,
        ByVal SchemeGuid As Guid,
        ByVal SubGroupOfPowerSettingsGuid As IntPtr,
        ByVal PowerSettingGuid As IntPtr,
        ByVal Buffer As StringBuilder,
        ByRef BufferSize As UInteger) As UInteger
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function LocalFree(ByVal hMem As IntPtr) As IntPtr
    End Function

    Public Shared Function GetActiveSchemeGuid() As Guid
        Dim guidPtr As IntPtr = IntPtr.Zero
        Dim result As UInteger = PowerGetActiveScheme(IntPtr.Zero, guidPtr)

        If result = 0 AndAlso guidPtr <> IntPtr.Zero Then
            Dim activeGuid As Guid = Marshal.PtrToStructure(Of Guid)(guidPtr)
            LocalFree(guidPtr)
            Return activeGuid
        End If

        Return Guid.Empty
    End Function

    Public Shared Function SetActiveScheme(ByVal schemeGuid As Guid) As Boolean
        Dim result As UInteger = PowerSetActiveScheme(IntPtr.Zero, schemeGuid)
        Return result = 0
    End Function

    Public Shared Sub GenerateRadioButtons(ByVal targetPanel As Panel)
        targetPanel.Controls.Clear()

        Dim activeGuid As Guid = GetActiveSchemeGuid()
        Dim index As UInteger = 0
        Dim bufferSize As UInteger = Convert.ToUInt32(Marshal.SizeOf(GetType(Guid)))
        Dim buffer As IntPtr = Marshal.AllocHGlobal(Convert.ToInt32(bufferSize))

        Dim currentTop As Integer = 10

        Try
            ' ACCESS_SCHEME = 16 (0x10)
            While PowerEnumerate(IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, 16, index, buffer, bufferSize) = 0
                Dim schemeGuid As Guid = Marshal.PtrToStructure(Of Guid)(buffer)

                Dim nameSize As UInteger = 256
                Dim nameBuffer As New StringBuilder(256)
                Dim nameResult As UInteger = PowerReadFriendlyName(IntPtr.Zero, schemeGuid, IntPtr.Zero, IntPtr.Zero, nameBuffer, nameSize)

                Dim friendlyName As String = If(nameResult = 0, nameBuffer.ToString(), "Unknown Scheme")

                Dim rb As New RadioButton()
                rb.Text = friendlyName
                rb.Tag = schemeGuid
                rb.Top = currentTop
                rb.Left = 10
                rb.AutoSize = True

                If schemeGuid = activeGuid Then
                    rb.Checked = True
                End If

                AddHandler rb.CheckedChanged, AddressOf RadioButton_CheckedChanged
                AddHandler rb.KeyDown, AddressOf PowerControl.PowerControl_KeyDown

                targetPanel.Controls.Add(rb)
                currentTop += 25
                index += 1UI
            End While
        Finally
            Marshal.FreeHGlobal(buffer)
        End Try
    End Sub

    Private Shared Sub RadioButton_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim rb As RadioButton = TryCast(sender, RadioButton)
        If rb IsNot Nothing AndAlso rb.Checked Then
            Dim schemeGuid As Guid = CType(rb.Tag, Guid)
            SetActiveScheme(schemeGuid)
        End If
    End Sub
End Class