Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices

Public Class CapacityManager
    <StructLayout(LayoutKind.Sequential)>
    Private Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure SYSTEM_POWER_STATUS
        Public ACLineStatus As Byte
        Public BatteryFlag As Byte
        Public BatteryLifePercent As Byte
        Public SystemStatusFlag As Byte
        Public BatteryLifeTime As Integer
        Public BatteryFullLifeTime As Integer
    End Structure

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function GetSystemPowerStatus(ByRef lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Boolean
    End Function

    <DllImport("batmeter.dll", SetLastError:=True)>
    Public Shared Function IsBatteryBad() As Boolean
    End Function

    <DllImport("batmeter.dll", SetLastError:=True)>
    Public Shared Function IsBatteryLevelLow() As Boolean
    End Function

    <DllImport("batmeter.dll", SetLastError:=True)>
    Public Shared Function IsBatteryLevelReserve() As Boolean
    End Function

    <DllImport("batmeter.dll", SetLastError:=True)>
    Public Shared Function IsBatteryLevelCritical() As Boolean
    End Function

    Private Shared Function GetStatus() As SYSTEM_POWER_STATUS
        Dim status As New SYSTEM_POWER_STATUS()
        GetSystemPowerStatus(status)
        Return status
    End Function

    Public Shared Function HasBattery() As Boolean
        Dim status = GetStatus()
        Return (status.BatteryFlag <> 128) AndAlso (status.BatteryLifePercent <> 255)
    End Function

    Public Shared Function IsCharging() As Boolean
        Dim status = GetStatus()
        If Not HasBattery() Then Return False
        Return (status.BatteryFlag And 8) = 8
    End Function

    Public Shared Function GetBatteryPercentage() As Integer
        Dim status = GetStatus()
        If Not HasBattery() Then Return 0
        If status.BatteryLifePercent <= 100 Then
            Return Convert.ToInt32(status.BatteryLifePercent)
        End If
        Return 0
    End Function

    Public Shared Function GetPowerIcon() As Bitmap
        Dim status = GetStatus()
        Dim acConnected As Boolean = (status.ACLineStatus = 1)
        Dim batteryPresent As Boolean = HasBattery()

        Dim mainBateryIcon As Bitmap = My.Resources.Battery
        Dim isCharging As Boolean = False

        If acConnected AndAlso Not batteryPresent Then
            Return My.Resources.ACPower
        End If

        If acConnected AndAlso batteryPresent Then
            Dim pct As Integer = GetBatteryPercentage()

            If status.BatteryLifePercent = 255 Then
                Return GenerateBatteryIcon(pct, mainBateryIcon, My.Resources.Electricity)
            End If

            isCharging = True
        End If

        If Not acConnected AndAlso batteryPresent Then
            Dim pct As Integer = GetBatteryPercentage()

            Return GenerateBatteryIcon(pct, mainBateryIcon, If(isCharging, My.Resources.Electricity, If(IsBatteryBad(), My.Resources.Tray_Warning, Nothing)))
        End If

        Return GenerateBatteryIcon(0, mainBateryIcon, My.Resources.Tray_Error)
    End Function

    Public Shared Function GenerateBatteryIcon(ByVal percentage As Integer, ByVal template As Image, Optional ByVal overlayImage As Image = Nothing) As Bitmap
        Dim bmp As New Bitmap(template.Width, template.Height)

        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = SmoothingMode.AntiAlias

            Dim fillPercent As Integer = Math.Max(0, Math.Min(percentage, 100))

            Dim innerX As Integer = 9
            Dim innerY As Integer = 6
            Dim innerWidth As Integer = 12
            Dim innerMaxHeight As Integer = 21

            Dim fillHeight As Integer = Convert.ToInt32((fillPercent / 100.0) * innerMaxHeight)
            Dim fillY As Integer = innerY + (innerMaxHeight - fillHeight)

            If fillHeight > 0 Then
                Dim color1 As Color
                Dim color2 As Color

                If fillPercent > 50 Then
                    Dim factor As Double = (fillPercent - 50) / 50.0
                    color1 = InterpolateColor(Color.Yellow, Color.LimeGreen, factor)
                    color2 = InterpolateColor(Color.Orange, Color.Green, factor)
                Else
                    Dim factor As Double = fillPercent / 50.0
                    color1 = InterpolateColor(Color.Red, Color.Yellow, factor)
                    color2 = InterpolateColor(Color.DarkRed, Color.Orange, factor)
                End If

                Using brush As New LinearGradientBrush(New Rectangle(innerX, fillY, innerWidth, fillHeight), color1, color2, LinearGradientMode.Vertical)
                    g.FillRectangle(brush, innerX, fillY, innerWidth, fillHeight)
                End Using
            End If

            g.DrawImage(template, 0, 0, template.Width, template.Height)

            If overlayImage IsNot Nothing Then
                g.DrawImage(overlayImage, 0, 0, template.Width, template.Height)
            End If
        End Using

        Return bmp
    End Function

    Private Shared Function InterpolateColor(ByVal c1 As Color, ByVal c2 As Color, ByVal factor As Double) As Color
        Dim r As Integer = Convert.ToInt32(CInt(c1.R) + (CInt(c2.R) - CInt(c1.R)) * factor)
        Dim g As Integer = Convert.ToInt32(CInt(c1.G) + (CInt(c2.G) - CInt(c1.G)) * factor)
        Dim b As Integer = Convert.ToInt32(CInt(c1.B) + (CInt(c2.B) - CInt(c1.B)) * factor)

        Return Color.FromArgb(
        Math.Max(0, Math.Min(255, r)),
        Math.Max(0, Math.Min(255, g)),
        Math.Max(0, Math.Min(255, b))
    )
    End Function
End Class
