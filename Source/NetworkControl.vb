Imports System.Net.NetworkInformation
Imports System.Runtime.InteropServices
Imports System.Security
Imports System.Text
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports KrrShell.NetworkControl
Imports KrrShell.WlanNative

Public Class NetworkControl
    Public Function GetPrimaryConnectionMethod() As String
        Dim interfaces As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()

        For Each ni As NetworkInterface In interfaces
            If ni.OperationalStatus = OperationalStatus.Up AndAlso ni.NetworkInterfaceType <> NetworkInterfaceType.Loopback Then
                If ni.NetworkInterfaceType = NetworkInterfaceType.Ethernet Then
                    Return "LAN"
                ElseIf ni.NetworkInterfaceType = NetworkInterfaceType.Wireless80211 Then
                    Return "Wi-Fi"
                End If
            End If
        Next

        Return "Disconnected"
    End Function

    Public Function ConnectToWifi(ssid As String, securePassword As SecureString) As Boolean
        Dim clientHandle As IntPtr = IntPtr.Zero
        Dim negotiatedVersion As UInteger = 0

        If WlanNative.WlanOpenHandle(2, IntPtr.Zero, negotiatedVersion, clientHandle) <> 0 Then Return False

        Dim password As String = ""
        If securePassword IsNot Nothing AndAlso securePassword.Length > 0 Then
            Dim unmanagedString As IntPtr = IntPtr.Zero
            Try
                unmanagedString = Marshal.SecureStringToGlobalAllocUnicode(securePassword)
                password = Marshal.PtrToStringUni(unmanagedString)
            Finally
                Marshal.ZeroFreeGlobalAllocUnicode(unmanagedString)
            End Try
        End If

        Try
            Dim interfaceListPtr As IntPtr = IntPtr.Zero
            If WlanNative.WlanEnumInterfaces(clientHandle, IntPtr.Zero, interfaceListPtr) <> 0 Then Return False

            Dim listHeader As WlanNative.WLAN_INTERFACE_INFO_LIST = Marshal.PtrToStructure(Of WlanNative.WLAN_INTERFACE_INFO_LIST)(interfaceListPtr)
            Dim interfaceGuid As Guid = listHeader.InterfaceInfo.InterfaceGuid
            WlanNative.WlanFreeMemory(interfaceListPtr)

            Dim ssidBytes As Byte() = Encoding.UTF8.GetBytes(ssid)
            Dim ssidHex As String = BitConverter.ToString(ssidBytes).Replace("-", "")

            Dim securityXml As String = ""
            If String.IsNullOrEmpty(password) Then
                securityXml = "<security><authEncryption><authentication>open</authentication><encryption>none</encryption><useOneX>false</useOneX></authEncryption></security>"
            Else
                securityXml = $"<security><authEncryption><authentication>WPA2PSK</authentication><encryption>AES</encryption><useOneX>false</useOneX></authEncryption><sharedKey><keyType>passPhrase</keyType><protected>false</protected><keyMaterial>{password}</keyMaterial></sharedKey></security>"
            End If

            Dim profileXml As String =
        $"<?xml version=""1.0""?>" &
        $"<WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1"">" &
            $"<name>{ssid}</name>" &
            $"<SSIDConfig><SSID><hex>{ssidHex}</hex><name>{ssid}</name></SSID></SSIDConfig>" &
            $"<connectionType>ESS</connectionType>" &
            $"<connectionMode>manual</connectionMode>" &
            $"<MSM>{securityXml}</MSM>" &
        $"</WLANProfile>"

            Dim reasonCode As UInteger = 0
            Dim setProfileResult As UInteger = WlanNative.WlanSetProfile(clientHandle, interfaceGuid, 0, profileXml, Nothing, True, IntPtr.Zero, reasonCode)

            If Not String.IsNullOrEmpty(password) Then
                password = New String("*"c, password.Length)
            End If
            password = Nothing
            profileXml = Nothing

            If setProfileResult <> 0 Then Return False

            Dim ssidStructure As New WlanNative.DOT11_SSID()
            ssidStructure.ucSSID = New Byte(31) {}
            ssidStructure.uSSIDLength = Convert.ToUInt32(ssidBytes.Length)
            Array.Copy(ssidBytes, ssidStructure.ucSSID, ssidBytes.Length)

            Dim pSsid As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(ssidStructure))
            Marshal.StructureToPtr(ssidStructure, pSsid, False)

            WlanNative.WlanDisconnect(clientHandle, interfaceGuid, IntPtr.Zero)

            Dim connectResult As UInteger = 1
            Try
                Dim cParams As New WlanNative.WLAN_CONNECTION_PARAMETERS()
                cParams.wlanConnectionMode = 0
                cParams.strProfile = ssid
                cParams.pDot11Ssid = pSsid
                cParams.pDesiredBssidList = IntPtr.Zero
                cParams.dot11BssType = 1
                cParams.dwFlags = 0

                connectResult = WlanNative.WlanConnect(clientHandle, interfaceGuid, cParams, IntPtr.Zero)
            Finally
                Marshal.FreeHGlobal(pSsid)
            End Try

            Return (connectResult = 0)
        Finally
            WlanNative.WlanCloseHandle(clientHandle, IntPtr.Zero)
        End Try
    End Function

    Public Function ConnectToExistingProfile(ssid As String) As Boolean
        Dim clientHandle As IntPtr = IntPtr.Zero
        Dim negotiatedVersion As UInteger = 0

        If WlanNative.WlanOpenHandle(2, IntPtr.Zero, negotiatedVersion, clientHandle) <> 0 Then Return False

        Try
            Dim interfaceListPtr As IntPtr = IntPtr.Zero
            If WlanNative.WlanEnumInterfaces(clientHandle, IntPtr.Zero, interfaceListPtr) <> 0 Then Return False

            Dim listHeader As WlanNative.WLAN_INTERFACE_INFO_LIST = Marshal.PtrToStructure(Of WlanNative.WLAN_INTERFACE_INFO_LIST)(interfaceListPtr)
            Dim interfaceGuid As Guid = listHeader.InterfaceInfo.InterfaceGuid
            WlanNative.WlanFreeMemory(interfaceListPtr)

            Dim ssidBytes As Byte() = Encoding.UTF8.GetBytes(ssid)
            Dim ssidStructure As New WlanNative.DOT11_SSID()
            ssidStructure.ucSSID = New Byte(31) {}
            ssidStructure.uSSIDLength = Convert.ToUInt32(ssidBytes.Length)
            Array.Copy(ssidBytes, ssidStructure.ucSSID, ssidBytes.Length)

            Dim pSsid As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(ssidStructure))
            Marshal.StructureToPtr(ssidStructure, pSsid, False)

            WlanNative.WlanDisconnect(clientHandle, interfaceGuid, IntPtr.Zero)

            Dim connectResult As UInteger = 1
            Try
                Dim cParams As New WlanNative.WLAN_CONNECTION_PARAMETERS()
                cParams.wlanConnectionMode = 0
                cParams.strProfile = ssid
                cParams.pDot11Ssid = pSsid
                cParams.pDesiredBssidList = IntPtr.Zero
                cParams.dot11BssType = 1
                cParams.dwFlags = 0

                connectResult = WlanNative.WlanConnect(clientHandle, interfaceGuid, cParams, IntPtr.Zero)
            Finally
                Marshal.FreeHGlobal(pSsid)
            End Try

            Return (connectResult = 0)
        Finally
            WlanNative.WlanCloseHandle(clientHandle, IntPtr.Zero)
        End Try
    End Function

    Public Class WifiNetwork
        Public Property Ssid As String
        Public Property SignalQuality As UInteger
        Public Property IsSecured As Boolean
        Public Property FrequencyKhz As UInteger
        Public Property Channel As Integer
        Public Property Band As String ' "2.4 GHz", "5 GHz", "6 GHz"
        Public Property Bssid As String ' MAC address AA:BB:CC...
        Public Property WifiStandard As String ' "Wi-Fi 4 (n)", "Wi-Fi 5 (ac)", "Wi-Fi 6 (ax)"...
        Public Property IsConnected As Boolean
        Public Property IsLan As Boolean = False
    End Class

    Public Shared Function GetChannelFromFrequency(freqKhz As UInteger) As Integer
        Dim mhz As Double = freqKhz / 1000

        If mhz = 2484 Then Return 14
        If mhz >= 2412 AndAlso mhz <= 2472 Then Return CInt((mhz - 2412) / 5) + 1
        If mhz >= 5180 AndAlso mhz <= 5825 Then Return CInt((mhz - 5180) / 5) + 36
        If mhz >= 5955 AndAlso mhz <= 7115 Then Return CInt((mhz - 5955) / 5) + 1

        Return 0
    End Function

    ' Převod PHY typu na lidský název standardu
    Public Shared Function GetWifiStandardName(phyType As Integer) As String
        Select Case phyType
            Case 1 : Return "802.11b"
            Case 2 : Return "802.11g"
            Case 3 : Return "Wi-Fi 1 (802.11a)"
            Case 4 : Return "Wi-Fi 4 (802.11n)"
            Case 6 : Return "Wi-Fi 5 (802.11ac)"
            Case 8 : Return "Wi-Fi 6 (802.11ax)"
            Case Else : Return $"Unknown Standard Name ({phyType})"
        End Select
    End Function

    Public Shared Function GetActiveNetworkLive() As WifiNetwork
        Try
            ' 1. CHECKING WI-FI
            Dim clientHandle As IntPtr = IntPtr.Zero
            Dim negotiatedVersion As UInteger = 0

            If WlanNative.WlanOpenHandle(2, IntPtr.Zero, negotiatedVersion, clientHandle) = 0 Then
                Try
                    Dim interfaceListPtr As IntPtr = IntPtr.Zero
                    If WlanNative.WlanEnumInterfaces(clientHandle, IntPtr.Zero, interfaceListPtr) = 0 Then
                        Dim listHeader As WlanNative.WLAN_INTERFACE_INFO_LIST = Marshal.PtrToStructure(Of WlanNative.WLAN_INTERFACE_INFO_LIST)(interfaceListPtr)
                        Dim interfaceGuid As Guid = listHeader.InterfaceInfo.InterfaceGuid
                        WlanNative.WlanFreeMemory(interfaceListPtr)

                        Dim dataSize As UInteger = 0
                        Dim dataPtr As IntPtr = IntPtr.Zero

                        ' Opcode 7
                        Dim result As UInteger = WlanNative.WlanQueryInterface(clientHandle, interfaceGuid, 7, IntPtr.Zero, dataSize, dataPtr, IntPtr.Zero)

                        If result = 0 AndAlso dataPtr <> IntPtr.Zero Then
                            Dim connAttr As WlanNative.WLAN_CONNECTION_ATTRIBUTES = Marshal.PtrToStructure(Of WlanNative.WLAN_CONNECTION_ATTRIBUTES)(dataPtr)
                            WlanNative.WlanFreeMemory(dataPtr)

                            If connAttr.isState = 1 Then
                                Dim assoc = connAttr.wlanAssociationAttributes
                                Dim ssidLength As Integer = Convert.ToInt32(assoc.dot11Ssid.uSSIDLength)
                                Dim ssidString As String = Encoding.UTF8.GetString(assoc.dot11Ssid.ucSSID, 0, ssidLength)

                                Return New WifiNetwork() With {
        .Ssid = ssidString,
        .SignalQuality = assoc.wlanSignalQuality,
        .IsConnected = True,
        .IsLan = False
    }
                            End If
                        End If
                    End If
                Finally
                    WlanNative.WlanCloseHandle(clientHandle, IntPtr.Zero)
                End Try
            End If

            ' 2. CHECKING LAN
            Dim activeLan = NetworkInterface.GetAllNetworkInterfaces().
            FirstOrDefault(Function(ni) ni.NetworkInterfaceType = NetworkInterfaceType.Ethernet AndAlso
                                        ni.NetworkInterfaceType <> NetworkInterfaceType.Loopback AndAlso
                                        ni.OperationalStatus = OperationalStatus.Up AndAlso
                                        Not ni.Description.ToLower().Contains("virtual") AndAlso
                                        Not ni.Description.ToLower().Contains("vmware") AndAlso
                                        Not ni.Description.ToLower().Contains("vbox"))

            If activeLan IsNot Nothing Then
                Return New WifiNetwork() With {
                .Ssid = "LAN",
                .IsConnected = True,
                .IsLan = True,
                .SignalQuality = 100
            }
            End If

        Catch
        End Try

        Return Nothing ' No Internet
    End Function

    Public Function IsEthernetDriverPresent() As Boolean
        Return NetworkInterface.GetAllNetworkInterfaces().
        Any(Function(ni) ni.NetworkInterfaceType = NetworkInterfaceType.Ethernet AndAlso
                         ni.NetworkInterfaceType <> NetworkInterfaceType.Loopback)
    End Function

    Public Shared Function IsWifiDriverPresent() As Boolean
        Return NetworkInterface.GetAllNetworkInterfaces().
        Any(Function(ni) ni.NetworkInterfaceType = NetworkInterfaceType.Wireless80211)
    End Function

    Public Shared Async Function GetAvailableWifiNetworksAsync() As Task(Of List(Of WifiNetwork))
        Dim networks As New List(Of WifiNetwork)()

        ' --- CHECK: Is PC connected to the LAN? ---
        Dim ethernetInterfaces = NetworkInterface.GetAllNetworkInterfaces().
        Where(Function(ni) ni.NetworkInterfaceType = NetworkInterfaceType.Ethernet AndAlso ni.NetworkInterfaceType <> NetworkInterfaceType.Loopback)

        For Each lan In ethernetInterfaces
            If lan.OperationalStatus = OperationalStatus.Up Then
                networks.Add(New WifiNetwork() With {
                .Ssid = If(String.IsNullOrEmpty(lan.Description), "Local Network (LAN)", lan.Description),
                .Band = "LAN",
                .WifiStandard = "Ethernet",
                .SignalQuality = 100,
                .IsSecured = False,
                .IsConnected = True,
                .IsLan = True
            })
            End If
        Next

        ' --- CHECK: Does this PC have a Wi-fi card? ---
        Dim hasWifi As Boolean = IsWifiDriverPresent()

        If Not hasWifi Then Return networks

        Dim clientHandle As IntPtr = IntPtr.Zero
        Dim negotiatedVersion As UInteger = 0

        If WlanNative.WlanOpenHandle(2, IntPtr.Zero, negotiatedVersion, clientHandle) <> 0 Then Return networks

        Try
            Dim interfaceListPtr As IntPtr = IntPtr.Zero
            If WlanNative.WlanEnumInterfaces(clientHandle, IntPtr.Zero, interfaceListPtr) <> 0 Then Return networks

            Dim listHeader As WlanNative.WLAN_INTERFACE_INFO_LIST = Marshal.PtrToStructure(Of WlanNative.WLAN_INTERFACE_INFO_LIST)(interfaceListPtr)
            Dim interfaceGuid As Guid = listHeader.InterfaceInfo.InterfaceGuid
            WlanNative.WlanFreeMemory(interfaceListPtr)

            ' SCANNING WI-FIs
            WlanNative.WlanScan(clientHandle, interfaceGuid, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero)
            Await Task.Delay(300)


            Dim securityMap As New Dictionary(Of String, Boolean)()
            Dim connectionMap As New Dictionary(Of String, Boolean)()
            Dim netListPtr As IntPtr = IntPtr.Zero

            If WlanNative.WlanGetAvailableNetworkList(clientHandle, interfaceGuid, 3, IntPtr.Zero, netListPtr) = 0 Then
                Dim numNetItems As Integer = Marshal.ReadInt32(netListPtr)
                Dim netCurrentPtr As IntPtr = IntPtr.Add(netListPtr, 8)

                For i As Integer = 0 To numNetItems - 1
                    Dim net As WlanNative.WLAN_AVAILABLE_NETWORK = Marshal.PtrToStructure(Of WlanNative.WLAN_AVAILABLE_NETWORK)(netCurrentPtr)
                    Dim sLength As Integer = Convert.ToInt32(net.dot11Ssid.uSSIDLength)
                    If sLength > 0 AndAlso sLength <= 32 Then
                        Dim sString As String = Encoding.UTF8.GetString(net.dot11Ssid.ucSSID, 0, sLength)
                        If Not String.IsNullOrWhiteSpace(sString) Then
                            If Not securityMap.ContainsKey(sString) Then
                                securityMap.Add(sString, net.bSecurityEnabled)
                            End If

                            ' 1 = WLAN_AVAILABLE_NETWORK_CONNECTED
                            Dim isCurrentConnection As Boolean = ((net.dwFlags And 1) = 1)
                            If Not connectionMap.ContainsKey(sString) Then
                                connectionMap.Add(sString, isCurrentConnection)
                            End If
                        End If
                    End If
                    netCurrentPtr = IntPtr.Add(netCurrentPtr, Marshal.SizeOf(net))
                Next
                WlanNative.WlanFreeMemory(netListPtr)
            End If


            Dim bssListPtr As IntPtr = IntPtr.Zero
            If WlanNative.WlanGetNetworkBssList(clientHandle, interfaceGuid, IntPtr.Zero, 0, False, IntPtr.Zero, bssListPtr) <> 0 Then
                Return networks
            End If

            Dim numBssItems As Integer = Marshal.ReadInt32(bssListPtr, 4)
            Dim currentPtr As IntPtr = IntPtr.Add(bssListPtr, 8)
            Dim bssEntrySize As Integer = 360

            For i As Integer = 0 To numBssItems - 1
                Dim bssEntry As WlanNative.WLAN_BSS_ENTRY = Marshal.PtrToStructure(Of WlanNative.WLAN_BSS_ENTRY)(currentPtr)
                Dim ssidLength As Integer = Convert.ToInt32(bssEntry.dot11Ssid.uSSIDLength)

                If ssidLength > 0 AndAlso ssidLength <= 32 Then
                    Dim ssidString As String = Encoding.UTF8.GetString(bssEntry.dot11Ssid.ucSSID, 0, ssidLength)

                    If Not String.IsNullOrWhiteSpace(ssidString) Then
                        Dim isSecured As Boolean = False
                        If securityMap.ContainsKey(ssidString) Then
                            isSecured = securityMap(ssidString)
                        End If

                        Dim freq As UInteger = bssEntry.ulChCenterFrequency
                        Dim bandName As String = "Unknown"

                        ' Frequence
                        If freq >= 2400000 AndAlso freq < 2500000 Then
                            bandName = "2.4 GHz"
                        ElseIf freq >= 5000000 AndAlso freq < 6000000 Then
                            bandName = "5 GHz"
                        ElseIf freq >= 6000000 AndAlso freq < 7200000 Then
                            bandName = "6 GHz"
                        End If

                        ' Converting MAC Address (AA:BB:CC:DD:EE:FF)
                        Dim bssidStr As String = String.Join(":", bssEntry.dot11Bssid.Select(Function(b) b.ToString("X2")))

                        Dim isConnected As Boolean = False
                        If connectionMap.ContainsKey(ssidString) Then
                            isConnected = connectionMap(ssidString)
                        End If

                        networks.Add(New WifiNetwork() With {
    .Ssid = ssidString,
    .SignalQuality = bssEntry.uLinkQuality,
    .IsSecured = isSecured,
    .IsConnected = isConnected,
    .FrequencyKhz = freq,
    .Band = bandName,
    .Channel = GetChannelFromFrequency(freq),
    .Bssid = bssidStr,
    .WifiStandard = GetWifiStandardName(bssEntry.dot11BssPhyType),
    .IsLan = False
})
                    End If
                End If

                currentPtr = IntPtr.Add(currentPtr, bssEntrySize)
            Next

            WlanNative.WlanFreeMemory(bssListPtr)
        Finally
            WlanNative.WlanCloseHandle(clientHandle, IntPtr.Zero)
        End Try

        Return networks
    End Function

    Public Function DisconnectWifi() As Boolean
        Dim clientHandle As IntPtr = IntPtr.Zero
        Dim negotiatedVersion As UInteger = 0

        If WlanNative.WlanOpenHandle(2, IntPtr.Zero, negotiatedVersion, clientHandle) <> 0 Then Return False

        Try
            Dim interfaceListPtr As IntPtr = IntPtr.Zero
            If WlanNative.WlanEnumInterfaces(clientHandle, IntPtr.Zero, interfaceListPtr) <> 0 Then Return False

            Dim listHeader As WlanNative.WLAN_INTERFACE_INFO_LIST = Marshal.PtrToStructure(Of WlanNative.WLAN_INTERFACE_INFO_LIST)(interfaceListPtr)
            Dim interfaceGuid As Guid = listHeader.InterfaceInfo.InterfaceGuid
            WlanNative.WlanFreeMemory(interfaceListPtr)

            Dim result As UInteger = WlanNative.WlanDisconnect(clientHandle, interfaceGuid, IntPtr.Zero)

            Return (result = 0)
        Finally
            WlanNative.WlanCloseHandle(clientHandle, IntPtr.Zero)
        End Try
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function DestroyIcon(ByVal hIcon As IntPtr) As Boolean
    End Function

    Public Function ConvertBitmapToIcon(ByVal bmp As Bitmap) As Icon
        Dim hIcon As IntPtr = bmp.GetHicon()

        Dim newIcon As Icon = Icon.FromHandle(hIcon)

        Dim finalIcon As Icon = CType(newIcon.Clone(), Icon)

        DestroyIcon(hIcon)
        newIcon.Dispose()

        Return finalIcon
    End Function

    Public Sub LoadDefaultSettings()
        lbl_Tip.Visible = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Network", "ShowTip", True)

        chk_HidePassword.Checked = False
        chk_HideSSID.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Network", "HideSSID", True)
        chk_AutoUpdate.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Network", "AutoRefresh", True)

        Try
            Controller.Interval = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Shell\Network", "AutoRefreshInterval", 5000)
        Catch
            Controller.Interval = 5000
        End Try
    End Sub

    Private Sub NetworkControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Icon = ConvertBitmapToIcon(My.Resources.Network)

        ' Check if Drivers are installed
        If Not IsWifiDriverPresent() AndAlso Not IsEthernetDriverPresent() Then
            MessageBox.Show("Your OS didn't detected any network drivers installed on your PC!
To avoid corruption of the Shell and OS, this dialog will be disabled.
Please check your drivers, restart this shell and try that again.", "No Drivers found.", MessageBoxButtons.OK, MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1)

            Me.Close()
            Exit Sub
        End If

        LoadDefaultSettings()
        LoadImageListIcons()

        UpdateNetworkListAsync()
    End Sub

    Private Sub NetworkControl_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        AppBar.ToolStripButton5.Checked = False
    End Sub

    Public Sub LoadImageListIcons()
        il_Main.Images.Clear()

        il_Main.Images.Add(ConvertBitmapToIcon(My.Resources.NetworkError))
        il_Main.Images.Add(ConvertBitmapToIcon(My.Resources.NetworkLAN))
        il_Main.Images.Add(ConvertBitmapToIcon(My.Resources.NetworkMax))
        il_Main.Images.Add(ConvertBitmapToIcon(My.Resources.NetworkMaxW))
        il_Main.Images.Add(ConvertBitmapToIcon(My.Resources.NetworkMid))
        il_Main.Images.Add(ConvertBitmapToIcon(My.Resources.NetworkMidW))
        il_Main.Images.Add(ConvertBitmapToIcon(My.Resources.NetworkLow))
        il_Main.Images.Add(ConvertBitmapToIcon(My.Resources.NetworkLowW))
        il_Main.Images.Add(ConvertBitmapToIcon(My.Resources.Network))
        il_Main.Images.Add(ConvertBitmapToIcon(My.Resources.NetworkNone))
    End Sub

    Public Function GetImageIndexByInfo(ByVal speed As Integer, Optional connected As Boolean = False) As Integer
        If speed >= 75 Then
            Return If(connected, 2, 3)
        ElseIf speed >= 50 Then
            Return If(connected, 4, 5)
        ElseIf speed >= 25 Then
            Return If(connected, 6, 7)
        Else
            Return If(connected, 8, 9)
        End If

        Return 0
    End Function

    Dim hideSSIDnames As Boolean
    Public Async Sub UpdateNetworkListAsync()
        btn_Refresh.Enabled = False
        chk_HideSSID.Enabled = False
        lv_Main.Items.Clear()

        Dim allNetworks As List(Of WifiNetwork) = Await GetAvailableWifiNetworksAsync()

        Dim lanItems = allNetworks.Where(Function(n) n.IsLan)

        Dim uniqueWifiItems = allNetworks _
        .Where(Function(n) Not n.IsLan) _
        .GroupBy(Function(n) n.Ssid) _
        .Select(Function(g) g.OrderByDescending(Function(n) n.SignalQuality).First())

        Dim finalSequence = lanItems.Concat(uniqueWifiItems.OrderByDescending(Function(n) n.SignalQuality))

        Dim hiddenIndex As Integer = 0

        lv_Main.BeginUpdate()
        For Each ntw As WifiNetwork In finalSequence
            Dim lvi As New ListViewItem() With {
                .Tag = ntw,
                .ToolTipText = ntw.WifiStandard}

            If hideSSIDnames Then
                lvi.Text = "Network " & hiddenIndex
            Else
                lvi.Text = ntw.Ssid
            End If

            Dim finalIndex As Integer = 0
            If ntw.IsLan Then
                lvi.Text &= " (LAN)"
                finalIndex = 1
            ElseIf ntw.IsConnected Then
                lvi.Text &= " (Connected)"
                AppBar.currentNet = ntw
            End If

            If Not finalIndex = 1 Then finalIndex = GetImageIndexByInfo(ntw.SignalQuality, ntw.IsConnected)

            lvi.ImageIndex = finalIndex

            lvi.SubItems.Add(ntw.IsSecured)
            lvi.SubItems.Add(GetChannelFromFrequency(ntw.FrequencyKhz))
            lvi.SubItems.Add(ntw.SignalQuality & "%")

            If ntw.IsLan OrElse ntw.IsConnected Then
                lvi.Font = New Font(lv_Main.Font, FontStyle.Bold)
            End If

            lv_Main.Items.Add(lvi)

            hiddenIndex += 1
        Next

        lv_Main.EndUpdate()
        btn_Refresh.Enabled = True
        chk_HideSSID.Enabled = True

        lbl_WifiCount.Text = hiddenIndex & " Networks are in this area"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btn_Refresh.Click
        UpdateNetworkListAsync()
    End Sub

    Private Async Sub Button2_Click(sender As Object, e As EventArgs) Handles btn_Connect.Click
        btn_Connect.Enabled = False
        btn_Connect.Text = "Connecting..."

        If String.IsNullOrEmpty(txt_SSID.Text) Then
            MsgBox("SSID field is empty.", MsgBoxStyle.Exclamation)
            btn_Connect.Text = "Connect"
            btn_Connect.Enabled = True
            Exit Sub
        End If

        Dim isNetworkSecured As Boolean = True
        If lv_Main.SelectedItems.Count > 0 Then
            Dim selectedWifi = DirectCast(lv_Main.SelectedItems(0).Tag, WifiNetwork)
            isNetworkSecured = selectedWifi.IsSecured
        End If

        Dim requestSent As Boolean = False

        If String.IsNullOrEmpty(txt_Password.Text) AndAlso isNetworkSecured Then
            requestSent = ConnectToExistingProfile(txt_SSID.Text)
        Else
            Dim secPassword As New SecureString()
            For Each ch As Char In txt_Password.Text
                secPassword.AppendChar(ch)
            Next
            secPassword.MakeReadOnly()

            requestSent = ConnectToWifi(txt_SSID.Text, secPassword)
            secPassword.Dispose()
        End If

        txt_Password.Text = String.Empty

        If requestSent Then
            Await Task.Delay(2000)

            UpdateNetworkListAsync()
            AppBar.UpdateNetworkUI()
        Else
            MsgBox("Failed to send connection request.", MsgBoxStyle.Critical)
        End If

        btn_Connect.Text = "Connect"
        btn_Connect.Enabled = True
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lv_Main.SelectedIndexChanged
        If lv_Main.SelectedItems.Count = 0 Then
            btn_Disconnect.Enabled = False
            Return
        End If

        Dim selectedWifi As WifiNetwork = DirectCast(lv_Main.SelectedItems(0).Tag, WifiNetwork)

        If hideSSIDnames Then txt_SSID.PasswordChar = "●" Else txt_SSID.PasswordChar = ""
        txt_SSID.Text = selectedWifi.Ssid

        If selectedWifi.IsConnected AndAlso Not selectedWifi.IsLan Then
            btn_Disconnect.Enabled = True
        Else
            btn_Disconnect.Enabled = False
        End If
    End Sub

    Private lastSortColumn As Integer = -1
    Private lastSortOrder As SortOrder = SortOrder.None
    Private Sub ListView1_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles lv_Main.ColumnClick
        If e.Column = lastSortColumn Then
            If lastSortOrder = SortOrder.Ascending Then
                lastSortOrder = SortOrder.Descending
            Else
                lastSortOrder = SortOrder.Ascending
            End If
        Else
            lastSortOrder = SortOrder.Ascending
            lastSortColumn = e.Column
        End If

        lv_Main.ListViewItemSorter = New ListViewItemComparer(e.Column, lastSortOrder)
        lv_Main.Sort()

        lv_Main.ListViewItemSorter = Nothing
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles chk_HideSSID.CheckedChanged
        hideSSIDnames = chk_HideSSID.Checked

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Network", "HideSSID", chk_HideSSID.Checked)

        If hideSSIDnames Then txt_SSID.PasswordChar = "●" Else txt_SSID.PasswordChar = ""
        UpdateNetworkListAsync()
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles chk_HidePassword.CheckedChanged
        If Not chk_HidePassword.Checked Then txt_Password.PasswordChar = "●" Else txt_Password.PasswordChar = ""
    End Sub

    Private Sub btn_Disconnect_Click(sender As Object, e As EventArgs) Handles btn_Disconnect.Click
        btn_Disconnect.Enabled = False
        btn_Disconnect.Text = "Disconnecting..."

        If DisconnectWifi() Then
            MsgBox("Wi-Fi síť byla odpojena.", MsgBoxStyle.Information)
            UpdateNetworkListAsync()
            btn_Disconnect.Text = "Disconnect"
        Else
            MsgBox("Odpojení od sítě selhalo.", MsgBoxStyle.Critical)
            btn_Disconnect.Text = "Disconnect"
            btn_Disconnect.Enabled = True
        End If
    End Sub

    Private Sub chk_AutoUpdate_CheckedChanged(sender As Object, e As EventArgs) Handles chk_AutoUpdate.CheckedChanged
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Shell\Network", "AutoRefresh", chk_AutoUpdate.Checked)

        Controller.Enabled = chk_AutoUpdate.Checked
    End Sub

    Private Sub Controller_Tick(sender As Object, e As EventArgs) Handles Controller.Tick
        UpdateNetworkListAsync()
    End Sub

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub txt_SSID_TextChanged(sender As Object, e As EventArgs) Handles txt_SSID.TextChanged
        If Not String.IsNullOrEmpty(txt_SSID.Text) Then
            btn_Connect.Enabled = True
        Else
            btn_Connect.Enabled = False
        End If
    End Sub
End Class

Public Class WlanNative
    <DllImport("wlanapi.dll", SetLastError:=True)>
    Public Shared Function WlanOpenHandle(dwClientVersion As UInteger, pReserved As IntPtr, <Out> ByRef pdwNegotiatedVersion As UInteger, <Out> ByRef phClientHandle As IntPtr) As UInteger
    End Function

    <DllImport("wlanapi.dll", SetLastError:=True)>
    Public Shared Function WlanCloseHandle(hClientHandle As IntPtr, pReserved As IntPtr) As UInteger
    End Function

    <DllImport("wlanapi.dll", SetLastError:=True)>
    Public Shared Function WlanScan(hClientHandle As IntPtr, ByRef pInterfaceGuid As Guid, pDot11Ssid As IntPtr, pIeData As IntPtr, pReserved As IntPtr) As UInteger
    End Function

    <DllImport("wlanapi.dll", SetLastError:=True)>
    Public Shared Function WlanEnumInterfaces(hClientHandle As IntPtr, pReserved As IntPtr, <Out> ByRef ppInterfaceList As IntPtr) As UInteger
    End Function

    <DllImport("wlanapi.dll", SetLastError:=True)>
    Public Shared Function WlanGetAvailableNetworkList(hClientHandle As IntPtr, ByRef pInterfaceGuid As Guid, dwFlags As UInteger, pReserved As IntPtr, <Out> ByRef ppAvailableNetworkList As IntPtr) As UInteger
    End Function

    <DllImport("wlanapi.dll", SetLastError:=True)>
    Public Shared Sub WlanFreeMemory(pMemory As IntPtr)
    End Sub

    <DllImport("wlanapi.dll", SetLastError:=True)>
    Public Shared Function WlanQueryInterface(hClientHandle As IntPtr, ByRef pInterfaceGuid As Guid, opCode As Integer, pReserved As IntPtr, <Out> ByRef pdwDataSize As UInteger, <Out> ByRef ppData As IntPtr, pWlanOpcodeValueType As IntPtr) As UInteger
    End Function

    Public Const wlan_intf_opcode_current_connection As Integer = 7

    <DllImport("wlanapi.dll", SetLastError:=True)>
    Public Shared Function WlanSetProfile(hClientHandle As IntPtr, ByRef pInterfaceGuid As Guid, dwFlags As UInteger, <MarshalAs(UnmanagedType.LPWStr)> strProfileXml As String, <MarshalAs(UnmanagedType.LPWStr)> strAllUserProfileSecurity As String, bOverwrite As Boolean, pReserved As IntPtr, <Out> ByRef pdwReasonCode As UInteger) As UInteger
    End Function

    <DllImport("wlanapi.dll", SetLastError:=True)>
    Public Shared Function WlanConnect(hClientHandle As IntPtr, ByRef pInterfaceGuid As Guid, ByRef pConnectionParameters As WLAN_CONNECTION_PARAMETERS, pReserved As IntPtr) As UInteger
    End Function

    <DllImport("wlanapi.dll", SetLastError:=True)>
    Public Shared Function WlanDisconnect(hClientHandle As IntPtr, ByRef pInterfaceGuid As Guid, pReserved As IntPtr) As UInteger
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure WLAN_CONNECTION_PARAMETERS
        Public wlanConnectionMode As Integer ' 0 = wlan_connection_mode_profile
        <MarshalAs(UnmanagedType.LPWStr)> Public strProfile As String
        Public pDot11Ssid As IntPtr
        Public pDesiredBssidList As IntPtr
        Public dot11BssType As Integer ' 1 = dot11_BSS_type_infrastructure
        Public dwFlags As UInteger
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure WLAN_CONNECTION_ATTRIBUTES
        Public isState As Integer ' 4 = wlan_interface_state_connected
        Public wlanConnectionMode As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)> Public strProfileName As String
        Public wlanAssociationAttributes As WLAN_ASSOCIATION_ATTRIBUTES
        Public wlanSecurityAttributes As WLAN_SECURITY_ATTRIBUTES
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure WLAN_ASSOCIATION_ATTRIBUTES
        Public dot11Ssid As DOT11_SSID
        Public dot11BssType As Integer
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=6)> Public dot11Bssid As Byte()
        Public dot11PhyType As Integer
        Public uDot11PhyIndex As UInteger
        Public wlanSignalQuality As UInteger ' 0 to 100
        Public ulRxRate As UInteger
        Public ulTxRate As UInteger
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure WLAN_SECURITY_ATTRIBUTES
        Public bSecurityEnabled As Boolean
        Public bOneXEnabled As Boolean
        Public dot11AuthAlgorithm As Integer
        Public dot11CipherAlgorithm As Integer
    End Structure

    <DllImport("wlanapi.dll", SetLastError:=True)>
    Public Shared Function WlanGetNetworkBssList(hClientHandle As IntPtr, ByRef pInterfaceGuid As Guid, pDot11Ssid As IntPtr, dot11BssType As Integer, bSecurityEnabled As Boolean, pReserved As IntPtr, <Out> ByRef ppWlanBssList As IntPtr) As UInteger
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure WLAN_BSS_ENTRY
        Public dot11Ssid As DOT11_SSID
        Public uPhyId As UInteger
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=6)> Public dot11Bssid As Byte()
        Public dot11BssType As Integer
        Public dot11BssPhyType As Integer
        Public lRssi As Integer
        Public uLinkQuality As UInteger
        Public bInRegDomain As Boolean
        Public usBeaconPeriod As Short ' Usings / UInt16
        Public ullTimestamp As UInt64
        Public ullHostTimestamp As UInt64
        Public usCapabilityInformation As UInt16
        Public ulChCenterFrequency As UInteger
        Public wlanRateSet As WLAN_RATE_SET
        Public ulIeOffset As UInteger
        Public ulIeSize As UInteger
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure WLAN_RATE_SET
        Public uRateSetLength As UInteger
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=126)> Public usRateSet As UInt16()
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure WLAN_INTERFACE_INFO
        Public InterfaceGuid As Guid
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)> Public InterfaceDescription As String
        Public isState As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure WLAN_INTERFACE_INFO_LIST
        Public dwNumberOfItems As UInteger
        Public dwIndex As UInteger
        Public InterfaceInfo As WLAN_INTERFACE_INFO
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DOT11_SSID
        Public uSSIDLength As UInteger
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)> Public ucSSID As Byte()
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure WLAN_AVAILABLE_NETWORK
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)> Public strProfileName As String
        Public dot11Ssid As DOT11_SSID
        Public dot11BssType As Integer
        Public dwNumberOfBssids As UInteger
        Public bNetworkConnectable As Boolean
        Public wlanNotConnectableReason As UInteger
        Public dwNumberOfPhyTypes As UInteger
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)> Public dot11PhyTypes As UInteger()
        Public bMorePhyTypes As Boolean
        Public wlanSignalQuality As UInteger
        Public bSecurityEnabled As Boolean
        Public dot11DefaultAuthAlgorithm As UInteger
        Public dot11DefaultCipherAlgorithm As UInteger
        Public dwFlags As UInteger
        Public dwReserved As UInteger
    End Structure
End Class

Public Class ListViewItemComparer
    Implements IComparer

    Private _column As Integer
    Private _order As SortOrder

    Public Sub New(column As Integer, order As SortOrder)
        _column = column
        _order = order
    End Sub

    Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
        Dim returnVal As Integer = -1
        Dim itemX As ListViewItem = DirectCast(x, ListViewItem)
        Dim itemY As ListViewItem = DirectCast(y, ListViewItem)

        Dim wifiX As WifiNetwork = DirectCast(itemX.Tag, WifiNetwork)
        Dim wifiY As WifiNetwork = DirectCast(itemY.Tag, WifiNetwork)

        If wifiX Is Nothing OrElse wifiY Is Nothing Then
            Return String.Compare(itemX.SubItems(_column).Text, itemY.SubItems(_column).Text)
        End If

        Select Case _column
            Case 0 ' SSID
                returnVal = String.Compare(wifiX.Ssid, wifiY.Ssid)
            Case 1 ' Pásmo
                returnVal = String.Compare(wifiX.Band, wifiY.Band)
            Case 2 ' Standard
                returnVal = String.Compare(wifiX.WifiStandard, wifiY.WifiStandard)
            Case 3 ' Signal
                returnVal = wifiX.SignalQuality.CompareTo(wifiY.SignalQuality)
            Case 4 ' Secured
                returnVal = wifiX.IsSecured.CompareTo(wifiY.IsSecured)
            Case Else
                returnVal = String.Compare(itemX.SubItems(_column).Text, itemY.SubItems(_column).Text)
        End Select

        If _order = SortOrder.Descending Then
            returnVal *= -1
        End If

        Return returnVal
    End Function
End Class