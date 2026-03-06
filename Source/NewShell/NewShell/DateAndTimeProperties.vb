Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Class DateAndTimeProperties

    Private Const DWMWA_NCRENDERING_POLICY As Integer = 2
    Private Const DWMNCRP_DISABLED As Integer = 1

    <DllImport("dwmapi.dll")>
    Private Shared Function DwmSetWindowAttribute(hwnd As IntPtr, attr As Integer, ByRef attrValue As Integer, attrSize As Integer) As Integer
    End Function

    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        Dim policy As Integer = DWMNCRP_DISABLED
        DwmSetWindowAttribute(Me.Handle, DWMWA_NCRENDERING_POLICY, policy, Marshal.SizeOf(policy))
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DateAndTimeProperties_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If String.IsNullOrEmpty(ComboBox1.Text) Then
            MsgBox("The current inputed Alarm file is missing.")
        End If
    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        If String.IsNullOrWhiteSpace(ComboBox1.Text) Then
            OK_Button.Enabled = False
            GroupBox1.Enabled = False
        Else
            If File.Exists(ComboBox1.Text) Then
                Dim FI As New FileInfo(ComboBox1.Text)

                If FI.Extension.ToLower = ".wav" Then
                    OK_Button.Enabled = True
                    GroupBox1.Enabled = True
                Else
                    OK_Button.Enabled = False
                    GroupBox1.Enabled = False
                End If
            Else
                OK_Button.Enabled = False
                GroupBox1.Enabled = False
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim OFD As New OpenFileDialog With {
        .AutoUpgradeEnabled = AppBar.UseExplorerFP,
        .SupportMultiDottedExtensions = True,
        .Filter = "Wave Audio (*.wav)|*.wav;*.wave",
        .ValidateNames = True,
        .FileName = "",
        .CheckFileExists = True
        }

        If Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\Media") Then
            OFD.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\Media"
        Else
            OFD.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyMusic)
        End If

        If OFD.ShowDialog(Me) = DialogResult.OK Then
            ComboBox1.Text = OFD.FileName
        End If
    End Sub
End Class
