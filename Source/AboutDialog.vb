Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public NotInheritable Class AboutDialog
    <Runtime.InteropServices.DllImport("dwmapi.dll")> Public Shared Function DwmExtendFrameIntoClientArea(ByVal hWnd As IntPtr, ByRef pMarinset As side) As Integer
    End Function

    <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)> Public Structure side
        Public Left As Integer
        Public Right As Integer
        Public Top As Integer
        Public Bottom As Integer
    End Structure

    <DllImport("dwmapi.dll", PreserveSig:=False)>
    Public Shared Function DwmGetColorizationColor(ByRef pcrColorization As UInteger, ByRef pfOpaqueBlend As Boolean) As Integer
    End Function

    Private Sub AboutDialog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim side As side = side
            side.Left = -1
            side.Right = -1
            side.Top = -1
            side.Bottom = -1
            Dim result As Integer = DwmExtendFrameIntoClientArea(Me.Handle, side)
        Catch ex As Exception

        End Try

        ' Set the title of the form.
        Dim ApplicationTitle As String
        If My.Application.Info.Title <> "" Then
            ApplicationTitle = My.Application.Info.Title
        Else
            ApplicationTitle = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If
        Me.Text = String.Format("About {0}", ApplicationTitle)

        ' Initialize all of the text displayed on the About Box.
        ' TODO: Customize the application's assembly information in the "Application" pane of the project 
        '    properties dialog (under the "Project" menu).
        Me.LabelProductName.Text = "Product Name: " & My.Application.Info.ProductName
        Me.LabelVersion.Text = String.Format("Version: {0}", My.Application.Info.Version.ToString)
        Me.LabelCopyright.Text = My.Application.Info.Copyright
        Me.LabelCompanyName.Text = "Company Name: " & My.Application.Info.CompanyName
        Me.TextBoxDescription.Text = "Description:" & Environment.NewLine & Environment.NewLine & My.Application.Info.Description
    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
        Me.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Me.BackColor = SystemColors.Control
            Me.ForeColor = SystemColors.ControlText

            Dim side As side = side
            side.Left = 0
            side.Right = 0
            side.Top = 0
            side.Bottom = 0
            Dim result As Integer = DwmExtendFrameIntoClientArea(Me.Handle, side)
        Else
            Me.BackColor = Color.Black
            Me.ForeColor = SystemColors.HighlightText

            Dim side As side = side
            side.Left = -1
            side.Right = -1
            side.Top = -1
            side.Bottom = -1
            Dim result As Integer = DwmExtendFrameIntoClientArea(Me.Handle, side)
        End If
    End Sub

    Private Sub AboutDialog_BackColorChanged(sender As Object, e As EventArgs) Handles Me.BackColorChanged
        If Me.BackColor = Color.Black Then
            Me.LabelProductName.ForeColor = SystemColors.HighlightText
            Me.LabelVersion.ForeColor = SystemColors.HighlightText
            Me.LabelCopyright.ForeColor = SystemColors.HighlightText
            Me.LabelCompanyName.ForeColor = SystemColors.HighlightText
            Me.TextBoxDescription.ForeColor = SystemColors.HighlightText
            CheckBox1.ForeColor = SystemColors.HighlightText
            OKButton.ForeColor = SystemColors.HighlightText
            OKButton.BackColor = Color.Black
        Else
            Me.LabelProductName.ForeColor = SystemColors.ControlText
            Me.LabelVersion.ForeColor = SystemColors.ControlText
            Me.LabelCopyright.ForeColor = SystemColors.ControlText
            Me.LabelCompanyName.ForeColor = SystemColors.ControlText
            Me.TextBoxDescription.ForeColor = SystemColors.ControlText
            CheckBox1.ForeColor = SystemColors.ControlText
            OKButton.ForeColor = SystemColors.ControlText
            OKButton.BackColor = SystemColors.Control
        End If
    End Sub
End Class
