Imports System.Windows.Forms
Public Class NFD
    Public PATH As String = String.Empty
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        My.Computer.FileSystem.CreateDirectory(PATH & TextBox1.Text)
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Dialog2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Controller.Enabled = False
    End Sub

    Private Sub Dialog2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Controller.Enabled = True
        LinkLabel1.Text = PATH
    End Sub

    Private Sub Controller_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Controller.Tick
        If PATH = String.Empty Then
            OK_Button.Enabled = False
        Else
            If TextBox1.Text = "" Then
                OK_Button.Enabled = False
            Else
                If TextBox1.Text.Contains("*") OrElse TextBox1.Text.Contains("?") OrElse TextBox1.Text.Contains("""") OrElse TextBox1.Text.Contains(":") OrElse TextBox1.Text.Contains("/") OrElse TextBox1.Text.Contains("\") OrElse TextBox1.Text.Contains("<") OrElse TextBox1.Text.Contains(">") OrElse TextBox1.Text.Contains("|") OrElse TextBox1.Text.Contains("con") OrElse TextBox1.Text.Contains("aux") OrElse TextBox1.Text.Contains("prn2") OrElse TextBox1.Text.Contains("prn3") OrElse TextBox1.Text.Contains("prn4") OrElse TextBox1.Text.Contains("prn5") OrElse TextBox1.Text.Contains("prn6") OrElse TextBox1.Text.Contains("prn7") OrElse TextBox1.Text.Contains("prn8") OrElse TextBox1.Text.Contains("prn9") OrElse TextBox1.Text.Contains("com3") OrElse TextBox1.Text.Contains("com4") OrElse TextBox1.Text.Contains("com5") OrElse TextBox1.Text.Contains("com6") OrElse TextBox1.Text.Contains("com7") OrElse TextBox1.Text.Contains("com8") OrElse TextBox1.Text.Contains("com9") Then
                    OK_Button.Enabled = False
                Else
                    OK_Button.Enabled = True
                End If
            End If
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        If Not PATH = String.Empty Then
            System.Diagnostics.Process.Start(PATH)
        End If
    End Sub

    Private Sub LinkLabel1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel1.TextChanged
        If PATH = String.Empty Then
            LinkLabel1.Text = "(Path)"
        End If
    End Sub
End Class
