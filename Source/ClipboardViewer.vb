Imports System.Collections.Specialized
Imports System.IO

Public Class ClipboardViewer
    Private lastText As String = ""
    Private lastFileCount As Integer = -1
    Private soundPlayer As New Media.SoundPlayer()

    Private Sub ClipboardViewer_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Visible = False
        Controller.Enabled = False
    End Sub

    Private Sub ClipboardViewer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Visible = True
        Me.Location = New Point(My.Computer.Screen.WorkingArea.Width - Me.Width, My.Computer.Screen.WorkingArea.Height - Me.Height)
        OnTopToolStripMenuItem.Checked = Me.TopMost
        Controller.Enabled = True
        UpdateClipboardDisplay()
    End Sub

    Private Sub Controller_Tick(sender As Object, e As EventArgs) Handles Controller.Tick
        UpdateClipboardDisplay()
    End Sub

    Private Sub UpdateClipboardDisplay()
        ' 1. PRIORITA: RICH TEXT (Zachovává styly, barvy, fonty)
        If Clipboard.ContainsData(DataFormats.Rtf) Then
            Dim currentRtf As String = DirectCast(Clipboard.GetData(DataFormats.Rtf), String)
            If lastText <> currentRtf Then
                lastText = currentRtf
                ShowControl(RichTextBox1)
                TSSL1.Text = "Clipboard now contains Rich Text (Styled):"
                RichTextBox1.Rtf = currentRtf
            End If

            ' 2. HTML (Zdrojový kód z webu)
        ElseIf Clipboard.ContainsData(DataFormats.Html) Then
            Dim currentHtml As String = DirectCast(Clipboard.GetData(DataFormats.Html), String)
            If lastText <> currentHtml Then
                lastText = currentHtml
                ShowControl(RichTextBox1)
                TSSL1.Text = "Clipboard now contains HTML Source:"
                RichTextBox1.Text = currentHtml
            End If

            ' 3. PROSTÝ TEXT
        ElseIf Clipboard.ContainsText() Then
            Dim currentText As String = Clipboard.GetText()
            If lastText <> currentText Then
                lastText = currentText
                ShowControl(RichTextBox1)
                TSSL1.Text = "Clipboard now contains Plain Text:"
                RichTextBox1.Text = currentText
            End If

            ' 4. SOUBORY
        ElseIf Clipboard.ContainsFileDropList() Then
            Dim files As StringCollection = Clipboard.GetFileDropList()
            If lastFileCount <> files.Count Then
                lastFileCount = files.Count
                ShowControl(ListBox1)
                TSSL1.Text = "Clipboard now contains File Drop List (" & files.Count & "):"
                ListBox1.Items.Clear()
                For Each file As String In files
                    ListBox1.Items.Add(file)
                Next
            End If

            ' 5. OBRÁZKY
        ElseIf Clipboard.ContainsImage() Then
            If TSSL1.Text <> "Clipboard now contains Image:" Then
                ShowControl(PictureBox1)
                TSSL1.Text = "Clipboard now contains Image:"
                PictureBox1.Image = Clipboard.GetImage()
            End If

            ' 6. AUDIO
        ElseIf Clipboard.ContainsAudio() Then
            If TSSL1.Text <> "Clipboard now contains Audio Stream:" Then
                ShowControl(Panel1)
                TSSL1.Text = "Clipboard now contains Audio Stream:"
                soundPlayer.Stream = Clipboard.GetAudioStream()
                Label1.Text = "Stream: Audio Data Presence Detected"
            End If

            ' 7. PRÁZDNO
        Else
            If TSSL1.Text <> "Clipboard has nothing now to display." Then
                TSSL1.Text = "Clipboard has nothing now to display."
                HideAllControls()
            End If
            ResetTrackers()
        End If
    End Sub

    Private Sub ShowControl(ByVal visibleControl As Control)
        HideAllControls()
        visibleControl.Visible = True
    End Sub

    Private Sub HideAllControls()
        Panel1.Visible = False
        ListBox1.Visible = False
        PictureBox1.Visible = False
        RichTextBox1.Visible = False
    End Sub

    Private Sub ResetTrackers()
        lastText = ""
        lastFileCount = -1
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text = "Play" Then
            soundPlayer.PlayLooping()
            Button1.Text = "Stop"
        Else
            soundPlayer.Stop()
            Button1.Text = "Play"
        End If
    End Sub

    Private Sub ClearContentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearContentToolStripMenuItem.Click
        If MessageBox.Show("Are you sure you want to clear the clipboard data?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
            Clipboard.Clear()
            UpdateClipboardDisplay()
        End If
    End Sub

    Private Sub ReloadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReloadToolStripMenuItem.Click
        ResetTrackers()
        TSSL1.Text = "" ' Force refresh
        UpdateClipboardDisplay()
        TSSL2.Visible = True
        ReloadInfo.Enabled = True
    End Sub

    Private Sub ReloadInfo_Tick(sender As Object, e As EventArgs) Handles ReloadInfo.Tick
        TSSL2.Visible = False
        ReloadInfo.Enabled = False
    End Sub

    Private Sub RichTextBox1_LinkClicked(sender As Object, e As LinkClickedEventArgs) Handles RichTextBox1.LinkClicked
        If MessageBox.Show("Do you want to visit this link?" & Environment.NewLine & e.LinkText, "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Process.Start(New ProcessStartInfo(e.LinkText) With {.UseShellExecute = True})
        End If
    End Sub

    Private Sub OnTopToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OnTopToolStripMenuItem.CheckedChanged
        Me.TopMost = OnTopToolStripMenuItem.Checked
    End Sub

    Private Sub HideToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HideToolStripMenuItem.Click
        Me.Close()
    End Sub
End Class