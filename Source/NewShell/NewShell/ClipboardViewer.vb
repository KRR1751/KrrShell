Public Class ClipboardViewer

    Private Sub ClipboardViewer_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Visible = False
        Controller.Enabled = False
    End Sub

    Private Sub ClipboardViewer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Visible = True
        Me.Location = New Point(My.Computer.Screen.WorkingArea.Width - Me.Width, My.Computer.Screen.WorkingArea.Height - Me.Height)
        OnTopToolStripMenuItem.Checked = Me.TopMost
        Controller.Enabled = True
    End Sub

    Private Sub Controller_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Controller.Tick
        If Clipboard.ContainsText = True Then
            If RichTextBox1.Text = Clipboard.GetText Then
            Else
                TSSL1.Text = "Clipboard now contains Text:"
                Panel1.Visible = False
                ListBox1.Visible = False
                PictureBox1.Visible = False
                RichTextBox1.Visible = True
                RichTextBox1.Text = Clipboard.GetText
            End If
        ElseIf Clipboard.ContainsFileDropList = True Then
            If ListBox1.Items.Count = Clipboard.GetFileDropList.Count Then
            Else
                ListBox1.Items.Clear()
                TSSL1.Text = "Clipboard now contains File Drop List:"
                Panel1.Visible = False
                PictureBox1.Visible = False
                RichTextBox1.Visible = False
                ListBox1.Visible = True
                Dim i As Integer
                For i = 0 To Clipboard.GetFileDropList.Count - 1
                    ListBox1.Items.Add(Clipboard.GetFileDropList.Item(i))
                Next
            End If
        ElseIf Clipboard.ContainsImage = True Then
            If PictureBox1.Image Is Clipboard.GetImage Then
            Else
                TSSL1.Text = "Clipboard now contains Image:"
                Panel1.Visible = False
                ListBox1.Visible = False
                RichTextBox1.Visible = False
                PictureBox1.Visible = True
                PictureBox1.Image = Clipboard.GetImage
            End If
        ElseIf Clipboard.ContainsAudio = True Then
            A.Stream = Clipboard.GetAudioStream
            If A.Stream Is Clipboard.GetAudioStream Then
            Else
                TSSL1.Text = "Clipboard now contains Audio Stream:"
                ListBox1.Visible = False
                RichTextBox1.Visible = False
                PictureBox1.Visible = False
                Panel1.Visible = True
                Label1.Text = "Stream: " & A.Stream.ToString & Environment.NewLine & "Sound Location: " & A.SoundLocation
            End If
        ElseIf Clipboard.ContainsData(System.Windows.Forms.DataFormats.Rtf) = True Then

        ElseIf Clipboard.ContainsData(System.Windows.Forms.DataFormats.Html) = True Then
            If RichTextBox1.Text = Clipboard.GetData(System.Windows.Forms.DataFormats.Html) Then
            Else
                TSSL1.Text = "Clipboard now contains HTML:"
                Panel1.Visible = False
                ListBox1.Visible = False
                PictureBox1.Visible = False
                RichTextBox1.Visible = True
                RichTextBox1.Text = Clipboard.GetText
            End If
        Else
            If TSSL1.Text = "Clipboard has nothing now to display." Then
            Else
                TSSL1.Text = "Clipboard has nothing now to display."
                Panel1.Visible = False
                ListBox1.Visible = False
                PictureBox1.Visible = False
                RichTextBox1.Visible = False
            End If
        End If
    End Sub
    Dim A As New Media.SoundPlayer()
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Button1.Text = "Play" Then
            A.PlayLooping()
            Button1.Text = "Stop"
        Else
            A.Stop()
            Button1.Text = "Play"
        End If
    End Sub

    Private Sub ClearContentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearContentToolStripMenuItem.Click
        If MessageBox.Show("Are you sure do you want clear the clipboard data?", "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
            Clipboard.Clear()
        End If
    End Sub

    Private Sub HideToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HideToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub OnTopToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OnTopToolStripMenuItem.Click
        Me.TopMost = OnTopToolStripMenuItem.Checked
    End Sub

    Private Sub ReloadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReloadToolStripMenuItem.Click
        ReloadToolStripMenuItem.Enabled = False
        Controller.Enabled = False
        If Clipboard.ContainsText = True Then
            If Clipboard.ContainsData(System.Windows.Forms.DataFormats.Rtf) = True Then
                TSSL1.Text = "Clipboard now contains Rich Text:"
                Panel1.Visible = False
                ListBox1.Visible = False
                PictureBox1.Visible = False
                RichTextBox1.Visible = True
                RichTextBox1.Rtf = Clipboard.GetData(System.Windows.Forms.DataFormats.Rtf)
            Else
                TSSL1.Text = "Clipboard now contains Text:"
                Panel1.Visible = False
                ListBox1.Visible = False
                PictureBox1.Visible = False
                RichTextBox1.Visible = True
                RichTextBox1.Text = Clipboard.GetText
            End If
        ElseIf Clipboard.ContainsFileDropList = True Then
            ListBox1.Items.Clear()
            TSSL1.Text = "Clipboard now contains File Drop List:"
            Panel1.Visible = False
            PictureBox1.Visible = False
            RichTextBox1.Visible = False
            ListBox1.Visible = True
            Dim i As Integer
            For i = 0 To Clipboard.GetFileDropList.Count - 1
                ListBox1.Items.Add(Clipboard.GetFileDropList.Item(i))
            Next
        ElseIf Clipboard.ContainsImage = True Then
            TSSL1.Text = "Clipboard now contains Image:"
            Panel1.Visible = False
            ListBox1.Visible = False
            RichTextBox1.Visible = False
            PictureBox1.Visible = True
            PictureBox1.Image = Clipboard.GetImage
        ElseIf Clipboard.ContainsAudio = True Then
            A.Stream = Clipboard.GetAudioStream
            TSSL1.Text = "Clipboard now contains Audio Stream:"
            ListBox1.Visible = False
            RichTextBox1.Visible = False
            PictureBox1.Visible = False
            Panel1.Visible = True
            Label1.Text = "Stream: " & A.Stream.ToString & Environment.NewLine & "Sound Location: " & A.SoundLocation
        ElseIf Clipboard.ContainsData(System.Windows.Forms.DataFormats.Html) = True Then
            TSSL1.Text = "Clipboard now contains HTML:"
            Panel1.Visible = False
            ListBox1.Visible = False
            PictureBox1.Visible = False
            RichTextBox1.Visible = True
            RichTextBox1.Text = Clipboard.GetText
        Else
            TSSL1.Text = "Clipboard has nothing now to display."
            Panel1.Visible = False
            ListBox1.Visible = False
            PictureBox1.Visible = False
            RichTextBox1.Visible = False
        End If
            Controller.Enabled = True
            ReloadToolStripMenuItem.Enabled = True
            TSSL2.Visible = True
            ReloadInfo.Enabled = True
    End Sub

    Private Sub ReloadInfo_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReloadInfo.Tick
        TSSL2.Visible = False
        ReloadInfo.Enabled = False
    End Sub

    Private Sub RichTextBox1_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkClickedEventArgs) Handles RichTextBox1.LinkClicked
        On Error Resume Next
        If MessageBox.Show("Are you sure do you want visit this link?" & Environment.NewLine & Environment.NewLine & e.LinkText, "Confirm Box", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
            Process.Start(e.LinkText)
        End If
    End Sub
End Class