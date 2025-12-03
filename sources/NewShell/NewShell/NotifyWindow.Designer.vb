<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NW
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(NW))
        Me.B1 = New System.Windows.Forms.Button
        Me.B2 = New System.Windows.Forms.Button
        Me.B3 = New System.Windows.Forms.Button
        Me.PB = New System.Windows.Forms.PictureBox
        Me.TL = New System.Windows.Forms.Label
        Me.IP = New System.Windows.Forms.TextBox
        Me.AlCol = New System.Windows.Forms.ListBox
        CType(Me.PB, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'B1
        '
        Me.B1.Location = New System.Drawing.Point(12, 156)
        Me.B1.Name = "B1"
        Me.B1.Size = New System.Drawing.Size(337, 23)
        Me.B1.TabIndex = 0
        Me.B1.Text = "B1_TEXT"
        Me.B1.UseVisualStyleBackColor = True
        '
        'B2
        '
        Me.B2.Location = New System.Drawing.Point(12, 185)
        Me.B2.Name = "B2"
        Me.B2.Size = New System.Drawing.Size(165, 23)
        Me.B2.TabIndex = 1
        Me.B2.Text = "B2_TEXT"
        Me.B2.UseVisualStyleBackColor = True
        '
        'B3
        '
        Me.B3.Location = New System.Drawing.Point(184, 185)
        Me.B3.Name = "B3"
        Me.B3.Size = New System.Drawing.Size(165, 23)
        Me.B3.TabIndex = 2
        Me.B3.Text = "B3_TEXT"
        Me.B3.UseVisualStyleBackColor = True
        '
        'PB
        '
        Me.PB.BackColor = System.Drawing.Color.Transparent
        Me.PB.Image = Global.NewShell.My.Resources.Resources.Information
        Me.PB.Location = New System.Drawing.Point(12, 12)
        Me.PB.Name = "PB"
        Me.PB.Size = New System.Drawing.Size(32, 32)
        Me.PB.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PB.TabIndex = 3
        Me.PB.TabStop = False
        '
        'TL
        '
        Me.TL.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TL.BackColor = System.Drawing.Color.Transparent
        Me.TL.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.TL.Location = New System.Drawing.Point(47, 12)
        Me.TL.Name = "TL"
        Me.TL.Size = New System.Drawing.Size(302, 24)
        Me.TL.TabIndex = 4
        Me.TL.Text = "TITLE_TEXT"
        '
        'IP
        '
        Me.IP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.IP.Location = New System.Drawing.Point(50, 29)
        Me.IP.Multiline = True
        Me.IP.Name = "IP"
        Me.IP.ReadOnly = True
        Me.IP.ShortcutsEnabled = False
        Me.IP.Size = New System.Drawing.Size(299, 121)
        Me.IP.TabIndex = 5
        Me.IP.Text = "DESCRIPTION_TEXT"
        '
        'AlCol
        '
        Me.AlCol.FormattingEnabled = True
        Me.AlCol.Location = New System.Drawing.Point(252, 6)
        Me.AlCol.Name = "AlCol"
        Me.AlCol.Size = New System.Drawing.Size(58, 17)
        Me.AlCol.TabIndex = 6
        Me.AlCol.Visible = False
        '
        'NW
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.AutoValidate = System.Windows.Forms.AutoValidate.EnableAllowFocusChange
        Me.ClientSize = New System.Drawing.Size(361, 216)
        Me.Controls.Add(Me.AlCol)
        Me.Controls.Add(Me.IP)
        Me.Controls.Add(Me.TL)
        Me.Controls.Add(Me.PB)
        Me.Controls.Add(Me.B3)
        Me.Controls.Add(Me.B2)
        Me.Controls.Add(Me.B1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(367, 215)
        Me.Name = "NW"
        Me.RightToLeftLayout = True
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "NotifyWindow"
        Me.TopMost = True
        Me.TransparencyKey = System.Drawing.Color.LightBlue
        CType(Me.PB, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents B1 As System.Windows.Forms.Button
    Friend WithEvents B2 As System.Windows.Forms.Button
    Friend WithEvents B3 As System.Windows.Forms.Button
    Friend WithEvents PB As System.Windows.Forms.PictureBox
    Friend WithEvents TL As System.Windows.Forms.Label
    Friend WithEvents IP As System.Windows.Forms.TextBox
    Friend WithEvents AlCol As System.Windows.Forms.ListBox
End Class
