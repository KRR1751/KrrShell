<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Desktop
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.PictureBoxWallpaper = New System.Windows.Forms.PictureBox()
        Me.DesktopCM = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ViewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExtraLargeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LargeIconsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MediumIconsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SmallIconsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MiniIconsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.ArrangeIconsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GridSnappingToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RefreshToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.PasteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewFolderToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewLinkToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.PropertiesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.PictureBoxWallpaper, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DesktopCM.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBoxWallpaper
        '
        Me.PictureBoxWallpaper.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PictureBoxWallpaper.Location = New System.Drawing.Point(0, 0)
        Me.PictureBoxWallpaper.Name = "PictureBoxWallpaper"
        Me.PictureBoxWallpaper.Size = New System.Drawing.Size(800, 450)
        Me.PictureBoxWallpaper.TabIndex = 0
        Me.PictureBoxWallpaper.TabStop = False
        '
        'DesktopCM
        '
        Me.DesktopCM.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ViewToolStripMenuItem, Me.RefreshToolStripMenuItem, Me.ToolStripSeparator1, Me.PasteToolStripMenuItem, Me.NewToolStripMenuItem, Me.ToolStripSeparator2, Me.PropertiesToolStripMenuItem})
        Me.DesktopCM.Name = "DesktopCM"
        Me.DesktopCM.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.DesktopCM.Size = New System.Drawing.Size(128, 126)
        '
        'ViewToolStripMenuItem
        '
        Me.ViewToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExtraLargeToolStripMenuItem, Me.LargeIconsToolStripMenuItem, Me.MediumIconsToolStripMenuItem, Me.SmallIconsToolStripMenuItem, Me.MiniIconsToolStripMenuItem, Me.ToolStripSeparator3, Me.ArrangeIconsToolStripMenuItem, Me.GridSnappingToolStripMenuItem})
        Me.ViewToolStripMenuItem.Name = "ViewToolStripMenuItem"
        Me.ViewToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.ViewToolStripMenuItem.Text = "View"
        '
        'ExtraLargeToolStripMenuItem
        '
        Me.ExtraLargeToolStripMenuItem.Name = "ExtraLargeToolStripMenuItem"
        Me.ExtraLargeToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
        Me.ExtraLargeToolStripMenuItem.Text = "Extra Large Icons"
        '
        'LargeIconsToolStripMenuItem
        '
        Me.LargeIconsToolStripMenuItem.Name = "LargeIconsToolStripMenuItem"
        Me.LargeIconsToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
        Me.LargeIconsToolStripMenuItem.Text = "Large Icons"
        '
        'MediumIconsToolStripMenuItem
        '
        Me.MediumIconsToolStripMenuItem.Name = "MediumIconsToolStripMenuItem"
        Me.MediumIconsToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
        Me.MediumIconsToolStripMenuItem.Text = "Medium Icons"
        '
        'SmallIconsToolStripMenuItem
        '
        Me.SmallIconsToolStripMenuItem.Name = "SmallIconsToolStripMenuItem"
        Me.SmallIconsToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
        Me.SmallIconsToolStripMenuItem.Text = "Small Icons"
        '
        'MiniIconsToolStripMenuItem
        '
        Me.MiniIconsToolStripMenuItem.Name = "MiniIconsToolStripMenuItem"
        Me.MiniIconsToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
        Me.MiniIconsToolStripMenuItem.Text = "Mini Icons"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(160, 6)
        '
        'ArrangeIconsToolStripMenuItem
        '
        Me.ArrangeIconsToolStripMenuItem.Name = "ArrangeIconsToolStripMenuItem"
        Me.ArrangeIconsToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
        Me.ArrangeIconsToolStripMenuItem.Text = "Arrange Icons"
        '
        'GridSnappingToolStripMenuItem
        '
        Me.GridSnappingToolStripMenuItem.Name = "GridSnappingToolStripMenuItem"
        Me.GridSnappingToolStripMenuItem.Size = New System.Drawing.Size(163, 22)
        Me.GridSnappingToolStripMenuItem.Text = "Grid Snapping"
        '
        'RefreshToolStripMenuItem
        '
        Me.RefreshToolStripMenuItem.Name = "RefreshToolStripMenuItem"
        Me.RefreshToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.RefreshToolStripMenuItem.Text = "Refresh"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(124, 6)
        '
        'PasteToolStripMenuItem
        '
        Me.PasteToolStripMenuItem.Name = "PasteToolStripMenuItem"
        Me.PasteToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.PasteToolStripMenuItem.Text = "Paste"
        '
        'NewToolStripMenuItem
        '
        Me.NewToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewFileToolStripMenuItem, Me.NewFolderToolStripMenuItem, Me.NewLinkToolStripMenuItem})
        Me.NewToolStripMenuItem.Name = "NewToolStripMenuItem"
        Me.NewToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.NewToolStripMenuItem.Text = "New"
        '
        'NewFileToolStripMenuItem
        '
        Me.NewFileToolStripMenuItem.Name = "NewFileToolStripMenuItem"
        Me.NewFileToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.NewFileToolStripMenuItem.Text = "File"
        '
        'NewFolderToolStripMenuItem
        '
        Me.NewFolderToolStripMenuItem.Name = "NewFolderToolStripMenuItem"
        Me.NewFolderToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.NewFolderToolStripMenuItem.Text = "Folder"
        '
        'NewLinkToolStripMenuItem
        '
        Me.NewLinkToolStripMenuItem.Name = "NewLinkToolStripMenuItem"
        Me.NewLinkToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.NewLinkToolStripMenuItem.Text = "Link"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(124, 6)
        '
        'PropertiesToolStripMenuItem
        '
        Me.PropertiesToolStripMenuItem.Name = "PropertiesToolStripMenuItem"
        Me.PropertiesToolStripMenuItem.Size = New System.Drawing.Size(127, 22)
        Me.PropertiesToolStripMenuItem.Text = "Properties"
        '
        'Desktop
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Desktop
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.ControlBox = False
        Me.Controls.Add(Me.PictureBoxWallpaper)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Desktop"
        Me.ShowInTaskbar = False
        Me.Text = "Program Manager"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.PictureBoxWallpaper, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DesktopCM.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents PictureBoxWallpaper As PictureBox
    Friend WithEvents DesktopCM As ContextMenuStrip
    Friend WithEvents ViewToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents RefreshToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents PasteToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents NewToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents PropertiesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExtraLargeToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents LargeIconsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MediumIconsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SmallIconsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MiniIconsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents ArrangeIconsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents GridSnappingToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents NewFileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents NewFolderToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents NewLinkToolStripMenuItem As ToolStripMenuItem
End Class
