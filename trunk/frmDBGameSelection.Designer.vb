<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmDBGameSelection
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents imlIcons As System.Windows.Forms.ImageList
	Public WithEvents _lvGames_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
	Public WithEvents lvGames As System.Windows.Forms.ListView
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDBGameSelection))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.imlIcons = New System.Windows.Forms.ImageList
		Me.lvGames = New System.Windows.Forms.ListView
		Me._lvGames_ColumnHeader_1 = New System.Windows.Forms.ColumnHeader
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.Label1 = New System.Windows.Forms.Label
		Me.lvGames.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.BackColor = System.Drawing.SystemColors.MenuText
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
		Me.Text = "New Entry - Select Game"
		Me.ClientSize = New System.Drawing.Size(249, 240)
		Me.Location = New System.Drawing.Point(3, 21)
		Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmDBGameSelection"
		Me.imlIcons.ImageSize = New System.Drawing.Size(28, 14)
		Me.imlIcons.ImageStream = CType(resources.GetObject("imlIcons.ImageStream"), System.Windows.Forms.ImageListStreamer)
		Me.imlIcons.Images.SetKeyName(0, "")
		Me.imlIcons.Images.SetKeyName(1, "")
		Me.imlIcons.Images.SetKeyName(2, "")
		Me.imlIcons.Images.SetKeyName(3, "")
		Me.imlIcons.Images.SetKeyName(4, "")
		Me.imlIcons.Images.SetKeyName(5, "")
		Me.imlIcons.Images.SetKeyName(6, "")
		Me.imlIcons.Images.SetKeyName(7, "")
		Me.imlIcons.Images.SetKeyName(8, "")
		Me.imlIcons.Images.SetKeyName(9, "")
		Me.imlIcons.Images.SetKeyName(10, "")
		Me.imlIcons.Images.SetKeyName(11, "")
		Me.imlIcons.Images.SetKeyName(12, "")
		Me.lvGames.Size = New System.Drawing.Size(213, 161)
		Me.lvGames.Location = New System.Drawing.Point(16, 45)
		Me.lvGames.TabIndex = 0
		Me.lvGames.View = System.Windows.Forms.View.Details
		Me.lvGames.LabelEdit = False
		Me.lvGames.LabelWrap = True
		Me.lvGames.HideSelection = True
		Me.lvGames.FullRowSelect = True
		Me.lvGames.SmallImageList = imlIcons
		Me.lvGames.ForeColor = System.Drawing.Color.White
		Me.lvGames.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.lvGames.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lvGames.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lvGames.Name = "lvGames"
		Me._lvGames_ColumnHeader_1.Text = "Game"
		Me._lvGames_ColumnHeader_1.Width = 339
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(57, 17)
		Me.cmdCancel.Location = New System.Drawing.Point(120, 215)
		Me.cmdCancel.TabIndex = 2
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "&OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(57, 17)
		Me.cmdOK.Location = New System.Drawing.Point(174, 215)
		Me.cmdOK.TabIndex = 3
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label1.BackColor = System.Drawing.SystemColors.MenuText
		Me.Label1.Text = "Which game do you wish to create an entry for?"
		Me.Label1.ForeColor = System.Drawing.Color.White
		Me.Label1.Size = New System.Drawing.Size(185, 33)
		Me.Label1.Location = New System.Drawing.Point(30, 8)
		Me.Label1.TabIndex = 1
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Enabled = True
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(lvGames)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(Label1)
		Me.lvGames.Columns.Add(_lvGames_ColumnHeader_1)
		Me.lvGames.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class