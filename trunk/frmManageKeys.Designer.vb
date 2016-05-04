<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmManageKeys
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
	Public WithEvents cmdSetKey As System.Windows.Forms.Button
	Public WithEvents imlIcons As System.Windows.Forms.ImageList
	Public WithEvents _lvKeys_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
	Public WithEvents lvKeys As System.Windows.Forms.ListView
	Public WithEvents cmdDelete As System.Windows.Forms.Button
	Public WithEvents cmdEdit As System.Windows.Forms.Button
	Public WithEvents cmdAdd As System.Windows.Forms.Button
	Public WithEvents cmdDone As System.Windows.Forms.Button
	Public WithEvents txtActiveKey As System.Windows.Forms.TextBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmManageKeys))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdSetKey = New System.Windows.Forms.Button
		Me.imlIcons = New System.Windows.Forms.ImageList
		Me.lvKeys = New System.Windows.Forms.ListView
		Me._lvKeys_ColumnHeader_1 = New System.Windows.Forms.ColumnHeader
		Me.cmdDelete = New System.Windows.Forms.Button
		Me.cmdEdit = New System.Windows.Forms.Button
		Me.cmdAdd = New System.Windows.Forms.Button
		Me.cmdDone = New System.Windows.Forms.Button
		Me.txtActiveKey = New System.Windows.Forms.TextBox
		Me.lvKeys.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.BackColor = System.Drawing.Color.Black
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Manage CDKeys"
		Me.ClientSize = New System.Drawing.Size(377, 179)
		Me.Location = New System.Drawing.Point(13, 34)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmManageKeys"
		Me.cmdSetKey.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSetKey.Text = "Set Key"
		Me.cmdSetKey.Size = New System.Drawing.Size(65, 25)
		Me.cmdSetKey.Location = New System.Drawing.Point(304, 8)
		Me.cmdSetKey.TabIndex = 1
		Me.cmdSetKey.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSetKey.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSetKey.CausesValidation = True
		Me.cmdSetKey.Enabled = True
		Me.cmdSetKey.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSetKey.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSetKey.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSetKey.TabStop = True
		Me.cmdSetKey.Name = "cmdSetKey"
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
		Me.lvKeys.Size = New System.Drawing.Size(289, 137)
		Me.lvKeys.Location = New System.Drawing.Point(8, 8)
		Me.lvKeys.TabIndex = 0
		Me.lvKeys.View = System.Windows.Forms.View.Details
		Me.lvKeys.LabelEdit = False
		Me.lvKeys.LabelWrap = True
		Me.lvKeys.HideSelection = True
		Me.lvKeys.FullRowSelect = True
		Me.lvKeys.SmallImageList = imlIcons
		Me.lvKeys.ForeColor = System.Drawing.Color.White
		Me.lvKeys.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.lvKeys.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lvKeys.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lvKeys.Name = "lvKeys"
		Me._lvKeys_ColumnHeader_1.Text = "Key"
		Me._lvKeys_ColumnHeader_1.Width = 471
		Me.cmdDelete.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDelete.Text = "D&elete Selected"
		Me.cmdDelete.Size = New System.Drawing.Size(65, 33)
		Me.cmdDelete.Location = New System.Drawing.Point(304, 72)
		Me.cmdDelete.TabIndex = 6
		Me.cmdDelete.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDelete.CausesValidation = True
		Me.cmdDelete.Enabled = True
		Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDelete.TabStop = True
		Me.cmdDelete.Name = "cmdDelete"
		Me.cmdEdit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdEdit.Text = "&Edit Selected"
		Me.cmdEdit.Size = New System.Drawing.Size(65, 33)
		Me.cmdEdit.Location = New System.Drawing.Point(304, 112)
		Me.cmdEdit.TabIndex = 5
		Me.cmdEdit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdEdit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdEdit.CausesValidation = True
		Me.cmdEdit.Enabled = True
		Me.cmdEdit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdEdit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdEdit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdEdit.TabStop = True
		Me.cmdEdit.Name = "cmdEdit"
		Me.cmdAdd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAdd.Text = "&Add"
		Me.cmdAdd.Size = New System.Drawing.Size(57, 17)
		Me.cmdAdd.Location = New System.Drawing.Point(240, 152)
		Me.cmdAdd.TabIndex = 3
		Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAdd.CausesValidation = True
		Me.cmdAdd.Enabled = True
		Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAdd.TabStop = True
		Me.cmdAdd.Name = "cmdAdd"
		Me.cmdDone.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDone.Text = "&Done"
		Me.cmdDone.Size = New System.Drawing.Size(65, 17)
		Me.cmdDone.Location = New System.Drawing.Point(304, 152)
		Me.cmdDone.TabIndex = 4
		Me.cmdDone.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDone.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDone.CausesValidation = True
		Me.cmdDone.Enabled = True
		Me.cmdDone.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDone.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDone.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDone.TabStop = True
		Me.cmdDone.Name = "cmdDone"
		Me.txtActiveKey.AutoSize = False
		Me.txtActiveKey.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtActiveKey.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtActiveKey.ForeColor = System.Drawing.Color.White
		Me.txtActiveKey.Size = New System.Drawing.Size(225, 19)
		Me.txtActiveKey.Location = New System.Drawing.Point(8, 152)
		Me.txtActiveKey.TabIndex = 2
		Me.txtActiveKey.AcceptsReturn = True
		Me.txtActiveKey.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtActiveKey.CausesValidation = True
		Me.txtActiveKey.Enabled = True
		Me.txtActiveKey.HideSelection = True
		Me.txtActiveKey.ReadOnly = False
		Me.txtActiveKey.Maxlength = 0
		Me.txtActiveKey.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtActiveKey.MultiLine = False
		Me.txtActiveKey.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtActiveKey.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtActiveKey.TabStop = True
		Me.txtActiveKey.Visible = True
		Me.txtActiveKey.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtActiveKey.Name = "txtActiveKey"
		Me.Controls.Add(cmdSetKey)
		Me.Controls.Add(lvKeys)
		Me.Controls.Add(cmdDelete)
		Me.Controls.Add(cmdEdit)
		Me.Controls.Add(cmdAdd)
		Me.Controls.Add(cmdDone)
		Me.Controls.Add(txtActiveKey)
		Me.lvKeys.Columns.Add(_lvKeys_ColumnHeader_1)
		Me.lvKeys.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class