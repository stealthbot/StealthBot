<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmScript
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			Static fTerminateCalled As Boolean
			If Not fTerminateCalled Then
				Form_Terminate_renamed()
				fTerminateCalled = True
			End If
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents dummy As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents _prg_0 As System.Windows.Forms.ProgressBar
	Public WithEvents _trv_0 As System.Windows.Forms.TreeView
	Public WithEvents _cmd_0 As System.Windows.Forms.Button
	Public WithEvents _txt_0 As System.Windows.Forms.TextBox
	Public WithEvents _pic_0 As System.Windows.Forms.PictureBox
	Public WithEvents _chk_0 As System.Windows.Forms.CheckBox
	Public WithEvents _opt_0 As System.Windows.Forms.RadioButton
	Public WithEvents _cmb_0 As System.Windows.Forms.ComboBox
	Public WithEvents _lst_0 As System.Windows.Forms.ListBox
	Public WithEvents _rtb_0 As System.Windows.Forms.RichTextBox
	Public WithEvents _iml_0 As System.Windows.Forms.ImageList
	Public WithEvents _lsv_0 As System.Windows.Forms.ListView
	Public WithEvents _fra_0 As System.Windows.Forms.GroupBox
	Public WithEvents _lbl_0 As System.Windows.Forms.Label
	Public WithEvents _shp_0 As Microsoft.VisualBasic.PowerPacks.RectangleShape
	Public WithEvents _lin_0 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents chk As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	Public WithEvents cmb As Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray
	Public WithEvents cmd As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	Public WithEvents fra As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
	Public WithEvents iml As Microsoft.VisualBasic.Compatibility.VB6.ImageListArray
	Public WithEvents lbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lin As LineShapeArray
	Public WithEvents lst As Microsoft.VisualBasic.Compatibility.VB6.ListBoxArray
	Public WithEvents lsv As Microsoft.VisualBasic.Compatibility.VB6.ListViewArray
	Public WithEvents opt As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents pic As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
	Public WithEvents prg As Microsoft.VisualBasic.Compatibility.VB6.ProgressBarArray
	Public WithEvents rtb As Microsoft.VisualBasic.Compatibility.VB6.RichTextBoxArray
	Public WithEvents trv As Microsoft.VisualBasic.Compatibility.VB6.TreeViewArray
	Public WithEvents txt As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents shp As RectangleShapeArray
	Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmScript))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.dummy = New System.Windows.Forms.ToolStripMenuItem
		Me._prg_0 = New System.Windows.Forms.ProgressBar
		Me._trv_0 = New System.Windows.Forms.TreeView
		Me._cmd_0 = New System.Windows.Forms.Button
		Me._txt_0 = New System.Windows.Forms.TextBox
		Me._pic_0 = New System.Windows.Forms.PictureBox
		Me._chk_0 = New System.Windows.Forms.CheckBox
		Me._opt_0 = New System.Windows.Forms.RadioButton
		Me._cmb_0 = New System.Windows.Forms.ComboBox
		Me._lst_0 = New System.Windows.Forms.ListBox
		Me._rtb_0 = New System.Windows.Forms.RichTextBox
		Me._iml_0 = New System.Windows.Forms.ImageList
		Me._lsv_0 = New System.Windows.Forms.ListView
		Me._fra_0 = New System.Windows.Forms.GroupBox
		Me._lbl_0 = New System.Windows.Forms.Label
		Me._shp_0 = New Microsoft.VisualBasic.PowerPacks.RectangleShape
		Me._lin_0 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.chk = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.cmb = New Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray(components)
		Me.cmd = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(components)
		Me.fra = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(components)
		Me.iml = New Microsoft.VisualBasic.Compatibility.VB6.ImageListArray(components)
		Me.lbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lin = New LineShapeArray(components)
		Me.lst = New Microsoft.VisualBasic.Compatibility.VB6.ListBoxArray(components)
		Me.lsv = New Microsoft.VisualBasic.Compatibility.VB6.ListViewArray(components)
		Me.opt = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.pic = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
		Me.prg = New Microsoft.VisualBasic.Compatibility.VB6.ProgressBarArray(components)
		Me.rtb = New Microsoft.VisualBasic.Compatibility.VB6.RichTextBoxArray(components)
		Me.trv = New Microsoft.VisualBasic.Compatibility.VB6.TreeViewArray(components)
		Me.txt = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.shp = New RectangleShapeArray(components)
		Me.MainMenu1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.chk, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmb, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmd, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.fra, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.iml, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lin, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lst, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.opt, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.pic, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.rtb, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.txt, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.shp, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.Color.Black
		Me.Text = "Scripting UI"
		Me.ClientSize = New System.Drawing.Size(312, 230)
		Me.Location = New System.Drawing.Point(4, 30)
		Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmScript"
		Me.dummy.Name = "dummy"
		Me.dummy.Text = "dummy"
		Me.dummy.Enabled = False
		Me.dummy.Visible = False
		Me.dummy.Checked = False
		Me._prg_0.Size = New System.Drawing.Size(17, 17)
		Me._prg_0.Location = New System.Drawing.Point(112, 48)
		Me._prg_0.TabIndex = 11
		Me._prg_0.Visible = False
		Me._prg_0.Name = "_prg_0"
		Me._trv_0.LabelEdit = True
		Me._trv_0.CausesValidation = True
		Me._trv_0.Size = New System.Drawing.Size(49, 41)
		Me._trv_0.Location = New System.Drawing.Point(192, 24)
		Me._trv_0.TabIndex = 10
		Me._trv_0.Visible = False
		Me._trv_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._trv_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._trv_0.Name = "_trv_0"
		Me._cmd_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmd_0.Size = New System.Drawing.Size(17, 17)
		Me._cmd_0.Location = New System.Drawing.Point(0, 24)
		Me._cmd_0.TabIndex = 8
		Me._cmd_0.Visible = False
		Me._cmd_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmd_0.BackColor = System.Drawing.SystemColors.Control
		Me._cmd_0.CausesValidation = True
		Me._cmd_0.Enabled = True
		Me._cmd_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmd_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmd_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmd_0.TabStop = True
		Me._cmd_0.Name = "_cmd_0"
		Me._txt_0.AutoSize = False
		Me._txt_0.Size = New System.Drawing.Size(17, 19)
		Me._txt_0.Location = New System.Drawing.Point(40, 24)
		Me._txt_0.MultiLine = True
		Me._txt_0.TabIndex = 7
		Me._txt_0.Visible = False
		Me._txt_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txt_0.AcceptsReturn = True
		Me._txt_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txt_0.BackColor = System.Drawing.SystemColors.Window
		Me._txt_0.CausesValidation = True
		Me._txt_0.Enabled = True
		Me._txt_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txt_0.HideSelection = True
		Me._txt_0.ReadOnly = False
		Me._txt_0.Maxlength = 0
		Me._txt_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txt_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txt_0.TabStop = True
		Me._txt_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txt_0.Name = "_txt_0"
		Me._pic_0.Size = New System.Drawing.Size(17, 17)
		Me._pic_0.Location = New System.Drawing.Point(64, 24)
		Me._pic_0.TabIndex = 6
		Me._pic_0.Visible = False
		Me._pic_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._pic_0.Dock = System.Windows.Forms.DockStyle.None
		Me._pic_0.BackColor = System.Drawing.SystemColors.Control
		Me._pic_0.CausesValidation = True
		Me._pic_0.Enabled = True
		Me._pic_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._pic_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._pic_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._pic_0.TabStop = True
		Me._pic_0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me._pic_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._pic_0.Name = "_pic_0"
		Me._chk_0.Size = New System.Drawing.Size(17, 17)
		Me._chk_0.Location = New System.Drawing.Point(88, 24)
		Me._chk_0.TabIndex = 5
		Me._chk_0.Visible = False
		Me._chk_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chk_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chk_0.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chk_0.BackColor = System.Drawing.SystemColors.Control
		Me._chk_0.Text = ""
		Me._chk_0.CausesValidation = True
		Me._chk_0.Enabled = True
		Me._chk_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chk_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._chk_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chk_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chk_0.TabStop = True
		Me._chk_0.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chk_0.Name = "_chk_0"
		Me._opt_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._opt_0.Size = New System.Drawing.Size(17, 17)
		Me._opt_0.Location = New System.Drawing.Point(112, 24)
		Me._opt_0.TabIndex = 4
		Me._opt_0.Visible = False
		Me._opt_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._opt_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._opt_0.BackColor = System.Drawing.SystemColors.Control
		Me._opt_0.CausesValidation = True
		Me._opt_0.Enabled = True
		Me._opt_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._opt_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._opt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._opt_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._opt_0.TabStop = True
		Me._opt_0.Checked = False
		Me._opt_0.Name = "_opt_0"
		Me._cmb_0.Size = New System.Drawing.Size(26, 21)
		Me._cmb_0.Location = New System.Drawing.Point(136, 24)
		Me._cmb_0.TabIndex = 3
		Me._cmb_0.Visible = False
		Me._cmb_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmb_0.BackColor = System.Drawing.SystemColors.Window
		Me._cmb_0.CausesValidation = True
		Me._cmb_0.Enabled = True
		Me._cmb_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cmb_0.IntegralHeight = True
		Me._cmb_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmb_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmb_0.Sorted = False
		Me._cmb_0.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me._cmb_0.TabStop = True
		Me._cmb_0.Name = "_cmb_0"
		Me._lst_0.Size = New System.Drawing.Size(17, 20)
		Me._lst_0.Location = New System.Drawing.Point(168, 24)
		Me._lst_0.TabIndex = 2
		Me._lst_0.Visible = False
		Me._lst_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lst_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._lst_0.BackColor = System.Drawing.SystemColors.Window
		Me._lst_0.CausesValidation = True
		Me._lst_0.Enabled = True
		Me._lst_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._lst_0.IntegralHeight = True
		Me._lst_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lst_0.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me._lst_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lst_0.Sorted = False
		Me._lst_0.TabStop = True
		Me._lst_0.MultiColumn = False
		Me._lst_0.Name = "_lst_0"
		Me._rtb_0.Size = New System.Drawing.Size(33, 17)
		Me._rtb_0.Location = New System.Drawing.Point(24, 48)
		Me._rtb_0.TabIndex = 0
		Me._rtb_0.Visible = False
		Me._rtb_0.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
		Me._rtb_0.RTF = resources.GetString("_rtb_0.TextRTF")
		Me._rtb_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._rtb_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._rtb_0.Name = "_rtb_0"
		Me._iml_0.TransparentColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me._lsv_0.Size = New System.Drawing.Size(17, 25)
		Me._lsv_0.Location = New System.Drawing.Point(0, 48)
		Me._lsv_0.TabIndex = 1
		Me._lsv_0.Visible = False
		Me._lsv_0.View = System.Windows.Forms.View.Details
		Me._lsv_0.LabelEdit = False
		Me._lsv_0.LabelWrap = False
		Me._lsv_0.HideSelection = True
		Me._lsv_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._lsv_0.BackColor = System.Drawing.SystemColors.Window
		Me._lsv_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lsv_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._lsv_0.Name = "_lsv_0"
		Me._fra_0.BackColor = System.Drawing.Color.Black
		Me._fra_0.ForeColor = System.Drawing.Color.White
		Me._fra_0.Size = New System.Drawing.Size(25, 25)
		Me._fra_0.Location = New System.Drawing.Point(136, 48)
		Me._fra_0.TabIndex = 12
		Me._fra_0.Visible = False
		Me._fra_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fra_0.Enabled = True
		Me._fra_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fra_0.Padding = New System.Windows.Forms.Padding(0)
		Me._fra_0.Name = "_fra_0"
		Me._lbl_0.BackColor = System.Drawing.Color.Black
		Me._lbl_0.Text = "lbl"
		Me._lbl_0.ForeColor = System.Drawing.Color.White
		Me._lbl_0.Size = New System.Drawing.Size(17, 17)
		Me._lbl_0.Location = New System.Drawing.Point(24, 24)
		Me._lbl_0.TabIndex = 9
		Me._lbl_0.Visible = False
		Me._lbl_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lbl_0.Enabled = True
		Me._lbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lbl_0.UseMnemonic = True
		Me._lbl_0.AutoSize = False
		Me._lbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lbl_0.Name = "_lbl_0"
		Me._shp_0.Size = New System.Drawing.Size(17, 17)
		Me._shp_0.Location = New System.Drawing.Point(192, 24)
		Me._shp_0.Visible = False
		Me._shp_0.BackColor = System.Drawing.SystemColors.Window
		Me._shp_0.BackStyle = Microsoft.VisualBasic.PowerPacks.BackStyle.Transparent
		Me._shp_0.BorderColor = System.Drawing.SystemColors.WindowText
		Me._shp_0.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._shp_0.BorderWidth = 1
		Me._shp_0.FillColor = System.Drawing.Color.Black
		Me._shp_0.FillStyle = Microsoft.VisualBasic.PowerPacks.FillStyle.Transparent
		Me._shp_0.Name = "_shp_0"
		Me._lin_0.Visible = False
		Me._lin_0.X1 = 48
		Me._lin_0.X2 = 128
		Me._lin_0.Y1 = 32
		Me._lin_0.Y2 = 32
		Me._lin_0.BorderColor = System.Drawing.SystemColors.WindowText
		Me._lin_0.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._lin_0.BorderWidth = 1
		Me._lin_0.Name = "_lin_0"
		Me.Controls.Add(_prg_0)
		Me.Controls.Add(_trv_0)
		Me.Controls.Add(_cmd_0)
		Me.Controls.Add(_txt_0)
		Me.Controls.Add(_pic_0)
		Me.Controls.Add(_chk_0)
		Me.Controls.Add(_opt_0)
		Me.Controls.Add(_cmb_0)
		Me.Controls.Add(_lst_0)
		Me.Controls.Add(_rtb_0)
		Me.Controls.Add(_lsv_0)
		Me.Controls.Add(_fra_0)
		Me.Controls.Add(_lbl_0)
		Me.ShapeContainer1.Shapes.Add(_shp_0)
		Me.ShapeContainer1.Shapes.Add(_lin_0)
		Me.Controls.Add(ShapeContainer1)
		Me.chk.SetIndex(_chk_0, CType(0, Short))
		Me.cmb.SetIndex(_cmb_0, CType(0, Short))
		Me.cmd.SetIndex(_cmd_0, CType(0, Short))
		Me.fra.SetIndex(_fra_0, CType(0, Short))
		Me.iml.SetIndex(_iml_0, CType(0, Short))
		Me.lbl.SetIndex(_lbl_0, CType(0, Short))
		Me.lin.SetIndex(_lin_0, CType(0, Short))
		Me.lst.SetIndex(_lst_0, CType(0, Short))
		Me.lsv.SetIndex(_lsv_0, CType(0, Short))
		Me.opt.SetIndex(_opt_0, CType(0, Short))
		Me.pic.SetIndex(_pic_0, CType(0, Short))
		Me.prg.SetIndex(_prg_0, CType(0, Short))
		Me.rtb.SetIndex(_rtb_0, CType(0, Short))
		Me.trv.SetIndex(_trv_0, CType(0, Short))
		Me.txt.SetIndex(_txt_0, CType(0, Short))
		Me.shp.SetIndex(_shp_0, CType(0, Short))
		CType(Me.shp, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.txt, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.rtb, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.pic, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.opt, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lst, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lin, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.iml, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.fra, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmd, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmb, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.chk, System.ComponentModel.ISupportInitialize).EndInit()
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.dummy})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class