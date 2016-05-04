<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFilters
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
	Public WithEvents _txtOutAdd_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtOutAdd_0 As System.Windows.Forms.TextBox
	Public WithEvents _lvReplace_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
	Public WithEvents _lvReplace_ColumnHeader_2 As System.Windows.Forms.ColumnHeader
	Public WithEvents lvReplace As System.Windows.Forms.ListView
	Public WithEvents _optType_1 As System.Windows.Forms.RadioButton
	Public WithEvents _optType_0 As System.Windows.Forms.RadioButton
	Public WithEvents cmdDone As System.Windows.Forms.Button
	Public WithEvents cmdRem As System.Windows.Forms.Button
	Public WithEvents cmdAdd As System.Windows.Forms.Button
	Public WithEvents txtAdd As System.Windows.Forms.TextBox
	Public WithEvents lbBlock As System.Windows.Forms.ListBox
	Public WithEvents lbText As System.Windows.Forms.ListBox
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdEdit As System.Windows.Forms.Button
	Public WithEvents optText As System.Windows.Forms.RadioButton
	Public WithEvents optBlock As System.Windows.Forms.RadioButton
	Public WithEvents cmdOutRem As System.Windows.Forms.Button
	Public WithEvents cmdOutAdd As System.Windows.Forms.Button
	Public WithEvents _IncomingLbl_2 As System.Windows.Forms.Label
	Public WithEvents _IncomingLbl_3 As System.Windows.Forms.Label
	Public WithEvents _OutgoingLbl_0 As System.Windows.Forms.Label
	Public WithEvents lblInfo As System.Windows.Forms.Label
	Public WithEvents _IncomingLbl_1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents _IncomingLbl_0 As System.Windows.Forms.Label
	Public WithEvents lblMI As System.Windows.Forms.Label
	Public WithEvents _OutgoingLbl_1 As System.Windows.Forms.Label
	Public WithEvents IncomingLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents OutgoingLbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents optType As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents txtOutAdd As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmFilters))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me._txtOutAdd_1 = New System.Windows.Forms.TextBox
		Me._txtOutAdd_0 = New System.Windows.Forms.TextBox
		Me.lvReplace = New System.Windows.Forms.ListView
		Me._lvReplace_ColumnHeader_1 = New System.Windows.Forms.ColumnHeader
		Me._lvReplace_ColumnHeader_2 = New System.Windows.Forms.ColumnHeader
		Me._optType_1 = New System.Windows.Forms.RadioButton
		Me._optType_0 = New System.Windows.Forms.RadioButton
		Me.cmdDone = New System.Windows.Forms.Button
		Me.cmdRem = New System.Windows.Forms.Button
		Me.cmdAdd = New System.Windows.Forms.Button
		Me.txtAdd = New System.Windows.Forms.TextBox
		Me.lbBlock = New System.Windows.Forms.ListBox
		Me.lbText = New System.Windows.Forms.ListBox
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.cmdEdit = New System.Windows.Forms.Button
		Me.optText = New System.Windows.Forms.RadioButton
		Me.optBlock = New System.Windows.Forms.RadioButton
		Me.cmdOutRem = New System.Windows.Forms.Button
		Me.cmdOutAdd = New System.Windows.Forms.Button
		Me._IncomingLbl_2 = New System.Windows.Forms.Label
		Me._IncomingLbl_3 = New System.Windows.Forms.Label
		Me._OutgoingLbl_0 = New System.Windows.Forms.Label
		Me.lblInfo = New System.Windows.Forms.Label
		Me._IncomingLbl_1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me._IncomingLbl_0 = New System.Windows.Forms.Label
		Me.lblMI = New System.Windows.Forms.Label
		Me._OutgoingLbl_1 = New System.Windows.Forms.Label
		Me.IncomingLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.OutgoingLbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.optType = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.txtOutAdd = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.lvReplace.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.IncomingLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.OutgoingLbl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.optType, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.txtOutAdd, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.Color.Black
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Username and Text Filters"
		Me.ClientSize = New System.Drawing.Size(525, 450)
		Me.Location = New System.Drawing.Point(6, 32)
		Me.Icon = CType(resources.GetObject("frmFilters.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmFilters"
		Me._txtOutAdd_1.AutoSize = False
		Me._txtOutAdd_1.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me._txtOutAdd_1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtOutAdd_1.ForeColor = System.Drawing.Color.White
		Me._txtOutAdd_1.Size = New System.Drawing.Size(257, 19)
		Me._txtOutAdd_1.Location = New System.Drawing.Point(256, 312)
		Me._txtOutAdd_1.TabIndex = 22
		Me._txtOutAdd_1.AcceptsReturn = True
		Me._txtOutAdd_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtOutAdd_1.CausesValidation = True
		Me._txtOutAdd_1.Enabled = True
		Me._txtOutAdd_1.HideSelection = True
		Me._txtOutAdd_1.ReadOnly = False
		Me._txtOutAdd_1.Maxlength = 0
		Me._txtOutAdd_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtOutAdd_1.MultiLine = False
		Me._txtOutAdd_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtOutAdd_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtOutAdd_1.TabStop = True
		Me._txtOutAdd_1.Visible = True
		Me._txtOutAdd_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtOutAdd_1.Name = "_txtOutAdd_1"
		Me._txtOutAdd_0.AutoSize = False
		Me._txtOutAdd_0.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me._txtOutAdd_0.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtOutAdd_0.ForeColor = System.Drawing.Color.White
		Me._txtOutAdd_0.Size = New System.Drawing.Size(249, 19)
		Me._txtOutAdd_0.Location = New System.Drawing.Point(8, 312)
		Me._txtOutAdd_0.TabIndex = 21
		Me._txtOutAdd_0.AcceptsReturn = True
		Me._txtOutAdd_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtOutAdd_0.CausesValidation = True
		Me._txtOutAdd_0.Enabled = True
		Me._txtOutAdd_0.HideSelection = True
		Me._txtOutAdd_0.ReadOnly = False
		Me._txtOutAdd_0.Maxlength = 0
		Me._txtOutAdd_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtOutAdd_0.MultiLine = False
		Me._txtOutAdd_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtOutAdd_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtOutAdd_0.TabStop = True
		Me._txtOutAdd_0.Visible = True
		Me._txtOutAdd_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtOutAdd_0.Name = "_txtOutAdd_0"
		Me.lvReplace.Size = New System.Drawing.Size(505, 241)
		Me.lvReplace.Location = New System.Drawing.Point(8, 48)
		Me.lvReplace.TabIndex = 18
		Me.lvReplace.View = System.Windows.Forms.View.Details
		Me.lvReplace.Alignment = System.Windows.Forms.ListViewAlignment.Top
		Me.lvReplace.LabelWrap = False
		Me.lvReplace.HideSelection = True
		Me.lvReplace.FullRowSelect = True
		Me.lvReplace.ForeColor = System.Drawing.Color.White
		Me.lvReplace.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.lvReplace.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lvReplace.LabelEdit = True
		Me.lvReplace.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lvReplace.Name = "lvReplace"
		Me._lvReplace_ColumnHeader_1.Text = "Words To Replace"
		Me._lvReplace_ColumnHeader_1.Width = 436
		Me._lvReplace_ColumnHeader_2.Text = "Replace With"
		Me._lvReplace_ColumnHeader_2.Width = 436
		Me._optType_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._optType_1.BackColor = System.Drawing.Color.Black
		Me._optType_1.Text = "Outgoing Filters"
		Me._optType_1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optType_1.ForeColor = System.Drawing.Color.White
		Me._optType_1.Size = New System.Drawing.Size(145, 17)
		Me._optType_1.Location = New System.Drawing.Point(360, 16)
		Me._optType_1.Appearance = System.Windows.Forms.Appearance.Button
		Me._optType_1.TabIndex = 16
		Me._optType_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optType_1.CausesValidation = True
		Me._optType_1.Enabled = True
		Me._optType_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._optType_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optType_1.TabStop = True
		Me._optType_1.Checked = False
		Me._optType_1.Visible = True
		Me._optType_1.Name = "_optType_1"
		Me._optType_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._optType_0.BackColor = System.Drawing.Color.Black
		Me._optType_0.Text = "Incoming Filters"
		Me._optType_0.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optType_0.ForeColor = System.Drawing.Color.White
		Me._optType_0.Size = New System.Drawing.Size(145, 17)
		Me._optType_0.Location = New System.Drawing.Point(216, 16)
		Me._optType_0.Appearance = System.Windows.Forms.Appearance.Button
		Me._optType_0.TabIndex = 15
		Me._optType_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optType_0.CausesValidation = True
		Me._optType_0.Enabled = True
		Me._optType_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._optType_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optType_0.TabStop = True
		Me._optType_0.Checked = False
		Me._optType_0.Visible = True
		Me._optType_0.Name = "_optType_0"
		Me.cmdDone.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDone.Text = "&Done"
		Me.AcceptButton = Me.cmdDone
		Me.cmdDone.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDone.Size = New System.Drawing.Size(97, 17)
		Me.cmdDone.Location = New System.Drawing.Point(416, 352)
		Me.cmdDone.TabIndex = 14
		Me.cmdDone.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDone.CausesValidation = True
		Me.cmdDone.Enabled = True
		Me.cmdDone.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDone.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDone.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDone.TabStop = True
		Me.cmdDone.Name = "cmdDone"
		Me.cmdRem.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdRem.Text = "&Remove Selected Item(s)"
		Me.cmdRem.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdRem.Size = New System.Drawing.Size(153, 17)
		Me.cmdRem.Location = New System.Drawing.Point(264, 352)
		Me.cmdRem.TabIndex = 13
		Me.cmdRem.BackColor = System.Drawing.SystemColors.Control
		Me.cmdRem.CausesValidation = True
		Me.cmdRem.Enabled = True
		Me.cmdRem.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdRem.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdRem.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdRem.TabStop = True
		Me.cmdRem.Name = "cmdRem"
		Me.cmdAdd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAdd.Text = "&Add It!"
		Me.cmdAdd.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAdd.Size = New System.Drawing.Size(129, 17)
		Me.cmdAdd.Location = New System.Drawing.Point(136, 352)
		Me.cmdAdd.TabIndex = 12
		Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAdd.CausesValidation = True
		Me.cmdAdd.Enabled = True
		Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAdd.TabStop = True
		Me.cmdAdd.Name = "cmdAdd"
		Me.txtAdd.AutoSize = False
		Me.txtAdd.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtAdd.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtAdd.ForeColor = System.Drawing.Color.White
		Me.txtAdd.Size = New System.Drawing.Size(505, 19)
		Me.txtAdd.Location = New System.Drawing.Point(8, 304)
		Me.txtAdd.TabIndex = 8
		Me.txtAdd.AcceptsReturn = True
		Me.txtAdd.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtAdd.CausesValidation = True
		Me.txtAdd.Enabled = True
		Me.txtAdd.HideSelection = True
		Me.txtAdd.ReadOnly = False
		Me.txtAdd.Maxlength = 0
		Me.txtAdd.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtAdd.MultiLine = False
		Me.txtAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtAdd.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtAdd.TabStop = True
		Me.txtAdd.Visible = True
		Me.txtAdd.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtAdd.Name = "txtAdd"
		Me.lbBlock.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.lbBlock.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lbBlock.ForeColor = System.Drawing.Color.White
		Me.lbBlock.Size = New System.Drawing.Size(505, 111)
		Me.lbBlock.Location = New System.Drawing.Point(8, 176)
		Me.lbBlock.TabIndex = 1
		Me.lbBlock.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lbBlock.CausesValidation = True
		Me.lbBlock.Enabled = True
		Me.lbBlock.IntegralHeight = True
		Me.lbBlock.Cursor = System.Windows.Forms.Cursors.Default
		Me.lbBlock.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lbBlock.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lbBlock.Sorted = False
		Me.lbBlock.TabStop = True
		Me.lbBlock.Visible = True
		Me.lbBlock.MultiColumn = False
		Me.lbBlock.Name = "lbBlock"
		Me.lbText.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.lbText.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lbText.ForeColor = System.Drawing.Color.White
		Me.lbText.Size = New System.Drawing.Size(505, 111)
		Me.lbText.Location = New System.Drawing.Point(8, 48)
		Me.lbText.TabIndex = 0
		Me.lbText.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lbText.CausesValidation = True
		Me.lbText.Enabled = True
		Me.lbText.IntegralHeight = True
		Me.lbText.Cursor = System.Windows.Forms.Cursors.Default
		Me.lbText.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lbText.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lbText.Sorted = False
		Me.lbText.TabStop = True
		Me.lbText.Visible = True
		Me.lbText.MultiColumn = False
		Me.lbText.Name = "lbText"
		Me.Frame1.BackColor = System.Drawing.Color.Black
		Me.Frame1.ForeColor = System.Drawing.Color.White
		Me.Frame1.Size = New System.Drawing.Size(305, 41)
		Me.Frame1.Location = New System.Drawing.Point(208, 0)
		Me.Frame1.TabIndex = 17
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.Enabled = True
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.cmdEdit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdEdit.Text = "&Edit Selected Item"
		Me.cmdEdit.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdEdit.Size = New System.Drawing.Size(129, 17)
		Me.cmdEdit.Location = New System.Drawing.Point(8, 352)
		Me.cmdEdit.TabIndex = 25
		Me.cmdEdit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdEdit.CausesValidation = True
		Me.cmdEdit.Enabled = True
		Me.cmdEdit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdEdit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdEdit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdEdit.TabStop = True
		Me.cmdEdit.Name = "cmdEdit"
		Me.optText.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optText.BackColor = System.Drawing.Color.Black
		Me.optText.Text = "Message Filters"
		Me.optText.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optText.ForeColor = System.Drawing.Color.White
		Me.optText.Size = New System.Drawing.Size(105, 17)
		Me.optText.Location = New System.Drawing.Point(104, 328)
		Me.optText.Appearance = System.Windows.Forms.Appearance.Button
		Me.optText.TabIndex = 5
		Me.optText.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optText.CausesValidation = True
		Me.optText.Enabled = True
		Me.optText.Cursor = System.Windows.Forms.Cursors.Default
		Me.optText.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optText.TabStop = True
		Me.optText.Checked = False
		Me.optText.Visible = True
		Me.optText.Name = "optText"
		Me.optBlock.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optBlock.BackColor = System.Drawing.Color.Black
		Me.optBlock.Text = "Block List"
		Me.optBlock.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optBlock.ForeColor = System.Drawing.Color.White
		Me.optBlock.Size = New System.Drawing.Size(105, 17)
		Me.optBlock.Location = New System.Drawing.Point(208, 328)
		Me.optBlock.Appearance = System.Windows.Forms.Appearance.Button
		Me.optBlock.TabIndex = 6
		Me.optBlock.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optBlock.CausesValidation = True
		Me.optBlock.Enabled = True
		Me.optBlock.Cursor = System.Windows.Forms.Cursors.Default
		Me.optBlock.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optBlock.TabStop = True
		Me.optBlock.Checked = False
		Me.optBlock.Visible = True
		Me.optBlock.Name = "optBlock"
		Me.cmdOutRem.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOutRem.Text = "&Remove Selected Row"
		Me.cmdOutRem.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOutRem.Size = New System.Drawing.Size(153, 17)
		Me.cmdOutRem.Location = New System.Drawing.Point(264, 352)
		Me.cmdOutRem.TabIndex = 19
		Me.cmdOutRem.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOutRem.CausesValidation = True
		Me.cmdOutRem.Enabled = True
		Me.cmdOutRem.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOutRem.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOutRem.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOutRem.TabStop = True
		Me.cmdOutRem.Name = "cmdOutRem"
		Me.cmdOutAdd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOutAdd.Text = "&Add It!"
		Me.cmdOutAdd.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOutAdd.Size = New System.Drawing.Size(129, 17)
		Me.cmdOutAdd.Location = New System.Drawing.Point(136, 352)
		Me.cmdOutAdd.TabIndex = 20
		Me.cmdOutAdd.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOutAdd.CausesValidation = True
		Me.cmdOutAdd.Enabled = True
		Me.cmdOutAdd.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOutAdd.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOutAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOutAdd.TabStop = True
		Me.cmdOutAdd.Name = "cmdOutAdd"
		Me._IncomingLbl_2.BackColor = System.Drawing.SystemColors.ControlText
		Me._IncomingLbl_2.Text = "Add to which list?"
		Me._IncomingLbl_2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._IncomingLbl_2.ForeColor = System.Drawing.SystemColors.highlightText
		Me._IncomingLbl_2.Size = New System.Drawing.Size(89, 17)
		Me._IncomingLbl_2.Location = New System.Drawing.Point(8, 328)
		Me._IncomingLbl_2.TabIndex = 7
		Me._IncomingLbl_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._IncomingLbl_2.Enabled = True
		Me._IncomingLbl_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._IncomingLbl_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._IncomingLbl_2.UseMnemonic = True
		Me._IncomingLbl_2.Visible = True
		Me._IncomingLbl_2.AutoSize = False
		Me._IncomingLbl_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._IncomingLbl_2.Name = "_IncomingLbl_2"
		Me._IncomingLbl_3.BackColor = System.Drawing.SystemColors.ControlText
		Me._IncomingLbl_3.Text = "Username / Phrase to add"
		Me._IncomingLbl_3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._IncomingLbl_3.ForeColor = System.Drawing.SystemColors.highlightText
		Me._IncomingLbl_3.Size = New System.Drawing.Size(161, 17)
		Me._IncomingLbl_3.Location = New System.Drawing.Point(8, 288)
		Me._IncomingLbl_3.TabIndex = 9
		Me._IncomingLbl_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._IncomingLbl_3.Enabled = True
		Me._IncomingLbl_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._IncomingLbl_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._IncomingLbl_3.UseMnemonic = True
		Me._IncomingLbl_3.Visible = True
		Me._IncomingLbl_3.AutoSize = False
		Me._IncomingLbl_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._IncomingLbl_3.Name = "_IncomingLbl_3"
		Me._OutgoingLbl_0.BackColor = System.Drawing.SystemColors.ControlText
		Me._OutgoingLbl_0.Text = "Phrase to find:"
		Me._OutgoingLbl_0.ForeColor = System.Drawing.SystemColors.highlightText
		Me._OutgoingLbl_0.Size = New System.Drawing.Size(73, 17)
		Me._OutgoingLbl_0.Location = New System.Drawing.Point(8, 296)
		Me._OutgoingLbl_0.TabIndex = 23
		Me._OutgoingLbl_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._OutgoingLbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._OutgoingLbl_0.Enabled = True
		Me._OutgoingLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._OutgoingLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._OutgoingLbl_0.UseMnemonic = True
		Me._OutgoingLbl_0.Visible = True
		Me._OutgoingLbl_0.AutoSize = False
		Me._OutgoingLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._OutgoingLbl_0.Name = "_OutgoingLbl_0"
		Me.lblInfo.BackColor = System.Drawing.SystemColors.ControlText
		Me.lblInfo.Text = "More Information:"
		Me.lblInfo.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblInfo.ForeColor = System.Drawing.SystemColors.highlightText
		Me.lblInfo.Size = New System.Drawing.Size(89, 17)
		Me.lblInfo.Location = New System.Drawing.Point(8, 376)
		Me.lblInfo.TabIndex = 10
		Me.lblInfo.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblInfo.Enabled = True
		Me.lblInfo.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInfo.UseMnemonic = True
		Me.lblInfo.Visible = True
		Me.lblInfo.AutoSize = False
		Me.lblInfo.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInfo.Name = "lblInfo"
		Me._IncomingLbl_1.BackColor = System.Drawing.SystemColors.ControlText
		Me._IncomingLbl_1.Text = "Username-Based Block List"
		Me._IncomingLbl_1.ForeColor = System.Drawing.SystemColors.highlightText
		Me._IncomingLbl_1.Size = New System.Drawing.Size(137, 17)
		Me._IncomingLbl_1.Location = New System.Drawing.Point(8, 160)
		Me._IncomingLbl_1.TabIndex = 4
		Me._IncomingLbl_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._IncomingLbl_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._IncomingLbl_1.Enabled = True
		Me._IncomingLbl_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._IncomingLbl_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._IncomingLbl_1.UseMnemonic = True
		Me._IncomingLbl_1.Visible = True
		Me._IncomingLbl_1.AutoSize = False
		Me._IncomingLbl_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._IncomingLbl_1.Name = "_IncomingLbl_1"
		Me.Label2.BackColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Text = "-- StealthBot Custom Filters --"
		Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.ForeColor = System.Drawing.SystemColors.highlightText
		Me.Label2.Size = New System.Drawing.Size(193, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 8)
		Me.Label2.TabIndex = 3
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.Enabled = True
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me._IncomingLbl_0.BackColor = System.Drawing.SystemColors.ControlText
		Me._IncomingLbl_0.Text = "Text Message Filters"
		Me._IncomingLbl_0.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._IncomingLbl_0.ForeColor = System.Drawing.SystemColors.highlightText
		Me._IncomingLbl_0.Size = New System.Drawing.Size(105, 17)
		Me._IncomingLbl_0.Location = New System.Drawing.Point(8, 32)
		Me._IncomingLbl_0.TabIndex = 2
		Me._IncomingLbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._IncomingLbl_0.Enabled = True
		Me._IncomingLbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._IncomingLbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._IncomingLbl_0.UseMnemonic = True
		Me._IncomingLbl_0.Visible = True
		Me._IncomingLbl_0.AutoSize = False
		Me._IncomingLbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._IncomingLbl_0.Name = "_IncomingLbl_0"
		Me.lblMI.BackColor = System.Drawing.SystemColors.ControlText
		Me.lblMI.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblMI.ForeColor = System.Drawing.SystemColors.highlightText
		Me.lblMI.Size = New System.Drawing.Size(505, 89)
		Me.lblMI.Location = New System.Drawing.Point(8, 392)
		Me.lblMI.TabIndex = 11
		Me.lblMI.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblMI.Enabled = True
		Me.lblMI.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblMI.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblMI.UseMnemonic = True
		Me.lblMI.Visible = True
		Me.lblMI.AutoSize = False
		Me.lblMI.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblMI.Name = "lblMI"
		Me._OutgoingLbl_1.BackColor = System.Drawing.SystemColors.ControlText
		Me._OutgoingLbl_1.Text = "Phrase to replace with:"
		Me._OutgoingLbl_1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._OutgoingLbl_1.ForeColor = System.Drawing.SystemColors.highlightText
		Me._OutgoingLbl_1.Size = New System.Drawing.Size(113, 17)
		Me._OutgoingLbl_1.Location = New System.Drawing.Point(256, 296)
		Me._OutgoingLbl_1.TabIndex = 24
		Me._OutgoingLbl_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._OutgoingLbl_1.Enabled = True
		Me._OutgoingLbl_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._OutgoingLbl_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._OutgoingLbl_1.UseMnemonic = True
		Me._OutgoingLbl_1.Visible = True
		Me._OutgoingLbl_1.AutoSize = False
		Me._OutgoingLbl_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._OutgoingLbl_1.Name = "_OutgoingLbl_1"
		Me.Controls.Add(_txtOutAdd_1)
		Me.Controls.Add(_txtOutAdd_0)
		Me.Controls.Add(lvReplace)
		Me.Controls.Add(_optType_1)
		Me.Controls.Add(_optType_0)
		Me.Controls.Add(cmdDone)
		Me.Controls.Add(cmdRem)
		Me.Controls.Add(cmdAdd)
		Me.Controls.Add(txtAdd)
		Me.Controls.Add(lbBlock)
		Me.Controls.Add(lbText)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(cmdEdit)
		Me.Controls.Add(optText)
		Me.Controls.Add(optBlock)
		Me.Controls.Add(cmdOutRem)
		Me.Controls.Add(cmdOutAdd)
		Me.Controls.Add(_IncomingLbl_2)
		Me.Controls.Add(_IncomingLbl_3)
		Me.Controls.Add(_OutgoingLbl_0)
		Me.Controls.Add(lblInfo)
		Me.Controls.Add(_IncomingLbl_1)
		Me.Controls.Add(Label2)
		Me.Controls.Add(_IncomingLbl_0)
		Me.Controls.Add(lblMI)
		Me.Controls.Add(_OutgoingLbl_1)
		Me.lvReplace.Columns.Add(_lvReplace_ColumnHeader_1)
		Me.lvReplace.Columns.Add(_lvReplace_ColumnHeader_2)
		Me.IncomingLbl.SetIndex(_IncomingLbl_2, CType(2, Short))
		Me.IncomingLbl.SetIndex(_IncomingLbl_3, CType(3, Short))
		Me.IncomingLbl.SetIndex(_IncomingLbl_1, CType(1, Short))
		Me.IncomingLbl.SetIndex(_IncomingLbl_0, CType(0, Short))
		Me.OutgoingLbl.SetIndex(_OutgoingLbl_0, CType(0, Short))
		Me.OutgoingLbl.SetIndex(_OutgoingLbl_1, CType(1, Short))
		Me.optType.SetIndex(_optType_1, CType(1, Short))
		Me.optType.SetIndex(_optType_0, CType(0, Short))
		Me.txtOutAdd.SetIndex(_txtOutAdd_1, CType(1, Short))
		Me.txtOutAdd.SetIndex(_txtOutAdd_0, CType(0, Short))
		CType(Me.txtOutAdd, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.optType, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.OutgoingLbl, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.IncomingLbl, System.ComponentModel.ISupportInitialize).EndInit()
		Me.lvReplace.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class