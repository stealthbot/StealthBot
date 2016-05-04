<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmCommands
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
	Public WithEvents lblRequirements As System.Windows.Forms.Label
	Public WithEvents lblSyntax As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents trvCommands As AxvbalTreeViewLib6.AxvbalTreeView
	Public WithEvents cboCommandGroup As System.Windows.Forms.ComboBox
	Public WithEvents cmdDeleteCommand As System.Windows.Forms.Button
	Public WithEvents cmdFlagRemove As System.Windows.Forms.Button
	Public WithEvents cmdAliasAdd As System.Windows.Forms.Button
	Public WithEvents cmdDiscard As System.Windows.Forms.Button
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents cboFlags As System.Windows.Forms.ComboBox
	Public WithEvents cboAlias As System.Windows.Forms.ComboBox
	Public WithEvents txtRank As System.Windows.Forms.TextBox
	Public WithEvents chkDisable As System.Windows.Forms.CheckBox
	Public WithEvents txtDescription As System.Windows.Forms.TextBox
	Public WithEvents txtSpecialNotes As System.Windows.Forms.TextBox
	Public WithEvents cmdFlagAdd As System.Windows.Forms.Button
	Public WithEvents cmdAliasRemove As System.Windows.Forms.Button
	Public WithEvents lblAlias As System.Windows.Forms.Label
	Public WithEvents lblRank As System.Windows.Forms.Label
	Public WithEvents lblFlags As System.Windows.Forms.Label
	Public WithEvents lblDescription As System.Windows.Forms.Label
	Public WithEvents lblSpecialNotes As System.Windows.Forms.Label
	Public WithEvents fraCommand As System.Windows.Forms.GroupBox
	Public WithEvents lblCommandList As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCommands))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.lblRequirements = New System.Windows.Forms.Label
		Me.lblSyntax = New System.Windows.Forms.Label
		Me.trvCommands = New AxvbalTreeViewLib6.AxvbalTreeView
		Me.cboCommandGroup = New System.Windows.Forms.ComboBox
		Me.fraCommand = New System.Windows.Forms.GroupBox
		Me.cmdDeleteCommand = New System.Windows.Forms.Button
		Me.cmdFlagRemove = New System.Windows.Forms.Button
		Me.cmdAliasAdd = New System.Windows.Forms.Button
		Me.cmdDiscard = New System.Windows.Forms.Button
		Me.cmdSave = New System.Windows.Forms.Button
		Me.cboFlags = New System.Windows.Forms.ComboBox
		Me.cboAlias = New System.Windows.Forms.ComboBox
		Me.txtRank = New System.Windows.Forms.TextBox
		Me.chkDisable = New System.Windows.Forms.CheckBox
		Me.txtDescription = New System.Windows.Forms.TextBox
		Me.txtSpecialNotes = New System.Windows.Forms.TextBox
		Me.cmdFlagAdd = New System.Windows.Forms.Button
		Me.cmdAliasRemove = New System.Windows.Forms.Button
		Me.lblAlias = New System.Windows.Forms.Label
		Me.lblRank = New System.Windows.Forms.Label
		Me.lblFlags = New System.Windows.Forms.Label
		Me.lblDescription = New System.Windows.Forms.Label
		Me.lblSpecialNotes = New System.Windows.Forms.Label
		Me.lblCommandList = New System.Windows.Forms.Label
		Me.Frame1.SuspendLayout()
		Me.fraCommand.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.trvCommands, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.Color.Black
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Command Manager"
		Me.ClientSize = New System.Drawing.Size(622, 479)
		Me.Location = New System.Drawing.Point(3, 29)
		Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Icon = CType(resources.GetObject("frmCommands.Icon"), System.Drawing.Icon)
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
		Me.Name = "frmCommands"
		Me.Frame1.BackColor = System.Drawing.Color.Black
		Me.Frame1.Text = "Command Syntax"
		Me.Frame1.ForeColor = System.Drawing.Color.White
		Me.Frame1.Size = New System.Drawing.Size(609, 65)
		Me.Frame1.Location = New System.Drawing.Point(8, 408)
		Me.Frame1.TabIndex = 22
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.Enabled = True
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.lblRequirements.Text = "Command Requirements"
		Me.lblRequirements.ForeColor = System.Drawing.Color.White
		Me.lblRequirements.Size = New System.Drawing.Size(577, 28)
		Me.lblRequirements.Location = New System.Drawing.Point(16, 29)
		Me.lblRequirements.TabIndex = 24
		Me.lblRequirements.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblRequirements.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblRequirements.BackColor = System.Drawing.Color.Transparent
		Me.lblRequirements.Enabled = True
		Me.lblRequirements.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblRequirements.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblRequirements.UseMnemonic = True
		Me.lblRequirements.Visible = True
		Me.lblRequirements.AutoSize = False
		Me.lblRequirements.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblRequirements.Name = "lblRequirements"
		Me.lblSyntax.Text = "Command Syntax"
		Me.lblSyntax.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblSyntax.ForeColor = System.Drawing.Color.FromARGB(0, 128, 128)
		Me.lblSyntax.Size = New System.Drawing.Size(577, 17)
		Me.lblSyntax.Location = New System.Drawing.Point(18, 16)
		Me.lblSyntax.TabIndex = 23
		Me.lblSyntax.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblSyntax.BackColor = System.Drawing.Color.Transparent
		Me.lblSyntax.Enabled = True
		Me.lblSyntax.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblSyntax.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSyntax.UseMnemonic = True
		Me.lblSyntax.Visible = True
		Me.lblSyntax.AutoSize = False
		Me.lblSyntax.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblSyntax.Name = "lblSyntax"
		trvCommands.OcxState = CType(resources.GetObject("trvCommands.OcxState"), System.Windows.Forms.AxHost.State)
		Me.trvCommands.Size = New System.Drawing.Size(257, 345)
		Me.trvCommands.Location = New System.Drawing.Point(8, 56)
		Me.trvCommands.TabIndex = 1
		Me.trvCommands.Name = "trvCommands"
		Me.cboCommandGroup.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.cboCommandGroup.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboCommandGroup.ForeColor = System.Drawing.Color.White
		Me.cboCommandGroup.Size = New System.Drawing.Size(257, 24)
		Me.cboCommandGroup.Location = New System.Drawing.Point(8, 24)
		Me.cboCommandGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboCommandGroup.TabIndex = 0
		Me.cboCommandGroup.CausesValidation = True
		Me.cboCommandGroup.Enabled = True
		Me.cboCommandGroup.IntegralHeight = True
		Me.cboCommandGroup.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboCommandGroup.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboCommandGroup.Sorted = False
		Me.cboCommandGroup.TabStop = True
		Me.cboCommandGroup.Visible = True
		Me.cboCommandGroup.Name = "cboCommandGroup"
		Me.fraCommand.BackColor = System.Drawing.Color.Black
		Me.fraCommand.ForeColor = System.Drawing.Color.White
		Me.fraCommand.Size = New System.Drawing.Size(337, 393)
		Me.fraCommand.Location = New System.Drawing.Point(280, 8)
		Me.fraCommand.TabIndex = 12
		Me.fraCommand.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraCommand.Enabled = True
		Me.fraCommand.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraCommand.Visible = True
		Me.fraCommand.Padding = New System.Windows.Forms.Padding(0)
		Me.fraCommand.Name = "fraCommand"
		Me.cmdDeleteCommand.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDeleteCommand.Text = "&Delete Command"
		Me.cmdDeleteCommand.Size = New System.Drawing.Size(97, 20)
		Me.cmdDeleteCommand.Location = New System.Drawing.Point(232, 360)
		Me.cmdDeleteCommand.TabIndex = 14
		Me.cmdDeleteCommand.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDeleteCommand.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDeleteCommand.CausesValidation = True
		Me.cmdDeleteCommand.Enabled = True
		Me.cmdDeleteCommand.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDeleteCommand.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDeleteCommand.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDeleteCommand.TabStop = True
		Me.cmdDeleteCommand.Name = "cmdDeleteCommand"
		Me.cmdFlagRemove.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdFlagRemove.Text = "-"
		Me.cmdFlagRemove.Size = New System.Drawing.Size(18, 21)
		Me.cmdFlagRemove.Location = New System.Drawing.Point(312, 40)
		Me.cmdFlagRemove.TabIndex = 8
		Me.cmdFlagRemove.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdFlagRemove.BackColor = System.Drawing.SystemColors.Control
		Me.cmdFlagRemove.CausesValidation = True
		Me.cmdFlagRemove.Enabled = True
		Me.cmdFlagRemove.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdFlagRemove.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdFlagRemove.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdFlagRemove.TabStop = True
		Me.cmdFlagRemove.Name = "cmdFlagRemove"
		Me.cmdAliasAdd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAliasAdd.Text = "+"
		Me.cmdAliasAdd.Size = New System.Drawing.Size(18, 21)
		Me.cmdAliasAdd.Location = New System.Drawing.Point(194, 40)
		Me.cmdAliasAdd.TabIndex = 4
		Me.cmdAliasAdd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAliasAdd.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAliasAdd.CausesValidation = True
		Me.cmdAliasAdd.Enabled = True
		Me.cmdAliasAdd.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAliasAdd.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAliasAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAliasAdd.TabStop = True
		Me.cmdAliasAdd.Name = "cmdAliasAdd"
		Me.cmdDiscard.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDiscard.Text = "Di&scard Changes"
		Me.cmdDiscard.Size = New System.Drawing.Size(97, 20)
		Me.cmdDiscard.Location = New System.Drawing.Point(124, 360)
		Me.cmdDiscard.TabIndex = 13
		Me.cmdDiscard.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDiscard.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDiscard.CausesValidation = True
		Me.cmdDiscard.Enabled = True
		Me.cmdDiscard.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDiscard.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDiscard.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDiscard.TabStop = True
		Me.cmdDiscard.Name = "cmdDiscard"
		Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSave.Text = "&Save Changes"
		Me.cmdSave.Size = New System.Drawing.Size(97, 20)
		Me.cmdSave.Location = New System.Drawing.Point(16, 360)
		Me.cmdSave.TabIndex = 15
		Me.cmdSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSave.CausesValidation = True
		Me.cmdSave.Enabled = True
		Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSave.TabStop = True
		Me.cmdSave.Name = "cmdSave"
		Me.cboFlags.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.cboFlags.ForeColor = System.Drawing.Color.White
		Me.cboFlags.Size = New System.Drawing.Size(47, 21)
		Me.cboFlags.Location = New System.Drawing.Point(240, 40)
		Me.cboFlags.TabIndex = 6
		Me.cboFlags.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboFlags.CausesValidation = True
		Me.cboFlags.Enabled = True
		Me.cboFlags.IntegralHeight = True
		Me.cboFlags.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboFlags.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboFlags.Sorted = False
		Me.cboFlags.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboFlags.TabStop = True
		Me.cboFlags.Visible = True
		Me.cboFlags.Name = "cboFlags"
		Me.cboAlias.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.cboAlias.ForeColor = System.Drawing.Color.White
		Me.cboAlias.Size = New System.Drawing.Size(83, 21)
		Me.cboAlias.Location = New System.Drawing.Point(107, 40)
		Me.cboAlias.TabIndex = 3
		Me.cboAlias.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboAlias.CausesValidation = True
		Me.cboAlias.Enabled = True
		Me.cboAlias.IntegralHeight = True
		Me.cboAlias.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboAlias.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboAlias.Sorted = False
		Me.cboAlias.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboAlias.TabStop = True
		Me.cboAlias.Visible = True
		Me.cboAlias.Name = "cboAlias"
		Me.txtRank.AutoSize = False
		Me.txtRank.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtRank.ForeColor = System.Drawing.Color.White
		Me.txtRank.Size = New System.Drawing.Size(81, 21)
		Me.txtRank.Location = New System.Drawing.Point(16, 41)
		Me.txtRank.Maxlength = 25
		Me.txtRank.TabIndex = 2
		Me.txtRank.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtRank.AcceptsReturn = True
		Me.txtRank.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtRank.CausesValidation = True
		Me.txtRank.Enabled = True
		Me.txtRank.HideSelection = True
		Me.txtRank.ReadOnly = False
		Me.txtRank.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtRank.MultiLine = False
		Me.txtRank.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtRank.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtRank.TabStop = True
		Me.txtRank.Visible = True
		Me.txtRank.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtRank.Name = "txtRank"
		Me.chkDisable.BackColor = System.Drawing.Color.Black
		Me.chkDisable.Text = "Disable"
		Me.chkDisable.ForeColor = System.Drawing.Color.White
		Me.chkDisable.Size = New System.Drawing.Size(313, 33)
		Me.chkDisable.Location = New System.Drawing.Point(16, 328)
		Me.chkDisable.TabIndex = 11
		Me.chkDisable.Visible = False
		Me.chkDisable.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkDisable.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkDisable.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkDisable.CausesValidation = True
		Me.chkDisable.Enabled = True
		Me.chkDisable.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkDisable.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkDisable.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkDisable.TabStop = True
		Me.chkDisable.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkDisable.Name = "chkDisable"
		Me.txtDescription.AutoSize = False
		Me.txtDescription.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtDescription.ForeColor = System.Drawing.Color.White
		Me.txtDescription.Size = New System.Drawing.Size(313, 105)
		Me.txtDescription.Location = New System.Drawing.Point(16, 80)
		Me.txtDescription.MultiLine = True
		Me.txtDescription.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtDescription.TabIndex = 9
		Me.txtDescription.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDescription.AcceptsReturn = True
		Me.txtDescription.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDescription.CausesValidation = True
		Me.txtDescription.Enabled = True
		Me.txtDescription.HideSelection = True
		Me.txtDescription.ReadOnly = False
		Me.txtDescription.Maxlength = 0
		Me.txtDescription.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDescription.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDescription.TabStop = True
		Me.txtDescription.Visible = True
		Me.txtDescription.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDescription.Name = "txtDescription"
		Me.txtSpecialNotes.AutoSize = False
		Me.txtSpecialNotes.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtSpecialNotes.ForeColor = System.Drawing.Color.White
		Me.txtSpecialNotes.Size = New System.Drawing.Size(313, 113)
		Me.txtSpecialNotes.Location = New System.Drawing.Point(16, 208)
		Me.txtSpecialNotes.MultiLine = True
		Me.txtSpecialNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtSpecialNotes.TabIndex = 10
		Me.txtSpecialNotes.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSpecialNotes.AcceptsReturn = True
		Me.txtSpecialNotes.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSpecialNotes.CausesValidation = True
		Me.txtSpecialNotes.Enabled = True
		Me.txtSpecialNotes.HideSelection = True
		Me.txtSpecialNotes.ReadOnly = False
		Me.txtSpecialNotes.Maxlength = 0
		Me.txtSpecialNotes.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSpecialNotes.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSpecialNotes.TabStop = True
		Me.txtSpecialNotes.Visible = True
		Me.txtSpecialNotes.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSpecialNotes.Name = "txtSpecialNotes"
		Me.cmdFlagAdd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdFlagAdd.Text = "+"
		Me.cmdFlagAdd.Size = New System.Drawing.Size(18, 21)
		Me.cmdFlagAdd.Location = New System.Drawing.Point(292, 40)
		Me.cmdFlagAdd.TabIndex = 7
		Me.cmdFlagAdd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdFlagAdd.BackColor = System.Drawing.SystemColors.Control
		Me.cmdFlagAdd.CausesValidation = True
		Me.cmdFlagAdd.Enabled = True
		Me.cmdFlagAdd.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdFlagAdd.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdFlagAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdFlagAdd.TabStop = True
		Me.cmdFlagAdd.Name = "cmdFlagAdd"
		Me.cmdAliasRemove.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAliasRemove.Text = "-"
		Me.cmdAliasRemove.Size = New System.Drawing.Size(18, 21)
		Me.cmdAliasRemove.Location = New System.Drawing.Point(214, 40)
		Me.cmdAliasRemove.TabIndex = 5
		Me.cmdAliasRemove.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAliasRemove.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAliasRemove.CausesValidation = True
		Me.cmdAliasRemove.Enabled = True
		Me.cmdAliasRemove.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAliasRemove.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAliasRemove.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAliasRemove.TabStop = True
		Me.cmdAliasRemove.Name = "cmdAliasRemove"
		Me.lblAlias.Text = "Custom aliases:"
		Me.lblAlias.ForeColor = System.Drawing.Color.White
		Me.lblAlias.Size = New System.Drawing.Size(81, 17)
		Me.lblAlias.Location = New System.Drawing.Point(107, 24)
		Me.lblAlias.TabIndex = 20
		Me.lblAlias.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblAlias.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblAlias.BackColor = System.Drawing.Color.Transparent
		Me.lblAlias.Enabled = True
		Me.lblAlias.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblAlias.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblAlias.UseMnemonic = True
		Me.lblAlias.Visible = True
		Me.lblAlias.AutoSize = False
		Me.lblAlias.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblAlias.Name = "lblAlias"
		Me.lblRank.Text = "Rank (1 - 200):"
		Me.lblRank.ForeColor = System.Drawing.Color.White
		Me.lblRank.Size = New System.Drawing.Size(81, 17)
		Me.lblRank.Location = New System.Drawing.Point(16, 24)
		Me.lblRank.TabIndex = 19
		Me.lblRank.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblRank.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblRank.BackColor = System.Drawing.Color.Transparent
		Me.lblRank.Enabled = True
		Me.lblRank.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblRank.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblRank.UseMnemonic = True
		Me.lblRank.Visible = True
		Me.lblRank.AutoSize = False
		Me.lblRank.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblRank.Name = "lblRank"
		Me.lblFlags.Text = "Flags:"
		Me.lblFlags.ForeColor = System.Drawing.Color.White
		Me.lblFlags.Size = New System.Drawing.Size(65, 17)
		Me.lblFlags.Location = New System.Drawing.Point(240, 24)
		Me.lblFlags.TabIndex = 18
		Me.lblFlags.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblFlags.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblFlags.BackColor = System.Drawing.Color.Transparent
		Me.lblFlags.Enabled = True
		Me.lblFlags.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFlags.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFlags.UseMnemonic = True
		Me.lblFlags.Visible = True
		Me.lblFlags.AutoSize = False
		Me.lblFlags.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblFlags.Name = "lblFlags"
		Me.lblDescription.Text = "Description:"
		Me.lblDescription.ForeColor = System.Drawing.Color.White
		Me.lblDescription.Size = New System.Drawing.Size(145, 17)
		Me.lblDescription.Location = New System.Drawing.Point(16, 64)
		Me.lblDescription.TabIndex = 17
		Me.lblDescription.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDescription.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDescription.BackColor = System.Drawing.Color.Transparent
		Me.lblDescription.Enabled = True
		Me.lblDescription.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDescription.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDescription.UseMnemonic = True
		Me.lblDescription.Visible = True
		Me.lblDescription.AutoSize = False
		Me.lblDescription.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDescription.Name = "lblDescription"
		Me.lblSpecialNotes.Text = "Special notes:"
		Me.lblSpecialNotes.ForeColor = System.Drawing.Color.White
		Me.lblSpecialNotes.Size = New System.Drawing.Size(145, 17)
		Me.lblSpecialNotes.Location = New System.Drawing.Point(16, 192)
		Me.lblSpecialNotes.TabIndex = 16
		Me.lblSpecialNotes.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblSpecialNotes.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblSpecialNotes.BackColor = System.Drawing.Color.Transparent
		Me.lblSpecialNotes.Enabled = True
		Me.lblSpecialNotes.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblSpecialNotes.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSpecialNotes.UseMnemonic = True
		Me.lblSpecialNotes.Visible = True
		Me.lblSpecialNotes.AutoSize = False
		Me.lblSpecialNotes.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblSpecialNotes.Name = "lblSpecialNotes"
		Me.lblCommandList.Text = "Command List"
		Me.lblCommandList.ForeColor = System.Drawing.Color.White
		Me.lblCommandList.Size = New System.Drawing.Size(66, 13)
		Me.lblCommandList.Location = New System.Drawing.Point(10, 11)
		Me.lblCommandList.TabIndex = 21
		Me.lblCommandList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCommandList.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCommandList.BackColor = System.Drawing.Color.Transparent
		Me.lblCommandList.Enabled = True
		Me.lblCommandList.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCommandList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCommandList.UseMnemonic = True
		Me.lblCommandList.Visible = True
		Me.lblCommandList.AutoSize = True
		Me.lblCommandList.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCommandList.Name = "lblCommandList"
		Me.Controls.Add(Frame1)
		Me.Controls.Add(trvCommands)
		Me.Controls.Add(cboCommandGroup)
		Me.Controls.Add(fraCommand)
		Me.Controls.Add(lblCommandList)
		Me.Frame1.Controls.Add(lblRequirements)
		Me.Frame1.Controls.Add(lblSyntax)
		Me.fraCommand.Controls.Add(cmdDeleteCommand)
		Me.fraCommand.Controls.Add(cmdFlagRemove)
		Me.fraCommand.Controls.Add(cmdAliasAdd)
		Me.fraCommand.Controls.Add(cmdDiscard)
		Me.fraCommand.Controls.Add(cmdSave)
		Me.fraCommand.Controls.Add(cboFlags)
		Me.fraCommand.Controls.Add(cboAlias)
		Me.fraCommand.Controls.Add(txtRank)
		Me.fraCommand.Controls.Add(chkDisable)
		Me.fraCommand.Controls.Add(txtDescription)
		Me.fraCommand.Controls.Add(txtSpecialNotes)
		Me.fraCommand.Controls.Add(cmdFlagAdd)
		Me.fraCommand.Controls.Add(cmdAliasRemove)
		Me.fraCommand.Controls.Add(lblAlias)
		Me.fraCommand.Controls.Add(lblRank)
		Me.fraCommand.Controls.Add(lblFlags)
		Me.fraCommand.Controls.Add(lblDescription)
		Me.fraCommand.Controls.Add(lblSpecialNotes)
		CType(Me.trvCommands, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Frame1.ResumeLayout(False)
		Me.fraCommand.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class