<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmRealm
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
	Public WithEvents mnuPopDelete As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopUpgrade As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPop As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents cboOtherRealms As System.Windows.Forms.ComboBox
	Public WithEvents btnDisconnect As System.Windows.Forms.Button
	Public WithEvents btnChoose As System.Windows.Forms.Button
	Public WithEvents tmrLoginTimeout As System.Windows.Forms.Timer
	Public WithEvents optCreateNew As System.Windows.Forms.RadioButton
	Public WithEvents optViewExisting As System.Windows.Forms.RadioButton
	Public WithEvents imlChars As System.Windows.Forms.ImageList
	Public WithEvents txtCharName As System.Windows.Forms.TextBox
	Public WithEvents chkLadder As System.Windows.Forms.CheckBox
	Public WithEvents chkHardcore As System.Windows.Forms.CheckBox
	Public WithEvents chkExpansion As System.Windows.Forms.CheckBox
	Public WithEvents cmdCreate As System.Windows.Forms.Button
	Public WithEvents _optNewCharType_7 As System.Windows.Forms.RadioButton
	Public WithEvents _optNewCharType_6 As System.Windows.Forms.RadioButton
	Public WithEvents _optNewCharType_5 As System.Windows.Forms.RadioButton
	Public WithEvents _optNewCharType_4 As System.Windows.Forms.RadioButton
	Public WithEvents _optNewCharType_3 As System.Windows.Forms.RadioButton
	Public WithEvents _optNewCharType_2 As System.Windows.Forms.RadioButton
	Public WithEvents _optNewCharType_1 As System.Windows.Forms.RadioButton
	Public WithEvents lblCopy As System.Windows.Forms.Label
	Public WithEvents lblCharName As System.Windows.Forms.Label
	Public WithEvents imgCharPortrait As System.Windows.Forms.PictureBox
	Public WithEvents fraCreateNew As System.Windows.Forms.GroupBox
	Public WithEvents lvwChars As System.Windows.Forms.ListView
	Public WithEvents btnUpgrade As System.Windows.Forms.Button
	Public WithEvents btnDelete As System.Windows.Forms.Button
	Public WithEvents _lblRealm_5 As System.Windows.Forms.Label
	Public WithEvents _lblRealm_1 As System.Windows.Forms.Label
	Public WithEvents _lblRealm_0 As System.Windows.Forms.Label
	Public WithEvents _lblRealm_4 As System.Windows.Forms.Label
	Public WithEvents _lblRealm_3 As System.Windows.Forms.Label
	Public WithEvents _lblRealm_2 As System.Windows.Forms.Label
	Public WithEvents lblRealm As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents optNewCharType As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmRealm))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.mnuPop = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopDelete = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopUpgrade = New System.Windows.Forms.ToolStripMenuItem
		Me.cboOtherRealms = New System.Windows.Forms.ComboBox
		Me.btnDisconnect = New System.Windows.Forms.Button
		Me.btnChoose = New System.Windows.Forms.Button
		Me.tmrLoginTimeout = New System.Windows.Forms.Timer(components)
		Me.optCreateNew = New System.Windows.Forms.RadioButton
		Me.optViewExisting = New System.Windows.Forms.RadioButton
		Me.imlChars = New System.Windows.Forms.ImageList
		Me.fraCreateNew = New System.Windows.Forms.GroupBox
		Me.txtCharName = New System.Windows.Forms.TextBox
		Me.chkLadder = New System.Windows.Forms.CheckBox
		Me.chkHardcore = New System.Windows.Forms.CheckBox
		Me.chkExpansion = New System.Windows.Forms.CheckBox
		Me.cmdCreate = New System.Windows.Forms.Button
		Me._optNewCharType_7 = New System.Windows.Forms.RadioButton
		Me._optNewCharType_6 = New System.Windows.Forms.RadioButton
		Me._optNewCharType_5 = New System.Windows.Forms.RadioButton
		Me._optNewCharType_4 = New System.Windows.Forms.RadioButton
		Me._optNewCharType_3 = New System.Windows.Forms.RadioButton
		Me._optNewCharType_2 = New System.Windows.Forms.RadioButton
		Me._optNewCharType_1 = New System.Windows.Forms.RadioButton
		Me.lblCopy = New System.Windows.Forms.Label
		Me.lblCharName = New System.Windows.Forms.Label
		Me.imgCharPortrait = New System.Windows.Forms.PictureBox
		Me.lvwChars = New System.Windows.Forms.ListView
		Me.btnUpgrade = New System.Windows.Forms.Button
		Me.btnDelete = New System.Windows.Forms.Button
		Me._lblRealm_5 = New System.Windows.Forms.Label
		Me._lblRealm_1 = New System.Windows.Forms.Label
		Me._lblRealm_0 = New System.Windows.Forms.Label
		Me._lblRealm_4 = New System.Windows.Forms.Label
		Me._lblRealm_3 = New System.Windows.Forms.Label
		Me._lblRealm_2 = New System.Windows.Forms.Label
		Me.lblRealm = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.optNewCharType = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.MainMenu1.SuspendLayout()
		Me.fraCreateNew.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.lblRealm, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.optNewCharType, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.Color.Black
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Diablo II Realm Login"
		Me.ClientSize = New System.Drawing.Size(728, 341)
		Me.Location = New System.Drawing.Point(35, 56)
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
		Me.Name = "frmRealm"
		Me.mnuPop.Name = "mnuPop"
		Me.mnuPop.Text = "mnuPop"
		Me.mnuPop.Visible = False
		Me.mnuPop.Checked = False
		Me.mnuPop.Enabled = True
		Me.mnuPopDelete.Name = "mnuPopDelete"
		Me.mnuPopDelete.Text = "&Delete"
		Me.mnuPopDelete.ShortcutKeys = CType(System.Windows.Forms.Keys.Delete, System.Windows.Forms.Keys)
		Me.mnuPopDelete.Checked = False
		Me.mnuPopDelete.Enabled = True
		Me.mnuPopDelete.Visible = True
		Me.mnuPopUpgrade.Name = "mnuPopUpgrade"
		Me.mnuPopUpgrade.Text = "&Upgrade to Expansion"
		Me.mnuPopUpgrade.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.U, System.Windows.Forms.Keys)
		Me.mnuPopUpgrade.Visible = False
		Me.mnuPopUpgrade.Checked = False
		Me.mnuPopUpgrade.Enabled = True
		Me.cboOtherRealms.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.cboOtherRealms.Enabled = False
		Me.cboOtherRealms.ForeColor = System.Drawing.Color.White
		Me.cboOtherRealms.Size = New System.Drawing.Size(89, 21)
		Me.cboOtherRealms.Location = New System.Drawing.Point(632, 128)
		Me.cboOtherRealms.TabIndex = 7
		Me.cboOtherRealms.Text = "Combo1"
		Me.cboOtherRealms.Visible = False
		Me.cboOtherRealms.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboOtherRealms.CausesValidation = True
		Me.cboOtherRealms.IntegralHeight = True
		Me.cboOtherRealms.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboOtherRealms.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboOtherRealms.Sorted = False
		Me.cboOtherRealms.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboOtherRealms.TabStop = True
		Me.cboOtherRealms.Name = "cboOtherRealms"
		Me.btnDisconnect.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.btnDisconnect
		Me.btnDisconnect.Text = "&Disconnect"
		Me.btnDisconnect.Size = New System.Drawing.Size(97, 20)
		Me.btnDisconnect.Location = New System.Drawing.Point(536, 312)
		Me.btnDisconnect.TabIndex = 2
		Me.btnDisconnect.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnDisconnect.BackColor = System.Drawing.SystemColors.Control
		Me.btnDisconnect.CausesValidation = True
		Me.btnDisconnect.Enabled = True
		Me.btnDisconnect.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnDisconnect.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnDisconnect.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnDisconnect.TabStop = True
		Me.btnDisconnect.Name = "btnDisconnect"
		Me.btnChoose.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnChoose.Text = "&Choose This"
		Me.AcceptButton = Me.btnChoose
		Me.btnChoose.Size = New System.Drawing.Size(89, 20)
		Me.btnChoose.Location = New System.Drawing.Point(632, 312)
		Me.btnChoose.TabIndex = 1
		Me.btnChoose.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnChoose.BackColor = System.Drawing.SystemColors.Control
		Me.btnChoose.CausesValidation = True
		Me.btnChoose.Enabled = True
		Me.btnChoose.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnChoose.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnChoose.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnChoose.TabStop = True
		Me.btnChoose.Name = "btnChoose"
		Me.tmrLoginTimeout.Enabled = False
		Me.tmrLoginTimeout.Interval = 1000
		Me.optCreateNew.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optCreateNew.BackColor = System.Drawing.Color.Black
		Me.optCreateNew.Text = "Create New Character"
		Me.optCreateNew.Enabled = False
		Me.optCreateNew.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optCreateNew.ForeColor = System.Drawing.Color.White
		Me.optCreateNew.Size = New System.Drawing.Size(89, 33)
		Me.optCreateNew.Location = New System.Drawing.Point(632, 72)
		Me.optCreateNew.Appearance = System.Windows.Forms.Appearance.Button
		Me.optCreateNew.TabIndex = 6
		Me.optCreateNew.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optCreateNew.CausesValidation = True
		Me.optCreateNew.Cursor = System.Windows.Forms.Cursors.Default
		Me.optCreateNew.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optCreateNew.TabStop = True
		Me.optCreateNew.Checked = False
		Me.optCreateNew.Visible = True
		Me.optCreateNew.Name = "optCreateNew"
		Me.optViewExisting.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optViewExisting.BackColor = System.Drawing.Color.Black
		Me.optViewExisting.Text = "View Existing Characters"
		Me.optViewExisting.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optViewExisting.ForeColor = System.Drawing.Color.White
		Me.optViewExisting.Size = New System.Drawing.Size(89, 33)
		Me.optViewExisting.Location = New System.Drawing.Point(632, 32)
		Me.optViewExisting.Appearance = System.Windows.Forms.Appearance.Button
		Me.optViewExisting.TabIndex = 5
		Me.optViewExisting.Checked = True
		Me.optViewExisting.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optViewExisting.CausesValidation = True
		Me.optViewExisting.Enabled = True
		Me.optViewExisting.Cursor = System.Windows.Forms.Cursors.Default
		Me.optViewExisting.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optViewExisting.TabStop = True
		Me.optViewExisting.Visible = True
		Me.optViewExisting.Name = "optViewExisting"
		Me.imlChars.ImageSize = New System.Drawing.Size(103, 201)
		Me.imlChars.TransparentColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.imlChars.ImageStream = CType(resources.GetObject("imlChars.ImageStream"), System.Windows.Forms.ImageListStreamer)
		Me.imlChars.Images.SetKeyName(0, "")
		Me.imlChars.Images.SetKeyName(1, "")
		Me.imlChars.Images.SetKeyName(2, "")
		Me.imlChars.Images.SetKeyName(3, "")
		Me.imlChars.Images.SetKeyName(4, "")
		Me.imlChars.Images.SetKeyName(5, "")
		Me.imlChars.Images.SetKeyName(6, "")
		Me.imlChars.Images.SetKeyName(7, "")
		Me.fraCreateNew.BackColor = System.Drawing.Color.Black
		Me.fraCreateNew.ForeColor = System.Drawing.Color.White
		Me.fraCreateNew.Size = New System.Drawing.Size(617, 273)
		Me.fraCreateNew.Location = New System.Drawing.Point(8, 32)
		Me.fraCreateNew.TabIndex = 8
		Me.fraCreateNew.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraCreateNew.Enabled = True
		Me.fraCreateNew.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraCreateNew.Visible = True
		Me.fraCreateNew.Padding = New System.Windows.Forms.Padding(0)
		Me.fraCreateNew.Name = "fraCreateNew"
		Me.txtCharName.AutoSize = False
		Me.txtCharName.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtCharName.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtCharName.ForeColor = System.Drawing.Color.White
		Me.txtCharName.Size = New System.Drawing.Size(153, 19)
		Me.txtCharName.Location = New System.Drawing.Point(424, 128)
		Me.txtCharName.Maxlength = 15
		Me.txtCharName.TabIndex = 19
		Me.txtCharName.AcceptsReturn = True
		Me.txtCharName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCharName.CausesValidation = True
		Me.txtCharName.Enabled = True
		Me.txtCharName.HideSelection = True
		Me.txtCharName.ReadOnly = False
		Me.txtCharName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCharName.MultiLine = False
		Me.txtCharName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCharName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCharName.TabStop = True
		Me.txtCharName.Visible = True
		Me.txtCharName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtCharName.Name = "txtCharName"
		Me.chkLadder.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.chkLadder.BackColor = System.Drawing.Color.Black
		Me.chkLadder.Text = "Ladder"
		Me.chkLadder.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkLadder.ForeColor = System.Drawing.Color.White
		Me.chkLadder.Size = New System.Drawing.Size(89, 25)
		Me.chkLadder.Location = New System.Drawing.Point(304, 152)
		Me.chkLadder.Appearance = System.Windows.Forms.Appearance.Button
		Me.chkLadder.TabIndex = 18
		Me.chkLadder.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLadder.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLadder.CausesValidation = True
		Me.chkLadder.Enabled = True
		Me.chkLadder.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLadder.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLadder.TabStop = True
		Me.chkLadder.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLadder.Visible = True
		Me.chkLadder.Name = "chkLadder"
		Me.chkHardcore.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.chkHardcore.BackColor = System.Drawing.Color.Black
		Me.chkHardcore.Text = "Hardcore"
		Me.chkHardcore.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkHardcore.ForeColor = System.Drawing.Color.White
		Me.chkHardcore.Size = New System.Drawing.Size(89, 25)
		Me.chkHardcore.Location = New System.Drawing.Point(304, 120)
		Me.chkHardcore.Appearance = System.Windows.Forms.Appearance.Button
		Me.chkHardcore.TabIndex = 17
		Me.chkHardcore.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkHardcore.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkHardcore.CausesValidation = True
		Me.chkHardcore.Enabled = True
		Me.chkHardcore.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkHardcore.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkHardcore.TabStop = True
		Me.chkHardcore.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkHardcore.Visible = True
		Me.chkHardcore.Name = "chkHardcore"
		Me.chkExpansion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.chkExpansion.BackColor = System.Drawing.Color.Black
		Me.chkExpansion.Text = "Expansion"
		Me.chkExpansion.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkExpansion.ForeColor = System.Drawing.Color.White
		Me.chkExpansion.Size = New System.Drawing.Size(89, 25)
		Me.chkExpansion.Location = New System.Drawing.Point(304, 88)
		Me.chkExpansion.Appearance = System.Windows.Forms.Appearance.Button
		Me.chkExpansion.TabIndex = 16
		Me.chkExpansion.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkExpansion.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkExpansion.CausesValidation = True
		Me.chkExpansion.Enabled = True
		Me.chkExpansion.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkExpansion.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkExpansion.TabStop = True
		Me.chkExpansion.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkExpansion.Visible = True
		Me.chkExpansion.Name = "chkExpansion"
		Me.cmdCreate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCreate.Text = "&Create"
		Me.cmdCreate.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCreate.Size = New System.Drawing.Size(97, 25)
		Me.cmdCreate.Location = New System.Drawing.Point(480, 160)
		Me.cmdCreate.TabIndex = 20
		Me.cmdCreate.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCreate.CausesValidation = True
		Me.cmdCreate.Enabled = True
		Me.cmdCreate.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCreate.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCreate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCreate.TabStop = True
		Me.cmdCreate.Name = "cmdCreate"
		Me._optNewCharType_7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._optNewCharType_7.BackColor = System.Drawing.Color.Black
		Me._optNewCharType_7.Text = "Assassin"
		Me._optNewCharType_7.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optNewCharType_7.ForeColor = System.Drawing.Color.White
		Me._optNewCharType_7.Size = New System.Drawing.Size(105, 17)
		Me._optNewCharType_7.Location = New System.Drawing.Point(56, 192)
		Me._optNewCharType_7.Appearance = System.Windows.Forms.Appearance.Button
		Me._optNewCharType_7.TabIndex = 15
		Me._optNewCharType_7.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optNewCharType_7.CausesValidation = True
		Me._optNewCharType_7.Enabled = True
		Me._optNewCharType_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._optNewCharType_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optNewCharType_7.TabStop = True
		Me._optNewCharType_7.Checked = False
		Me._optNewCharType_7.Visible = True
		Me._optNewCharType_7.Name = "_optNewCharType_7"
		Me._optNewCharType_6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._optNewCharType_6.BackColor = System.Drawing.Color.Black
		Me._optNewCharType_6.Text = "Druid"
		Me._optNewCharType_6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optNewCharType_6.ForeColor = System.Drawing.Color.White
		Me._optNewCharType_6.Size = New System.Drawing.Size(105, 17)
		Me._optNewCharType_6.Location = New System.Drawing.Point(56, 168)
		Me._optNewCharType_6.Appearance = System.Windows.Forms.Appearance.Button
		Me._optNewCharType_6.TabIndex = 14
		Me._optNewCharType_6.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optNewCharType_6.CausesValidation = True
		Me._optNewCharType_6.Enabled = True
		Me._optNewCharType_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._optNewCharType_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optNewCharType_6.TabStop = True
		Me._optNewCharType_6.Checked = False
		Me._optNewCharType_6.Visible = True
		Me._optNewCharType_6.Name = "_optNewCharType_6"
		Me._optNewCharType_5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._optNewCharType_5.BackColor = System.Drawing.Color.Black
		Me._optNewCharType_5.Text = "Barbarian"
		Me._optNewCharType_5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optNewCharType_5.ForeColor = System.Drawing.Color.White
		Me._optNewCharType_5.Size = New System.Drawing.Size(105, 17)
		Me._optNewCharType_5.Location = New System.Drawing.Point(56, 144)
		Me._optNewCharType_5.Appearance = System.Windows.Forms.Appearance.Button
		Me._optNewCharType_5.TabIndex = 13
		Me._optNewCharType_5.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optNewCharType_5.CausesValidation = True
		Me._optNewCharType_5.Enabled = True
		Me._optNewCharType_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._optNewCharType_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optNewCharType_5.TabStop = True
		Me._optNewCharType_5.Checked = False
		Me._optNewCharType_5.Visible = True
		Me._optNewCharType_5.Name = "_optNewCharType_5"
		Me._optNewCharType_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._optNewCharType_4.BackColor = System.Drawing.Color.Black
		Me._optNewCharType_4.Text = "Paladin"
		Me._optNewCharType_4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optNewCharType_4.ForeColor = System.Drawing.Color.White
		Me._optNewCharType_4.Size = New System.Drawing.Size(105, 17)
		Me._optNewCharType_4.Location = New System.Drawing.Point(56, 120)
		Me._optNewCharType_4.Appearance = System.Windows.Forms.Appearance.Button
		Me._optNewCharType_4.TabIndex = 12
		Me._optNewCharType_4.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optNewCharType_4.CausesValidation = True
		Me._optNewCharType_4.Enabled = True
		Me._optNewCharType_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._optNewCharType_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optNewCharType_4.TabStop = True
		Me._optNewCharType_4.Checked = False
		Me._optNewCharType_4.Visible = True
		Me._optNewCharType_4.Name = "_optNewCharType_4"
		Me._optNewCharType_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._optNewCharType_3.BackColor = System.Drawing.Color.Black
		Me._optNewCharType_3.Text = "Necromancer"
		Me._optNewCharType_3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optNewCharType_3.ForeColor = System.Drawing.Color.White
		Me._optNewCharType_3.Size = New System.Drawing.Size(105, 17)
		Me._optNewCharType_3.Location = New System.Drawing.Point(56, 96)
		Me._optNewCharType_3.Appearance = System.Windows.Forms.Appearance.Button
		Me._optNewCharType_3.TabIndex = 11
		Me._optNewCharType_3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optNewCharType_3.CausesValidation = True
		Me._optNewCharType_3.Enabled = True
		Me._optNewCharType_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._optNewCharType_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optNewCharType_3.TabStop = True
		Me._optNewCharType_3.Checked = False
		Me._optNewCharType_3.Visible = True
		Me._optNewCharType_3.Name = "_optNewCharType_3"
		Me._optNewCharType_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._optNewCharType_2.BackColor = System.Drawing.Color.Black
		Me._optNewCharType_2.Text = "Sorceress"
		Me._optNewCharType_2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optNewCharType_2.ForeColor = System.Drawing.Color.White
		Me._optNewCharType_2.Size = New System.Drawing.Size(105, 17)
		Me._optNewCharType_2.Location = New System.Drawing.Point(56, 72)
		Me._optNewCharType_2.Appearance = System.Windows.Forms.Appearance.Button
		Me._optNewCharType_2.TabIndex = 10
		Me._optNewCharType_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optNewCharType_2.CausesValidation = True
		Me._optNewCharType_2.Enabled = True
		Me._optNewCharType_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._optNewCharType_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optNewCharType_2.TabStop = True
		Me._optNewCharType_2.Checked = False
		Me._optNewCharType_2.Visible = True
		Me._optNewCharType_2.Name = "_optNewCharType_2"
		Me._optNewCharType_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._optNewCharType_1.BackColor = System.Drawing.Color.Black
		Me._optNewCharType_1.Text = "Amazon"
		Me._optNewCharType_1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optNewCharType_1.ForeColor = System.Drawing.Color.White
		Me._optNewCharType_1.Size = New System.Drawing.Size(105, 17)
		Me._optNewCharType_1.Location = New System.Drawing.Point(56, 48)
		Me._optNewCharType_1.Appearance = System.Windows.Forms.Appearance.Button
		Me._optNewCharType_1.TabIndex = 9
		Me._optNewCharType_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optNewCharType_1.CausesValidation = True
		Me._optNewCharType_1.Enabled = True
		Me._optNewCharType_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._optNewCharType_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optNewCharType_1.TabStop = True
		Me._optNewCharType_1.Checked = False
		Me._optNewCharType_1.Visible = True
		Me._optNewCharType_1.Name = "_optNewCharType_1"
		Me.lblCopy.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblCopy.BackColor = System.Drawing.Color.Black
		Me.lblCopy.Text = "Images (c) Blizzard Entertainment"
		Me.lblCopy.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCopy.ForeColor = System.Drawing.Color.White
		Me.lblCopy.Size = New System.Drawing.Size(169, 17)
		Me.lblCopy.Location = New System.Drawing.Point(152, 240)
		Me.lblCopy.TabIndex = 26
		Me.lblCopy.Enabled = True
		Me.lblCopy.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCopy.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCopy.UseMnemonic = True
		Me.lblCopy.Visible = True
		Me.lblCopy.AutoSize = False
		Me.lblCopy.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCopy.Name = "lblCopy"
		Me.lblCharName.BackColor = System.Drawing.Color.Black
		Me.lblCharName.Text = "Desired character name:"
		Me.lblCharName.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCharName.ForeColor = System.Drawing.Color.White
		Me.lblCharName.Size = New System.Drawing.Size(145, 17)
		Me.lblCharName.Location = New System.Drawing.Point(424, 104)
		Me.lblCharName.TabIndex = 24
		Me.lblCharName.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCharName.Enabled = True
		Me.lblCharName.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCharName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCharName.UseMnemonic = True
		Me.lblCharName.Visible = True
		Me.lblCharName.AutoSize = False
		Me.lblCharName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCharName.Name = "lblCharName"
		Me.imgCharPortrait.Size = New System.Drawing.Size(103, 201)
		Me.imgCharPortrait.Location = New System.Drawing.Point(184, 32)
		Me.imgCharPortrait.Image = CType(resources.GetObject("imgCharPortrait.Image"), System.Drawing.Image)
		Me.imgCharPortrait.Enabled = True
		Me.imgCharPortrait.Cursor = System.Windows.Forms.Cursors.Default
		Me.imgCharPortrait.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.imgCharPortrait.Visible = True
		Me.imgCharPortrait.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.imgCharPortrait.Name = "imgCharPortrait"
		Me.lvwChars.Size = New System.Drawing.Size(617, 273)
		Me.lvwChars.Location = New System.Drawing.Point(8, 32)
		Me.lvwChars.TabIndex = 0
		Me.lvwChars.Alignment = System.Windows.Forms.ListViewAlignment.Left
		Me.lvwChars.LabelWrap = True
		Me.lvwChars.HideSelection = False
		Me.lvwChars.LargeImageList = imlChars
		Me.lvwChars.ForeColor = System.Drawing.Color.White
		Me.lvwChars.BackColor = System.Drawing.Color.Black
		Me.lvwChars.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lvwChars.LabelEdit = True
		Me.lvwChars.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lvwChars.Name = "lvwChars"
		Me.btnUpgrade.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnUpgrade.Text = "&Upgrade This"
		Me.btnUpgrade.Size = New System.Drawing.Size(89, 20)
		Me.btnUpgrade.Location = New System.Drawing.Point(632, 184)
		Me.btnUpgrade.TabIndex = 4
		Me.btnUpgrade.Visible = False
		Me.btnUpgrade.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnUpgrade.BackColor = System.Drawing.SystemColors.Control
		Me.btnUpgrade.CausesValidation = True
		Me.btnUpgrade.Enabled = True
		Me.btnUpgrade.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnUpgrade.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnUpgrade.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnUpgrade.TabStop = True
		Me.btnUpgrade.Name = "btnUpgrade"
		Me.btnDelete.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnDelete.Text = "&Delete This"
		Me.btnDelete.Size = New System.Drawing.Size(89, 20)
		Me.btnDelete.Location = New System.Drawing.Point(632, 208)
		Me.btnDelete.TabIndex = 3
		Me.btnDelete.Visible = False
		Me.btnDelete.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnDelete.BackColor = System.Drawing.SystemColors.Control
		Me.btnDelete.CausesValidation = True
		Me.btnDelete.Enabled = True
		Me.btnDelete.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnDelete.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnDelete.TabStop = True
		Me.btnDelete.Name = "btnDelete"
		Me._lblRealm_5.BackColor = System.Drawing.Color.Black
		Me._lblRealm_5.Text = "Other Realms:"
		Me._lblRealm_5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblRealm_5.ForeColor = System.Drawing.Color.White
		Me._lblRealm_5.Size = New System.Drawing.Size(89, 17)
		Me._lblRealm_5.Location = New System.Drawing.Point(632, 112)
		Me._lblRealm_5.TabIndex = 28
		Me._lblRealm_5.Visible = False
		Me._lblRealm_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblRealm_5.Enabled = True
		Me._lblRealm_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblRealm_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblRealm_5.UseMnemonic = True
		Me._lblRealm_5.AutoSize = False
		Me._lblRealm_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblRealm_5.Name = "_lblRealm_5"
		Me._lblRealm_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblRealm_1.BackColor = System.Drawing.Color.Black
		Me._lblRealm_1.Text = "Expires:"
		Me._lblRealm_1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblRealm_1.ForeColor = System.Drawing.Color.Yellow
		Me._lblRealm_1.Size = New System.Drawing.Size(89, 57)
		Me._lblRealm_1.Location = New System.Drawing.Point(632, 248)
		Me._lblRealm_1.TabIndex = 27
		Me._lblRealm_1.Enabled = True
		Me._lblRealm_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblRealm_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblRealm_1.UseMnemonic = True
		Me._lblRealm_1.Visible = True
		Me._lblRealm_1.AutoSize = False
		Me._lblRealm_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblRealm_1.Name = "_lblRealm_1"
		Me._lblRealm_0.BackColor = System.Drawing.Color.Black
		Me._lblRealm_0.Text = "{0}{1} is a {2} {3} on Realm {4}"
		Me._lblRealm_0.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblRealm_0.ForeColor = System.Drawing.Color.White
		Me._lblRealm_0.Size = New System.Drawing.Size(513, 17)
		Me._lblRealm_0.Location = New System.Drawing.Point(8, 316)
		Me._lblRealm_0.TabIndex = 25
		Me._lblRealm_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblRealm_0.Enabled = True
		Me._lblRealm_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblRealm_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblRealm_0.UseMnemonic = True
		Me._lblRealm_0.Visible = True
		Me._lblRealm_0.AutoSize = False
		Me._lblRealm_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblRealm_0.Name = "_lblRealm_0"
		Me._lblRealm_4.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblRealm_4.BackColor = System.Drawing.Color.Black
		Me._lblRealm_4.Text = "seconds."
		Me._lblRealm_4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblRealm_4.ForeColor = System.Drawing.Color.White
		Me._lblRealm_4.Size = New System.Drawing.Size(89, 17)
		Me._lblRealm_4.Location = New System.Drawing.Point(632, 224)
		Me._lblRealm_4.TabIndex = 23
		Me._lblRealm_4.Enabled = True
		Me._lblRealm_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblRealm_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblRealm_4.UseMnemonic = True
		Me._lblRealm_4.Visible = True
		Me._lblRealm_4.AutoSize = False
		Me._lblRealm_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblRealm_4.Name = "_lblRealm_4"
		Me._lblRealm_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblRealm_3.BackColor = System.Drawing.Color.Black
		Me._lblRealm_3.Text = "#"
		Me._lblRealm_3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblRealm_3.ForeColor = System.Drawing.Color.White
		Me._lblRealm_3.Size = New System.Drawing.Size(89, 17)
		Me._lblRealm_3.Location = New System.Drawing.Point(632, 200)
		Me._lblRealm_3.TabIndex = 22
		Me._lblRealm_3.Enabled = True
		Me._lblRealm_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblRealm_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblRealm_3.UseMnemonic = True
		Me._lblRealm_3.Visible = True
		Me._lblRealm_3.AutoSize = False
		Me._lblRealm_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblRealm_3.Name = "_lblRealm_3"
		Me._lblRealm_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblRealm_2.BackColor = System.Drawing.Color.Black
		Me._lblRealm_2.Text = "Auto-choose X in"
		Me._lblRealm_2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblRealm_2.ForeColor = System.Drawing.Color.White
		Me._lblRealm_2.Size = New System.Drawing.Size(89, 49)
		Me._lblRealm_2.Location = New System.Drawing.Point(632, 152)
		Me._lblRealm_2.TabIndex = 21
		Me._lblRealm_2.Enabled = True
		Me._lblRealm_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblRealm_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblRealm_2.UseMnemonic = True
		Me._lblRealm_2.Visible = True
		Me._lblRealm_2.AutoSize = False
		Me._lblRealm_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblRealm_2.Name = "_lblRealm_2"
		Me.Controls.Add(cboOtherRealms)
		Me.Controls.Add(btnDisconnect)
		Me.Controls.Add(btnChoose)
		Me.Controls.Add(optCreateNew)
		Me.Controls.Add(optViewExisting)
		Me.Controls.Add(fraCreateNew)
		Me.Controls.Add(lvwChars)
		Me.Controls.Add(btnUpgrade)
		Me.Controls.Add(btnDelete)
		Me.Controls.Add(_lblRealm_5)
		Me.Controls.Add(_lblRealm_1)
		Me.Controls.Add(_lblRealm_0)
		Me.Controls.Add(_lblRealm_4)
		Me.Controls.Add(_lblRealm_3)
		Me.Controls.Add(_lblRealm_2)
		Me.fraCreateNew.Controls.Add(txtCharName)
		Me.fraCreateNew.Controls.Add(chkLadder)
		Me.fraCreateNew.Controls.Add(chkHardcore)
		Me.fraCreateNew.Controls.Add(chkExpansion)
		Me.fraCreateNew.Controls.Add(cmdCreate)
		Me.fraCreateNew.Controls.Add(_optNewCharType_7)
		Me.fraCreateNew.Controls.Add(_optNewCharType_6)
		Me.fraCreateNew.Controls.Add(_optNewCharType_5)
		Me.fraCreateNew.Controls.Add(_optNewCharType_4)
		Me.fraCreateNew.Controls.Add(_optNewCharType_3)
		Me.fraCreateNew.Controls.Add(_optNewCharType_2)
		Me.fraCreateNew.Controls.Add(_optNewCharType_1)
		Me.fraCreateNew.Controls.Add(lblCopy)
		Me.fraCreateNew.Controls.Add(lblCharName)
		Me.fraCreateNew.Controls.Add(imgCharPortrait)
		Me.lblRealm.SetIndex(_lblRealm_5, CType(5, Short))
		Me.lblRealm.SetIndex(_lblRealm_1, CType(1, Short))
		Me.lblRealm.SetIndex(_lblRealm_0, CType(0, Short))
		Me.lblRealm.SetIndex(_lblRealm_4, CType(4, Short))
		Me.lblRealm.SetIndex(_lblRealm_3, CType(3, Short))
		Me.lblRealm.SetIndex(_lblRealm_2, CType(2, Short))
		Me.optNewCharType.SetIndex(_optNewCharType_7, CType(7, Short))
		Me.optNewCharType.SetIndex(_optNewCharType_6, CType(6, Short))
		Me.optNewCharType.SetIndex(_optNewCharType_5, CType(5, Short))
		Me.optNewCharType.SetIndex(_optNewCharType_4, CType(4, Short))
		Me.optNewCharType.SetIndex(_optNewCharType_3, CType(3, Short))
		Me.optNewCharType.SetIndex(_optNewCharType_2, CType(2, Short))
		Me.optNewCharType.SetIndex(_optNewCharType_1, CType(1, Short))
		CType(Me.optNewCharType, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblRealm, System.ComponentModel.ISupportInitialize).EndInit()
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuPop})
		mnuPop.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuPopDelete, Me.mnuPopUpgrade})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.fraCreateNew.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class