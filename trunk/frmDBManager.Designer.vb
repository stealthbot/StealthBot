<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmDBManager
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
	Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuRename As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDelete As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSetPrimary As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuContext As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents btnCreateGame As System.Windows.Forms.Button
	Public WithEvents btnCreateClan As System.Windows.Forms.Button
	Public WithEvents icons As System.Windows.Forms.ImageList
	Public CommonDialogOpen As System.Windows.Forms.OpenFileDialog
	Public CommonDialogSave As System.Windows.Forms.SaveFileDialog
	Public CommonDialogFont As System.Windows.Forms.FontDialog
	Public CommonDialogColor As System.Windows.Forms.ColorDialog
	Public CommonDialogPrint As System.Windows.Forms.PrintDialog
	Public WithEvents btnCreateGroup As System.Windows.Forms.Button
	Public WithEvents btnCreateUser As System.Windows.Forms.Button
	Public WithEvents btnSaveForm As System.Windows.Forms.Button
	Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents trvUsers As AxvbalTreeViewLib6.AxvbalTreeView
	Public WithEvents btnSaveUser As System.Windows.Forms.Button
	Public WithEvents btnDelete As System.Windows.Forms.Button
	Public WithEvents btnRename As System.Windows.Forms.Button
	Public WithEvents txtBanMessage As System.Windows.Forms.TextBox
	Public WithEvents _lvGroups_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
	Public WithEvents lvGroups As System.Windows.Forms.ListView
	Public WithEvents txtFlags As System.Windows.Forms.TextBox
	Public WithEvents txtRank As System.Windows.Forms.TextBox
	Public WithEvents lblInherit As System.Windows.Forms.Label
	Public WithEvents lblBanMessage As System.Windows.Forms.Label
	Public WithEvents lblModifiedOn As System.Windows.Forms.Label
	Public WithEvents lblGroups As System.Windows.Forms.Label
	Public WithEvents lblModifiedBy As System.Windows.Forms.Label
	Public WithEvents lblCreatedBy As System.Windows.Forms.Label
	Public WithEvents lblFlags As System.Windows.Forms.Label
	Public WithEvents lblRank As System.Windows.Forms.Label
	Public WithEvents lblCreatedOn As System.Windows.Forms.Label
	Public WithEvents lblCreated As System.Windows.Forms.Label
	Public WithEvents lblLastMod As System.Windows.Forms.Label
	Public WithEvents frmDatabase As System.Windows.Forms.GroupBox
	Public WithEvents lblDB As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDBManager))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuContext = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuRename = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuDelete = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSetPrimary = New System.Windows.Forms.ToolStripMenuItem
		Me.btnCreateGame = New System.Windows.Forms.Button
		Me.btnCreateClan = New System.Windows.Forms.Button
		Me.icons = New System.Windows.Forms.ImageList
		Me.CommonDialogOpen = New System.Windows.Forms.OpenFileDialog
		Me.CommonDialogSave = New System.Windows.Forms.SaveFileDialog
		Me.CommonDialogFont = New System.Windows.Forms.FontDialog
		Me.CommonDialogColor = New System.Windows.Forms.ColorDialog
		Me.CommonDialogPrint = New System.Windows.Forms.PrintDialog
		Me.btnCreateGroup = New System.Windows.Forms.Button
		Me.btnCreateUser = New System.Windows.Forms.Button
		Me.btnSaveForm = New System.Windows.Forms.Button
		Me.btnCancel = New System.Windows.Forms.Button
		Me.trvUsers = New AxvbalTreeViewLib6.AxvbalTreeView
		Me.frmDatabase = New System.Windows.Forms.GroupBox
		Me.btnSaveUser = New System.Windows.Forms.Button
		Me.btnDelete = New System.Windows.Forms.Button
		Me.btnRename = New System.Windows.Forms.Button
		Me.txtBanMessage = New System.Windows.Forms.TextBox
		Me.lvGroups = New System.Windows.Forms.ListView
		Me._lvGroups_ColumnHeader_1 = New System.Windows.Forms.ColumnHeader
		Me.txtFlags = New System.Windows.Forms.TextBox
		Me.txtRank = New System.Windows.Forms.TextBox
		Me.lblInherit = New System.Windows.Forms.Label
		Me.lblBanMessage = New System.Windows.Forms.Label
		Me.lblModifiedOn = New System.Windows.Forms.Label
		Me.lblGroups = New System.Windows.Forms.Label
		Me.lblModifiedBy = New System.Windows.Forms.Label
		Me.lblCreatedBy = New System.Windows.Forms.Label
		Me.lblFlags = New System.Windows.Forms.Label
		Me.lblRank = New System.Windows.Forms.Label
		Me.lblCreatedOn = New System.Windows.Forms.Label
		Me.lblCreated = New System.Windows.Forms.Label
		Me.lblLastMod = New System.Windows.Forms.Label
		Me.lblDB = New System.Windows.Forms.Label
		Me.MainMenu1.SuspendLayout()
		Me.frmDatabase.SuspendLayout()
		Me.lvGroups.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.trvUsers, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.SystemColors.MenuText
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Database Manager"
		Me.ClientSize = New System.Drawing.Size(488, 441)
		Me.Location = New System.Drawing.Point(3, 29)
		Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmDBManager"
		Me.mnuFile.Name = "mnuFile"
		Me.mnuFile.Text = "File"
		Me.mnuFile.Visible = False
		Me.mnuFile.Checked = False
		Me.mnuFile.Enabled = True
		Me.mnuHelp.Name = "mnuHelp"
		Me.mnuHelp.Text = "Help"
		Me.mnuHelp.Visible = False
		Me.mnuHelp.Checked = False
		Me.mnuHelp.Enabled = True
		Me.mnuContext.Name = "mnuContext"
		Me.mnuContext.Text = "mnuContext"
		Me.mnuContext.Visible = False
		Me.mnuContext.Checked = False
		Me.mnuContext.Enabled = True
		Me.mnuRename.Name = "mnuRename"
		Me.mnuRename.Text = "Rename"
		Me.mnuRename.Enabled = False
		Me.mnuRename.ShortcutKeys = CType(System.Windows.Forms.Keys.F2, System.Windows.Forms.Keys)
		Me.mnuRename.Checked = False
		Me.mnuRename.Visible = True
		Me.mnuDelete.Name = "mnuDelete"
		Me.mnuDelete.Text = "Delete"
		Me.mnuDelete.Enabled = False
		Me.mnuDelete.ShortcutKeys = CType(System.Windows.Forms.Keys.Delete, System.Windows.Forms.Keys)
		Me.mnuDelete.Checked = False
		Me.mnuDelete.Visible = True
		Me.mnuSetPrimary.Name = "mnuSetPrimary"
		Me.mnuSetPrimary.Text = "Set This Primary"
		Me.mnuSetPrimary.Enabled = False
		Me.mnuSetPrimary.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.P, System.Windows.Forms.Keys)
		Me.mnuSetPrimary.Visible = False
		Me.mnuSetPrimary.Checked = False
		Me.btnCreateGame.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnCreateGame.Text = "Ga&me"
		Me.btnCreateGame.Size = New System.Drawing.Size(49, 25)
		Me.btnCreateGame.Location = New System.Drawing.Point(216, 382)
		Me.btnCreateGame.Image = CType(resources.GetObject("btnCreateGame.Image"), System.Drawing.Image)
		Me.btnCreateGame.TabIndex = 23
		Me.ToolTip1.SetToolTip(Me.btnCreateGame, "Create Group")
		Me.btnCreateGame.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnCreateGame.BackColor = System.Drawing.SystemColors.Control
		Me.btnCreateGame.CausesValidation = True
		Me.btnCreateGame.Enabled = True
		Me.btnCreateGame.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCreateGame.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCreateGame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCreateGame.TabStop = True
		Me.btnCreateGame.Name = "btnCreateGame"
		Me.btnCreateClan.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnCreateClan.Text = "C&lan"
		Me.btnCreateClan.Size = New System.Drawing.Size(49, 25)
		Me.btnCreateClan.Location = New System.Drawing.Point(168, 382)
		Me.btnCreateClan.Image = CType(resources.GetObject("btnCreateClan.Image"), System.Drawing.Image)
		Me.btnCreateClan.TabIndex = 24
		Me.ToolTip1.SetToolTip(Me.btnCreateClan, "Create Group")
		Me.btnCreateClan.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnCreateClan.BackColor = System.Drawing.SystemColors.Control
		Me.btnCreateClan.CausesValidation = True
		Me.btnCreateClan.Enabled = True
		Me.btnCreateClan.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCreateClan.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCreateClan.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCreateClan.TabStop = True
		Me.btnCreateClan.Name = "btnCreateClan"
		Me.icons.ImageSize = New System.Drawing.Size(16, 16)
		Me.icons.TransparentColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.icons.ImageStream = CType(resources.GetObject("icons.ImageStream"), System.Windows.Forms.ImageListStreamer)
		Me.icons.Images.SetKeyName(0, "")
		Me.icons.Images.SetKeyName(1, "")
		Me.icons.Images.SetKeyName(2, "")
		Me.icons.Images.SetKeyName(3, "")
		Me.icons.Images.SetKeyName(4, "")
		Me.icons.Images.SetKeyName(5, "")
		Me.icons.Images.SetKeyName(6, "")
		Me.icons.Images.SetKeyName(7, "")
		Me.icons.Images.SetKeyName(8, "")
		Me.CommonDialogOpen.Filter = "*.txt"
		Me.btnCreateGroup.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnCreateGroup.Text = "&Group"
		Me.btnCreateGroup.Size = New System.Drawing.Size(49, 25)
		Me.btnCreateGroup.Location = New System.Drawing.Point(120, 382)
		Me.btnCreateGroup.TabIndex = 2
		Me.ToolTip1.SetToolTip(Me.btnCreateGroup, "Create Group")
		Me.btnCreateGroup.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnCreateGroup.BackColor = System.Drawing.SystemColors.Control
		Me.btnCreateGroup.CausesValidation = True
		Me.btnCreateGroup.Enabled = True
		Me.btnCreateGroup.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCreateGroup.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCreateGroup.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCreateGroup.TabStop = True
		Me.btnCreateGroup.Name = "btnCreateGroup"
		Me.btnCreateUser.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnCreateUser.Text = "Create &User..."
		Me.btnCreateUser.Size = New System.Drawing.Size(113, 25)
		Me.btnCreateUser.Location = New System.Drawing.Point(8, 382)
		Me.btnCreateUser.TabIndex = 1
		Me.ToolTip1.SetToolTip(Me.btnCreateUser, "Create User")
		Me.btnCreateUser.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnCreateUser.BackColor = System.Drawing.SystemColors.Control
		Me.btnCreateUser.CausesValidation = True
		Me.btnCreateUser.Enabled = True
		Me.btnCreateUser.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCreateUser.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCreateUser.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCreateUser.TabStop = True
		Me.btnCreateUser.Name = "btnCreateUser"
		Me.btnSaveForm.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnSaveForm.Text = "Apply and Cl&ose"
		Me.AcceptButton = Me.btnSaveForm
		Me.btnSaveForm.Size = New System.Drawing.Size(89, 20)
		Me.btnSaveForm.Location = New System.Drawing.Point(392, 416)
		Me.btnSaveForm.TabIndex = 4
		Me.btnSaveForm.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnSaveForm.BackColor = System.Drawing.SystemColors.Control
		Me.btnSaveForm.CausesValidation = True
		Me.btnSaveForm.Enabled = True
		Me.btnSaveForm.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnSaveForm.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnSaveForm.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnSaveForm.TabStop = True
		Me.btnSaveForm.Name = "btnSaveForm"
		Me.btnCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.btnCancel
		Me.btnCancel.Text = "&Cancel"
		Me.btnCancel.Size = New System.Drawing.Size(49, 20)
		Me.btnCancel.Location = New System.Drawing.Point(344, 416)
		Me.btnCancel.TabIndex = 3
		Me.btnCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
		Me.btnCancel.CausesValidation = True
		Me.btnCancel.Enabled = True
		Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCancel.TabStop = True
		Me.btnCancel.Name = "btnCancel"
		trvUsers.OcxState = CType(resources.GetObject("trvUsers.OcxState"), System.Windows.Forms.AxHost.State)
		Me.trvUsers.Size = New System.Drawing.Size(257, 331)
		Me.trvUsers.Location = New System.Drawing.Point(8, 48)
		Me.trvUsers.TabIndex = 0
		Me.trvUsers.Name = "trvUsers"
		Me.frmDatabase.BackColor = System.Drawing.SystemColors.MenuText
		Me.frmDatabase.Text = "Eric[nK]"
		Me.frmDatabase.ForeColor = System.Drawing.Color.White
		Me.frmDatabase.Size = New System.Drawing.Size(202, 377)
		Me.frmDatabase.Location = New System.Drawing.Point(280, 32)
		Me.frmDatabase.TabIndex = 5
		Me.frmDatabase.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmDatabase.Enabled = True
		Me.frmDatabase.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmDatabase.Visible = True
		Me.frmDatabase.Padding = New System.Windows.Forms.Padding(0)
		Me.frmDatabase.Name = "frmDatabase"
		Me.btnSaveUser.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnSaveUser.Text = "&Save"
		Me.btnSaveUser.Enabled = False
		Me.btnSaveUser.Size = New System.Drawing.Size(57, 20)
		Me.btnSaveUser.Location = New System.Drawing.Point(128, 336)
		Me.btnSaveUser.TabIndex = 8
		Me.btnSaveUser.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnSaveUser.BackColor = System.Drawing.SystemColors.Control
		Me.btnSaveUser.CausesValidation = True
		Me.btnSaveUser.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnSaveUser.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnSaveUser.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnSaveUser.TabStop = True
		Me.btnSaveUser.Name = "btnSaveUser"
		Me.btnDelete.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnDelete.Text = "&Delete"
		Me.btnDelete.Enabled = False
		Me.btnDelete.Size = New System.Drawing.Size(57, 20)
		Me.btnDelete.Location = New System.Drawing.Point(72, 336)
		Me.btnDelete.TabIndex = 9
		Me.btnDelete.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnDelete.BackColor = System.Drawing.SystemColors.Control
		Me.btnDelete.CausesValidation = True
		Me.btnDelete.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnDelete.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnDelete.TabStop = True
		Me.btnDelete.Name = "btnDelete"
		Me.btnRename.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnRename.Text = "&Rename"
		Me.btnRename.Enabled = False
		Me.btnRename.Size = New System.Drawing.Size(57, 20)
		Me.btnRename.Location = New System.Drawing.Point(16, 336)
		Me.btnRename.TabIndex = 22
		Me.btnRename.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnRename.BackColor = System.Drawing.SystemColors.Control
		Me.btnRename.CausesValidation = True
		Me.btnRename.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnRename.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnRename.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnRename.TabStop = True
		Me.btnRename.Name = "btnRename"
		Me.txtBanMessage.AutoSize = False
		Me.txtBanMessage.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtBanMessage.ForeColor = System.Drawing.Color.White
		Me.txtBanMessage.Size = New System.Drawing.Size(169, 19)
		Me.txtBanMessage.Location = New System.Drawing.Point(16, 280)
		Me.txtBanMessage.TabIndex = 20
		Me.txtBanMessage.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtBanMessage.AcceptsReturn = True
		Me.txtBanMessage.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtBanMessage.CausesValidation = True
		Me.txtBanMessage.Enabled = True
		Me.txtBanMessage.HideSelection = True
		Me.txtBanMessage.ReadOnly = False
		Me.txtBanMessage.Maxlength = 0
		Me.txtBanMessage.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtBanMessage.MultiLine = False
		Me.txtBanMessage.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtBanMessage.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtBanMessage.TabStop = True
		Me.txtBanMessage.Visible = True
		Me.txtBanMessage.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtBanMessage.Name = "txtBanMessage"
		Me.lvGroups.Size = New System.Drawing.Size(169, 88)
		Me.lvGroups.Location = New System.Drawing.Point(16, 168)
		Me.lvGroups.TabIndex = 18
		Me.lvGroups.View = System.Windows.Forms.View.SmallIcon
		Me.lvGroups.LabelEdit = False
		Me.lvGroups.LabelWrap = True
		Me.lvGroups.HideSelection = True
		Me.lvGroups.SmallImageList = icons
		Me.lvGroups.ForeColor = System.Drawing.Color.White
		Me.lvGroups.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.lvGroups.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lvGroups.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lvGroups.Name = "lvGroups"
		Me._lvGroups_ColumnHeader_1.Width = 265
		Me.txtFlags.AutoSize = False
		Me.txtFlags.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtFlags.Enabled = False
		Me.txtFlags.ForeColor = System.Drawing.Color.White
		Me.txtFlags.Size = New System.Drawing.Size(81, 19)
		Me.txtFlags.Location = New System.Drawing.Point(104, 32)
		Me.txtFlags.Maxlength = 25
		Me.txtFlags.TabIndex = 7
		Me.txtFlags.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtFlags.AcceptsReturn = True
		Me.txtFlags.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFlags.CausesValidation = True
		Me.txtFlags.HideSelection = True
		Me.txtFlags.ReadOnly = False
		Me.txtFlags.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFlags.MultiLine = False
		Me.txtFlags.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFlags.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFlags.TabStop = True
		Me.txtFlags.Visible = True
		Me.txtFlags.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFlags.Name = "txtFlags"
		Me.txtRank.AutoSize = False
		Me.txtRank.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtRank.Enabled = False
		Me.txtRank.ForeColor = System.Drawing.Color.White
		Me.txtRank.Size = New System.Drawing.Size(81, 19)
		Me.txtRank.Location = New System.Drawing.Point(16, 32)
		Me.txtRank.Maxlength = 25
		Me.txtRank.TabIndex = 6
		Me.txtRank.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtRank.AcceptsReturn = True
		Me.txtRank.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtRank.CausesValidation = True
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
		Me.lblInherit.BackColor = System.Drawing.Color.Black
		Me.lblInherit.Text = "Inherits:"
		Me.lblInherit.ForeColor = System.Drawing.Color.White
		Me.lblInherit.Size = New System.Drawing.Size(169, 33)
		Me.lblInherit.Location = New System.Drawing.Point(16, 304)
		Me.lblInherit.TabIndex = 26
		Me.lblInherit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblInherit.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblInherit.Enabled = True
		Me.lblInherit.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInherit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInherit.UseMnemonic = True
		Me.lblInherit.Visible = True
		Me.lblInherit.AutoSize = False
		Me.lblInherit.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInherit.Name = "lblInherit"
		Me.lblBanMessage.BackColor = System.Drawing.Color.Black
		Me.lblBanMessage.Text = "Ban message:"
		Me.lblBanMessage.ForeColor = System.Drawing.Color.White
		Me.lblBanMessage.Size = New System.Drawing.Size(169, 17)
		Me.lblBanMessage.Location = New System.Drawing.Point(16, 264)
		Me.lblBanMessage.TabIndex = 21
		Me.lblBanMessage.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblBanMessage.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblBanMessage.Enabled = True
		Me.lblBanMessage.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblBanMessage.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblBanMessage.UseMnemonic = True
		Me.lblBanMessage.Visible = True
		Me.lblBanMessage.AutoSize = False
		Me.lblBanMessage.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblBanMessage.Name = "lblBanMessage"
		Me.lblModifiedOn.BackColor = System.Drawing.Color.Black
		Me.lblModifiedOn.Text = "(not applicable)"
		Me.lblModifiedOn.Font = New System.Drawing.Font("Tahoma", 6!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblModifiedOn.ForeColor = System.Drawing.Color.White
		Me.lblModifiedOn.Size = New System.Drawing.Size(161, 9)
		Me.lblModifiedOn.Location = New System.Drawing.Point(24, 120)
		Me.lblModifiedOn.TabIndex = 15
		Me.lblModifiedOn.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblModifiedOn.Enabled = True
		Me.lblModifiedOn.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblModifiedOn.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblModifiedOn.UseMnemonic = True
		Me.lblModifiedOn.Visible = True
		Me.lblModifiedOn.AutoSize = False
		Me.lblModifiedOn.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblModifiedOn.Name = "lblModifiedOn"
		Me.lblGroups.BackColor = System.Drawing.Color.Black
		Me.lblGroups.Text = "Groups:"
		Me.lblGroups.ForeColor = System.Drawing.Color.White
		Me.lblGroups.Size = New System.Drawing.Size(169, 17)
		Me.lblGroups.Location = New System.Drawing.Point(16, 152)
		Me.lblGroups.TabIndex = 19
		Me.lblGroups.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblGroups.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblGroups.Enabled = True
		Me.lblGroups.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblGroups.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblGroups.UseMnemonic = True
		Me.lblGroups.Visible = True
		Me.lblGroups.AutoSize = False
		Me.lblGroups.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblGroups.Name = "lblGroups"
		Me.lblModifiedBy.BackColor = System.Drawing.Color.Black
		Me.lblModifiedBy.Text = "(modified by)"
		Me.lblModifiedBy.Font = New System.Drawing.Font("Tahoma", 6!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblModifiedBy.ForeColor = System.Drawing.Color.White
		Me.lblModifiedBy.Size = New System.Drawing.Size(161, 9)
		Me.lblModifiedBy.Location = New System.Drawing.Point(32, 134)
		Me.lblModifiedBy.TabIndex = 17
		Me.lblModifiedBy.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblModifiedBy.Enabled = True
		Me.lblModifiedBy.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblModifiedBy.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblModifiedBy.UseMnemonic = True
		Me.lblModifiedBy.Visible = True
		Me.lblModifiedBy.AutoSize = False
		Me.lblModifiedBy.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblModifiedBy.Name = "lblModifiedBy"
		Me.lblCreatedBy.BackColor = System.Drawing.Color.Black
		Me.lblCreatedBy.Text = "(created by)"
		Me.lblCreatedBy.Font = New System.Drawing.Font("Tahoma", 6!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCreatedBy.ForeColor = System.Drawing.Color.White
		Me.lblCreatedBy.Size = New System.Drawing.Size(161, 9)
		Me.lblCreatedBy.Location = New System.Drawing.Point(32, 87)
		Me.lblCreatedBy.TabIndex = 16
		Me.lblCreatedBy.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCreatedBy.Enabled = True
		Me.lblCreatedBy.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCreatedBy.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCreatedBy.UseMnemonic = True
		Me.lblCreatedBy.Visible = True
		Me.lblCreatedBy.AutoSize = False
		Me.lblCreatedBy.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCreatedBy.Name = "lblCreatedBy"
		Me.lblFlags.BackColor = System.Drawing.Color.Black
		Me.lblFlags.Text = "Flags:"
		Me.lblFlags.ForeColor = System.Drawing.Color.White
		Me.lblFlags.Size = New System.Drawing.Size(81, 17)
		Me.lblFlags.Location = New System.Drawing.Point(104, 16)
		Me.lblFlags.TabIndex = 11
		Me.lblFlags.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblFlags.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblFlags.Enabled = True
		Me.lblFlags.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFlags.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFlags.UseMnemonic = True
		Me.lblFlags.Visible = True
		Me.lblFlags.AutoSize = False
		Me.lblFlags.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblFlags.Name = "lblFlags"
		Me.lblRank.BackColor = System.Drawing.Color.Black
		Me.lblRank.Text = "Rank (1 - 200):"
		Me.lblRank.ForeColor = System.Drawing.Color.White
		Me.lblRank.Size = New System.Drawing.Size(81, 17)
		Me.lblRank.Location = New System.Drawing.Point(16, 16)
		Me.lblRank.TabIndex = 10
		Me.lblRank.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblRank.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblRank.Enabled = True
		Me.lblRank.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblRank.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblRank.UseMnemonic = True
		Me.lblRank.Visible = True
		Me.lblRank.AutoSize = False
		Me.lblRank.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblRank.Name = "lblRank"
		Me.lblCreatedOn.BackColor = System.Drawing.Color.Black
		Me.lblCreatedOn.Text = "(not applicable)"
		Me.lblCreatedOn.Font = New System.Drawing.Font("Tahoma", 6!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCreatedOn.ForeColor = System.Drawing.Color.White
		Me.lblCreatedOn.Size = New System.Drawing.Size(161, 9)
		Me.lblCreatedOn.Location = New System.Drawing.Point(24, 74)
		Me.lblCreatedOn.TabIndex = 12
		Me.lblCreatedOn.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCreatedOn.Enabled = True
		Me.lblCreatedOn.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCreatedOn.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCreatedOn.UseMnemonic = True
		Me.lblCreatedOn.Visible = True
		Me.lblCreatedOn.AutoSize = False
		Me.lblCreatedOn.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCreatedOn.Name = "lblCreatedOn"
		Me.lblCreated.BackColor = System.Drawing.Color.Black
		Me.lblCreated.Text = "Created on:"
		Me.lblCreated.Font = New System.Drawing.Font("Tahoma", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCreated.ForeColor = System.Drawing.Color.White
		Me.lblCreated.Size = New System.Drawing.Size(169, 9)
		Me.lblCreated.Location = New System.Drawing.Point(16, 60)
		Me.lblCreated.TabIndex = 14
		Me.lblCreated.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCreated.Enabled = True
		Me.lblCreated.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCreated.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCreated.UseMnemonic = True
		Me.lblCreated.Visible = True
		Me.lblCreated.AutoSize = False
		Me.lblCreated.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCreated.Name = "lblCreated"
		Me.lblLastMod.BackColor = System.Drawing.Color.Black
		Me.lblLastMod.Text = "Last modified on:"
		Me.lblLastMod.Font = New System.Drawing.Font("Tahoma", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLastMod.ForeColor = System.Drawing.Color.White
		Me.lblLastMod.Size = New System.Drawing.Size(169, 9)
		Me.lblLastMod.Location = New System.Drawing.Point(16, 107)
		Me.lblLastMod.TabIndex = 13
		Me.lblLastMod.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblLastMod.Enabled = True
		Me.lblLastMod.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLastMod.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLastMod.UseMnemonic = True
		Me.lblLastMod.Visible = True
		Me.lblLastMod.AutoSize = False
		Me.lblLastMod.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLastMod.Name = "lblLastMod"
		Me.lblDB.BackColor = System.Drawing.Color.Black
		Me.lblDB.Text = "User Database"
		Me.lblDB.ForeColor = System.Drawing.Color.White
		Me.lblDB.Size = New System.Drawing.Size(121, 17)
		Me.lblDB.Location = New System.Drawing.Point(8, 32)
		Me.lblDB.TabIndex = 25
		Me.lblDB.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDB.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDB.Enabled = True
		Me.lblDB.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDB.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDB.UseMnemonic = True
		Me.lblDB.Visible = True
		Me.lblDB.AutoSize = False
		Me.lblDB.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDB.Name = "lblDB"
		Me.Controls.Add(btnCreateGame)
		Me.Controls.Add(btnCreateClan)
		Me.Controls.Add(btnCreateGroup)
		Me.Controls.Add(btnCreateUser)
		Me.Controls.Add(btnSaveForm)
		Me.Controls.Add(btnCancel)
		Me.Controls.Add(trvUsers)
		Me.Controls.Add(frmDatabase)
		Me.Controls.Add(lblDB)
		Me.frmDatabase.Controls.Add(btnSaveUser)
		Me.frmDatabase.Controls.Add(btnDelete)
		Me.frmDatabase.Controls.Add(btnRename)
		Me.frmDatabase.Controls.Add(txtBanMessage)
		Me.frmDatabase.Controls.Add(lvGroups)
		Me.frmDatabase.Controls.Add(txtFlags)
		Me.frmDatabase.Controls.Add(txtRank)
		Me.frmDatabase.Controls.Add(lblInherit)
		Me.frmDatabase.Controls.Add(lblBanMessage)
		Me.frmDatabase.Controls.Add(lblModifiedOn)
		Me.frmDatabase.Controls.Add(lblGroups)
		Me.frmDatabase.Controls.Add(lblModifiedBy)
		Me.frmDatabase.Controls.Add(lblCreatedBy)
		Me.frmDatabase.Controls.Add(lblFlags)
		Me.frmDatabase.Controls.Add(lblRank)
		Me.frmDatabase.Controls.Add(lblCreatedOn)
		Me.frmDatabase.Controls.Add(lblCreated)
		Me.frmDatabase.Controls.Add(lblLastMod)
		Me.lvGroups.Columns.Add(_lvGroups_ColumnHeader_1)
		CType(Me.trvUsers, System.ComponentModel.ISupportInitialize).EndInit()
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuFile, Me.mnuHelp, Me.mnuContext})
		mnuContext.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuRename, Me.mnuDelete, Me.mnuSetPrimary})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.frmDatabase.ResumeLayout(False)
		Me.lvGroups.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class