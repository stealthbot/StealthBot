<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmQuickChannel
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
	Public WithEvents _Channel_8 As System.Windows.Forms.TextBox
	Public WithEvents _Channel_7 As System.Windows.Forms.TextBox
	Public WithEvents _Channel_6 As System.Windows.Forms.TextBox
	Public WithEvents _Channel_5 As System.Windows.Forms.TextBox
	Public WithEvents _Channel_4 As System.Windows.Forms.TextBox
	Public WithEvents _Channel_3 As System.Windows.Forms.TextBox
	Public WithEvents _Channel_2 As System.Windows.Forms.TextBox
	Public WithEvents _Channel_1 As System.Windows.Forms.TextBox
	Public WithEvents _Channel_0 As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdDone As System.Windows.Forms.Button
	Public WithEvents fraSample4 As System.Windows.Forms.GroupBox
	Public WithEvents _picOptions_3 As System.Windows.Forms.Panel
	Public WithEvents fraSample3 As System.Windows.Forms.GroupBox
	Public WithEvents _picOptions_2 As System.Windows.Forms.Panel
	Public WithEvents txtUsername As System.Windows.Forms.TextBox
	Public WithEvents fraSample2 As System.Windows.Forms.GroupBox
	Public WithEvents _picOptions_1 As System.Windows.Forms.Panel
	Public WithEvents Label10 As System.Windows.Forms.Label
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Channel As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents picOptions As Microsoft.VisualBasic.Compatibility.VB6.PanelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmQuickChannel))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me._Channel_8 = New System.Windows.Forms.TextBox
		Me._Channel_7 = New System.Windows.Forms.TextBox
		Me._Channel_6 = New System.Windows.Forms.TextBox
		Me._Channel_5 = New System.Windows.Forms.TextBox
		Me._Channel_4 = New System.Windows.Forms.TextBox
		Me._Channel_3 = New System.Windows.Forms.TextBox
		Me._Channel_2 = New System.Windows.Forms.TextBox
		Me._Channel_1 = New System.Windows.Forms.TextBox
		Me._Channel_0 = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdDone = New System.Windows.Forms.Button
		Me._picOptions_3 = New System.Windows.Forms.Panel
		Me.fraSample4 = New System.Windows.Forms.GroupBox
		Me._picOptions_2 = New System.Windows.Forms.Panel
		Me.fraSample3 = New System.Windows.Forms.GroupBox
		Me._picOptions_1 = New System.Windows.Forms.Panel
		Me.fraSample2 = New System.Windows.Forms.GroupBox
		Me.txtUsername = New System.Windows.Forms.TextBox
		Me.Label10 = New System.Windows.Forms.Label
		Me.Label9 = New System.Windows.Forms.Label
		Me.Label8 = New System.Windows.Forms.Label
		Me.Label7 = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Channel = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.picOptions = New Microsoft.VisualBasic.Compatibility.VB6.PanelArray(components)
		Me._picOptions_3.SuspendLayout()
		Me._picOptions_2.SuspendLayout()
		Me._picOptions_1.SuspendLayout()
		Me.fraSample2.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Channel, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.picOptions, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.SystemColors.MenuText
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "QuickChannel Manager"
		Me.ClientSize = New System.Drawing.Size(323, 211)
		Me.Location = New System.Drawing.Point(174, 103)
		Me.KeyPreview = True
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ControlBox = True
		Me.Enabled = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmQuickChannel"
		Me._Channel_8.AutoSize = False
		Me._Channel_8.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me._Channel_8.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Channel_8.ForeColor = System.Drawing.Color.White
		Me._Channel_8.Size = New System.Drawing.Size(273, 19)
		Me._Channel_8.Location = New System.Drawing.Point(40, 160)
		Me._Channel_8.Maxlength = 31
		Me._Channel_8.TabIndex = 9
		Me._Channel_8.AcceptsReturn = True
		Me._Channel_8.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._Channel_8.CausesValidation = True
		Me._Channel_8.Enabled = True
		Me._Channel_8.HideSelection = True
		Me._Channel_8.ReadOnly = False
		Me._Channel_8.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._Channel_8.MultiLine = False
		Me._Channel_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Channel_8.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._Channel_8.TabStop = True
		Me._Channel_8.Visible = True
		Me._Channel_8.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Channel_8.Name = "_Channel_8"
		Me._Channel_7.AutoSize = False
		Me._Channel_7.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me._Channel_7.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Channel_7.ForeColor = System.Drawing.Color.White
		Me._Channel_7.Size = New System.Drawing.Size(273, 19)
		Me._Channel_7.Location = New System.Drawing.Point(40, 144)
		Me._Channel_7.Maxlength = 31
		Me._Channel_7.TabIndex = 8
		Me._Channel_7.AcceptsReturn = True
		Me._Channel_7.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._Channel_7.CausesValidation = True
		Me._Channel_7.Enabled = True
		Me._Channel_7.HideSelection = True
		Me._Channel_7.ReadOnly = False
		Me._Channel_7.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._Channel_7.MultiLine = False
		Me._Channel_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Channel_7.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._Channel_7.TabStop = True
		Me._Channel_7.Visible = True
		Me._Channel_7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Channel_7.Name = "_Channel_7"
		Me._Channel_6.AutoSize = False
		Me._Channel_6.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me._Channel_6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Channel_6.ForeColor = System.Drawing.Color.White
		Me._Channel_6.Size = New System.Drawing.Size(273, 19)
		Me._Channel_6.Location = New System.Drawing.Point(40, 128)
		Me._Channel_6.Maxlength = 31
		Me._Channel_6.TabIndex = 7
		Me._Channel_6.AcceptsReturn = True
		Me._Channel_6.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._Channel_6.CausesValidation = True
		Me._Channel_6.Enabled = True
		Me._Channel_6.HideSelection = True
		Me._Channel_6.ReadOnly = False
		Me._Channel_6.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._Channel_6.MultiLine = False
		Me._Channel_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Channel_6.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._Channel_6.TabStop = True
		Me._Channel_6.Visible = True
		Me._Channel_6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Channel_6.Name = "_Channel_6"
		Me._Channel_5.AutoSize = False
		Me._Channel_5.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me._Channel_5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Channel_5.ForeColor = System.Drawing.Color.White
		Me._Channel_5.Size = New System.Drawing.Size(273, 19)
		Me._Channel_5.Location = New System.Drawing.Point(40, 112)
		Me._Channel_5.Maxlength = 31
		Me._Channel_5.TabIndex = 6
		Me._Channel_5.AcceptsReturn = True
		Me._Channel_5.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._Channel_5.CausesValidation = True
		Me._Channel_5.Enabled = True
		Me._Channel_5.HideSelection = True
		Me._Channel_5.ReadOnly = False
		Me._Channel_5.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._Channel_5.MultiLine = False
		Me._Channel_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Channel_5.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._Channel_5.TabStop = True
		Me._Channel_5.Visible = True
		Me._Channel_5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Channel_5.Name = "_Channel_5"
		Me._Channel_4.AutoSize = False
		Me._Channel_4.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me._Channel_4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Channel_4.ForeColor = System.Drawing.Color.White
		Me._Channel_4.Size = New System.Drawing.Size(273, 19)
		Me._Channel_4.Location = New System.Drawing.Point(40, 96)
		Me._Channel_4.Maxlength = 31
		Me._Channel_4.TabIndex = 5
		Me._Channel_4.AcceptsReturn = True
		Me._Channel_4.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._Channel_4.CausesValidation = True
		Me._Channel_4.Enabled = True
		Me._Channel_4.HideSelection = True
		Me._Channel_4.ReadOnly = False
		Me._Channel_4.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._Channel_4.MultiLine = False
		Me._Channel_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Channel_4.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._Channel_4.TabStop = True
		Me._Channel_4.Visible = True
		Me._Channel_4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Channel_4.Name = "_Channel_4"
		Me._Channel_3.AutoSize = False
		Me._Channel_3.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me._Channel_3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Channel_3.ForeColor = System.Drawing.Color.White
		Me._Channel_3.Size = New System.Drawing.Size(273, 19)
		Me._Channel_3.Location = New System.Drawing.Point(40, 80)
		Me._Channel_3.Maxlength = 31
		Me._Channel_3.TabIndex = 4
		Me._Channel_3.AcceptsReturn = True
		Me._Channel_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._Channel_3.CausesValidation = True
		Me._Channel_3.Enabled = True
		Me._Channel_3.HideSelection = True
		Me._Channel_3.ReadOnly = False
		Me._Channel_3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._Channel_3.MultiLine = False
		Me._Channel_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Channel_3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._Channel_3.TabStop = True
		Me._Channel_3.Visible = True
		Me._Channel_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Channel_3.Name = "_Channel_3"
		Me._Channel_2.AutoSize = False
		Me._Channel_2.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me._Channel_2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Channel_2.ForeColor = System.Drawing.Color.White
		Me._Channel_2.Size = New System.Drawing.Size(273, 19)
		Me._Channel_2.Location = New System.Drawing.Point(40, 64)
		Me._Channel_2.Maxlength = 31
		Me._Channel_2.TabIndex = 3
		Me._Channel_2.AcceptsReturn = True
		Me._Channel_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._Channel_2.CausesValidation = True
		Me._Channel_2.Enabled = True
		Me._Channel_2.HideSelection = True
		Me._Channel_2.ReadOnly = False
		Me._Channel_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._Channel_2.MultiLine = False
		Me._Channel_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Channel_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._Channel_2.TabStop = True
		Me._Channel_2.Visible = True
		Me._Channel_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Channel_2.Name = "_Channel_2"
		Me._Channel_1.AutoSize = False
		Me._Channel_1.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me._Channel_1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Channel_1.ForeColor = System.Drawing.Color.White
		Me._Channel_1.Size = New System.Drawing.Size(273, 19)
		Me._Channel_1.Location = New System.Drawing.Point(40, 48)
		Me._Channel_1.Maxlength = 31
		Me._Channel_1.TabIndex = 2
		Me._Channel_1.AcceptsReturn = True
		Me._Channel_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._Channel_1.CausesValidation = True
		Me._Channel_1.Enabled = True
		Me._Channel_1.HideSelection = True
		Me._Channel_1.ReadOnly = False
		Me._Channel_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._Channel_1.MultiLine = False
		Me._Channel_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Channel_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._Channel_1.TabStop = True
		Me._Channel_1.Visible = True
		Me._Channel_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Channel_1.Name = "_Channel_1"
		Me._Channel_0.AutoSize = False
		Me._Channel_0.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me._Channel_0.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Channel_0.ForeColor = System.Drawing.Color.White
		Me._Channel_0.Size = New System.Drawing.Size(273, 19)
		Me._Channel_0.Location = New System.Drawing.Point(40, 32)
		Me._Channel_0.Maxlength = 31
		Me._Channel_0.TabIndex = 1
		Me._Channel_0.AcceptsReturn = True
		Me._Channel_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._Channel_0.CausesValidation = True
		Me._Channel_0.Enabled = True
		Me._Channel_0.HideSelection = True
		Me._Channel_0.ReadOnly = False
		Me._Channel_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._Channel_0.MultiLine = False
		Me._Channel_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Channel_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._Channel_0.TabStop = True
		Me._Channel_0.Visible = True
		Me._Channel_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Channel_0.Name = "_Channel_0"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "&Cancel"
		Me.cmdCancel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.Size = New System.Drawing.Size(97, 17)
		Me.cmdCancel.Location = New System.Drawing.Point(216, 184)
		Me.cmdCancel.TabIndex = 11
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdDone.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDone.Text = "&Done"
		Me.AcceptButton = Me.cmdDone
		Me.cmdDone.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDone.Size = New System.Drawing.Size(209, 17)
		Me.cmdDone.Location = New System.Drawing.Point(8, 184)
		Me.cmdDone.TabIndex = 10
		Me.cmdDone.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDone.CausesValidation = True
		Me.cmdDone.Enabled = True
		Me.cmdDone.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDone.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDone.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDone.TabStop = True
		Me.cmdDone.Name = "cmdDone"
		Me._picOptions_3.Size = New System.Drawing.Size(379, 252)
		Me._picOptions_3.Location = New System.Drawing.Point(-1333, 32)
		Me._picOptions_3.TabIndex = 13
		Me._picOptions_3.TabStop = False
		Me._picOptions_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picOptions_3.Dock = System.Windows.Forms.DockStyle.None
		Me._picOptions_3.BackColor = System.Drawing.SystemColors.Control
		Me._picOptions_3.CausesValidation = True
		Me._picOptions_3.Enabled = True
		Me._picOptions_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._picOptions_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._picOptions_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picOptions_3.Visible = True
		Me._picOptions_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picOptions_3.Name = "_picOptions_3"
		Me.fraSample4.Text = "Sample 4"
		Me.fraSample4.Size = New System.Drawing.Size(137, 119)
		Me.fraSample4.Location = New System.Drawing.Point(140, 56)
		Me.fraSample4.TabIndex = 16
		Me.fraSample4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraSample4.BackColor = System.Drawing.SystemColors.Control
		Me.fraSample4.Enabled = True
		Me.fraSample4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraSample4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraSample4.Visible = True
		Me.fraSample4.Padding = New System.Windows.Forms.Padding(0)
		Me.fraSample4.Name = "fraSample4"
		Me._picOptions_2.Size = New System.Drawing.Size(379, 252)
		Me._picOptions_2.Location = New System.Drawing.Point(-1333, 32)
		Me._picOptions_2.TabIndex = 12
		Me._picOptions_2.TabStop = False
		Me._picOptions_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picOptions_2.Dock = System.Windows.Forms.DockStyle.None
		Me._picOptions_2.BackColor = System.Drawing.SystemColors.Control
		Me._picOptions_2.CausesValidation = True
		Me._picOptions_2.Enabled = True
		Me._picOptions_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._picOptions_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._picOptions_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picOptions_2.Visible = True
		Me._picOptions_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picOptions_2.Name = "_picOptions_2"
		Me.fraSample3.Text = "Sample 3"
		Me.fraSample3.Size = New System.Drawing.Size(137, 119)
		Me.fraSample3.Location = New System.Drawing.Point(103, 45)
		Me.fraSample3.TabIndex = 15
		Me.fraSample3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraSample3.BackColor = System.Drawing.SystemColors.Control
		Me.fraSample3.Enabled = True
		Me.fraSample3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraSample3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraSample3.Visible = True
		Me.fraSample3.Padding = New System.Windows.Forms.Padding(0)
		Me.fraSample3.Name = "fraSample3"
		Me._picOptions_1.Size = New System.Drawing.Size(379, 252)
		Me._picOptions_1.Location = New System.Drawing.Point(-1333, 32)
		Me._picOptions_1.TabIndex = 0
		Me._picOptions_1.TabStop = False
		Me._picOptions_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picOptions_1.Dock = System.Windows.Forms.DockStyle.None
		Me._picOptions_1.BackColor = System.Drawing.SystemColors.Control
		Me._picOptions_1.CausesValidation = True
		Me._picOptions_1.Enabled = True
		Me._picOptions_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._picOptions_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._picOptions_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picOptions_1.Visible = True
		Me._picOptions_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picOptions_1.Name = "_picOptions_1"
		Me.fraSample2.Text = "Sample 2"
		Me.fraSample2.Size = New System.Drawing.Size(137, 119)
		Me.fraSample2.Location = New System.Drawing.Point(43, 20)
		Me.fraSample2.TabIndex = 14
		Me.fraSample2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraSample2.BackColor = System.Drawing.SystemColors.Control
		Me.fraSample2.Enabled = True
		Me.fraSample2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraSample2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraSample2.Visible = True
		Me.fraSample2.Padding = New System.Windows.Forms.Padding(0)
		Me.fraSample2.Name = "fraSample2"
		Me.txtUsername.AutoSize = False
		Me.txtUsername.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtUsername.ForeColor = System.Drawing.Color.White
		Me.txtUsername.Size = New System.Drawing.Size(105, 19)
		Me.txtUsername.Location = New System.Drawing.Point(0, 0)
		Me.txtUsername.Maxlength = 15
		Me.txtUsername.TabIndex = 17
		Me.txtUsername.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtUsername.AcceptsReturn = True
		Me.txtUsername.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtUsername.CausesValidation = True
		Me.txtUsername.Enabled = True
		Me.txtUsername.HideSelection = True
		Me.txtUsername.ReadOnly = False
		Me.txtUsername.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtUsername.MultiLine = False
		Me.txtUsername.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtUsername.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtUsername.TabStop = True
		Me.txtUsername.Visible = True
		Me.txtUsername.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtUsername.Name = "txtUsername"
		Me.Label10.BackColor = System.Drawing.Color.Black
		Me.Label10.Text = "F9"
		Me.Label10.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label10.ForeColor = System.Drawing.Color.White
		Me.Label10.Size = New System.Drawing.Size(25, 17)
		Me.Label10.Location = New System.Drawing.Point(8, 160)
		Me.Label10.TabIndex = 27
		Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label10.Enabled = True
		Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label10.UseMnemonic = True
		Me.Label10.Visible = True
		Me.Label10.AutoSize = False
		Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label10.Name = "Label10"
		Me.Label9.BackColor = System.Drawing.Color.Black
		Me.Label9.Text = "F8"
		Me.Label9.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label9.ForeColor = System.Drawing.Color.White
		Me.Label9.Size = New System.Drawing.Size(25, 17)
		Me.Label9.Location = New System.Drawing.Point(8, 144)
		Me.Label9.TabIndex = 26
		Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label9.Enabled = True
		Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label9.UseMnemonic = True
		Me.Label9.Visible = True
		Me.Label9.AutoSize = False
		Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label9.Name = "Label9"
		Me.Label8.BackColor = System.Drawing.Color.Black
		Me.Label8.Text = "F7"
		Me.Label8.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label8.ForeColor = System.Drawing.Color.White
		Me.Label8.Size = New System.Drawing.Size(25, 17)
		Me.Label8.Location = New System.Drawing.Point(8, 128)
		Me.Label8.TabIndex = 25
		Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label8.Enabled = True
		Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label8.UseMnemonic = True
		Me.Label8.Visible = True
		Me.Label8.AutoSize = False
		Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label8.Name = "Label8"
		Me.Label7.BackColor = System.Drawing.Color.Black
		Me.Label7.Text = "F6"
		Me.Label7.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label7.ForeColor = System.Drawing.Color.White
		Me.Label7.Size = New System.Drawing.Size(25, 17)
		Me.Label7.Location = New System.Drawing.Point(8, 112)
		Me.Label7.TabIndex = 24
		Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label7.Enabled = True
		Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label7.UseMnemonic = True
		Me.Label7.Visible = True
		Me.Label7.AutoSize = False
		Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label7.Name = "Label7"
		Me.Label6.BackColor = System.Drawing.Color.Black
		Me.Label6.Text = "F5"
		Me.Label6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.ForeColor = System.Drawing.Color.White
		Me.Label6.Size = New System.Drawing.Size(25, 17)
		Me.Label6.Location = New System.Drawing.Point(8, 96)
		Me.Label6.TabIndex = 23
		Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label6.Enabled = True
		Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label6.UseMnemonic = True
		Me.Label6.Visible = True
		Me.Label6.AutoSize = False
		Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label6.Name = "Label6"
		Me.Label5.BackColor = System.Drawing.Color.Black
		Me.Label5.Text = "F4"
		Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.ForeColor = System.Drawing.Color.White
		Me.Label5.Size = New System.Drawing.Size(25, 17)
		Me.Label5.Location = New System.Drawing.Point(8, 80)
		Me.Label5.TabIndex = 22
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label5.Enabled = True
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label5.Name = "Label5"
		Me.Label4.BackColor = System.Drawing.Color.Black
		Me.Label4.Text = "F3"
		Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.ForeColor = System.Drawing.Color.White
		Me.Label4.Size = New System.Drawing.Size(25, 17)
		Me.Label4.Location = New System.Drawing.Point(8, 64)
		Me.Label4.TabIndex = 21
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.Enabled = True
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label3.BackColor = System.Drawing.Color.Black
		Me.Label3.Text = "F2"
		Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.ForeColor = System.Drawing.Color.White
		Me.Label3.Size = New System.Drawing.Size(25, 17)
		Me.Label3.Location = New System.Drawing.Point(8, 48)
		Me.Label3.TabIndex = 20
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.Enabled = True
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label2.BackColor = System.Drawing.Color.Black
		Me.Label2.Text = "F1  "
		Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.ForeColor = System.Drawing.Color.White
		Me.Label2.Size = New System.Drawing.Size(25, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 32)
		Me.Label2.TabIndex = 19
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.Enabled = True
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.BackColor = System.Drawing.Color.Black
		Me.Label1.Text = "Enter the nine channels you would like on your QuickChannel list:"
		Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.ForeColor = System.Drawing.Color.White
		Me.Label1.Size = New System.Drawing.Size(313, 17)
		Me.Label1.Location = New System.Drawing.Point(8, 8)
		Me.Label1.TabIndex = 18
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.Enabled = True
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(_Channel_8)
		Me.Controls.Add(_Channel_7)
		Me.Controls.Add(_Channel_6)
		Me.Controls.Add(_Channel_5)
		Me.Controls.Add(_Channel_4)
		Me.Controls.Add(_Channel_3)
		Me.Controls.Add(_Channel_2)
		Me.Controls.Add(_Channel_1)
		Me.Controls.Add(_Channel_0)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdDone)
		Me.Controls.Add(_picOptions_3)
		Me.Controls.Add(_picOptions_2)
		Me.Controls.Add(_picOptions_1)
		Me.Controls.Add(Label10)
		Me.Controls.Add(Label9)
		Me.Controls.Add(Label8)
		Me.Controls.Add(Label7)
		Me.Controls.Add(Label6)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me._picOptions_3.Controls.Add(fraSample4)
		Me._picOptions_2.Controls.Add(fraSample3)
		Me._picOptions_1.Controls.Add(fraSample2)
		Me.fraSample2.Controls.Add(txtUsername)
		Me.Channel.SetIndex(_Channel_8, CType(8, Short))
		Me.Channel.SetIndex(_Channel_7, CType(7, Short))
		Me.Channel.SetIndex(_Channel_6, CType(6, Short))
		Me.Channel.SetIndex(_Channel_5, CType(5, Short))
		Me.Channel.SetIndex(_Channel_4, CType(4, Short))
		Me.Channel.SetIndex(_Channel_3, CType(3, Short))
		Me.Channel.SetIndex(_Channel_2, CType(2, Short))
		Me.Channel.SetIndex(_Channel_1, CType(1, Short))
		Me.Channel.SetIndex(_Channel_0, CType(0, Short))
		Me.picOptions.SetIndex(_picOptions_3, CType(3, Short))
		Me.picOptions.SetIndex(_picOptions_2, CType(2, Short))
		Me.picOptions.SetIndex(_picOptions_1, CType(1, Short))
		CType(Me.picOptions, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Channel, System.ComponentModel.ISupportInitialize).EndInit()
		Me._picOptions_3.ResumeLayout(False)
		Me._picOptions_2.ResumeLayout(False)
		Me._picOptions_1.ResumeLayout(False)
		Me.fraSample2.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class