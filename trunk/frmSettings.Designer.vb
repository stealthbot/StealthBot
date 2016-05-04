<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSettings
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
	Public WithEvents cboProfile As System.Windows.Forms.ComboBox
	Public WithEvents cmdWebsite As System.Windows.Forms.Button
	Public WithEvents cmdReadme As System.Windows.Forms.Button
	Public WithEvents cmdStepByStep As System.Windows.Forms.Button
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents tvw As AxvbalTreeViewLib6.AxvbalTreeView
	Public WithEvents txtCDKey As System.Windows.Forms.TextBox
	Public WithEvents chkSHR As System.Windows.Forms.CheckBox
	Public WithEvents chkJPN As System.Windows.Forms.CheckBox
	Public WithEvents optDRTL As System.Windows.Forms.RadioButton
	Public WithEvents optW3XP As System.Windows.Forms.RadioButton
	Public WithEvents optSTAR As System.Windows.Forms.RadioButton
	Public WithEvents optSEXP As System.Windows.Forms.RadioButton
	Public WithEvents optD2DV As System.Windows.Forms.RadioButton
	Public WithEvents optD2XP As System.Windows.Forms.RadioButton
	Public WithEvents optW2BN As System.Windows.Forms.RadioButton
	Public WithEvents optWAR3 As System.Windows.Forms.RadioButton
	Public WithEvents txtUsername As System.Windows.Forms.TextBox
	Public WithEvents txtPassword As System.Windows.Forms.TextBox
	Public WithEvents txtHomeChan As System.Windows.Forms.TextBox
	Public WithEvents cboServer As System.Windows.Forms.ComboBox
	Public WithEvents txtExpKey As System.Windows.Forms.TextBox
	Public WithEvents chkUseRealm As System.Windows.Forms.CheckBox
	Public WithEvents chkSpawn As System.Windows.Forms.CheckBox
	Public WithEvents lblAddCurrentKey As System.Windows.Forms.Label
	Public WithEvents lblManageKeys As System.Windows.Forms.Label
	Public WithEvents _Label1_10 As System.Windows.Forms.Label
	Public WithEvents _Line3_2 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Line3_1 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Line3_0 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line2 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Label1_9 As System.Windows.Forms.Label
	Public WithEvents _Line1_1 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Label1_6 As System.Windows.Forms.Label
	Public WithEvents _Label1_4 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents _Label1_3 As System.Windows.Forms.Label
	Public WithEvents _Label1_5 As System.Windows.Forms.Label
	Public WithEvents _Label1_7 As System.Windows.Forms.Label
	Public WithEvents _fraPanel_0 As System.Windows.Forms.GroupBox
	Public WithEvents chkMinimizeOnStartup As System.Windows.Forms.CheckBox
	Public WithEvents chkURLDetect As System.Windows.Forms.CheckBox
	Public WithEvents chkDisablePrefix As System.Windows.Forms.CheckBox
	Public WithEvents chkDisableSuffix As System.Windows.Forms.CheckBox
	Public WithEvents chkShowUserGameStatsIcons As System.Windows.Forms.CheckBox
	Public WithEvents chkShowUserFlagsIcons As System.Windows.Forms.CheckBox
	Public WithEvents chkNoColoring As System.Windows.Forms.CheckBox
	Public WithEvents chkNoAutocomplete As System.Windows.Forms.CheckBox
	Public WithEvents chkNoTray As System.Windows.Forms.CheckBox
	Public WithEvents chkFlash As System.Windows.Forms.CheckBox
	Public WithEvents chkUTF8 As System.Windows.Forms.CheckBox
	Public WithEvents cboTimestamp As System.Windows.Forms.ComboBox
	Public WithEvents chkSplash As System.Windows.Forms.CheckBox
	Public WithEvents chkFilter As System.Windows.Forms.CheckBox
	Public WithEvents chkJoinLeaves As System.Windows.Forms.CheckBox
	Public WithEvents Line12 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line11 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line10 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line9 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Label8_8 As System.Windows.Forms.Label
	Public WithEvents _Line1_3 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Label1_12 As System.Windows.Forms.Label
	Public WithEvents _fraPanel_2 As System.Windows.Forms.GroupBox
	Public WithEvents cmdSaveColor As System.Windows.Forms.Button
	Public WithEvents cboColorList As System.Windows.Forms.ComboBox
	Public WithEvents txtValue As System.Windows.Forms.TextBox
	Public WithEvents cmdColorPicker As System.Windows.Forms.Button
	Public WithEvents txtR As System.Windows.Forms.TextBox
	Public WithEvents txtG As System.Windows.Forms.TextBox
	Public WithEvents txtB As System.Windows.Forms.TextBox
	Public WithEvents cmdGetRGB As System.Windows.Forms.Button
	Public WithEvents cmdImport As System.Windows.Forms.Button
	Public WithEvents cmdDefaults As System.Windows.Forms.Button
	Public WithEvents cmdExport As System.Windows.Forms.Button
	Public WithEvents txtHTML As System.Windows.Forms.TextBox
	Public WithEvents cmdHTMLGen As System.Windows.Forms.Button
	Public WithEvents txtChanFont As System.Windows.Forms.TextBox
	Public WithEvents txtChatFont As System.Windows.Forms.TextBox
	Public WithEvents txtChanSize As System.Windows.Forms.TextBox
	Public WithEvents txtChatSize As System.Windows.Forms.TextBox
	Public cDLGOpen As System.Windows.Forms.OpenFileDialog
	Public cDLGSave As System.Windows.Forms.SaveFileDialog
	Public cDLGColor As System.Windows.Forms.ColorDialog
	Public WithEvents lblCurrentValue As System.Windows.Forms.Label
	Public WithEvents lblEg As System.Windows.Forms.Label
	Public WithEvents _Label1_14 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents _Label8_5 As System.Windows.Forms.Label
	Public WithEvents _Label8_0 As System.Windows.Forms.Label
	Public WithEvents _Label8_2 As System.Windows.Forms.Label
	Public WithEvents _Label8_7 As System.Windows.Forms.Label
	Public WithEvents _Label8_1 As System.Windows.Forms.Label
	Public WithEvents _Label8_3 As System.Windows.Forms.Label
	Public WithEvents _Label8_4 As System.Windows.Forms.Label
	Public WithEvents _Line1_4 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Label1_13 As System.Windows.Forms.Label
	Public WithEvents lblColorStatus As System.Windows.Forms.Label
	Public WithEvents _fraPanel_3 As System.Windows.Forms.GroupBox
	Public WithEvents chkPhraseKick As System.Windows.Forms.CheckBox
	Public WithEvents chkQuietKick As System.Windows.Forms.CheckBox
	Public WithEvents chkPingBan As System.Windows.Forms.CheckBox
	Public WithEvents txtPingLevel As System.Windows.Forms.TextBox
	Public WithEvents chkBanEvasion As System.Windows.Forms.CheckBox
	Public WithEvents chkIdleKick As System.Windows.Forms.CheckBox
	Public WithEvents chkPeonbans As System.Windows.Forms.CheckBox
	Public WithEvents txtLevelBanMsg As System.Windows.Forms.TextBox
	Public WithEvents _chkCBan_5 As System.Windows.Forms.CheckBox
	Public WithEvents _chkCBan_3 As System.Windows.Forms.CheckBox
	Public WithEvents chkPhrasebans As System.Windows.Forms.CheckBox
	Public WithEvents chkIPBans As System.Windows.Forms.CheckBox
	Public WithEvents _chkCBan_0 As System.Windows.Forms.CheckBox
	Public WithEvents _chkCBan_1 As System.Windows.Forms.CheckBox
	Public WithEvents _chkCBan_2 As System.Windows.Forms.CheckBox
	Public WithEvents _chkCBan_6 As System.Windows.Forms.CheckBox
	Public WithEvents _chkCBan_4 As System.Windows.Forms.CheckBox
	Public WithEvents chkQuiet As System.Windows.Forms.CheckBox
	Public WithEvents txtProtectMsg As System.Windows.Forms.TextBox
	Public WithEvents chkProtect As System.Windows.Forms.CheckBox
	Public WithEvents chkKOY As System.Windows.Forms.CheckBox
	Public WithEvents txtBanW3 As System.Windows.Forms.TextBox
	Public WithEvents chkPlugban As System.Windows.Forms.CheckBox
	Public WithEvents txtBanD2 As System.Windows.Forms.TextBox
	Public WithEvents chkIdlebans As System.Windows.Forms.CheckBox
	Public WithEvents txtIdleBanDelay As System.Windows.Forms.TextBox
	Public WithEvents _lblMod_7 As System.Windows.Forms.Label
	Public WithEvents Line16 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line15 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line14 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line13 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _lblMod_0 As System.Windows.Forms.Label
	Public WithEvents _lblMod_2 As System.Windows.Forms.Label
	Public WithEvents _lblMod_3 As System.Windows.Forms.Label
	Public WithEvents _lblMod_5 As System.Windows.Forms.Label
	Public WithEvents _lblMod_1 As System.Windows.Forms.Label
	Public WithEvents _lblMod_4 As System.Windows.Forms.Label
	Public WithEvents _lblMod_6 As System.Windows.Forms.Label
	Public WithEvents _Line1_5 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Label1_15 As System.Windows.Forms.Label
	Public WithEvents _fraPanel_4 As System.Windows.Forms.GroupBox
	Public WithEvents optMsg As System.Windows.Forms.RadioButton
	Public WithEvents optQuote As System.Windows.Forms.RadioButton
	Public WithEvents optUptime As System.Windows.Forms.RadioButton
	Public WithEvents optMP3 As System.Windows.Forms.RadioButton
	Public WithEvents chkIdles As System.Windows.Forms.CheckBox
	Public WithEvents txtIdleMsg As System.Windows.Forms.TextBox
	Public WithEvents txtIdleWait As System.Windows.Forms.TextBox
	Public WithEvents _Line5_4 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Line5_3 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Line5_2 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Line5_1 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Line5_0 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _lblIdle_3 As System.Windows.Forms.Label
	Public WithEvents _lblIdle_0 As System.Windows.Forms.Label
	Public WithEvents _lblIdle_1 As System.Windows.Forms.Label
	Public WithEvents _Line1_7 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents lblIdleVars As System.Windows.Forms.Label
	Public WithEvents _Label1_17 As System.Windows.Forms.Label
	Public WithEvents _fraPanel_6 As System.Windows.Forms.GroupBox
	Public WithEvents txtOwner As System.Windows.Forms.TextBox
	Public WithEvents txtTrigger As System.Windows.Forms.TextBox
	Public WithEvents chkD2Naming As System.Windows.Forms.CheckBox
	Public WithEvents _optNaming_3 As System.Windows.Forms.RadioButton
	Public WithEvents _optNaming_2 As System.Windows.Forms.RadioButton
	Public WithEvents _optNaming_1 As System.Windows.Forms.RadioButton
	Public WithEvents _optNaming_0 As System.Windows.Forms.RadioButton
	Public WithEvents chkShowOffline As System.Windows.Forms.CheckBox
	Public WithEvents txtBackupChan As System.Windows.Forms.TextBox
	Public WithEvents chkAllowMP3 As System.Windows.Forms.CheckBox
	Public WithEvents chkWhisperCmds As System.Windows.Forms.CheckBox
	Public WithEvents chkPAmp As System.Windows.Forms.CheckBox
	Public WithEvents chkMail As System.Windows.Forms.CheckBox
	Public WithEvents chkBackup As System.Windows.Forms.CheckBox
	Public WithEvents _Label1_19 As System.Windows.Forms.Label
	Public WithEvents _Label1_8 As System.Windows.Forms.Label
	Public WithEvents _Label8_6 As System.Windows.Forms.Label
	Public WithEvents _Label8_10 As System.Windows.Forms.Label
	Public WithEvents _Line1_9 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Line1_8 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Label1_18 As System.Windows.Forms.Label
	Public WithEvents _fraPanel_7 As System.Windows.Forms.GroupBox
	Public WithEvents chkConnectOnStartup As System.Windows.Forms.CheckBox
	Public WithEvents txtEmail As System.Windows.Forms.TextBox
	Public WithEvents cboBNLSServer As System.Windows.Forms.ComboBox
	Public WithEvents txtReconDelay As System.Windows.Forms.TextBox
	Public WithEvents optSocks5 As System.Windows.Forms.RadioButton
	Public WithEvents optSocks4 As System.Windows.Forms.RadioButton
	Public WithEvents txtProxyPort As System.Windows.Forms.TextBox
	Public WithEvents txtProxyIP As System.Windows.Forms.TextBox
	Public WithEvents chkUseProxies As System.Windows.Forms.CheckBox
	Public WithEvents chkUDP As System.Windows.Forms.CheckBox
	Public WithEvents cboSpoof As System.Windows.Forms.ComboBox
	Public WithEvents cboConnMethod As System.Windows.Forms.ComboBox
	Public WithEvents Line7 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _lbl5_13 As System.Windows.Forms.Label
	Public WithEvents _lbl5_1 As System.Windows.Forms.Label
	Public WithEvents Line6 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _lbl5_12 As System.Windows.Forms.Label
	Public WithEvents _lbl5_11 As System.Windows.Forms.Label
	Public WithEvents _lbl5_10 As System.Windows.Forms.Label
	Public WithEvents _lbl5_9 As System.Windows.Forms.Label
	Public WithEvents _lbl5_8 As System.Windows.Forms.Label
	Public WithEvents Line4 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _lbl5_0 As System.Windows.Forms.Label
	Public WithEvents lblHashPath As System.Windows.Forms.Label
	Public WithEvents _lbl5_2 As System.Windows.Forms.Label
	Public WithEvents _lbl5_3 As System.Windows.Forms.Label
	Public WithEvents _Line1_2 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Label1_11 As System.Windows.Forms.Label
	Public WithEvents _fraPanel_1 As System.Windows.Forms.GroupBox
	Public WithEvents chkWhisperGreet As System.Windows.Forms.CheckBox
	Public WithEvents txtGreetMsg As System.Windows.Forms.TextBox
	Public WithEvents chkGreetMsg As System.Windows.Forms.CheckBox
	Public WithEvents lblGreetVars As System.Windows.Forms.Label
	Public WithEvents _lblIdle_2 As System.Windows.Forms.Label
	Public WithEvents _Line1_6 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Label1_16 As System.Windows.Forms.Label
	Public WithEvents _fraPanel_5 As System.Windows.Forms.GroupBox
	Public WithEvents chkLogDBActions As System.Windows.Forms.CheckBox
	Public WithEvents chkLogAllCommands As System.Windows.Forms.CheckBox
	Public WithEvents txtMaxLogSize As System.Windows.Forms.TextBox
	Public WithEvents cboLogging As System.Windows.Forms.ComboBox
	Public WithEvents txtMaxBackLogSize As System.Windows.Forms.TextBox
	Public WithEvents lbl9 As System.Windows.Forms.Label
	Public WithEvents Line8 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _lbl5_5 As System.Windows.Forms.Label
	Public WithEvents _lbl5_4 As System.Windows.Forms.Label
	Public WithEvents _lbl5_6 As System.Windows.Forms.Label
	Public WithEvents _lbl5_7 As System.Windows.Forms.Label
	Public WithEvents lblBacklogSize As System.Windows.Forms.Label
	Public WithEvents lblBacklog As System.Windows.Forms.Label
	Public WithEvents _Label1_20 As System.Windows.Forms.Label
	Public WithEvents _Line1_10 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _fraPanel_9 As System.Windows.Forms.GroupBox
	Public WithEvents lblSplash As System.Windows.Forms.Label
	Public WithEvents _Line1_0 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _fraPanel_8 As System.Windows.Forms.GroupBox
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents Label8 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents Line1 As LineShapeArray
	Public WithEvents Line3 As LineShapeArray
	Public WithEvents Line5 As LineShapeArray
	Public WithEvents chkCBan As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	Public WithEvents fraPanel As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
	Public WithEvents lbl5 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblIdle As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblMod As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents optNaming As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents ShapeContainer10 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	Public WithEvents ShapeContainer9 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	Public WithEvents ShapeContainer8 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	Public WithEvents ShapeContainer7 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	Public WithEvents ShapeContainer6 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	Public WithEvents ShapeContainer5 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	Public WithEvents ShapeContainer4 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	Public WithEvents ShapeContainer3 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	Public WithEvents ShapeContainer2 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSettings))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ShapeContainer10 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.ShapeContainer9 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.ShapeContainer8 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.ShapeContainer7 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.ShapeContainer6 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.ShapeContainer5 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.ShapeContainer4 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.ShapeContainer3 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.ShapeContainer2 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.cboProfile = New System.Windows.Forms.ComboBox
		Me.cmdWebsite = New System.Windows.Forms.Button
		Me.cmdReadme = New System.Windows.Forms.Button
		Me.cmdStepByStep = New System.Windows.Forms.Button
		Me.cmdSave = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.tvw = New AxvbalTreeViewLib6.AxvbalTreeView
		Me._fraPanel_0 = New System.Windows.Forms.GroupBox
		Me.txtCDKey = New System.Windows.Forms.TextBox
		Me.chkSHR = New System.Windows.Forms.CheckBox
		Me.chkJPN = New System.Windows.Forms.CheckBox
		Me.optDRTL = New System.Windows.Forms.RadioButton
		Me.optW3XP = New System.Windows.Forms.RadioButton
		Me.optSTAR = New System.Windows.Forms.RadioButton
		Me.optSEXP = New System.Windows.Forms.RadioButton
		Me.optD2DV = New System.Windows.Forms.RadioButton
		Me.optD2XP = New System.Windows.Forms.RadioButton
		Me.optW2BN = New System.Windows.Forms.RadioButton
		Me.optWAR3 = New System.Windows.Forms.RadioButton
		Me.txtUsername = New System.Windows.Forms.TextBox
		Me.txtPassword = New System.Windows.Forms.TextBox
		Me.txtHomeChan = New System.Windows.Forms.TextBox
		Me.cboServer = New System.Windows.Forms.ComboBox
		Me.txtExpKey = New System.Windows.Forms.TextBox
		Me.chkUseRealm = New System.Windows.Forms.CheckBox
		Me.chkSpawn = New System.Windows.Forms.CheckBox
		Me.lblAddCurrentKey = New System.Windows.Forms.Label
		Me.lblManageKeys = New System.Windows.Forms.Label
		Me._Label1_10 = New System.Windows.Forms.Label
		Me._Line3_2 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Line3_1 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Line3_0 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line2 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Label1_9 = New System.Windows.Forms.Label
		Me._Line1_1 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Label1_6 = New System.Windows.Forms.Label
		Me._Label1_4 = New System.Windows.Forms.Label
		Me._Label1_1 = New System.Windows.Forms.Label
		Me._Label1_2 = New System.Windows.Forms.Label
		Me._Label1_3 = New System.Windows.Forms.Label
		Me._Label1_5 = New System.Windows.Forms.Label
		Me._Label1_7 = New System.Windows.Forms.Label
		Me._fraPanel_2 = New System.Windows.Forms.GroupBox
		Me.chkMinimizeOnStartup = New System.Windows.Forms.CheckBox
		Me.chkURLDetect = New System.Windows.Forms.CheckBox
		Me.chkDisablePrefix = New System.Windows.Forms.CheckBox
		Me.chkDisableSuffix = New System.Windows.Forms.CheckBox
		Me.chkShowUserGameStatsIcons = New System.Windows.Forms.CheckBox
		Me.chkShowUserFlagsIcons = New System.Windows.Forms.CheckBox
		Me.chkNoColoring = New System.Windows.Forms.CheckBox
		Me.chkNoAutocomplete = New System.Windows.Forms.CheckBox
		Me.chkNoTray = New System.Windows.Forms.CheckBox
		Me.chkFlash = New System.Windows.Forms.CheckBox
		Me.chkUTF8 = New System.Windows.Forms.CheckBox
		Me.cboTimestamp = New System.Windows.Forms.ComboBox
		Me.chkSplash = New System.Windows.Forms.CheckBox
		Me.chkFilter = New System.Windows.Forms.CheckBox
		Me.chkJoinLeaves = New System.Windows.Forms.CheckBox
		Me.Line12 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line11 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line10 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line9 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Label8_8 = New System.Windows.Forms.Label
		Me._Line1_3 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Label1_12 = New System.Windows.Forms.Label
		Me._fraPanel_3 = New System.Windows.Forms.GroupBox
		Me.cmdSaveColor = New System.Windows.Forms.Button
		Me.cboColorList = New System.Windows.Forms.ComboBox
		Me.txtValue = New System.Windows.Forms.TextBox
		Me.cmdColorPicker = New System.Windows.Forms.Button
		Me.txtR = New System.Windows.Forms.TextBox
		Me.txtG = New System.Windows.Forms.TextBox
		Me.txtB = New System.Windows.Forms.TextBox
		Me.cmdGetRGB = New System.Windows.Forms.Button
		Me.cmdImport = New System.Windows.Forms.Button
		Me.cmdDefaults = New System.Windows.Forms.Button
		Me.cmdExport = New System.Windows.Forms.Button
		Me.txtHTML = New System.Windows.Forms.TextBox
		Me.cmdHTMLGen = New System.Windows.Forms.Button
		Me.txtChanFont = New System.Windows.Forms.TextBox
		Me.txtChatFont = New System.Windows.Forms.TextBox
		Me.txtChanSize = New System.Windows.Forms.TextBox
		Me.txtChatSize = New System.Windows.Forms.TextBox
		Me.cDLGOpen = New System.Windows.Forms.OpenFileDialog
		Me.cDLGSave = New System.Windows.Forms.SaveFileDialog
		Me.cDLGColor = New System.Windows.Forms.ColorDialog
		Me.lblCurrentValue = New System.Windows.Forms.Label
		Me.lblEg = New System.Windows.Forms.Label
		Me._Label1_14 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.Label7 = New System.Windows.Forms.Label
		Me._Label8_5 = New System.Windows.Forms.Label
		Me._Label8_0 = New System.Windows.Forms.Label
		Me._Label8_2 = New System.Windows.Forms.Label
		Me._Label8_7 = New System.Windows.Forms.Label
		Me._Label8_1 = New System.Windows.Forms.Label
		Me._Label8_3 = New System.Windows.Forms.Label
		Me._Label8_4 = New System.Windows.Forms.Label
		Me._Line1_4 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Label1_13 = New System.Windows.Forms.Label
		Me.lblColorStatus = New System.Windows.Forms.Label
		Me._fraPanel_4 = New System.Windows.Forms.GroupBox
		Me.chkPhraseKick = New System.Windows.Forms.CheckBox
		Me.chkQuietKick = New System.Windows.Forms.CheckBox
		Me.chkPingBan = New System.Windows.Forms.CheckBox
		Me.txtPingLevel = New System.Windows.Forms.TextBox
		Me.chkBanEvasion = New System.Windows.Forms.CheckBox
		Me.chkIdleKick = New System.Windows.Forms.CheckBox
		Me.chkPeonbans = New System.Windows.Forms.CheckBox
		Me.txtLevelBanMsg = New System.Windows.Forms.TextBox
		Me._chkCBan_5 = New System.Windows.Forms.CheckBox
		Me._chkCBan_3 = New System.Windows.Forms.CheckBox
		Me.chkPhrasebans = New System.Windows.Forms.CheckBox
		Me.chkIPBans = New System.Windows.Forms.CheckBox
		Me._chkCBan_0 = New System.Windows.Forms.CheckBox
		Me._chkCBan_1 = New System.Windows.Forms.CheckBox
		Me._chkCBan_2 = New System.Windows.Forms.CheckBox
		Me._chkCBan_6 = New System.Windows.Forms.CheckBox
		Me._chkCBan_4 = New System.Windows.Forms.CheckBox
		Me.chkQuiet = New System.Windows.Forms.CheckBox
		Me.txtProtectMsg = New System.Windows.Forms.TextBox
		Me.chkProtect = New System.Windows.Forms.CheckBox
		Me.chkKOY = New System.Windows.Forms.CheckBox
		Me.txtBanW3 = New System.Windows.Forms.TextBox
		Me.chkPlugban = New System.Windows.Forms.CheckBox
		Me.txtBanD2 = New System.Windows.Forms.TextBox
		Me.chkIdlebans = New System.Windows.Forms.CheckBox
		Me.txtIdleBanDelay = New System.Windows.Forms.TextBox
		Me._lblMod_7 = New System.Windows.Forms.Label
		Me.Line16 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line15 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line14 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line13 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._lblMod_0 = New System.Windows.Forms.Label
		Me._lblMod_2 = New System.Windows.Forms.Label
		Me._lblMod_3 = New System.Windows.Forms.Label
		Me._lblMod_5 = New System.Windows.Forms.Label
		Me._lblMod_1 = New System.Windows.Forms.Label
		Me._lblMod_4 = New System.Windows.Forms.Label
		Me._lblMod_6 = New System.Windows.Forms.Label
		Me._Line1_5 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Label1_15 = New System.Windows.Forms.Label
		Me._fraPanel_6 = New System.Windows.Forms.GroupBox
		Me.optMsg = New System.Windows.Forms.RadioButton
		Me.optQuote = New System.Windows.Forms.RadioButton
		Me.optUptime = New System.Windows.Forms.RadioButton
		Me.optMP3 = New System.Windows.Forms.RadioButton
		Me.chkIdles = New System.Windows.Forms.CheckBox
		Me.txtIdleMsg = New System.Windows.Forms.TextBox
		Me.txtIdleWait = New System.Windows.Forms.TextBox
		Me._Line5_4 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Line5_3 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Line5_2 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Line5_1 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Line5_0 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._lblIdle_3 = New System.Windows.Forms.Label
		Me._lblIdle_0 = New System.Windows.Forms.Label
		Me._lblIdle_1 = New System.Windows.Forms.Label
		Me._Line1_7 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.lblIdleVars = New System.Windows.Forms.Label
		Me._Label1_17 = New System.Windows.Forms.Label
		Me._fraPanel_7 = New System.Windows.Forms.GroupBox
		Me.txtOwner = New System.Windows.Forms.TextBox
		Me.txtTrigger = New System.Windows.Forms.TextBox
		Me.chkD2Naming = New System.Windows.Forms.CheckBox
		Me._optNaming_3 = New System.Windows.Forms.RadioButton
		Me._optNaming_2 = New System.Windows.Forms.RadioButton
		Me._optNaming_1 = New System.Windows.Forms.RadioButton
		Me._optNaming_0 = New System.Windows.Forms.RadioButton
		Me.chkShowOffline = New System.Windows.Forms.CheckBox
		Me.txtBackupChan = New System.Windows.Forms.TextBox
		Me.chkAllowMP3 = New System.Windows.Forms.CheckBox
		Me.chkWhisperCmds = New System.Windows.Forms.CheckBox
		Me.chkPAmp = New System.Windows.Forms.CheckBox
		Me.chkMail = New System.Windows.Forms.CheckBox
		Me.chkBackup = New System.Windows.Forms.CheckBox
		Me._Label1_19 = New System.Windows.Forms.Label
		Me._Label1_8 = New System.Windows.Forms.Label
		Me._Label8_6 = New System.Windows.Forms.Label
		Me._Label8_10 = New System.Windows.Forms.Label
		Me._Line1_9 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Line1_8 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Label1_18 = New System.Windows.Forms.Label
		Me._fraPanel_1 = New System.Windows.Forms.GroupBox
		Me.chkConnectOnStartup = New System.Windows.Forms.CheckBox
		Me.txtEmail = New System.Windows.Forms.TextBox
		Me.cboBNLSServer = New System.Windows.Forms.ComboBox
		Me.txtReconDelay = New System.Windows.Forms.TextBox
		Me.optSocks5 = New System.Windows.Forms.RadioButton
		Me.optSocks4 = New System.Windows.Forms.RadioButton
		Me.txtProxyPort = New System.Windows.Forms.TextBox
		Me.txtProxyIP = New System.Windows.Forms.TextBox
		Me.chkUseProxies = New System.Windows.Forms.CheckBox
		Me.chkUDP = New System.Windows.Forms.CheckBox
		Me.cboSpoof = New System.Windows.Forms.ComboBox
		Me.cboConnMethod = New System.Windows.Forms.ComboBox
		Me.Line7 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._lbl5_13 = New System.Windows.Forms.Label
		Me._lbl5_1 = New System.Windows.Forms.Label
		Me.Line6 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._lbl5_12 = New System.Windows.Forms.Label
		Me._lbl5_11 = New System.Windows.Forms.Label
		Me._lbl5_10 = New System.Windows.Forms.Label
		Me._lbl5_9 = New System.Windows.Forms.Label
		Me._lbl5_8 = New System.Windows.Forms.Label
		Me.Line4 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._lbl5_0 = New System.Windows.Forms.Label
		Me.lblHashPath = New System.Windows.Forms.Label
		Me._lbl5_2 = New System.Windows.Forms.Label
		Me._lbl5_3 = New System.Windows.Forms.Label
		Me._Line1_2 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Label1_11 = New System.Windows.Forms.Label
		Me._fraPanel_5 = New System.Windows.Forms.GroupBox
		Me.chkWhisperGreet = New System.Windows.Forms.CheckBox
		Me.txtGreetMsg = New System.Windows.Forms.TextBox
		Me.chkGreetMsg = New System.Windows.Forms.CheckBox
		Me.lblGreetVars = New System.Windows.Forms.Label
		Me._lblIdle_2 = New System.Windows.Forms.Label
		Me._Line1_6 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Label1_16 = New System.Windows.Forms.Label
		Me._fraPanel_9 = New System.Windows.Forms.GroupBox
		Me.chkLogDBActions = New System.Windows.Forms.CheckBox
		Me.chkLogAllCommands = New System.Windows.Forms.CheckBox
		Me.txtMaxLogSize = New System.Windows.Forms.TextBox
		Me.cboLogging = New System.Windows.Forms.ComboBox
		Me.txtMaxBackLogSize = New System.Windows.Forms.TextBox
		Me.lbl9 = New System.Windows.Forms.Label
		Me.Line8 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._lbl5_5 = New System.Windows.Forms.Label
		Me._lbl5_4 = New System.Windows.Forms.Label
		Me._lbl5_6 = New System.Windows.Forms.Label
		Me._lbl5_7 = New System.Windows.Forms.Label
		Me.lblBacklogSize = New System.Windows.Forms.Label
		Me.lblBacklog = New System.Windows.Forms.Label
		Me._Label1_20 = New System.Windows.Forms.Label
		Me._Line1_10 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._fraPanel_8 = New System.Windows.Forms.GroupBox
		Me.lblSplash = New System.Windows.Forms.Label
		Me._Line1_0 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._Label1_0 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Label8 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Line1 = New LineShapeArray(components)
		Me.Line3 = New LineShapeArray(components)
		Me.Line5 = New LineShapeArray(components)
		Me.chkCBan = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.fraPanel = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(components)
		Me.lbl5 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblIdle = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblMod = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.optNaming = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me._fraPanel_0.SuspendLayout()
		Me._fraPanel_2.SuspendLayout()
		Me._fraPanel_3.SuspendLayout()
		Me._fraPanel_4.SuspendLayout()
		Me._fraPanel_6.SuspendLayout()
		Me._fraPanel_7.SuspendLayout()
		Me._fraPanel_1.SuspendLayout()
		Me._fraPanel_5.SuspendLayout()
		Me._fraPanel_9.SuspendLayout()
		Me._fraPanel_8.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.tvw, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label8, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Line3, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Line5, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.chkCBan, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.fraPanel, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lbl5, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblIdle, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblMod, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.optNaming, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.Color.Black
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "StealthBot Settings"
		Me.ClientSize = New System.Drawing.Size(649, 354)
		Me.Location = New System.Drawing.Point(105, 129)
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
		Me.Name = "frmSettings"
		Me.cboProfile.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.cboProfile.Enabled = False
		Me.cboProfile.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboProfile.ForeColor = System.Drawing.Color.White
		Me.cboProfile.Size = New System.Drawing.Size(185, 21)
		Me.cboProfile.Location = New System.Drawing.Point(8, 8)
		Me.cboProfile.TabIndex = 191
		Me.cboProfile.Text = "Profile Selector"
		Me.cboProfile.CausesValidation = True
		Me.cboProfile.IntegralHeight = True
		Me.cboProfile.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboProfile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboProfile.Sorted = False
		Me.cboProfile.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboProfile.TabStop = True
		Me.cboProfile.Visible = True
		Me.cboProfile.Name = "cboProfile"
		Me.cmdWebsite.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdWebsite.Text = "&Forum"
		Me.cmdWebsite.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdWebsite.Size = New System.Drawing.Size(57, 17)
		Me.cmdWebsite.Location = New System.Drawing.Point(256, 328)
		Me.cmdWebsite.TabIndex = 130
		Me.cmdWebsite.BackColor = System.Drawing.SystemColors.Control
		Me.cmdWebsite.CausesValidation = True
		Me.cmdWebsite.Enabled = True
		Me.cmdWebsite.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdWebsite.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdWebsite.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdWebsite.TabStop = True
		Me.cmdWebsite.Name = "cmdWebsite"
		Me.cmdReadme.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdReadme.Text = "&Wiki"
		Me.cmdReadme.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdReadme.Size = New System.Drawing.Size(57, 17)
		Me.cmdReadme.Location = New System.Drawing.Point(200, 328)
		Me.cmdReadme.TabIndex = 129
		Me.cmdReadme.BackColor = System.Drawing.SystemColors.Control
		Me.cmdReadme.CausesValidation = True
		Me.cmdReadme.Enabled = True
		Me.cmdReadme.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdReadme.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdReadme.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdReadme.TabStop = True
		Me.cmdReadme.Name = "cmdReadme"
		Me.cmdStepByStep.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdStepByStep.Text = "&Step-By-Step Configuration"
		Me.cmdStepByStep.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdStepByStep.Size = New System.Drawing.Size(169, 17)
		Me.cmdStepByStep.Location = New System.Drawing.Point(320, 328)
		Me.cmdStepByStep.TabIndex = 131
		Me.cmdStepByStep.BackColor = System.Drawing.SystemColors.Control
		Me.cmdStepByStep.CausesValidation = True
		Me.cmdStepByStep.Enabled = True
		Me.cmdStepByStep.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdStepByStep.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdStepByStep.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdStepByStep.TabStop = True
		Me.cmdStepByStep.Name = "cmdStepByStep"
		Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSave.Text = "Apply and Cl&ose"
		Me.AcceptButton = Me.cmdSave
		Me.cmdSave.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSave.Size = New System.Drawing.Size(89, 17)
		Me.cmdSave.Location = New System.Drawing.Point(552, 328)
		Me.cmdSave.TabIndex = 132
		Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSave.CausesValidation = True
		Me.cmdSave.Enabled = True
		Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSave.TabStop = True
		Me.cmdSave.Name = "cmdSave"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "&Cancel"
		Me.cmdCancel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.Size = New System.Drawing.Size(57, 17)
		Me.cmdCancel.Location = New System.Drawing.Point(496, 328)
		Me.cmdCancel.TabIndex = 133
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		tvw.OcxState = CType(resources.GetObject("tvw.OcxState"), System.Windows.Forms.AxHost.State)
		Me.tvw.Size = New System.Drawing.Size(185, 308)
		Me.tvw.Location = New System.Drawing.Point(8, 37)
		Me.tvw.TabIndex = 0
		Me.tvw.Name = "tvw"
		Me._fraPanel_0.BackColor = System.Drawing.Color.Black
		Me._fraPanel_0.ForeColor = System.Drawing.Color.White
		Me._fraPanel_0.Size = New System.Drawing.Size(441, 321)
		Me._fraPanel_0.Location = New System.Drawing.Point(200, 0)
		Me._fraPanel_0.TabIndex = 118
		Me._fraPanel_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraPanel_0.Enabled = True
		Me._fraPanel_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraPanel_0.Visible = True
		Me._fraPanel_0.Padding = New System.Windows.Forms.Padding(0)
		Me._fraPanel_0.Name = "_fraPanel_0"
		Me.txtCDKey.AutoSize = False
		Me.txtCDKey.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtCDKey.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtCDKey.ForeColor = System.Drawing.Color.White
		Me.txtCDKey.Size = New System.Drawing.Size(169, 19)
		Me.txtCDKey.Location = New System.Drawing.Point(16, 152)
		Me.txtCDKey.TabIndex = 3
		Me.txtCDKey.AcceptsReturn = True
		Me.txtCDKey.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCDKey.CausesValidation = True
		Me.txtCDKey.Enabled = True
		Me.txtCDKey.HideSelection = True
		Me.txtCDKey.ReadOnly = False
		Me.txtCDKey.Maxlength = 0
		Me.txtCDKey.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCDKey.MultiLine = False
		Me.txtCDKey.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCDKey.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCDKey.TabStop = True
		Me.txtCDKey.Visible = True
		Me.txtCDKey.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtCDKey.Name = "txtCDKey"
		Me.chkSHR.BackColor = System.Drawing.Color.Black
		Me.chkSHR.Text = "Shareware"
		Me.chkSHR.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSHR.ForeColor = System.Drawing.Color.White
		Me.chkSHR.Size = New System.Drawing.Size(81, 17)
		Me.chkSHR.Location = New System.Drawing.Point(320, 200)
		Me.chkSHR.TabIndex = 15
		Me.chkSHR.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSHR.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSHR.CausesValidation = True
		Me.chkSHR.Enabled = True
		Me.chkSHR.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSHR.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSHR.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSHR.TabStop = True
		Me.chkSHR.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSHR.Visible = True
		Me.chkSHR.Name = "chkSHR"
		Me.chkJPN.BackColor = System.Drawing.Color.Black
		Me.chkJPN.Text = "Japanese"
		Me.chkJPN.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkJPN.ForeColor = System.Drawing.Color.White
		Me.chkJPN.Size = New System.Drawing.Size(81, 17)
		Me.chkJPN.Location = New System.Drawing.Point(320, 216)
		Me.chkJPN.TabIndex = 16
		Me.chkJPN.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkJPN.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkJPN.CausesValidation = True
		Me.chkJPN.Enabled = True
		Me.chkJPN.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkJPN.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkJPN.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkJPN.TabStop = True
		Me.chkJPN.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkJPN.Visible = True
		Me.chkJPN.Name = "chkJPN"
		Me.optDRTL.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optDRTL.BackColor = System.Drawing.Color.Black
		Me.optDRTL.Text = "Diablo"
		Me.optDRTL.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optDRTL.ForeColor = System.Drawing.Color.White
		Me.optDRTL.Size = New System.Drawing.Size(97, 17)
		Me.optDRTL.Location = New System.Drawing.Point(216, 216)
		Me.optDRTL.Appearance = System.Windows.Forms.Appearance.Button
		Me.optDRTL.TabIndex = 14
		Me.optDRTL.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optDRTL.CausesValidation = True
		Me.optDRTL.Enabled = True
		Me.optDRTL.Cursor = System.Windows.Forms.Cursors.Default
		Me.optDRTL.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optDRTL.TabStop = True
		Me.optDRTL.Checked = False
		Me.optDRTL.Visible = True
		Me.optDRTL.Name = "optDRTL"
		Me.optW3XP.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optW3XP.BackColor = System.Drawing.Color.Black
		Me.optW3XP.Text = "The Frozen Throne"
		Me.optW3XP.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optW3XP.ForeColor = System.Drawing.Color.White
		Me.optW3XP.Size = New System.Drawing.Size(105, 17)
		Me.optW3XP.Location = New System.Drawing.Point(320, 176)
		Me.optW3XP.Appearance = System.Windows.Forms.Appearance.Button
		Me.optW3XP.TabIndex = 12
		Me.optW3XP.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optW3XP.CausesValidation = True
		Me.optW3XP.Enabled = True
		Me.optW3XP.Cursor = System.Windows.Forms.Cursors.Default
		Me.optW3XP.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optW3XP.TabStop = True
		Me.optW3XP.Checked = False
		Me.optW3XP.Visible = True
		Me.optW3XP.Name = "optW3XP"
		Me.optSTAR.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optSTAR.BackColor = System.Drawing.Color.Black
		Me.optSTAR.Text = "StarCraft"
		Me.optSTAR.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optSTAR.ForeColor = System.Drawing.Color.White
		Me.optSTAR.Size = New System.Drawing.Size(97, 17)
		Me.optSTAR.Location = New System.Drawing.Point(216, 120)
		Me.optSTAR.Appearance = System.Windows.Forms.Appearance.Button
		Me.optSTAR.TabIndex = 7
		Me.optSTAR.Checked = True
		Me.optSTAR.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optSTAR.CausesValidation = True
		Me.optSTAR.Enabled = True
		Me.optSTAR.Cursor = System.Windows.Forms.Cursors.Default
		Me.optSTAR.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optSTAR.TabStop = True
		Me.optSTAR.Visible = True
		Me.optSTAR.Name = "optSTAR"
		Me.optSEXP.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optSEXP.BackColor = System.Drawing.Color.Black
		Me.optSEXP.Text = "Brood War"
		Me.optSEXP.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optSEXP.ForeColor = System.Drawing.Color.White
		Me.optSEXP.Size = New System.Drawing.Size(105, 17)
		Me.optSEXP.Location = New System.Drawing.Point(320, 128)
		Me.optSEXP.Appearance = System.Windows.Forms.Appearance.Button
		Me.optSEXP.TabIndex = 8
		Me.optSEXP.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optSEXP.CausesValidation = True
		Me.optSEXP.Enabled = True
		Me.optSEXP.Cursor = System.Windows.Forms.Cursors.Default
		Me.optSEXP.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optSEXP.TabStop = True
		Me.optSEXP.Checked = False
		Me.optSEXP.Visible = True
		Me.optSEXP.Name = "optSEXP"
		Me.optD2DV.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optD2DV.BackColor = System.Drawing.Color.Black
		Me.optD2DV.Text = "Diablo II"
		Me.optD2DV.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optD2DV.ForeColor = System.Drawing.Color.White
		Me.optD2DV.Size = New System.Drawing.Size(97, 17)
		Me.optD2DV.Location = New System.Drawing.Point(216, 144)
		Me.optD2DV.Appearance = System.Windows.Forms.Appearance.Button
		Me.optD2DV.TabIndex = 9
		Me.optD2DV.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optD2DV.CausesValidation = True
		Me.optD2DV.Enabled = True
		Me.optD2DV.Cursor = System.Windows.Forms.Cursors.Default
		Me.optD2DV.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optD2DV.TabStop = True
		Me.optD2DV.Checked = False
		Me.optD2DV.Visible = True
		Me.optD2DV.Name = "optD2DV"
		Me.optD2XP.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optD2XP.BackColor = System.Drawing.Color.Black
		Me.optD2XP.Text = "Lord of Destruction"
		Me.optD2XP.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optD2XP.ForeColor = System.Drawing.Color.White
		Me.optD2XP.Size = New System.Drawing.Size(105, 17)
		Me.optD2XP.Location = New System.Drawing.Point(320, 152)
		Me.optD2XP.Appearance = System.Windows.Forms.Appearance.Button
		Me.optD2XP.TabIndex = 10
		Me.optD2XP.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optD2XP.CausesValidation = True
		Me.optD2XP.Enabled = True
		Me.optD2XP.Cursor = System.Windows.Forms.Cursors.Default
		Me.optD2XP.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optD2XP.TabStop = True
		Me.optD2XP.Checked = False
		Me.optD2XP.Visible = True
		Me.optD2XP.Name = "optD2XP"
		Me.optW2BN.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optW2BN.BackColor = System.Drawing.Color.Black
		Me.optW2BN.Text = "WarCraft II BNE"
		Me.optW2BN.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optW2BN.ForeColor = System.Drawing.Color.White
		Me.optW2BN.Size = New System.Drawing.Size(97, 17)
		Me.optW2BN.Location = New System.Drawing.Point(216, 192)
		Me.optW2BN.Appearance = System.Windows.Forms.Appearance.Button
		Me.optW2BN.TabIndex = 13
		Me.optW2BN.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optW2BN.CausesValidation = True
		Me.optW2BN.Enabled = True
		Me.optW2BN.Cursor = System.Windows.Forms.Cursors.Default
		Me.optW2BN.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optW2BN.TabStop = True
		Me.optW2BN.Checked = False
		Me.optW2BN.Visible = True
		Me.optW2BN.Name = "optW2BN"
		Me.optWAR3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optWAR3.BackColor = System.Drawing.Color.Black
		Me.optWAR3.Text = "WarCraft III"
		Me.optWAR3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optWAR3.ForeColor = System.Drawing.Color.White
		Me.optWAR3.Size = New System.Drawing.Size(97, 17)
		Me.optWAR3.Location = New System.Drawing.Point(216, 168)
		Me.optWAR3.Appearance = System.Windows.Forms.Appearance.Button
		Me.optWAR3.TabIndex = 11
		Me.optWAR3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optWAR3.CausesValidation = True
		Me.optWAR3.Enabled = True
		Me.optWAR3.Cursor = System.Windows.Forms.Cursors.Default
		Me.optWAR3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optWAR3.TabStop = True
		Me.optWAR3.Checked = False
		Me.optWAR3.Visible = True
		Me.optWAR3.Name = "optWAR3"
		Me.txtUsername.AutoSize = False
		Me.txtUsername.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtUsername.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtUsername.ForeColor = System.Drawing.Color.White
		Me.txtUsername.Size = New System.Drawing.Size(169, 19)
		Me.txtUsername.Location = New System.Drawing.Point(16, 72)
		Me.txtUsername.Maxlength = 15
		Me.txtUsername.TabIndex = 1
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
		Me.txtPassword.AutoSize = False
		Me.txtPassword.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtPassword.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPassword.ForeColor = System.Drawing.Color.White
		Me.txtPassword.Size = New System.Drawing.Size(169, 19)
		Me.txtPassword.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.txtPassword.Location = New System.Drawing.Point(16, 112)
		Me.txtPassword.PasswordChar = ChrW(42)
		Me.txtPassword.TabIndex = 2
		Me.txtPassword.AcceptsReturn = True
		Me.txtPassword.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPassword.CausesValidation = True
		Me.txtPassword.Enabled = True
		Me.txtPassword.HideSelection = True
		Me.txtPassword.ReadOnly = False
		Me.txtPassword.Maxlength = 0
		Me.txtPassword.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPassword.MultiLine = False
		Me.txtPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPassword.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPassword.TabStop = True
		Me.txtPassword.Visible = True
		Me.txtPassword.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPassword.Name = "txtPassword"
		Me.txtHomeChan.AutoSize = False
		Me.txtHomeChan.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtHomeChan.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtHomeChan.ForeColor = System.Drawing.Color.White
		Me.txtHomeChan.Size = New System.Drawing.Size(169, 19)
		Me.txtHomeChan.Location = New System.Drawing.Point(16, 232)
		Me.txtHomeChan.Maxlength = 31
		Me.txtHomeChan.TabIndex = 5
		Me.ToolTip1.SetToolTip(Me.txtHomeChan, "This is the channel that the bot will attempt to join when it logs on.")
		Me.txtHomeChan.AcceptsReturn = True
		Me.txtHomeChan.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtHomeChan.CausesValidation = True
		Me.txtHomeChan.Enabled = True
		Me.txtHomeChan.HideSelection = True
		Me.txtHomeChan.ReadOnly = False
		Me.txtHomeChan.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtHomeChan.MultiLine = False
		Me.txtHomeChan.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtHomeChan.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtHomeChan.TabStop = True
		Me.txtHomeChan.Visible = True
		Me.txtHomeChan.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtHomeChan.Name = "txtHomeChan"
		Me.cboServer.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.cboServer.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboServer.ForeColor = System.Drawing.Color.White
		Me.cboServer.Size = New System.Drawing.Size(161, 21)
		Me.cboServer.Location = New System.Drawing.Point(216, 72)
		Me.cboServer.TabIndex = 6
		Me.cboServer.Text = "Choose One"
		Me.cboServer.CausesValidation = True
		Me.cboServer.Enabled = True
		Me.cboServer.IntegralHeight = True
		Me.cboServer.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboServer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboServer.Sorted = False
		Me.cboServer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboServer.TabStop = True
		Me.cboServer.Visible = True
		Me.cboServer.Name = "cboServer"
		Me.txtExpKey.AutoSize = False
		Me.txtExpKey.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtExpKey.Enabled = False
		Me.txtExpKey.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtExpKey.ForeColor = System.Drawing.Color.White
		Me.txtExpKey.Size = New System.Drawing.Size(169, 19)
		Me.txtExpKey.Location = New System.Drawing.Point(16, 192)
		Me.txtExpKey.TabIndex = 4
		Me.ToolTip1.SetToolTip(Me.txtExpKey, "Only required for Lord of Destruction and The Frozen Throne.")
		Me.txtExpKey.AcceptsReturn = True
		Me.txtExpKey.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtExpKey.CausesValidation = True
		Me.txtExpKey.HideSelection = True
		Me.txtExpKey.ReadOnly = False
		Me.txtExpKey.Maxlength = 0
		Me.txtExpKey.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtExpKey.MultiLine = False
		Me.txtExpKey.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtExpKey.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtExpKey.TabStop = True
		Me.txtExpKey.Visible = True
		Me.txtExpKey.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtExpKey.Name = "txtExpKey"
		Me.chkUseRealm.BackColor = System.Drawing.Color.Black
		Me.chkUseRealm.Text = "Use Diablo II Realms"
		Me.chkUseRealm.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkUseRealm.ForeColor = System.Drawing.Color.White
		Me.chkUseRealm.Size = New System.Drawing.Size(129, 17)
		Me.chkUseRealm.Location = New System.Drawing.Point(216, 256)
		Me.chkUseRealm.TabIndex = 18
		Me.chkUseRealm.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkUseRealm.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkUseRealm.CausesValidation = True
		Me.chkUseRealm.Enabled = True
		Me.chkUseRealm.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkUseRealm.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkUseRealm.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkUseRealm.TabStop = True
		Me.chkUseRealm.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkUseRealm.Visible = True
		Me.chkUseRealm.Name = "chkUseRealm"
		Me.chkSpawn.BackColor = System.Drawing.Color.Black
		Me.chkSpawn.Text = "Use Key as Spawned Client"
		Me.chkSpawn.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSpawn.ForeColor = System.Drawing.Color.White
		Me.chkSpawn.Size = New System.Drawing.Size(217, 17)
		Me.chkSpawn.Location = New System.Drawing.Point(216, 240)
		Me.chkSpawn.TabIndex = 17
		Me.chkSpawn.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSpawn.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSpawn.CausesValidation = True
		Me.chkSpawn.Enabled = True
		Me.chkSpawn.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSpawn.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSpawn.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSpawn.TabStop = True
		Me.chkSpawn.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSpawn.Visible = True
		Me.chkSpawn.Name = "chkSpawn"
		Me.lblAddCurrentKey.BackColor = System.Drawing.Color.Black
		Me.lblAddCurrentKey.Text = "( add current )"
		Me.lblAddCurrentKey.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblAddCurrentKey.ForeColor = System.Drawing.Color.White
		Me.lblAddCurrentKey.Size = New System.Drawing.Size(73, 17)
		Me.lblAddCurrentKey.Location = New System.Drawing.Point(56, 136)
		Me.lblAddCurrentKey.TabIndex = 187
		Me.lblAddCurrentKey.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblAddCurrentKey.Enabled = True
		Me.lblAddCurrentKey.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblAddCurrentKey.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblAddCurrentKey.UseMnemonic = True
		Me.lblAddCurrentKey.Visible = True
		Me.lblAddCurrentKey.AutoSize = False
		Me.lblAddCurrentKey.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblAddCurrentKey.Name = "lblAddCurrentKey"
		Me.lblManageKeys.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblManageKeys.BackColor = System.Drawing.Color.Black
		Me.lblManageKeys.Text = "( manage )"
		Me.lblManageKeys.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblManageKeys.ForeColor = System.Drawing.Color.White
		Me.lblManageKeys.Size = New System.Drawing.Size(57, 17)
		Me.lblManageKeys.Location = New System.Drawing.Point(128, 136)
		Me.lblManageKeys.TabIndex = 186
		Me.lblManageKeys.Enabled = True
		Me.lblManageKeys.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblManageKeys.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblManageKeys.UseMnemonic = True
		Me.lblManageKeys.Visible = True
		Me.lblManageKeys.AutoSize = False
		Me.lblManageKeys.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblManageKeys.Name = "lblManageKeys"
		Me._Label1_10.BackColor = System.Drawing.Color.Black
		Me._Label1_10.Text = "Expansion"
		Me._Label1_10.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_10.ForeColor = System.Drawing.Color.White
		Me._Label1_10.Size = New System.Drawing.Size(49, 17)
		Me._Label1_10.Location = New System.Drawing.Point(320, 112)
		Me._Label1_10.TabIndex = 142
		Me._Label1_10.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_10.Enabled = True
		Me._Label1_10.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_10.UseMnemonic = True
		Me._Label1_10.Visible = True
		Me._Label1_10.AutoSize = False
		Me._Label1_10.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_10.Name = "_Label1_10"
		Me._Line3_2.BorderColor = System.Drawing.Color.White
		Me._Line3_2.X1 = 312
		Me._Line3_2.X2 = 320
		Me._Line3_2.Y1 = 163
		Me._Line3_2.Y2 = 171
		Me._Line3_2.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line3_2.BorderWidth = 1
		Me._Line3_2.Visible = True
		Me._Line3_2.Name = "_Line3_2"
		Me._Line3_1.BorderColor = System.Drawing.Color.White
		Me._Line3_1.X1 = 312
		Me._Line3_1.X2 = 320
		Me._Line3_1.Y1 = 139
		Me._Line3_1.Y2 = 147
		Me._Line3_1.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line3_1.BorderWidth = 1
		Me._Line3_1.Visible = True
		Me._Line3_1.Name = "_Line3_1"
		Me._Line3_0.BorderColor = System.Drawing.Color.White
		Me._Line3_0.X1 = 312
		Me._Line3_0.X2 = 320
		Me._Line3_0.Y1 = 115
		Me._Line3_0.Y2 = 123
		Me._Line3_0.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line3_0.BorderWidth = 1
		Me._Line3_0.Visible = True
		Me._Line3_0.Name = "_Line3_0"
		Me.Line2.BorderColor = System.Drawing.Color.White
		Me.Line2.X1 = 200
		Me.Line2.X2 = 200
		Me.Line2.Y1 = 51
		Me.Line2.Y2 = 259
		Me.Line2.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line2.BorderWidth = 1
		Me.Line2.Visible = True
		Me.Line2.Name = "Line2"
		Me._Label1_9.BackColor = System.Drawing.Color.Black
		Me._Label1_9.Text = "Product"
		Me._Label1_9.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_9.ForeColor = System.Drawing.Color.White
		Me._Label1_9.Size = New System.Drawing.Size(41, 17)
		Me._Label1_9.Location = New System.Drawing.Point(216, 104)
		Me._Label1_9.TabIndex = 141
		Me._Label1_9.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_9.Enabled = True
		Me._Label1_9.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_9.UseMnemonic = True
		Me._Label1_9.Visible = True
		Me._Label1_9.AutoSize = False
		Me._Label1_9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_9.Name = "_Label1_9"
		Me._Line1_1.BorderColor = System.Drawing.Color.White
		Me._Line1_1.X1 = 24
		Me._Line1_1.X2 = 416
		Me._Line1_1.Y1 = 27
		Me._Line1_1.Y2 = 27
		Me._Line1_1.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line1_1.BorderWidth = 1
		Me._Line1_1.Visible = True
		Me._Line1_1.Name = "_Line1_1"
		Me._Label1_6.BackColor = System.Drawing.Color.Black
		Me._Label1_6.Text = "Server"
		Me._Label1_6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_6.ForeColor = System.Drawing.Color.White
		Me._Label1_6.Size = New System.Drawing.Size(33, 17)
		Me._Label1_6.Location = New System.Drawing.Point(216, 56)
		Me._Label1_6.TabIndex = 139
		Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_6.Enabled = True
		Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_6.UseMnemonic = True
		Me._Label1_6.Visible = True
		Me._Label1_6.AutoSize = False
		Me._Label1_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_6.Name = "_Label1_6"
		Me._Label1_4.BackColor = System.Drawing.Color.Black
		Me._Label1_4.Text = "Username"
		Me._Label1_4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_4.ForeColor = System.Drawing.Color.White
		Me._Label1_4.Size = New System.Drawing.Size(49, 17)
		Me._Label1_4.Location = New System.Drawing.Point(16, 56)
		Me._Label1_4.TabIndex = 138
		Me._Label1_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_4.Enabled = True
		Me._Label1_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_4.UseMnemonic = True
		Me._Label1_4.Visible = True
		Me._Label1_4.AutoSize = False
		Me._Label1_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_4.Name = "_Label1_4"
		Me._Label1_1.BackColor = System.Drawing.Color.Black
		Me._Label1_1.Text = "Password"
		Me._Label1_1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_1.ForeColor = System.Drawing.Color.White
		Me._Label1_1.Size = New System.Drawing.Size(49, 17)
		Me._Label1_1.Location = New System.Drawing.Point(16, 96)
		Me._Label1_1.TabIndex = 137
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_1.Enabled = True
		Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_1.UseMnemonic = True
		Me._Label1_1.Visible = True
		Me._Label1_1.AutoSize = False
		Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_1.Name = "_Label1_1"
		Me._Label1_2.BackColor = System.Drawing.Color.Black
		Me._Label1_2.Text = "CDKey"
		Me._Label1_2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_2.ForeColor = System.Drawing.Color.White
		Me._Label1_2.Size = New System.Drawing.Size(33, 17)
		Me._Label1_2.Location = New System.Drawing.Point(16, 136)
		Me._Label1_2.TabIndex = 136
		Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_2.Enabled = True
		Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_2.UseMnemonic = True
		Me._Label1_2.Visible = True
		Me._Label1_2.AutoSize = False
		Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_2.Name = "_Label1_2"
		Me._Label1_3.BackColor = System.Drawing.Color.Black
		Me._Label1_3.Text = "Home Channel"
		Me._Label1_3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_3.ForeColor = System.Drawing.Color.White
		Me._Label1_3.Size = New System.Drawing.Size(73, 17)
		Me._Label1_3.Location = New System.Drawing.Point(16, 216)
		Me._Label1_3.TabIndex = 135
		Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_3.Enabled = True
		Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_3.UseMnemonic = True
		Me._Label1_3.Visible = True
		Me._Label1_3.AutoSize = False
		Me._Label1_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_3.Name = "_Label1_3"
		Me._Label1_5.BackColor = System.Drawing.Color.Black
		Me._Label1_5.Text = "Expansion CDKey"
		Me._Label1_5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_5.ForeColor = System.Drawing.Color.White
		Me._Label1_5.Size = New System.Drawing.Size(89, 17)
		Me._Label1_5.Location = New System.Drawing.Point(16, 176)
		Me._Label1_5.TabIndex = 134
		Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_5.Enabled = True
		Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_5.UseMnemonic = True
		Me._Label1_5.Visible = True
		Me._Label1_5.AutoSize = False
		Me._Label1_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_5.Name = "_Label1_5"
		Me._Label1_7.BackColor = System.Drawing.Color.Black
		Me._Label1_7.Text = "Basic configuration"
		Me._Label1_7.Font = New System.Drawing.Font("Tahoma", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_7.ForeColor = System.Drawing.Color.White
		Me._Label1_7.Size = New System.Drawing.Size(321, 25)
		Me._Label1_7.Location = New System.Drawing.Point(24, 16)
		Me._Label1_7.TabIndex = 140
		Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_7.Enabled = True
		Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_7.UseMnemonic = True
		Me._Label1_7.Visible = True
		Me._Label1_7.AutoSize = False
		Me._Label1_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_7.Name = "_Label1_7"
		Me._fraPanel_2.BackColor = System.Drawing.Color.Black
		Me._fraPanel_2.ForeColor = System.Drawing.Color.White
		Me._fraPanel_2.Size = New System.Drawing.Size(441, 321)
		Me._fraPanel_2.Location = New System.Drawing.Point(200, 0)
		Me._fraPanel_2.TabIndex = 120
		Me._fraPanel_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraPanel_2.Enabled = True
		Me._fraPanel_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraPanel_2.Visible = True
		Me._fraPanel_2.Padding = New System.Windows.Forms.Padding(0)
		Me._fraPanel_2.Name = "_fraPanel_2"
		Me.chkMinimizeOnStartup.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkMinimizeOnStartup.BackColor = System.Drawing.Color.Black
		Me.chkMinimizeOnStartup.Text = "Minimize on startup"
		Me.chkMinimizeOnStartup.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkMinimizeOnStartup.ForeColor = System.Drawing.Color.White
		Me.chkMinimizeOnStartup.Size = New System.Drawing.Size(169, 17)
		Me.chkMinimizeOnStartup.Location = New System.Drawing.Point(248, 80)
		Me.chkMinimizeOnStartup.TabIndex = 34
		Me.ToolTip1.SetToolTip(Me.chkMinimizeOnStartup, "Automatically minimize on startup.")
		Me.chkMinimizeOnStartup.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkMinimizeOnStartup.CausesValidation = True
		Me.chkMinimizeOnStartup.Enabled = True
		Me.chkMinimizeOnStartup.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkMinimizeOnStartup.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkMinimizeOnStartup.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkMinimizeOnStartup.TabStop = True
		Me.chkMinimizeOnStartup.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkMinimizeOnStartup.Visible = True
		Me.chkMinimizeOnStartup.Name = "chkMinimizeOnStartup"
		Me.chkURLDetect.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkURLDetect.BackColor = System.Drawing.Color.Black
		Me.chkURLDetect.Text = "Enable URL detection"
		Me.chkURLDetect.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkURLDetect.ForeColor = System.Drawing.Color.White
		Me.chkURLDetect.Size = New System.Drawing.Size(169, 17)
		Me.chkURLDetect.Location = New System.Drawing.Point(248, 136)
		Me.chkURLDetect.TabIndex = 38
		Me.ToolTip1.SetToolTip(Me.chkURLDetect, "Enables automatic URL detection and highlighting in the main chat window.")
		Me.chkURLDetect.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkURLDetect.CausesValidation = True
		Me.chkURLDetect.Enabled = True
		Me.chkURLDetect.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkURLDetect.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkURLDetect.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkURLDetect.TabStop = True
		Me.chkURLDetect.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkURLDetect.Visible = True
		Me.chkURLDetect.Name = "chkURLDetect"
		Me.chkDisablePrefix.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkDisablePrefix.BackColor = System.Drawing.Color.Black
		Me.chkDisablePrefix.Text = "Disable prefix box"
		Me.chkDisablePrefix.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkDisablePrefix.ForeColor = System.Drawing.Color.White
		Me.chkDisablePrefix.Size = New System.Drawing.Size(121, 17)
		Me.chkDisablePrefix.Location = New System.Drawing.Point(24, 256)
		Me.chkDisablePrefix.TabIndex = 43
		Me.ToolTip1.SetToolTip(Me.chkDisablePrefix, "Disables the smaller prefix box to the left of the box you type in to send text to Battle.net")
		Me.chkDisablePrefix.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkDisablePrefix.CausesValidation = True
		Me.chkDisablePrefix.Enabled = True
		Me.chkDisablePrefix.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkDisablePrefix.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkDisablePrefix.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkDisablePrefix.TabStop = True
		Me.chkDisablePrefix.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkDisablePrefix.Visible = True
		Me.chkDisablePrefix.Name = "chkDisablePrefix"
		Me.chkDisableSuffix.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkDisableSuffix.BackColor = System.Drawing.Color.Black
		Me.chkDisableSuffix.Text = "Disable suffix box"
		Me.chkDisableSuffix.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkDisableSuffix.ForeColor = System.Drawing.Color.White
		Me.chkDisableSuffix.Size = New System.Drawing.Size(121, 17)
		Me.chkDisableSuffix.Location = New System.Drawing.Point(24, 280)
		Me.chkDisableSuffix.TabIndex = 44
		Me.ToolTip1.SetToolTip(Me.chkDisableSuffix, "Disables the smaller suffix box to the right of the box you type in to send text to Battle.net")
		Me.chkDisableSuffix.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkDisableSuffix.CausesValidation = True
		Me.chkDisableSuffix.Enabled = True
		Me.chkDisableSuffix.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkDisableSuffix.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkDisableSuffix.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkDisableSuffix.TabStop = True
		Me.chkDisableSuffix.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkDisableSuffix.Visible = True
		Me.chkDisableSuffix.Name = "chkDisableSuffix"
		Me.chkShowUserGameStatsIcons.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkShowUserGameStatsIcons.BackColor = System.Drawing.Color.Black
		Me.chkShowUserGameStatsIcons.Text = "Show user game stats icons"
		Me.chkShowUserGameStatsIcons.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkShowUserGameStatsIcons.ForeColor = System.Drawing.Color.White
		Me.chkShowUserGameStatsIcons.Size = New System.Drawing.Size(169, 25)
		Me.chkShowUserGameStatsIcons.Location = New System.Drawing.Point(248, 192)
		Me.chkShowUserGameStatsIcons.TabIndex = 41
		Me.chkShowUserGameStatsIcons.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkShowUserGameStatsIcons.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkShowUserGameStatsIcons.CausesValidation = True
		Me.chkShowUserGameStatsIcons.Enabled = True
		Me.chkShowUserGameStatsIcons.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkShowUserGameStatsIcons.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkShowUserGameStatsIcons.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkShowUserGameStatsIcons.TabStop = True
		Me.chkShowUserGameStatsIcons.Visible = True
		Me.chkShowUserGameStatsIcons.Name = "chkShowUserGameStatsIcons"
		Me.chkShowUserFlagsIcons.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkShowUserFlagsIcons.BackColor = System.Drawing.Color.Black
		Me.chkShowUserFlagsIcons.Text = "Show user flag-based icons"
		Me.chkShowUserFlagsIcons.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkShowUserFlagsIcons.ForeColor = System.Drawing.Color.White
		Me.chkShowUserFlagsIcons.Size = New System.Drawing.Size(169, 17)
		Me.chkShowUserFlagsIcons.Location = New System.Drawing.Point(248, 216)
		Me.chkShowUserFlagsIcons.TabIndex = 42
		Me.chkShowUserFlagsIcons.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkShowUserFlagsIcons.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkShowUserFlagsIcons.CausesValidation = True
		Me.chkShowUserFlagsIcons.Enabled = True
		Me.chkShowUserFlagsIcons.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkShowUserFlagsIcons.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkShowUserFlagsIcons.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkShowUserFlagsIcons.TabStop = True
		Me.chkShowUserFlagsIcons.Visible = True
		Me.chkShowUserFlagsIcons.Name = "chkShowUserFlagsIcons"
		Me.chkNoColoring.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkNoColoring.BackColor = System.Drawing.Color.Black
		Me.chkNoColoring.Text = "Disable channel list name coloring"
		Me.chkNoColoring.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkNoColoring.ForeColor = System.Drawing.Color.White
		Me.chkNoColoring.Size = New System.Drawing.Size(201, 17)
		Me.chkNoColoring.Location = New System.Drawing.Point(24, 208)
		Me.chkNoColoring.TabIndex = 40
		Me.ToolTip1.SetToolTip(Me.chkNoColoring, "Name coloring changes the color of people's names in the channel list based on their status or activity.")
		Me.chkNoColoring.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkNoColoring.CausesValidation = True
		Me.chkNoColoring.Enabled = True
		Me.chkNoColoring.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkNoColoring.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkNoColoring.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkNoColoring.TabStop = True
		Me.chkNoColoring.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkNoColoring.Visible = True
		Me.chkNoColoring.Name = "chkNoColoring"
		Me.chkNoAutocomplete.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkNoAutocomplete.BackColor = System.Drawing.Color.Black
		Me.chkNoAutocomplete.Text = "Disable name autocompletion"
		Me.chkNoAutocomplete.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkNoAutocomplete.ForeColor = System.Drawing.Color.White
		Me.chkNoAutocomplete.Size = New System.Drawing.Size(169, 17)
		Me.chkNoAutocomplete.Location = New System.Drawing.Point(248, 160)
		Me.chkNoAutocomplete.TabIndex = 39
		Me.ToolTip1.SetToolTip(Me.chkNoAutocomplete, "Checking this box prevents the highlighted display of suggested usernames as you type in the send box.")
		Me.chkNoAutocomplete.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkNoAutocomplete.CausesValidation = True
		Me.chkNoAutocomplete.Enabled = True
		Me.chkNoAutocomplete.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkNoAutocomplete.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkNoAutocomplete.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkNoAutocomplete.TabStop = True
		Me.chkNoAutocomplete.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkNoAutocomplete.Visible = True
		Me.chkNoAutocomplete.Name = "chkNoAutocomplete"
		Me.chkNoTray.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkNoTray.BackColor = System.Drawing.Color.Black
		Me.chkNoTray.Text = "Do not minimize to the System Tray"
		Me.chkNoTray.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkNoTray.ForeColor = System.Drawing.Color.White
		Me.chkNoTray.Size = New System.Drawing.Size(201, 17)
		Me.chkNoTray.Location = New System.Drawing.Point(24, 80)
		Me.chkNoTray.TabIndex = 32
		Me.ToolTip1.SetToolTip(Me.chkNoTray, "Disable minimization to the System Tray (only to the Taskbar).")
		Me.chkNoTray.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkNoTray.CausesValidation = True
		Me.chkNoTray.Enabled = True
		Me.chkNoTray.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkNoTray.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkNoTray.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkNoTray.TabStop = True
		Me.chkNoTray.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkNoTray.Visible = True
		Me.chkNoTray.Name = "chkNoTray"
		Me.chkFlash.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkFlash.BackColor = System.Drawing.Color.Black
		Me.chkFlash.Text = "Flash window on events"
		Me.chkFlash.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkFlash.ForeColor = System.Drawing.Color.White
		Me.chkFlash.Size = New System.Drawing.Size(169, 17)
		Me.chkFlash.Location = New System.Drawing.Point(248, 56)
		Me.chkFlash.TabIndex = 33
		Me.ToolTip1.SetToolTip(Me.chkFlash, "Flash the bot window when events occur.")
		Me.chkFlash.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkFlash.CausesValidation = True
		Me.chkFlash.Enabled = True
		Me.chkFlash.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkFlash.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkFlash.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkFlash.TabStop = True
		Me.chkFlash.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkFlash.Visible = True
		Me.chkFlash.Name = "chkFlash"
		Me.chkUTF8.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkUTF8.BackColor = System.Drawing.Color.Black
		Me.chkUTF8.Text = "Use UTF-8 encoding/decoding when processing and sending messages"
		Me.chkUTF8.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkUTF8.ForeColor = System.Drawing.Color.White
		Me.chkUTF8.Size = New System.Drawing.Size(201, 33)
		Me.chkUTF8.Location = New System.Drawing.Point(24, 112)
		Me.chkUTF8.TabIndex = 35
		Me.ToolTip1.SetToolTip(Me.chkUTF8, "Blizzard games encode their messages to UTF-8 format. Enable this setting to properly see special characters sent by these games.")
		Me.chkUTF8.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkUTF8.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkUTF8.CausesValidation = True
		Me.chkUTF8.Enabled = True
		Me.chkUTF8.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkUTF8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkUTF8.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkUTF8.TabStop = True
		Me.chkUTF8.Visible = True
		Me.chkUTF8.Name = "chkUTF8"
		Me.cboTimestamp.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.cboTimestamp.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboTimestamp.ForeColor = System.Drawing.Color.White
		Me.cboTimestamp.Size = New System.Drawing.Size(233, 21)
		Me.cboTimestamp.Location = New System.Drawing.Point(184, 272)
		Me.cboTimestamp.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboTimestamp.TabIndex = 45
		Me.cboTimestamp.CausesValidation = True
		Me.cboTimestamp.Enabled = True
		Me.cboTimestamp.IntegralHeight = True
		Me.cboTimestamp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboTimestamp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboTimestamp.Sorted = False
		Me.cboTimestamp.TabStop = True
		Me.cboTimestamp.Visible = True
		Me.cboTimestamp.Name = "cboTimestamp"
		Me.chkSplash.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkSplash.BackColor = System.Drawing.Color.Black
		Me.chkSplash.Text = "Show splash screen on startup"
		Me.chkSplash.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSplash.ForeColor = System.Drawing.Color.White
		Me.chkSplash.Size = New System.Drawing.Size(201, 17)
		Me.chkSplash.Location = New System.Drawing.Point(24, 56)
		Me.chkSplash.TabIndex = 31
		Me.ToolTip1.SetToolTip(Me.chkSplash, "Enable/disable the splash screen on startup.")
		Me.chkSplash.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSplash.CausesValidation = True
		Me.chkSplash.Enabled = True
		Me.chkSplash.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSplash.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSplash.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSplash.TabStop = True
		Me.chkSplash.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSplash.Visible = True
		Me.chkSplash.Name = "chkSplash"
		Me.chkFilter.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkFilter.BackColor = System.Drawing.Color.Black
		Me.chkFilter.Text = "Use chat filtering"
		Me.chkFilter.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkFilter.ForeColor = System.Drawing.Color.White
		Me.chkFilter.Size = New System.Drawing.Size(169, 17)
		Me.chkFilter.Location = New System.Drawing.Point(248, 112)
		Me.chkFilter.TabIndex = 37
		Me.ToolTip1.SetToolTip(Me.chkFilter, "Enable/disable chat filtering (lowers CPU usage)")
		Me.chkFilter.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkFilter.CausesValidation = True
		Me.chkFilter.Enabled = True
		Me.chkFilter.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkFilter.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkFilter.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkFilter.TabStop = True
		Me.chkFilter.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkFilter.Visible = True
		Me.chkFilter.Name = "chkFilter"
		Me.chkJoinLeaves.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkJoinLeaves.BackColor = System.Drawing.Color.Black
		Me.chkJoinLeaves.Text = "Show join/leave notifications"
		Me.chkJoinLeaves.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkJoinLeaves.ForeColor = System.Drawing.Color.White
		Me.chkJoinLeaves.Size = New System.Drawing.Size(201, 17)
		Me.chkJoinLeaves.Location = New System.Drawing.Point(24, 152)
		Me.chkJoinLeaves.TabIndex = 36
		Me.ToolTip1.SetToolTip(Me.chkJoinLeaves, "Enable/disable Join and Leave messages")
		Me.chkJoinLeaves.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkJoinLeaves.CausesValidation = True
		Me.chkJoinLeaves.Enabled = True
		Me.chkJoinLeaves.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkJoinLeaves.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkJoinLeaves.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkJoinLeaves.TabStop = True
		Me.chkJoinLeaves.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkJoinLeaves.Visible = True
		Me.chkJoinLeaves.Name = "chkJoinLeaves"
		Me.Line12.BorderColor = System.Drawing.Color.White
		Me.Line12.X1 = 168
		Me.Line12.X2 = 168
		Me.Line12.Y1 = 235
		Me.Line12.Y2 = 299
		Me.Line12.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line12.BorderWidth = 1
		Me.Line12.Visible = True
		Me.Line12.Name = "Line12"
		Me.Line11.BorderColor = System.Drawing.Color.White
		Me.Line11.X1 = 416
		Me.Line11.X2 = 24
		Me.Line11.Y1 = 227
		Me.Line11.Y2 = 227
		Me.Line11.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line11.BorderWidth = 1
		Me.Line11.Visible = True
		Me.Line11.Name = "Line11"
		Me.Line10.BorderColor = System.Drawing.Color.White
		Me.Line10.X1 = 416
		Me.Line10.X2 = 24
		Me.Line10.Y1 = 171
		Me.Line10.Y2 = 171
		Me.Line10.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line10.BorderWidth = 1
		Me.Line10.Visible = True
		Me.Line10.Name = "Line10"
		Me.Line9.BorderColor = System.Drawing.Color.White
		Me.Line9.X1 = 24
		Me.Line9.X2 = 416
		Me.Line9.Y1 = 91
		Me.Line9.Y2 = 91
		Me.Line9.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line9.BorderWidth = 1
		Me.Line9.Visible = True
		Me.Line9.Name = "Line9"
		Me._Label8_8.BackColor = System.Drawing.Color.Black
		Me._Label8_8.Text = "Timestamp Settings"
		Me._Label8_8.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label8_8.ForeColor = System.Drawing.Color.White
		Me._Label8_8.Size = New System.Drawing.Size(105, 17)
		Me._Label8_8.Location = New System.Drawing.Point(184, 256)
		Me._Label8_8.TabIndex = 156
		Me._Label8_8.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label8_8.Enabled = True
		Me._Label8_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label8_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label8_8.UseMnemonic = True
		Me._Label8_8.Visible = True
		Me._Label8_8.AutoSize = False
		Me._Label8_8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label8_8.Name = "_Label8_8"
		Me._Line1_3.BorderColor = System.Drawing.Color.White
		Me._Line1_3.X1 = 24
		Me._Line1_3.X2 = 416
		Me._Line1_3.Y1 = 27
		Me._Line1_3.Y2 = 27
		Me._Line1_3.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line1_3.BorderWidth = 1
		Me._Line1_3.Visible = True
		Me._Line1_3.Name = "_Line1_3"
		Me._Label1_12.BackColor = System.Drawing.Color.Black
		Me._Label1_12.Text = "General interface settings"
		Me._Label1_12.Font = New System.Drawing.Font("Tahoma", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_12.ForeColor = System.Drawing.Color.White
		Me._Label1_12.Size = New System.Drawing.Size(321, 25)
		Me._Label1_12.Location = New System.Drawing.Point(24, 16)
		Me._Label1_12.TabIndex = 144
		Me._Label1_12.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_12.Enabled = True
		Me._Label1_12.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_12.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_12.UseMnemonic = True
		Me._Label1_12.Visible = True
		Me._Label1_12.AutoSize = False
		Me._Label1_12.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_12.Name = "_Label1_12"
		Me._fraPanel_3.BackColor = System.Drawing.Color.Black
		Me._fraPanel_3.ForeColor = System.Drawing.Color.White
		Me._fraPanel_3.Size = New System.Drawing.Size(441, 321)
		Me._fraPanel_3.Location = New System.Drawing.Point(200, 0)
		Me._fraPanel_3.TabIndex = 121
		Me._fraPanel_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraPanel_3.Enabled = True
		Me._fraPanel_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraPanel_3.Visible = True
		Me._fraPanel_3.Padding = New System.Windows.Forms.Padding(0)
		Me._fraPanel_3.Name = "_fraPanel_3"
		Me.cmdSaveColor.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSaveColor.Text = "Sa&ve changes to this color"
		Me.cmdSaveColor.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSaveColor.Size = New System.Drawing.Size(137, 17)
		Me.cmdSaveColor.Location = New System.Drawing.Point(280, 144)
		Me.cmdSaveColor.TabIndex = 52
		Me.cmdSaveColor.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSaveColor.CausesValidation = True
		Me.cmdSaveColor.Enabled = True
		Me.cmdSaveColor.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSaveColor.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSaveColor.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSaveColor.TabStop = True
		Me.cmdSaveColor.Name = "cmdSaveColor"
		Me.cboColorList.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.cboColorList.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboColorList.ForeColor = System.Drawing.Color.White
		Me.cboColorList.Size = New System.Drawing.Size(393, 21)
		Me.cboColorList.Location = New System.Drawing.Point(24, 120)
		Me.cboColorList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboColorList.TabIndex = 51
		Me.cboColorList.CausesValidation = True
		Me.cboColorList.Enabled = True
		Me.cboColorList.IntegralHeight = True
		Me.cboColorList.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboColorList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboColorList.Sorted = False
		Me.cboColorList.TabStop = True
		Me.cboColorList.Visible = True
		Me.cboColorList.Name = "cboColorList"
		Me.txtValue.AutoSize = False
		Me.txtValue.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtValue.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtValue.ForeColor = System.Drawing.Color.White
		Me.txtValue.Size = New System.Drawing.Size(105, 19)
		Me.txtValue.Location = New System.Drawing.Point(24, 168)
		Me.txtValue.TabIndex = 53
		Me.txtValue.AcceptsReturn = True
		Me.txtValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtValue.CausesValidation = True
		Me.txtValue.Enabled = True
		Me.txtValue.HideSelection = True
		Me.txtValue.ReadOnly = False
		Me.txtValue.Maxlength = 0
		Me.txtValue.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtValue.MultiLine = False
		Me.txtValue.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtValue.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtValue.TabStop = True
		Me.txtValue.Visible = True
		Me.txtValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtValue.Name = "txtValue"
		Me.cmdColorPicker.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdColorPicker.Text = "Color Picker"
		Me.cmdColorPicker.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdColorPicker.Size = New System.Drawing.Size(105, 17)
		Me.cmdColorPicker.Location = New System.Drawing.Point(24, 200)
		Me.cmdColorPicker.TabIndex = 54
		Me.cmdColorPicker.BackColor = System.Drawing.SystemColors.Control
		Me.cmdColorPicker.CausesValidation = True
		Me.cmdColorPicker.Enabled = True
		Me.cmdColorPicker.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdColorPicker.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdColorPicker.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdColorPicker.TabStop = True
		Me.cmdColorPicker.Name = "cmdColorPicker"
		Me.txtR.AutoSize = False
		Me.txtR.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtR.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtR.ForeColor = System.Drawing.Color.White
		Me.txtR.Size = New System.Drawing.Size(41, 19)
		Me.txtR.Location = New System.Drawing.Point(40, 224)
		Me.txtR.TabIndex = 55
		Me.txtR.AcceptsReturn = True
		Me.txtR.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtR.CausesValidation = True
		Me.txtR.Enabled = True
		Me.txtR.HideSelection = True
		Me.txtR.ReadOnly = False
		Me.txtR.Maxlength = 0
		Me.txtR.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtR.MultiLine = False
		Me.txtR.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtR.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtR.TabStop = True
		Me.txtR.Visible = True
		Me.txtR.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtR.Name = "txtR"
		Me.txtG.AutoSize = False
		Me.txtG.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtG.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtG.ForeColor = System.Drawing.Color.White
		Me.txtG.Size = New System.Drawing.Size(41, 19)
		Me.txtG.Location = New System.Drawing.Point(104, 224)
		Me.txtG.TabIndex = 56
		Me.txtG.AcceptsReturn = True
		Me.txtG.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtG.CausesValidation = True
		Me.txtG.Enabled = True
		Me.txtG.HideSelection = True
		Me.txtG.ReadOnly = False
		Me.txtG.Maxlength = 0
		Me.txtG.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtG.MultiLine = False
		Me.txtG.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtG.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtG.TabStop = True
		Me.txtG.Visible = True
		Me.txtG.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtG.Name = "txtG"
		Me.txtB.AutoSize = False
		Me.txtB.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtB.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtB.ForeColor = System.Drawing.Color.White
		Me.txtB.Size = New System.Drawing.Size(41, 19)
		Me.txtB.Location = New System.Drawing.Point(168, 224)
		Me.txtB.TabIndex = 57
		Me.txtB.AcceptsReturn = True
		Me.txtB.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtB.CausesValidation = True
		Me.txtB.Enabled = True
		Me.txtB.HideSelection = True
		Me.txtB.ReadOnly = False
		Me.txtB.Maxlength = 0
		Me.txtB.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtB.MultiLine = False
		Me.txtB.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtB.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtB.TabStop = True
		Me.txtB.Visible = True
		Me.txtB.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtB.Name = "txtB"
		Me.cmdGetRGB.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdGetRGB.Text = "Generate New Value from RGB"
		Me.cmdGetRGB.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdGetRGB.Size = New System.Drawing.Size(185, 17)
		Me.cmdGetRGB.Location = New System.Drawing.Point(24, 248)
		Me.cmdGetRGB.TabIndex = 60
		Me.cmdGetRGB.BackColor = System.Drawing.SystemColors.Control
		Me.cmdGetRGB.CausesValidation = True
		Me.cmdGetRGB.Enabled = True
		Me.cmdGetRGB.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdGetRGB.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdGetRGB.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdGetRGB.TabStop = True
		Me.cmdGetRGB.Name = "cmdGetRGB"
		Me.cmdImport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdImport.Text = "&Import ColorList"
		Me.cmdImport.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdImport.Size = New System.Drawing.Size(97, 17)
		Me.cmdImport.Location = New System.Drawing.Point(216, 248)
		Me.cmdImport.TabIndex = 61
		Me.cmdImport.BackColor = System.Drawing.SystemColors.Control
		Me.cmdImport.CausesValidation = True
		Me.cmdImport.Enabled = True
		Me.cmdImport.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdImport.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdImport.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdImport.TabStop = True
		Me.cmdImport.Name = "cmdImport"
		Me.cmdDefaults.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDefaults.Text = "Restore &Default Colors"
		Me.cmdDefaults.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDefaults.Size = New System.Drawing.Size(137, 17)
		Me.cmdDefaults.Location = New System.Drawing.Point(280, 104)
		Me.cmdDefaults.TabIndex = 50
		Me.cmdDefaults.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDefaults.CausesValidation = True
		Me.cmdDefaults.Enabled = True
		Me.cmdDefaults.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDefaults.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDefaults.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDefaults.TabStop = True
		Me.cmdDefaults.Name = "cmdDefaults"
		Me.cmdExport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdExport.Text = "&Export ColorList"
		Me.cmdExport.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdExport.Size = New System.Drawing.Size(97, 17)
		Me.cmdExport.Location = New System.Drawing.Point(320, 248)
		Me.cmdExport.TabIndex = 62
		Me.cmdExport.BackColor = System.Drawing.SystemColors.Control
		Me.cmdExport.CausesValidation = True
		Me.cmdExport.Enabled = True
		Me.cmdExport.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdExport.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdExport.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdExport.TabStop = True
		Me.cmdExport.Name = "cmdExport"
		Me.txtHTML.AutoSize = False
		Me.txtHTML.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtHTML.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtHTML.ForeColor = System.Drawing.Color.White
		Me.txtHTML.Size = New System.Drawing.Size(97, 19)
		Me.txtHTML.Location = New System.Drawing.Point(240, 216)
		Me.txtHTML.Maxlength = 6
		Me.txtHTML.TabIndex = 58
		Me.txtHTML.AcceptsReturn = True
		Me.txtHTML.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtHTML.CausesValidation = True
		Me.txtHTML.Enabled = True
		Me.txtHTML.HideSelection = True
		Me.txtHTML.ReadOnly = False
		Me.txtHTML.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtHTML.MultiLine = False
		Me.txtHTML.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtHTML.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtHTML.TabStop = True
		Me.txtHTML.Visible = True
		Me.txtHTML.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtHTML.Name = "txtHTML"
		Me.cmdHTMLGen.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHTMLGen.Text = "Generate"
		Me.cmdHTMLGen.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHTMLGen.Size = New System.Drawing.Size(73, 17)
		Me.cmdHTMLGen.Location = New System.Drawing.Point(344, 216)
		Me.cmdHTMLGen.TabIndex = 59
		Me.cmdHTMLGen.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHTMLGen.CausesValidation = True
		Me.cmdHTMLGen.Enabled = True
		Me.cmdHTMLGen.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHTMLGen.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHTMLGen.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHTMLGen.TabStop = True
		Me.cmdHTMLGen.Name = "cmdHTMLGen"
		Me.txtChanFont.AutoSize = False
		Me.txtChanFont.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtChanFont.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtChanFont.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtChanFont.ForeColor = System.Drawing.Color.White
		Me.txtChanFont.Size = New System.Drawing.Size(145, 19)
		Me.txtChanFont.Location = New System.Drawing.Point(128, 80)
		Me.txtChanFont.TabIndex = 48
		Me.txtChanFont.AcceptsReturn = True
		Me.txtChanFont.CausesValidation = True
		Me.txtChanFont.Enabled = True
		Me.txtChanFont.HideSelection = True
		Me.txtChanFont.ReadOnly = False
		Me.txtChanFont.Maxlength = 0
		Me.txtChanFont.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtChanFont.MultiLine = False
		Me.txtChanFont.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtChanFont.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtChanFont.TabStop = True
		Me.txtChanFont.Visible = True
		Me.txtChanFont.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtChanFont.Name = "txtChanFont"
		Me.txtChatFont.AutoSize = False
		Me.txtChatFont.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtChatFont.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtChatFont.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtChatFont.ForeColor = System.Drawing.Color.White
		Me.txtChatFont.Size = New System.Drawing.Size(145, 19)
		Me.txtChatFont.Location = New System.Drawing.Point(128, 56)
		Me.txtChatFont.TabIndex = 46
		Me.txtChatFont.AcceptsReturn = True
		Me.txtChatFont.CausesValidation = True
		Me.txtChatFont.Enabled = True
		Me.txtChatFont.HideSelection = True
		Me.txtChatFont.ReadOnly = False
		Me.txtChatFont.Maxlength = 0
		Me.txtChatFont.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtChatFont.MultiLine = False
		Me.txtChatFont.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtChatFont.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtChatFont.TabStop = True
		Me.txtChatFont.Visible = True
		Me.txtChatFont.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtChatFont.Name = "txtChatFont"
		Me.txtChanSize.AutoSize = False
		Me.txtChanSize.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtChanSize.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtChanSize.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtChanSize.ForeColor = System.Drawing.Color.White
		Me.txtChanSize.Size = New System.Drawing.Size(41, 19)
		Me.txtChanSize.Location = New System.Drawing.Point(304, 80)
		Me.txtChanSize.Maxlength = 2
		Me.txtChanSize.TabIndex = 49
		Me.txtChanSize.AcceptsReturn = True
		Me.txtChanSize.CausesValidation = True
		Me.txtChanSize.Enabled = True
		Me.txtChanSize.HideSelection = True
		Me.txtChanSize.ReadOnly = False
		Me.txtChanSize.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtChanSize.MultiLine = False
		Me.txtChanSize.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtChanSize.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtChanSize.TabStop = True
		Me.txtChanSize.Visible = True
		Me.txtChanSize.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtChanSize.Name = "txtChanSize"
		Me.txtChatSize.AutoSize = False
		Me.txtChatSize.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtChatSize.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtChatSize.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtChatSize.ForeColor = System.Drawing.Color.White
		Me.txtChatSize.Size = New System.Drawing.Size(41, 19)
		Me.txtChatSize.Location = New System.Drawing.Point(304, 56)
		Me.txtChatSize.Maxlength = 2
		Me.txtChatSize.TabIndex = 47
		Me.txtChatSize.AcceptsReturn = True
		Me.txtChatSize.CausesValidation = True
		Me.txtChatSize.Enabled = True
		Me.txtChatSize.HideSelection = True
		Me.txtChatSize.ReadOnly = False
		Me.txtChatSize.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtChatSize.MultiLine = False
		Me.txtChatSize.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtChatSize.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtChatSize.TabStop = True
		Me.txtChatSize.Visible = True
		Me.txtChatSize.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtChatSize.Name = "txtChatSize"
		Me.lblCurrentValue.BackColor = System.Drawing.Color.Black
		Me.lblCurrentValue.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCurrentValue.ForeColor = System.Drawing.Color.White
		Me.lblCurrentValue.Size = New System.Drawing.Size(81, 17)
		Me.lblCurrentValue.Location = New System.Drawing.Point(136, 176)
		Me.lblCurrentValue.TabIndex = 165
		Me.lblCurrentValue.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCurrentValue.Enabled = True
		Me.lblCurrentValue.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCurrentValue.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCurrentValue.UseMnemonic = True
		Me.lblCurrentValue.Visible = True
		Me.lblCurrentValue.AutoSize = False
		Me.lblCurrentValue.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCurrentValue.Name = "lblCurrentValue"
		Me.lblEg.Size = New System.Drawing.Size(193, 25)
		Me.lblEg.Location = New System.Drawing.Point(224, 168)
		Me.lblEg.TabIndex = 164
		Me.lblEg.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblEg.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblEg.BackColor = System.Drawing.SystemColors.Control
		Me.lblEg.Enabled = True
		Me.lblEg.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblEg.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblEg.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblEg.UseMnemonic = True
		Me.lblEg.Visible = True
		Me.lblEg.AutoSize = False
		Me.lblEg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lblEg.Name = "lblEg"
		Me._Label1_14.BackColor = System.Drawing.Color.Black
		Me._Label1_14.Text = "Color to modify:"
		Me._Label1_14.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_14.ForeColor = System.Drawing.Color.White
		Me._Label1_14.Size = New System.Drawing.Size(81, 17)
		Me._Label1_14.Location = New System.Drawing.Point(24, 104)
		Me._Label1_14.TabIndex = 163
		Me._Label1_14.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_14.Enabled = True
		Me._Label1_14.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_14.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_14.UseMnemonic = True
		Me._Label1_14.Visible = True
		Me._Label1_14.AutoSize = False
		Me._Label1_14.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_14.Name = "_Label1_14"
		Me.Label3.BackColor = System.Drawing.Color.Black
		Me.Label3.Text = "New Value:                   Current Value:       Example:"
		Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.ForeColor = System.Drawing.Color.White
		Me.Label3.Size = New System.Drawing.Size(289, 17)
		Me.Label3.Location = New System.Drawing.Point(24, 152)
		Me.Label3.TabIndex = 162
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.Enabled = True
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label4.BackColor = System.Drawing.Color.Black
		Me.Label4.Text = "R:"
		Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.ForeColor = System.Drawing.Color.White
		Me.Label4.Size = New System.Drawing.Size(9, 17)
		Me.Label4.Location = New System.Drawing.Point(24, 224)
		Me.Label4.TabIndex = 161
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.Enabled = True
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label5.BackColor = System.Drawing.Color.Black
		Me.Label5.Text = "G:"
		Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.ForeColor = System.Drawing.Color.White
		Me.Label5.Size = New System.Drawing.Size(9, 17)
		Me.Label5.Location = New System.Drawing.Point(88, 224)
		Me.Label5.TabIndex = 160
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label5.Enabled = True
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label5.Name = "Label5"
		Me.Label6.BackColor = System.Drawing.Color.Black
		Me.Label6.Text = "B:"
		Me.Label6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.ForeColor = System.Drawing.Color.White
		Me.Label6.Size = New System.Drawing.Size(17, 17)
		Me.Label6.Location = New System.Drawing.Point(152, 224)
		Me.Label6.TabIndex = 159
		Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label6.Enabled = True
		Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label6.UseMnemonic = True
		Me.Label6.Visible = True
		Me.Label6.AutoSize = False
		Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label6.Name = "Label6"
		Me.Label7.BackColor = System.Drawing.Color.Black
		Me.Label7.Text = "Use HTML hexadecimal colors:"
		Me.Label7.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label7.ForeColor = System.Drawing.Color.White
		Me.Label7.Size = New System.Drawing.Size(161, 17)
		Me.Label7.Location = New System.Drawing.Point(224, 200)
		Me.Label7.TabIndex = 158
		Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label7.Enabled = True
		Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label7.UseMnemonic = True
		Me.Label7.Visible = True
		Me.Label7.AutoSize = False
		Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label7.Name = "Label7"
		Me._Label8_5.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._Label8_5.BackColor = System.Drawing.Color.Black
		Me._Label8_5.Text = "#"
		Me._Label8_5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label8_5.ForeColor = System.Drawing.Color.White
		Me._Label8_5.Size = New System.Drawing.Size(17, 17)
		Me._Label8_5.Location = New System.Drawing.Point(224, 216)
		Me._Label8_5.TabIndex = 157
		Me._Label8_5.Enabled = True
		Me._Label8_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label8_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label8_5.UseMnemonic = True
		Me._Label8_5.Visible = True
		Me._Label8_5.AutoSize = False
		Me._Label8_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label8_5.Name = "_Label8_5"
		Me._Label8_0.BackColor = System.Drawing.Color.Black
		Me._Label8_0.Text = "Size"
		Me._Label8_0.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label8_0.ForeColor = System.Drawing.Color.White
		Me._Label8_0.Size = New System.Drawing.Size(25, 17)
		Me._Label8_0.Location = New System.Drawing.Point(280, 56)
		Me._Label8_0.TabIndex = 155
		Me._Label8_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label8_0.Enabled = True
		Me._Label8_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label8_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label8_0.UseMnemonic = True
		Me._Label8_0.Visible = True
		Me._Label8_0.AutoSize = False
		Me._Label8_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label8_0.Name = "_Label8_0"
		Me._Label8_2.BackColor = System.Drawing.Color.Black
		Me._Label8_2.Text = "Channel List"
		Me._Label8_2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label8_2.ForeColor = System.Drawing.Color.White
		Me._Label8_2.Size = New System.Drawing.Size(65, 17)
		Me._Label8_2.Location = New System.Drawing.Point(24, 80)
		Me._Label8_2.TabIndex = 154
		Me.ToolTip1.SetToolTip(Me._Label8_2, "Changes the font settings for the channel list.")
		Me._Label8_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label8_2.Enabled = True
		Me._Label8_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label8_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label8_2.UseMnemonic = True
		Me._Label8_2.Visible = True
		Me._Label8_2.AutoSize = False
		Me._Label8_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label8_2.Name = "_Label8_2"
		Me._Label8_7.BackColor = System.Drawing.Color.Black
		Me._Label8_7.Text = "Chat Window"
		Me._Label8_7.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label8_7.ForeColor = System.Drawing.Color.White
		Me._Label8_7.Size = New System.Drawing.Size(65, 17)
		Me._Label8_7.Location = New System.Drawing.Point(24, 56)
		Me._Label8_7.TabIndex = 153
		Me.ToolTip1.SetToolTip(Me._Label8_7, "Changes the font setting for the main chat window.")
		Me._Label8_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label8_7.Enabled = True
		Me._Label8_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label8_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label8_7.UseMnemonic = True
		Me._Label8_7.Visible = True
		Me._Label8_7.AutoSize = False
		Me._Label8_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label8_7.Name = "_Label8_7"
		Me._Label8_1.BackColor = System.Drawing.Color.Black
		Me._Label8_1.Text = "Font"
		Me._Label8_1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label8_1.ForeColor = System.Drawing.Color.White
		Me._Label8_1.Size = New System.Drawing.Size(33, 17)
		Me._Label8_1.Location = New System.Drawing.Point(96, 56)
		Me._Label8_1.TabIndex = 152
		Me._Label8_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label8_1.Enabled = True
		Me._Label8_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label8_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label8_1.UseMnemonic = True
		Me._Label8_1.Visible = True
		Me._Label8_1.AutoSize = False
		Me._Label8_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label8_1.Name = "_Label8_1"
		Me._Label8_3.BackColor = System.Drawing.Color.Black
		Me._Label8_3.Text = "Size"
		Me._Label8_3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label8_3.ForeColor = System.Drawing.Color.White
		Me._Label8_3.Size = New System.Drawing.Size(25, 17)
		Me._Label8_3.Location = New System.Drawing.Point(280, 80)
		Me._Label8_3.TabIndex = 151
		Me._Label8_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label8_3.Enabled = True
		Me._Label8_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label8_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label8_3.UseMnemonic = True
		Me._Label8_3.Visible = True
		Me._Label8_3.AutoSize = False
		Me._Label8_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label8_3.Name = "_Label8_3"
		Me._Label8_4.BackColor = System.Drawing.Color.Black
		Me._Label8_4.Text = "Font"
		Me._Label8_4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label8_4.ForeColor = System.Drawing.Color.White
		Me._Label8_4.Size = New System.Drawing.Size(33, 17)
		Me._Label8_4.Location = New System.Drawing.Point(96, 80)
		Me._Label8_4.TabIndex = 150
		Me._Label8_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label8_4.Enabled = True
		Me._Label8_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label8_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label8_4.UseMnemonic = True
		Me._Label8_4.Visible = True
		Me._Label8_4.AutoSize = False
		Me._Label8_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label8_4.Name = "_Label8_4"
		Me._Line1_4.BorderColor = System.Drawing.Color.White
		Me._Line1_4.X1 = 24
		Me._Line1_4.X2 = 416
		Me._Line1_4.Y1 = 27
		Me._Line1_4.Y2 = 27
		Me._Line1_4.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line1_4.BorderWidth = 1
		Me._Line1_4.Visible = True
		Me._Line1_4.Name = "_Line1_4"
		Me._Label1_13.BackColor = System.Drawing.Color.Black
		Me._Label1_13.Text = "Interface font and color settings"
		Me._Label1_13.Font = New System.Drawing.Font("Tahoma", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_13.ForeColor = System.Drawing.Color.White
		Me._Label1_13.Size = New System.Drawing.Size(321, 25)
		Me._Label1_13.Location = New System.Drawing.Point(24, 16)
		Me._Label1_13.TabIndex = 149
		Me._Label1_13.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_13.Enabled = True
		Me._Label1_13.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_13.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_13.UseMnemonic = True
		Me._Label1_13.Visible = True
		Me._Label1_13.AutoSize = False
		Me._Label1_13.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_13.Name = "_Label1_13"
		Me.lblColorStatus.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblColorStatus.BackColor = System.Drawing.Color.Black
		Me.lblColorStatus.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblColorStatus.ForeColor = System.Drawing.Color.White
		Me.lblColorStatus.Size = New System.Drawing.Size(393, 17)
		Me.lblColorStatus.Location = New System.Drawing.Point(24, 280)
		Me.lblColorStatus.TabIndex = 190
		Me.lblColorStatus.Enabled = True
		Me.lblColorStatus.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblColorStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblColorStatus.UseMnemonic = True
		Me.lblColorStatus.Visible = True
		Me.lblColorStatus.AutoSize = False
		Me.lblColorStatus.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblColorStatus.Name = "lblColorStatus"
		Me._fraPanel_4.BackColor = System.Drawing.Color.Black
		Me._fraPanel_4.ForeColor = System.Drawing.Color.White
		Me._fraPanel_4.Size = New System.Drawing.Size(441, 321)
		Me._fraPanel_4.Location = New System.Drawing.Point(200, 0)
		Me._fraPanel_4.TabIndex = 122
		Me._fraPanel_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraPanel_4.Enabled = True
		Me._fraPanel_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraPanel_4.Visible = True
		Me._fraPanel_4.Padding = New System.Windows.Forms.Padding(0)
		Me._fraPanel_4.Name = "_fraPanel_4"
		Me.chkPhraseKick.BackColor = System.Drawing.Color.Black
		Me.chkPhraseKick.Text = "Kick"
		Me.chkPhraseKick.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkPhraseKick.ForeColor = System.Drawing.Color.White
		Me.chkPhraseKick.Size = New System.Drawing.Size(41, 17)
		Me.chkPhraseKick.Location = New System.Drawing.Point(144, 176)
		Me.chkPhraseKick.TabIndex = 70
		Me.ToolTip1.SetToolTip(Me.chkPhraseKick, "Instead of banning idle users, the bot will simply kick them.")
		Me.chkPhraseKick.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkPhraseKick.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkPhraseKick.CausesValidation = True
		Me.chkPhraseKick.Enabled = True
		Me.chkPhraseKick.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkPhraseKick.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkPhraseKick.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkPhraseKick.TabStop = True
		Me.chkPhraseKick.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkPhraseKick.Visible = True
		Me.chkPhraseKick.Name = "chkPhraseKick"
		Me.chkQuietKick.BackColor = System.Drawing.Color.Black
		Me.chkQuietKick.Text = "Kick"
		Me.chkQuietKick.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkQuietKick.ForeColor = System.Drawing.Color.White
		Me.chkQuietKick.Size = New System.Drawing.Size(41, 17)
		Me.chkQuietKick.Location = New System.Drawing.Point(144, 152)
		Me.chkQuietKick.TabIndex = 68
		Me.ToolTip1.SetToolTip(Me.chkQuietKick, "Instead of banning idle users, the bot will simply kick them.")
		Me.chkQuietKick.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkQuietKick.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkQuietKick.CausesValidation = True
		Me.chkQuietKick.Enabled = True
		Me.chkQuietKick.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkQuietKick.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkQuietKick.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkQuietKick.TabStop = True
		Me.chkQuietKick.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkQuietKick.Visible = True
		Me.chkQuietKick.Name = "chkQuietKick"
		Me.chkPingBan.BackColor = System.Drawing.Color.Black
		Me.chkPingBan.Text = "PingBan"
		Me.chkPingBan.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkPingBan.ForeColor = System.Drawing.Color.White
		Me.chkPingBan.Size = New System.Drawing.Size(65, 17)
		Me.chkPingBan.Location = New System.Drawing.Point(24, 208)
		Me.chkPingBan.TabIndex = 71
		Me.ToolTip1.SetToolTip(Me.chkPingBan, "Ban unsafelisted users who state banned phrases.")
		Me.chkPingBan.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkPingBan.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkPingBan.CausesValidation = True
		Me.chkPingBan.Enabled = True
		Me.chkPingBan.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkPingBan.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkPingBan.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkPingBan.TabStop = True
		Me.chkPingBan.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkPingBan.Visible = True
		Me.chkPingBan.Name = "chkPingBan"
		Me.txtPingLevel.AutoSize = False
		Me.txtPingLevel.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtPingLevel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPingLevel.ForeColor = System.Drawing.Color.White
		Me.txtPingLevel.Size = New System.Drawing.Size(49, 19)
		Me.txtPingLevel.Location = New System.Drawing.Point(136, 208)
		Me.txtPingLevel.Maxlength = 25
		Me.txtPingLevel.TabIndex = 72
		Me.txtPingLevel.AcceptsReturn = True
		Me.txtPingLevel.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPingLevel.CausesValidation = True
		Me.txtPingLevel.Enabled = True
		Me.txtPingLevel.HideSelection = True
		Me.txtPingLevel.ReadOnly = False
		Me.txtPingLevel.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPingLevel.MultiLine = False
		Me.txtPingLevel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPingLevel.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPingLevel.TabStop = True
		Me.txtPingLevel.Visible = True
		Me.txtPingLevel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPingLevel.Name = "txtPingLevel"
		Me.chkBanEvasion.BackColor = System.Drawing.Color.Black
		Me.chkBanEvasion.Text = "Use Ban Evasion"
		Me.chkBanEvasion.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkBanEvasion.ForeColor = System.Drawing.Color.White
		Me.chkBanEvasion.Size = New System.Drawing.Size(153, 17)
		Me.chkBanEvasion.Location = New System.Drawing.Point(24, 120)
		Me.chkBanEvasion.TabIndex = 66
		Me.ToolTip1.SetToolTip(Me.chkBanEvasion, "Ban Evasion attempts to keep people who are banned out of your channel.")
		Me.chkBanEvasion.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkBanEvasion.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkBanEvasion.CausesValidation = True
		Me.chkBanEvasion.Enabled = True
		Me.chkBanEvasion.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkBanEvasion.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkBanEvasion.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkBanEvasion.TabStop = True
		Me.chkBanEvasion.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkBanEvasion.Visible = True
		Me.chkBanEvasion.Name = "chkBanEvasion"
		Me.chkIdleKick.BackColor = System.Drawing.Color.Black
		Me.chkIdleKick.Text = "Kick instead of ban"
		Me.chkIdleKick.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkIdleKick.ForeColor = System.Drawing.Color.White
		Me.chkIdleKick.Size = New System.Drawing.Size(121, 17)
		Me.chkIdleKick.Location = New System.Drawing.Point(312, 48)
		Me.chkIdleKick.TabIndex = 76
		Me.ToolTip1.SetToolTip(Me.chkIdleKick, "Instead of banning idle users, the bot will simply kick them.")
		Me.chkIdleKick.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkIdleKick.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkIdleKick.CausesValidation = True
		Me.chkIdleKick.Enabled = True
		Me.chkIdleKick.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkIdleKick.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkIdleKick.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkIdleKick.TabStop = True
		Me.chkIdleKick.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkIdleKick.Visible = True
		Me.chkIdleKick.Name = "chkIdleKick"
		Me.chkPeonbans.BackColor = System.Drawing.Color.Black
		Me.chkPeonbans.Text = "Ban Warcraft III Peons"
		Me.chkPeonbans.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkPeonbans.ForeColor = System.Drawing.Color.White
		Me.chkPeonbans.Size = New System.Drawing.Size(153, 17)
		Me.chkPeonbans.Location = New System.Drawing.Point(216, 208)
		Me.chkPeonbans.TabIndex = 85
		Me.ToolTip1.SetToolTip(Me.chkPeonbans, "Ban Warcraft III users who have the Peon icon.")
		Me.chkPeonbans.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkPeonbans.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkPeonbans.CausesValidation = True
		Me.chkPeonbans.Enabled = True
		Me.chkPeonbans.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkPeonbans.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkPeonbans.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkPeonbans.TabStop = True
		Me.chkPeonbans.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkPeonbans.Visible = True
		Me.chkPeonbans.Name = "chkPeonbans"
		Me.txtLevelBanMsg.AutoSize = False
		Me.txtLevelBanMsg.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtLevelBanMsg.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLevelBanMsg.ForeColor = System.Drawing.Color.White
		Me.txtLevelBanMsg.Size = New System.Drawing.Size(217, 19)
		Me.txtLevelBanMsg.Location = New System.Drawing.Point(208, 287)
		Me.txtLevelBanMsg.Maxlength = 180
		Me.txtLevelBanMsg.TabIndex = 88
		Me.txtLevelBanMsg.Text = "You are below the required level for entry."
		Me.txtLevelBanMsg.AcceptsReturn = True
		Me.txtLevelBanMsg.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLevelBanMsg.CausesValidation = True
		Me.txtLevelBanMsg.Enabled = True
		Me.txtLevelBanMsg.HideSelection = True
		Me.txtLevelBanMsg.ReadOnly = False
		Me.txtLevelBanMsg.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLevelBanMsg.MultiLine = False
		Me.txtLevelBanMsg.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLevelBanMsg.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLevelBanMsg.TabStop = True
		Me.txtLevelBanMsg.Visible = True
		Me.txtLevelBanMsg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLevelBanMsg.Name = "txtLevelBanMsg"
		Me._chkCBan_5.BackColor = System.Drawing.Color.Black
		Me._chkCBan_5.Text = "The Frozen Throne"
		Me._chkCBan_5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkCBan_5.ForeColor = System.Drawing.Color.White
		Me._chkCBan_5.Size = New System.Drawing.Size(121, 17)
		Me._chkCBan_5.Location = New System.Drawing.Point(304, 160)
		Me._chkCBan_5.TabIndex = 83
		Me._chkCBan_5.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkCBan_5.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkCBan_5.CausesValidation = True
		Me._chkCBan_5.Enabled = True
		Me._chkCBan_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkCBan_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkCBan_5.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkCBan_5.TabStop = True
		Me._chkCBan_5.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkCBan_5.Visible = True
		Me._chkCBan_5.Name = "_chkCBan_5"
		Me._chkCBan_3.BackColor = System.Drawing.Color.Black
		Me._chkCBan_3.Text = "Lord of Destruction"
		Me._chkCBan_3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkCBan_3.ForeColor = System.Drawing.Color.White
		Me._chkCBan_3.Size = New System.Drawing.Size(121, 17)
		Me._chkCBan_3.Location = New System.Drawing.Point(304, 144)
		Me._chkCBan_3.TabIndex = 81
		Me._chkCBan_3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkCBan_3.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkCBan_3.CausesValidation = True
		Me._chkCBan_3.Enabled = True
		Me._chkCBan_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkCBan_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkCBan_3.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkCBan_3.TabStop = True
		Me._chkCBan_3.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkCBan_3.Visible = True
		Me._chkCBan_3.Name = "_chkCBan_3"
		Me.chkPhrasebans.BackColor = System.Drawing.Color.Black
		Me.chkPhrasebans.Text = "Enable Phrasebans"
		Me.chkPhrasebans.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkPhrasebans.ForeColor = System.Drawing.Color.White
		Me.chkPhrasebans.Size = New System.Drawing.Size(113, 17)
		Me.chkPhrasebans.Location = New System.Drawing.Point(24, 176)
		Me.chkPhrasebans.TabIndex = 69
		Me.ToolTip1.SetToolTip(Me.chkPhrasebans, "Ban unsafelisted users who state banned phrases.")
		Me.chkPhrasebans.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkPhrasebans.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkPhrasebans.CausesValidation = True
		Me.chkPhrasebans.Enabled = True
		Me.chkPhrasebans.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkPhrasebans.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkPhrasebans.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkPhrasebans.TabStop = True
		Me.chkPhrasebans.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkPhrasebans.Visible = True
		Me.chkPhrasebans.Name = "chkPhrasebans"
		Me.chkIPBans.BackColor = System.Drawing.Color.Black
		Me.chkIPBans.Text = "Enable IPBanning"
		Me.chkIPBans.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkIPBans.ForeColor = System.Drawing.Color.White
		Me.chkIPBans.Size = New System.Drawing.Size(129, 17)
		Me.chkIPBans.Location = New System.Drawing.Point(24, 48)
		Me.chkIPBans.TabIndex = 63
		Me.ToolTip1.SetToolTip(Me.chkIPBans, "Ban squelched users.")
		Me.chkIPBans.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkIPBans.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkIPBans.CausesValidation = True
		Me.chkIPBans.Enabled = True
		Me.chkIPBans.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkIPBans.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkIPBans.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkIPBans.TabStop = True
		Me.chkIPBans.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkIPBans.Visible = True
		Me.chkIPBans.Name = "chkIPBans"
		Me._chkCBan_0.BackColor = System.Drawing.Color.Black
		Me._chkCBan_0.Text = "Starcraft"
		Me._chkCBan_0.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkCBan_0.ForeColor = System.Drawing.Color.White
		Me._chkCBan_0.Size = New System.Drawing.Size(65, 17)
		Me._chkCBan_0.Location = New System.Drawing.Point(216, 128)
		Me._chkCBan_0.TabIndex = 78
		Me._chkCBan_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkCBan_0.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkCBan_0.CausesValidation = True
		Me._chkCBan_0.Enabled = True
		Me._chkCBan_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkCBan_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkCBan_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkCBan_0.TabStop = True
		Me._chkCBan_0.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkCBan_0.Visible = True
		Me._chkCBan_0.Name = "_chkCBan_0"
		Me._chkCBan_1.BackColor = System.Drawing.Color.Black
		Me._chkCBan_1.Text = "Brood War"
		Me._chkCBan_1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkCBan_1.ForeColor = System.Drawing.Color.White
		Me._chkCBan_1.Size = New System.Drawing.Size(81, 17)
		Me._chkCBan_1.Location = New System.Drawing.Point(304, 128)
		Me._chkCBan_1.TabIndex = 79
		Me._chkCBan_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkCBan_1.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkCBan_1.CausesValidation = True
		Me._chkCBan_1.Enabled = True
		Me._chkCBan_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkCBan_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkCBan_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkCBan_1.TabStop = True
		Me._chkCBan_1.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkCBan_1.Visible = True
		Me._chkCBan_1.Name = "_chkCBan_1"
		Me._chkCBan_2.BackColor = System.Drawing.Color.Black
		Me._chkCBan_2.Text = "Diablo II"
		Me._chkCBan_2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkCBan_2.ForeColor = System.Drawing.Color.White
		Me._chkCBan_2.Size = New System.Drawing.Size(65, 17)
		Me._chkCBan_2.Location = New System.Drawing.Point(216, 144)
		Me._chkCBan_2.TabIndex = 80
		Me._chkCBan_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkCBan_2.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkCBan_2.CausesValidation = True
		Me._chkCBan_2.Enabled = True
		Me._chkCBan_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkCBan_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkCBan_2.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkCBan_2.TabStop = True
		Me._chkCBan_2.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkCBan_2.Visible = True
		Me._chkCBan_2.Name = "_chkCBan_2"
		Me._chkCBan_6.BackColor = System.Drawing.Color.Black
		Me._chkCBan_6.Text = "Warcraft II"
		Me._chkCBan_6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkCBan_6.ForeColor = System.Drawing.Color.White
		Me._chkCBan_6.Size = New System.Drawing.Size(89, 17)
		Me._chkCBan_6.Location = New System.Drawing.Point(216, 176)
		Me._chkCBan_6.TabIndex = 84
		Me._chkCBan_6.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkCBan_6.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkCBan_6.CausesValidation = True
		Me._chkCBan_6.Enabled = True
		Me._chkCBan_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkCBan_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkCBan_6.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkCBan_6.TabStop = True
		Me._chkCBan_6.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkCBan_6.Visible = True
		Me._chkCBan_6.Name = "_chkCBan_6"
		Me._chkCBan_4.BackColor = System.Drawing.Color.Black
		Me._chkCBan_4.Text = "Warcraft III"
		Me._chkCBan_4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkCBan_4.ForeColor = System.Drawing.Color.White
		Me._chkCBan_4.Size = New System.Drawing.Size(89, 17)
		Me._chkCBan_4.Location = New System.Drawing.Point(216, 160)
		Me._chkCBan_4.TabIndex = 82
		Me._chkCBan_4.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkCBan_4.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkCBan_4.CausesValidation = True
		Me._chkCBan_4.Enabled = True
		Me._chkCBan_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkCBan_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkCBan_4.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkCBan_4.TabStop = True
		Me._chkCBan_4.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkCBan_4.Visible = True
		Me._chkCBan_4.Name = "_chkCBan_4"
		Me.chkQuiet.BackColor = System.Drawing.Color.Black
		Me.chkQuiet.Text = "Enable Quiet-Time"
		Me.chkQuiet.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkQuiet.ForeColor = System.Drawing.Color.White
		Me.chkQuiet.Size = New System.Drawing.Size(113, 17)
		Me.chkQuiet.Location = New System.Drawing.Point(24, 152)
		Me.chkQuiet.TabIndex = 67
		Me.ToolTip1.SetToolTip(Me.chkQuiet, "Ban unsafelisted users that talk.")
		Me.chkQuiet.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkQuiet.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkQuiet.CausesValidation = True
		Me.chkQuiet.Enabled = True
		Me.chkQuiet.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkQuiet.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkQuiet.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkQuiet.TabStop = True
		Me.chkQuiet.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkQuiet.Visible = True
		Me.chkQuiet.Name = "chkQuiet"
		Me.txtProtectMsg.AutoSize = False
		Me.txtProtectMsg.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtProtectMsg.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtProtectMsg.ForeColor = System.Drawing.Color.White
		Me.txtProtectMsg.Size = New System.Drawing.Size(161, 19)
		Me.txtProtectMsg.Location = New System.Drawing.Point(24, 287)
		Me.txtProtectMsg.Maxlength = 180
		Me.txtProtectMsg.TabIndex = 74
		Me.txtProtectMsg.Text = "Lockdown Enabled"
		Me.txtProtectMsg.AcceptsReturn = True
		Me.txtProtectMsg.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtProtectMsg.CausesValidation = True
		Me.txtProtectMsg.Enabled = True
		Me.txtProtectMsg.HideSelection = True
		Me.txtProtectMsg.ReadOnly = False
		Me.txtProtectMsg.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtProtectMsg.MultiLine = False
		Me.txtProtectMsg.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtProtectMsg.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtProtectMsg.TabStop = True
		Me.txtProtectMsg.Visible = True
		Me.txtProtectMsg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtProtectMsg.Name = "txtProtectMsg"
		Me.chkProtect.BackColor = System.Drawing.Color.Black
		Me.chkProtect.Text = "Enable Channel Protection"
		Me.chkProtect.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkProtect.ForeColor = System.Drawing.Color.White
		Me.chkProtect.Size = New System.Drawing.Size(153, 17)
		Me.chkProtect.Location = New System.Drawing.Point(24, 248)
		Me.chkProtect.TabIndex = 73
		Me.ToolTip1.SetToolTip(Me.chkProtect, "Ban unsafelisted users who join the channel.")
		Me.chkProtect.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkProtect.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkProtect.CausesValidation = True
		Me.chkProtect.Enabled = True
		Me.chkProtect.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkProtect.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkProtect.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkProtect.TabStop = True
		Me.chkProtect.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkProtect.Visible = True
		Me.chkProtect.Name = "chkProtect"
		Me.chkKOY.BackColor = System.Drawing.Color.Black
		Me.chkKOY.Text = "Enable Kick-On-Yell"
		Me.chkKOY.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkKOY.ForeColor = System.Drawing.Color.White
		Me.chkKOY.Size = New System.Drawing.Size(129, 17)
		Me.chkKOY.Location = New System.Drawing.Point(24, 96)
		Me.chkKOY.TabIndex = 65
		Me.ToolTip1.SetToolTip(Me.chkKOY, "Kick users who yell (uppercase message longer than 5 letters)")
		Me.chkKOY.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkKOY.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkKOY.CausesValidation = True
		Me.chkKOY.Enabled = True
		Me.chkKOY.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkKOY.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkKOY.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkKOY.TabStop = True
		Me.chkKOY.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkKOY.Visible = True
		Me.chkKOY.Name = "chkKOY"
		Me.txtBanW3.AutoSize = False
		Me.txtBanW3.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtBanW3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtBanW3.ForeColor = System.Drawing.Color.White
		Me.txtBanW3.Size = New System.Drawing.Size(33, 19)
		Me.txtBanW3.Location = New System.Drawing.Point(376, 248)
		Me.txtBanW3.Maxlength = 25
		Me.txtBanW3.TabIndex = 87
		Me.txtBanW3.AcceptsReturn = True
		Me.txtBanW3.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtBanW3.CausesValidation = True
		Me.txtBanW3.Enabled = True
		Me.txtBanW3.HideSelection = True
		Me.txtBanW3.ReadOnly = False
		Me.txtBanW3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtBanW3.MultiLine = False
		Me.txtBanW3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtBanW3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtBanW3.TabStop = True
		Me.txtBanW3.Visible = True
		Me.txtBanW3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtBanW3.Name = "txtBanW3"
		Me.chkPlugban.BackColor = System.Drawing.Color.Black
		Me.chkPlugban.Text = "Enable PlugBans"
		Me.chkPlugban.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkPlugban.ForeColor = System.Drawing.Color.White
		Me.chkPlugban.Size = New System.Drawing.Size(129, 17)
		Me.chkPlugban.Location = New System.Drawing.Point(24, 72)
		Me.chkPlugban.TabIndex = 64
		Me.ToolTip1.SetToolTip(Me.chkPlugban, "Ban users with a UDP plug.")
		Me.chkPlugban.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkPlugban.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkPlugban.CausesValidation = True
		Me.chkPlugban.Enabled = True
		Me.chkPlugban.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkPlugban.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkPlugban.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkPlugban.TabStop = True
		Me.chkPlugban.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkPlugban.Visible = True
		Me.chkPlugban.Name = "chkPlugban"
		Me.txtBanD2.AutoSize = False
		Me.txtBanD2.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtBanD2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtBanD2.ForeColor = System.Drawing.Color.White
		Me.txtBanD2.Size = New System.Drawing.Size(33, 19)
		Me.txtBanD2.Location = New System.Drawing.Point(272, 248)
		Me.txtBanD2.Maxlength = 25
		Me.txtBanD2.TabIndex = 86
		Me.txtBanD2.AcceptsReturn = True
		Me.txtBanD2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtBanD2.CausesValidation = True
		Me.txtBanD2.Enabled = True
		Me.txtBanD2.HideSelection = True
		Me.txtBanD2.ReadOnly = False
		Me.txtBanD2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtBanD2.MultiLine = False
		Me.txtBanD2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtBanD2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtBanD2.TabStop = True
		Me.txtBanD2.Visible = True
		Me.txtBanD2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtBanD2.Name = "txtBanD2"
		Me.chkIdlebans.BackColor = System.Drawing.Color.Black
		Me.chkIdlebans.Text = "Ban idle users"
		Me.chkIdlebans.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkIdlebans.ForeColor = System.Drawing.Color.White
		Me.chkIdlebans.Size = New System.Drawing.Size(97, 17)
		Me.chkIdlebans.Location = New System.Drawing.Point(216, 48)
		Me.chkIdlebans.TabIndex = 75
		Me.ToolTip1.SetToolTip(Me.chkIdlebans, "Ban users who have been idle for X seconds.")
		Me.chkIdlebans.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkIdlebans.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkIdlebans.CausesValidation = True
		Me.chkIdlebans.Enabled = True
		Me.chkIdlebans.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkIdlebans.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkIdlebans.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkIdlebans.TabStop = True
		Me.chkIdlebans.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkIdlebans.Visible = True
		Me.chkIdlebans.Name = "chkIdlebans"
		Me.txtIdleBanDelay.AutoSize = False
		Me.txtIdleBanDelay.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtIdleBanDelay.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtIdleBanDelay.ForeColor = System.Drawing.Color.White
		Me.txtIdleBanDelay.Size = New System.Drawing.Size(41, 19)
		Me.txtIdleBanDelay.Location = New System.Drawing.Point(344, 72)
		Me.txtIdleBanDelay.Maxlength = 25
		Me.txtIdleBanDelay.TabIndex = 77
		Me.txtIdleBanDelay.AcceptsReturn = True
		Me.txtIdleBanDelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtIdleBanDelay.CausesValidation = True
		Me.txtIdleBanDelay.Enabled = True
		Me.txtIdleBanDelay.HideSelection = True
		Me.txtIdleBanDelay.ReadOnly = False
		Me.txtIdleBanDelay.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtIdleBanDelay.MultiLine = False
		Me.txtIdleBanDelay.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtIdleBanDelay.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtIdleBanDelay.TabStop = True
		Me.txtIdleBanDelay.Visible = True
		Me.txtIdleBanDelay.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtIdleBanDelay.Name = "txtIdleBanDelay"
		Me._lblMod_7.BackColor = System.Drawing.Color.Black
		Me._lblMod_7.Text = "Level:"
		Me._lblMod_7.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblMod_7.ForeColor = System.Drawing.Color.White
		Me._lblMod_7.Size = New System.Drawing.Size(33, 17)
		Me._lblMod_7.Location = New System.Drawing.Point(104, 208)
		Me._lblMod_7.TabIndex = 207
		Me._lblMod_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblMod_7.Enabled = True
		Me._lblMod_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblMod_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblMod_7.UseMnemonic = True
		Me._lblMod_7.Visible = True
		Me._lblMod_7.AutoSize = False
		Me._lblMod_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblMod_7.Name = "_lblMod_7"
		Me.Line16.BorderColor = System.Drawing.Color.White
		Me.Line16.X1 = 192
		Me.Line16.X2 = 16
		Me.Line16.Y1 = 227
		Me.Line16.Y2 = 227
		Me.Line16.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line16.BorderWidth = 1
		Me.Line16.Visible = True
		Me.Line16.Name = "Line16"
		Me.Line15.BorderColor = System.Drawing.Color.White
		Me.Line15.X1 = 208
		Me.Line15.X2 = 424
		Me.Line15.Y1 = 187
		Me.Line15.Y2 = 187
		Me.Line15.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line15.BorderWidth = 1
		Me.Line15.Visible = True
		Me.Line15.Name = "Line15"
		Me.Line14.BorderColor = System.Drawing.Color.White
		Me.Line14.X1 = 208
		Me.Line14.X2 = 424
		Me.Line14.Y1 = 91
		Me.Line14.Y2 = 91
		Me.Line14.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line14.BorderWidth = 1
		Me.Line14.Visible = True
		Me.Line14.Name = "Line14"
		Me.Line13.BorderColor = System.Drawing.Color.White
		Me.Line13.X1 = 200
		Me.Line13.X2 = 200
		Me.Line13.Y1 = 291
		Me.Line13.Y2 = 35
		Me.Line13.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line13.BorderWidth = 1
		Me.Line13.Visible = True
		Me.Line13.Name = "Line13"
		Me._lblMod_0.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._lblMod_0.BackColor = System.Drawing.Color.Black
		Me._lblMod_0.Text = "Levelban message"
		Me._lblMod_0.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblMod_0.ForeColor = System.Drawing.Color.White
		Me._lblMod_0.Size = New System.Drawing.Size(89, 17)
		Me._lblMod_0.Location = New System.Drawing.Point(208, 271)
		Me._lblMod_0.TabIndex = 182
		Me._lblMod_0.Enabled = True
		Me._lblMod_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblMod_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblMod_0.UseMnemonic = True
		Me._lblMod_0.Visible = True
		Me._lblMod_0.AutoSize = False
		Me._lblMod_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblMod_0.Name = "_lblMod_0"
		Me._lblMod_2.BackColor = System.Drawing.Color.Black
		Me._lblMod_2.Text = "Seconds before ban:"
		Me._lblMod_2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblMod_2.ForeColor = System.Drawing.Color.White
		Me._lblMod_2.Size = New System.Drawing.Size(105, 17)
		Me._lblMod_2.Location = New System.Drawing.Point(240, 72)
		Me._lblMod_2.TabIndex = 172
		Me._lblMod_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblMod_2.Enabled = True
		Me._lblMod_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblMod_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblMod_2.UseMnemonic = True
		Me._lblMod_2.Visible = True
		Me._lblMod_2.AutoSize = False
		Me._lblMod_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblMod_2.Name = "_lblMod_2"
		Me._lblMod_3.BackColor = System.Drawing.Color.Black
		Me._lblMod_3.Text = "Clientbans"
		Me._lblMod_3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblMod_3.ForeColor = System.Drawing.Color.White
		Me._lblMod_3.Size = New System.Drawing.Size(57, 17)
		Me._lblMod_3.Location = New System.Drawing.Point(216, 112)
		Me._lblMod_3.TabIndex = 171
		Me._lblMod_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblMod_3.Enabled = True
		Me._lblMod_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblMod_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblMod_3.UseMnemonic = True
		Me._lblMod_3.Visible = True
		Me._lblMod_3.AutoSize = False
		Me._lblMod_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblMod_3.Name = "_lblMod_3"
		Me._lblMod_5.BackColor = System.Drawing.Color.Black
		Me._lblMod_5.Text = "Message"
		Me._lblMod_5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblMod_5.ForeColor = System.Drawing.Color.White
		Me._lblMod_5.Size = New System.Drawing.Size(129, 17)
		Me._lblMod_5.Location = New System.Drawing.Point(24, 271)
		Me._lblMod_5.TabIndex = 170
		Me.ToolTip1.SetToolTip(Me._lblMod_5, "Shorter is better")
		Me._lblMod_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblMod_5.Enabled = True
		Me._lblMod_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblMod_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblMod_5.UseMnemonic = True
		Me._lblMod_5.Visible = True
		Me._lblMod_5.AutoSize = False
		Me._lblMod_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblMod_5.Name = "_lblMod_5"
		Me._lblMod_1.BackColor = System.Drawing.Color.Black
		Me._lblMod_1.Text = "LevelBans: Set to 0 to disable."
		Me._lblMod_1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblMod_1.ForeColor = System.Drawing.Color.White
		Me._lblMod_1.Size = New System.Drawing.Size(153, 17)
		Me._lblMod_1.Location = New System.Drawing.Point(216, 232)
		Me._lblMod_1.TabIndex = 169
		Me._lblMod_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblMod_1.Enabled = True
		Me._lblMod_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblMod_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblMod_1.UseMnemonic = True
		Me._lblMod_1.Visible = True
		Me._lblMod_1.AutoSize = False
		Me._lblMod_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblMod_1.Name = "_lblMod_1"
		Me._lblMod_4.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._lblMod_4.BackColor = System.Drawing.Color.Black
		Me._lblMod_4.Text = "Diablo II"
		Me._lblMod_4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblMod_4.ForeColor = System.Drawing.Color.White
		Me._lblMod_4.Size = New System.Drawing.Size(49, 17)
		Me._lblMod_4.Location = New System.Drawing.Point(216, 248)
		Me._lblMod_4.TabIndex = 168
		Me._lblMod_4.Enabled = True
		Me._lblMod_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblMod_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblMod_4.UseMnemonic = True
		Me._lblMod_4.Visible = True
		Me._lblMod_4.AutoSize = False
		Me._lblMod_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblMod_4.Name = "_lblMod_4"
		Me._lblMod_6.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._lblMod_6.BackColor = System.Drawing.Color.Black
		Me._lblMod_6.Text = "Warcraft III"
		Me._lblMod_6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblMod_6.ForeColor = System.Drawing.Color.White
		Me._lblMod_6.Size = New System.Drawing.Size(65, 17)
		Me._lblMod_6.Location = New System.Drawing.Point(304, 248)
		Me._lblMod_6.TabIndex = 167
		Me._lblMod_6.Enabled = True
		Me._lblMod_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblMod_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblMod_6.UseMnemonic = True
		Me._lblMod_6.Visible = True
		Me._lblMod_6.AutoSize = False
		Me._lblMod_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblMod_6.Name = "_lblMod_6"
		Me._Line1_5.BorderColor = System.Drawing.Color.White
		Me._Line1_5.X1 = 24
		Me._Line1_5.X2 = 416
		Me._Line1_5.Y1 = 27
		Me._Line1_5.Y2 = 27
		Me._Line1_5.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line1_5.BorderWidth = 1
		Me._Line1_5.Visible = True
		Me._Line1_5.Name = "_Line1_5"
		Me._Label1_15.BackColor = System.Drawing.Color.Black
		Me._Label1_15.Text = "General moderation settings"
		Me._Label1_15.Font = New System.Drawing.Font("Tahoma", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_15.ForeColor = System.Drawing.Color.White
		Me._Label1_15.Size = New System.Drawing.Size(321, 25)
		Me._Label1_15.Location = New System.Drawing.Point(24, 16)
		Me._Label1_15.TabIndex = 166
		Me._Label1_15.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_15.Enabled = True
		Me._Label1_15.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_15.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_15.UseMnemonic = True
		Me._Label1_15.Visible = True
		Me._Label1_15.AutoSize = False
		Me._Label1_15.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_15.Name = "_Label1_15"
		Me._fraPanel_6.BackColor = System.Drawing.Color.Black
		Me._fraPanel_6.ForeColor = System.Drawing.Color.White
		Me._fraPanel_6.Size = New System.Drawing.Size(441, 321)
		Me._fraPanel_6.Location = New System.Drawing.Point(200, 0)
		Me._fraPanel_6.TabIndex = 124
		Me._fraPanel_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraPanel_6.Enabled = True
		Me._fraPanel_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraPanel_6.Visible = True
		Me._fraPanel_6.Padding = New System.Windows.Forms.Padding(0)
		Me._fraPanel_6.Name = "_fraPanel_6"
		Me.optMsg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optMsg.BackColor = System.Drawing.Color.Black
		Me.optMsg.Text = "Message"
		Me.optMsg.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optMsg.ForeColor = System.Drawing.Color.White
		Me.optMsg.Size = New System.Drawing.Size(65, 17)
		Me.optMsg.Location = New System.Drawing.Point(352, 56)
		Me.optMsg.Appearance = System.Windows.Forms.Appearance.Button
		Me.optMsg.TabIndex = 100
		Me.optMsg.Checked = True
		Me.optMsg.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optMsg.CausesValidation = True
		Me.optMsg.Enabled = True
		Me.optMsg.Cursor = System.Windows.Forms.Cursors.Default
		Me.optMsg.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optMsg.TabStop = True
		Me.optMsg.Visible = True
		Me.optMsg.Name = "optMsg"
		Me.optQuote.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optQuote.BackColor = System.Drawing.Color.Black
		Me.optQuote.Text = "Quote"
		Me.optQuote.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optQuote.ForeColor = System.Drawing.Color.White
		Me.optQuote.Size = New System.Drawing.Size(65, 17)
		Me.optQuote.Location = New System.Drawing.Point(352, 128)
		Me.optQuote.Appearance = System.Windows.Forms.Appearance.Button
		Me.optQuote.TabIndex = 103
		Me.optQuote.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optQuote.CausesValidation = True
		Me.optQuote.Enabled = True
		Me.optQuote.Cursor = System.Windows.Forms.Cursors.Default
		Me.optQuote.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optQuote.TabStop = True
		Me.optQuote.Checked = False
		Me.optQuote.Visible = True
		Me.optQuote.Name = "optQuote"
		Me.optUptime.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optUptime.BackColor = System.Drawing.Color.Black
		Me.optUptime.Text = "Uptime"
		Me.optUptime.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optUptime.ForeColor = System.Drawing.Color.White
		Me.optUptime.Size = New System.Drawing.Size(65, 17)
		Me.optUptime.Location = New System.Drawing.Point(352, 80)
		Me.optUptime.Appearance = System.Windows.Forms.Appearance.Button
		Me.optUptime.TabIndex = 101
		Me.optUptime.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optUptime.CausesValidation = True
		Me.optUptime.Enabled = True
		Me.optUptime.Cursor = System.Windows.Forms.Cursors.Default
		Me.optUptime.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optUptime.TabStop = True
		Me.optUptime.Checked = False
		Me.optUptime.Visible = True
		Me.optUptime.Name = "optUptime"
		Me.optMP3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optMP3.BackColor = System.Drawing.Color.Black
		Me.optMP3.Text = "MP3"
		Me.optMP3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optMP3.ForeColor = System.Drawing.Color.White
		Me.optMP3.Size = New System.Drawing.Size(65, 17)
		Me.optMP3.Location = New System.Drawing.Point(352, 104)
		Me.optMP3.Appearance = System.Windows.Forms.Appearance.Button
		Me.optMP3.TabIndex = 102
		Me.optMP3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optMP3.CausesValidation = True
		Me.optMP3.Enabled = True
		Me.optMP3.Cursor = System.Windows.Forms.Cursors.Default
		Me.optMP3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optMP3.TabStop = True
		Me.optMP3.Checked = False
		Me.optMP3.Visible = True
		Me.optMP3.Name = "optMP3"
		Me.chkIdles.BackColor = System.Drawing.Color.Black
		Me.chkIdles.Text = "Show anti-idle messages"
		Me.chkIdles.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkIdles.ForeColor = System.Drawing.Color.White
		Me.chkIdles.Size = New System.Drawing.Size(145, 17)
		Me.chkIdles.Location = New System.Drawing.Point(24, 56)
		Me.chkIdles.TabIndex = 97
		Me.chkIdles.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkIdles.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkIdles.CausesValidation = True
		Me.chkIdles.Enabled = True
		Me.chkIdles.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkIdles.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkIdles.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkIdles.TabStop = True
		Me.chkIdles.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkIdles.Visible = True
		Me.chkIdles.Name = "chkIdles"
		Me.txtIdleMsg.AutoSize = False
		Me.txtIdleMsg.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtIdleMsg.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtIdleMsg.ForeColor = System.Drawing.Color.White
		Me.txtIdleMsg.Size = New System.Drawing.Size(289, 19)
		Me.txtIdleMsg.Location = New System.Drawing.Point(24, 120)
		Me.txtIdleMsg.TabIndex = 99
		Me.txtIdleMsg.Text = "/me is a %ver"
		Me.txtIdleMsg.AcceptsReturn = True
		Me.txtIdleMsg.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtIdleMsg.CausesValidation = True
		Me.txtIdleMsg.Enabled = True
		Me.txtIdleMsg.HideSelection = True
		Me.txtIdleMsg.ReadOnly = False
		Me.txtIdleMsg.Maxlength = 0
		Me.txtIdleMsg.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtIdleMsg.MultiLine = False
		Me.txtIdleMsg.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtIdleMsg.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtIdleMsg.TabStop = True
		Me.txtIdleMsg.Visible = True
		Me.txtIdleMsg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtIdleMsg.Name = "txtIdleMsg"
		Me.txtIdleWait.AutoSize = False
		Me.txtIdleWait.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtIdleWait.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtIdleWait.ForeColor = System.Drawing.Color.White
		Me.txtIdleWait.Size = New System.Drawing.Size(33, 19)
		Me.txtIdleWait.Location = New System.Drawing.Point(200, 80)
		Me.txtIdleWait.TabIndex = 98
		Me.txtIdleWait.Text = "6"
		Me.txtIdleWait.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtIdleWait.AcceptsReturn = True
		Me.txtIdleWait.CausesValidation = True
		Me.txtIdleWait.Enabled = True
		Me.txtIdleWait.HideSelection = True
		Me.txtIdleWait.ReadOnly = False
		Me.txtIdleWait.Maxlength = 0
		Me.txtIdleWait.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtIdleWait.MultiLine = False
		Me.txtIdleWait.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtIdleWait.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtIdleWait.TabStop = True
		Me.txtIdleWait.Visible = True
		Me.txtIdleWait.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtIdleWait.Name = "txtIdleWait"
		Me._Line5_4.BorderColor = System.Drawing.Color.White
		Me._Line5_4.X1 = 256
		Me._Line5_4.X2 = 344
		Me._Line5_4.Y1 = 59
		Me._Line5_4.Y2 = 59
		Me._Line5_4.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line5_4.BorderWidth = 1
		Me._Line5_4.Visible = True
		Me._Line5_4.Name = "_Line5_4"
		Me._Line5_3.BorderColor = System.Drawing.Color.White
		Me._Line5_3.X1 = 336
		Me._Line5_3.X2 = 352
		Me._Line5_3.Y1 = 123
		Me._Line5_3.Y2 = 123
		Me._Line5_3.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line5_3.BorderWidth = 1
		Me._Line5_3.Visible = True
		Me._Line5_3.Name = "_Line5_3"
		Me._Line5_2.BorderColor = System.Drawing.Color.White
		Me._Line5_2.X1 = 336
		Me._Line5_2.X2 = 352
		Me._Line5_2.Y1 = 99
		Me._Line5_2.Y2 = 99
		Me._Line5_2.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line5_2.BorderWidth = 1
		Me._Line5_2.Visible = True
		Me._Line5_2.Name = "_Line5_2"
		Me._Line5_1.BorderColor = System.Drawing.Color.White
		Me._Line5_1.X1 = 336
		Me._Line5_1.X2 = 352
		Me._Line5_1.Y1 = 75
		Me._Line5_1.Y2 = 75
		Me._Line5_1.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line5_1.BorderWidth = 1
		Me._Line5_1.Visible = True
		Me._Line5_1.Name = "_Line5_1"
		Me._Line5_0.BorderColor = System.Drawing.Color.White
		Me._Line5_0.X1 = 336
		Me._Line5_0.X2 = 336
		Me._Line5_0.Y1 = 59
		Me._Line5_0.Y2 = 123
		Me._Line5_0.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line5_0.BorderWidth = 1
		Me._Line5_0.Visible = True
		Me._Line5_0.Name = "_Line5_0"
		Me._lblIdle_3.BackColor = System.Drawing.Color.Black
		Me._lblIdle_3.Text = "Idle message type"
		Me._lblIdle_3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblIdle_3.ForeColor = System.Drawing.Color.White
		Me._lblIdle_3.Size = New System.Drawing.Size(97, 17)
		Me._lblIdle_3.Location = New System.Drawing.Point(256, 56)
		Me._lblIdle_3.TabIndex = 180
		Me._lblIdle_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblIdle_3.Enabled = True
		Me._lblIdle_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblIdle_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblIdle_3.UseMnemonic = True
		Me._lblIdle_3.Visible = True
		Me._lblIdle_3.AutoSize = False
		Me._lblIdle_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblIdle_3.Name = "_lblIdle_3"
		Me._lblIdle_0.BackColor = System.Drawing.Color.Black
		Me._lblIdle_0.Text = "Delay between messages (minutes)"
		Me._lblIdle_0.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblIdle_0.ForeColor = System.Drawing.Color.White
		Me._lblIdle_0.Size = New System.Drawing.Size(201, 17)
		Me._lblIdle_0.Location = New System.Drawing.Point(24, 80)
		Me._lblIdle_0.TabIndex = 179
		Me._lblIdle_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblIdle_0.Enabled = True
		Me._lblIdle_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblIdle_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblIdle_0.UseMnemonic = True
		Me._lblIdle_0.Visible = True
		Me._lblIdle_0.AutoSize = False
		Me._lblIdle_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblIdle_0.Name = "_lblIdle_0"
		Me._lblIdle_1.BackColor = System.Drawing.Color.Black
		Me._lblIdle_1.Text = "Idle message"
		Me._lblIdle_1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblIdle_1.ForeColor = System.Drawing.Color.White
		Me._lblIdle_1.Size = New System.Drawing.Size(73, 17)
		Me._lblIdle_1.Location = New System.Drawing.Point(24, 104)
		Me._lblIdle_1.TabIndex = 178
		Me._lblIdle_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblIdle_1.Enabled = True
		Me._lblIdle_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblIdle_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblIdle_1.UseMnemonic = True
		Me._lblIdle_1.Visible = True
		Me._lblIdle_1.AutoSize = False
		Me._lblIdle_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblIdle_1.Name = "_lblIdle_1"
		Me._Line1_7.BorderColor = System.Drawing.Color.White
		Me._Line1_7.X1 = 24
		Me._Line1_7.X2 = 416
		Me._Line1_7.Y1 = 27
		Me._Line1_7.Y2 = 27
		Me._Line1_7.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line1_7.BorderWidth = 1
		Me._Line1_7.Visible = True
		Me._Line1_7.Name = "_Line1_7"
		Me.lblIdleVars.BackColor = System.Drawing.Color.Black
		Me.lblIdleVars.Text = "idle variable container label"
		Me.lblIdleVars.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblIdleVars.ForeColor = System.Drawing.Color.White
		Me.lblIdleVars.Size = New System.Drawing.Size(393, 129)
		Me.lblIdleVars.Location = New System.Drawing.Point(24, 152)
		Me.lblIdleVars.TabIndex = 176
		Me.lblIdleVars.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblIdleVars.Enabled = True
		Me.lblIdleVars.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblIdleVars.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblIdleVars.UseMnemonic = True
		Me.lblIdleVars.Visible = True
		Me.lblIdleVars.AutoSize = False
		Me.lblIdleVars.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblIdleVars.Name = "lblIdleVars"
		Me._Label1_17.BackColor = System.Drawing.Color.Black
		Me._Label1_17.Text = "Idle message settings"
		Me._Label1_17.Font = New System.Drawing.Font("Tahoma", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_17.ForeColor = System.Drawing.Color.White
		Me._Label1_17.Size = New System.Drawing.Size(321, 25)
		Me._Label1_17.Location = New System.Drawing.Point(24, 16)
		Me._Label1_17.TabIndex = 177
		Me._Label1_17.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_17.Enabled = True
		Me._Label1_17.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_17.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_17.UseMnemonic = True
		Me._Label1_17.Visible = True
		Me._Label1_17.AutoSize = False
		Me._Label1_17.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_17.Name = "_Label1_17"
		Me._fraPanel_7.BackColor = System.Drawing.Color.Black
		Me._fraPanel_7.ForeColor = System.Drawing.Color.White
		Me._fraPanel_7.Size = New System.Drawing.Size(441, 321)
		Me._fraPanel_7.Location = New System.Drawing.Point(200, 0)
		Me._fraPanel_7.TabIndex = 125
		Me._fraPanel_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraPanel_7.Enabled = True
		Me._fraPanel_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraPanel_7.Visible = True
		Me._fraPanel_7.Padding = New System.Windows.Forms.Padding(0)
		Me._fraPanel_7.Name = "_fraPanel_7"
		Me.txtOwner.AutoSize = False
		Me.txtOwner.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtOwner.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtOwner.ForeColor = System.Drawing.Color.White
		Me.txtOwner.Size = New System.Drawing.Size(161, 19)
		Me.txtOwner.Location = New System.Drawing.Point(24, 256)
		Me.txtOwner.Maxlength = 30
		Me.txtOwner.TabIndex = 116
		Me.ToolTip1.SetToolTip(Me.txtOwner, "This account has full control over the bot. Use carefully!")
		Me.txtOwner.AcceptsReturn = True
		Me.txtOwner.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtOwner.CausesValidation = True
		Me.txtOwner.Enabled = True
		Me.txtOwner.HideSelection = True
		Me.txtOwner.ReadOnly = False
		Me.txtOwner.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtOwner.MultiLine = False
		Me.txtOwner.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtOwner.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtOwner.TabStop = True
		Me.txtOwner.Visible = True
		Me.txtOwner.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtOwner.Name = "txtOwner"
		Me.txtTrigger.AutoSize = False
		Me.txtTrigger.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtTrigger.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtTrigger.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtTrigger.ForeColor = System.Drawing.Color.White
		Me.txtTrigger.Size = New System.Drawing.Size(41, 19)
		Me.txtTrigger.Location = New System.Drawing.Point(128, 288)
		Me.txtTrigger.TabIndex = 117
		Me.txtTrigger.Text = "."
		Me.txtTrigger.AcceptsReturn = True
		Me.txtTrigger.CausesValidation = True
		Me.txtTrigger.Enabled = True
		Me.txtTrigger.HideSelection = True
		Me.txtTrigger.ReadOnly = False
		Me.txtTrigger.Maxlength = 0
		Me.txtTrigger.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTrigger.MultiLine = False
		Me.txtTrigger.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTrigger.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTrigger.TabStop = True
		Me.txtTrigger.Visible = True
		Me.txtTrigger.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtTrigger.Name = "txtTrigger"
		Me.chkD2Naming.BackColor = System.Drawing.Color.Black
		Me.chkD2Naming.Text = "Use Diablo II naming conventions"
		Me.chkD2Naming.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkD2Naming.ForeColor = System.Drawing.Color.White
		Me.chkD2Naming.Size = New System.Drawing.Size(193, 17)
		Me.chkD2Naming.Location = New System.Drawing.Point(232, 288)
		Me.chkD2Naming.TabIndex = 115
		Me.ToolTip1.SetToolTip(Me.chkD2Naming, "Show usernames with Diablo II naming conventions.")
		Me.chkD2Naming.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkD2Naming.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkD2Naming.CausesValidation = True
		Me.chkD2Naming.Enabled = True
		Me.chkD2Naming.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkD2Naming.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkD2Naming.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkD2Naming.TabStop = True
		Me.chkD2Naming.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkD2Naming.Visible = True
		Me.chkD2Naming.Name = "chkD2Naming"
		Me._optNaming_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._optNaming_3.BackColor = System.Drawing.Color.Black
		Me._optNaming_3.Text = "Show all"
		Me._optNaming_3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optNaming_3.ForeColor = System.Drawing.Color.White
		Me._optNaming_3.Size = New System.Drawing.Size(81, 17)
		Me._optNaming_3.Location = New System.Drawing.Point(320, 264)
		Me._optNaming_3.Appearance = System.Windows.Forms.Appearance.Button
		Me._optNaming_3.TabIndex = 114
		Me.ToolTip1.SetToolTip(Me._optNaming_3, "Show usernames with all gateways.")
		Me._optNaming_3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optNaming_3.CausesValidation = True
		Me._optNaming_3.Enabled = True
		Me._optNaming_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._optNaming_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optNaming_3.TabStop = True
		Me._optNaming_3.Checked = False
		Me._optNaming_3.Visible = True
		Me._optNaming_3.Name = "_optNaming_3"
		Me._optNaming_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._optNaming_2.BackColor = System.Drawing.Color.Black
		Me._optNaming_2.Text = "WarCraft III"
		Me._optNaming_2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optNaming_2.ForeColor = System.Drawing.Color.White
		Me._optNaming_2.Size = New System.Drawing.Size(81, 17)
		Me._optNaming_2.Location = New System.Drawing.Point(232, 264)
		Me._optNaming_2.Appearance = System.Windows.Forms.Appearance.Button
		Me._optNaming_2.TabIndex = 113
		Me.ToolTip1.SetToolTip(Me._optNaming_2, "Show usernames as they would appear to a WarCraft III user.")
		Me._optNaming_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optNaming_2.CausesValidation = True
		Me._optNaming_2.Enabled = True
		Me._optNaming_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._optNaming_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optNaming_2.TabStop = True
		Me._optNaming_2.Checked = False
		Me._optNaming_2.Visible = True
		Me._optNaming_2.Name = "_optNaming_2"
		Me._optNaming_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._optNaming_1.BackColor = System.Drawing.Color.Black
		Me._optNaming_1.Text = "Legacy"
		Me._optNaming_1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optNaming_1.ForeColor = System.Drawing.Color.White
		Me._optNaming_1.Size = New System.Drawing.Size(81, 17)
		Me._optNaming_1.Location = New System.Drawing.Point(320, 240)
		Me._optNaming_1.Appearance = System.Windows.Forms.Appearance.Button
		Me._optNaming_1.TabIndex = 112
		Me.ToolTip1.SetToolTip(Me._optNaming_1, "Show usernames as they would appear to a StarCraft or WarCraft II user.")
		Me._optNaming_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optNaming_1.CausesValidation = True
		Me._optNaming_1.Enabled = True
		Me._optNaming_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._optNaming_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optNaming_1.TabStop = True
		Me._optNaming_1.Checked = False
		Me._optNaming_1.Visible = True
		Me._optNaming_1.Name = "_optNaming_1"
		Me._optNaming_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._optNaming_0.BackColor = System.Drawing.Color.Black
		Me._optNaming_0.Text = "Default"
		Me._optNaming_0.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optNaming_0.ForeColor = System.Drawing.Color.White
		Me._optNaming_0.Size = New System.Drawing.Size(81, 17)
		Me._optNaming_0.Location = New System.Drawing.Point(232, 240)
		Me._optNaming_0.Appearance = System.Windows.Forms.Appearance.Button
		Me._optNaming_0.TabIndex = 111
		Me.ToolTip1.SetToolTip(Me._optNaming_0, "Show usernames as they would appear to the selected client.")
		Me._optNaming_0.Checked = True
		Me._optNaming_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optNaming_0.CausesValidation = True
		Me._optNaming_0.Enabled = True
		Me._optNaming_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._optNaming_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optNaming_0.TabStop = True
		Me._optNaming_0.Visible = True
		Me._optNaming_0.Name = "_optNaming_0"
		Me.chkShowOffline.BackColor = System.Drawing.Color.Black
		Me.chkShowOffline.Text = "Show offline friends"
		Me.chkShowOffline.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkShowOffline.ForeColor = System.Drawing.Color.White
		Me.chkShowOffline.Size = New System.Drawing.Size(161, 17)
		Me.chkShowOffline.Location = New System.Drawing.Point(232, 80)
		Me.chkShowOffline.TabIndex = 108
		Me.ToolTip1.SetToolTip(Me.chkShowOffline, "Determines whether or not offline friends are hidden from /f list")
		Me.chkShowOffline.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkShowOffline.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkShowOffline.CausesValidation = True
		Me.chkShowOffline.Enabled = True
		Me.chkShowOffline.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkShowOffline.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkShowOffline.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkShowOffline.TabStop = True
		Me.chkShowOffline.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkShowOffline.Visible = True
		Me.chkShowOffline.Name = "chkShowOffline"
		Me.txtBackupChan.AutoSize = False
		Me.txtBackupChan.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtBackupChan.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtBackupChan.ForeColor = System.Drawing.Color.White
		Me.txtBackupChan.Size = New System.Drawing.Size(169, 19)
		Me.txtBackupChan.Location = New System.Drawing.Point(232, 200)
		Me.txtBackupChan.Maxlength = 31
		Me.txtBackupChan.TabIndex = 110
		Me.ToolTip1.SetToolTip(Me.txtBackupChan, "The channel to go to when kicked. Leave blank to stay in the void.")
		Me.txtBackupChan.AcceptsReturn = True
		Me.txtBackupChan.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtBackupChan.CausesValidation = True
		Me.txtBackupChan.Enabled = True
		Me.txtBackupChan.HideSelection = True
		Me.txtBackupChan.ReadOnly = False
		Me.txtBackupChan.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtBackupChan.MultiLine = False
		Me.txtBackupChan.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtBackupChan.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtBackupChan.TabStop = True
		Me.txtBackupChan.Visible = True
		Me.txtBackupChan.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtBackupChan.Name = "txtBackupChan"
		Me.chkAllowMP3.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkAllowMP3.BackColor = System.Drawing.Color.Black
		Me.chkAllowMP3.Text = "Allow MP3 commands"
		Me.chkAllowMP3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkAllowMP3.ForeColor = System.Drawing.Color.White
		Me.chkAllowMP3.Size = New System.Drawing.Size(161, 17)
		Me.chkAllowMP3.Location = New System.Drawing.Point(24, 56)
		Me.chkAllowMP3.TabIndex = 104
		Me.ToolTip1.SetToolTip(Me.chkAllowMP3, "Allow commands such as .next and .back that change your current Winamp song.")
		Me.chkAllowMP3.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkAllowMP3.CausesValidation = True
		Me.chkAllowMP3.Enabled = True
		Me.chkAllowMP3.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkAllowMP3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkAllowMP3.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkAllowMP3.TabStop = True
		Me.chkAllowMP3.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkAllowMP3.Visible = True
		Me.chkAllowMP3.Name = "chkAllowMP3"
		Me.chkWhisperCmds.BackColor = System.Drawing.Color.Black
		Me.chkWhisperCmds.Text = "Whisper command responses"
		Me.chkWhisperCmds.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkWhisperCmds.ForeColor = System.Drawing.Color.White
		Me.chkWhisperCmds.Size = New System.Drawing.Size(161, 17)
		Me.chkWhisperCmds.Location = New System.Drawing.Point(232, 56)
		Me.chkWhisperCmds.TabIndex = 107
		Me.ToolTip1.SetToolTip(Me.chkWhisperCmds, "Whisper the return messages of all bot commands.")
		Me.chkWhisperCmds.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkWhisperCmds.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkWhisperCmds.CausesValidation = True
		Me.chkWhisperCmds.Enabled = True
		Me.chkWhisperCmds.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkWhisperCmds.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkWhisperCmds.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkWhisperCmds.TabStop = True
		Me.chkWhisperCmds.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkWhisperCmds.Visible = True
		Me.chkWhisperCmds.Name = "chkWhisperCmds"
		Me.chkPAmp.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkPAmp.BackColor = System.Drawing.Color.Black
		Me.chkPAmp.Text = "Use ProfileAmp"
		Me.chkPAmp.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkPAmp.ForeColor = System.Drawing.Color.White
		Me.chkPAmp.Size = New System.Drawing.Size(161, 17)
		Me.chkPAmp.Location = New System.Drawing.Point(24, 80)
		Me.chkPAmp.TabIndex = 105
		Me.ToolTip1.SetToolTip(Me.chkPAmp, "ProfileAmp writes Winamp's currently played song to your profile every 30 seconds")
		Me.chkPAmp.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkPAmp.CausesValidation = True
		Me.chkPAmp.Enabled = True
		Me.chkPAmp.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkPAmp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkPAmp.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkPAmp.TabStop = True
		Me.chkPAmp.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkPAmp.Visible = True
		Me.chkPAmp.Name = "chkPAmp"
		Me.chkMail.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.chkMail.BackColor = System.Drawing.Color.Black
		Me.chkMail.Text = "Enable Mail System"
		Me.chkMail.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkMail.ForeColor = System.Drawing.Color.White
		Me.chkMail.Size = New System.Drawing.Size(161, 17)
		Me.chkMail.Location = New System.Drawing.Point(24, 104)
		Me.chkMail.TabIndex = 106
		Me.ToolTip1.SetToolTip(Me.chkMail, "Enable/disable checking of the mail.ini file when people join.")
		Me.chkMail.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkMail.CausesValidation = True
		Me.chkMail.Enabled = True
		Me.chkMail.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkMail.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkMail.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkMail.TabStop = True
		Me.chkMail.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkMail.Visible = True
		Me.chkMail.Name = "chkMail"
		Me.chkBackup.BackColor = System.Drawing.Color.Black
		Me.chkBackup.Text = "Join a backup channel when kicked"
		Me.chkBackup.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkBackup.ForeColor = System.Drawing.Color.White
		Me.chkBackup.Size = New System.Drawing.Size(193, 25)
		Me.chkBackup.Location = New System.Drawing.Point(232, 160)
		Me.chkBackup.TabIndex = 109
		Me.ToolTip1.SetToolTip(Me.chkBackup, "The bot will join a specified channel when kicked, instead of rejoining.")
		Me.chkBackup.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkBackup.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkBackup.CausesValidation = True
		Me.chkBackup.Enabled = True
		Me.chkBackup.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkBackup.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkBackup.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkBackup.TabStop = True
		Me.chkBackup.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkBackup.Visible = True
		Me.chkBackup.Name = "chkBackup"
		Me._Label1_19.BackColor = System.Drawing.Color.Black
		Me._Label1_19.Text = "Bot Owner"
		Me._Label1_19.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_19.ForeColor = System.Drawing.Color.White
		Me._Label1_19.Size = New System.Drawing.Size(81, 17)
		Me._Label1_19.Location = New System.Drawing.Point(24, 240)
		Me._Label1_19.TabIndex = 206
		Me._Label1_19.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_19.Enabled = True
		Me._Label1_19.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_19.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_19.UseMnemonic = True
		Me._Label1_19.Visible = True
		Me._Label1_19.AutoSize = False
		Me._Label1_19.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_19.Name = "_Label1_19"
		Me._Label1_8.BackColor = System.Drawing.Color.Black
		Me._Label1_8.Text = "Command trigger"
		Me._Label1_8.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_8.ForeColor = System.Drawing.Color.White
		Me._Label1_8.Size = New System.Drawing.Size(89, 17)
		Me._Label1_8.Location = New System.Drawing.Point(32, 288)
		Me._Label1_8.TabIndex = 205
		Me.ToolTip1.SetToolTip(Me._Label1_8, "The command trigger is used to identify a chat message as a command.")
		Me._Label1_8.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_8.Enabled = True
		Me._Label1_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_8.UseMnemonic = True
		Me._Label1_8.Visible = True
		Me._Label1_8.AutoSize = False
		Me._Label1_8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_8.Name = "_Label1_8"
		Me._Label8_6.Text = "Gateway naming convention:"
		Me._Label8_6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label8_6.ForeColor = System.Drawing.Color.White
		Me._Label8_6.Size = New System.Drawing.Size(140, 13)
		Me._Label8_6.Location = New System.Drawing.Point(232, 224)
		Me._Label8_6.TabIndex = 193
		Me._Label8_6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label8_6.BackColor = System.Drawing.Color.Transparent
		Me._Label8_6.Enabled = True
		Me._Label8_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label8_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label8_6.UseMnemonic = True
		Me._Label8_6.Visible = True
		Me._Label8_6.AutoSize = True
		Me._Label8_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label8_6.Name = "_Label8_6"
		Me._Label8_10.BackColor = System.Drawing.Color.Black
		Me._Label8_10.Text = "Backup channel:"
		Me._Label8_10.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label8_10.ForeColor = System.Drawing.Color.White
		Me._Label8_10.Size = New System.Drawing.Size(161, 17)
		Me._Label8_10.Location = New System.Drawing.Point(232, 184)
		Me._Label8_10.TabIndex = 185
		Me._Label8_10.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label8_10.Enabled = True
		Me._Label8_10.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label8_10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label8_10.UseMnemonic = True
		Me._Label8_10.Visible = True
		Me._Label8_10.AutoSize = False
		Me._Label8_10.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label8_10.Name = "_Label8_10"
		Me._Line1_9.BorderColor = System.Drawing.Color.White
		Me._Line1_9.X1 = 208
		Me._Line1_9.X2 = 208
		Me._Line1_9.Y1 = 43
		Me._Line1_9.Y2 = 291
		Me._Line1_9.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line1_9.BorderWidth = 1
		Me._Line1_9.Visible = True
		Me._Line1_9.Name = "_Line1_9"
		Me._Line1_8.BorderColor = System.Drawing.Color.White
		Me._Line1_8.X1 = 24
		Me._Line1_8.X2 = 416
		Me._Line1_8.Y1 = 27
		Me._Line1_8.Y2 = 27
		Me._Line1_8.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line1_8.BorderWidth = 1
		Me._Line1_8.Visible = True
		Me._Line1_8.Name = "_Line1_8"
		Me._Label1_18.BackColor = System.Drawing.Color.Black
		Me._Label1_18.Text = "Miscellaneous general settings"
		Me._Label1_18.Font = New System.Drawing.Font("Tahoma", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_18.ForeColor = System.Drawing.Color.White
		Me._Label1_18.Size = New System.Drawing.Size(321, 25)
		Me._Label1_18.Location = New System.Drawing.Point(24, 16)
		Me._Label1_18.TabIndex = 181
		Me._Label1_18.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_18.Enabled = True
		Me._Label1_18.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_18.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_18.UseMnemonic = True
		Me._Label1_18.Visible = True
		Me._Label1_18.AutoSize = False
		Me._Label1_18.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_18.Name = "_Label1_18"
		Me._fraPanel_1.BackColor = System.Drawing.Color.Black
		Me._fraPanel_1.ForeColor = System.Drawing.Color.White
		Me._fraPanel_1.Size = New System.Drawing.Size(441, 321)
		Me._fraPanel_1.Location = New System.Drawing.Point(200, 0)
		Me._fraPanel_1.TabIndex = 119
		Me._fraPanel_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraPanel_1.Enabled = True
		Me._fraPanel_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraPanel_1.Visible = True
		Me._fraPanel_1.Padding = New System.Windows.Forms.Padding(0)
		Me._fraPanel_1.Name = "_fraPanel_1"
		Me.chkConnectOnStartup.BackColor = System.Drawing.Color.Black
		Me.chkConnectOnStartup.Text = "Connect on startup"
		Me.chkConnectOnStartup.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkConnectOnStartup.ForeColor = System.Drawing.Color.White
		Me.chkConnectOnStartup.Size = New System.Drawing.Size(121, 17)
		Me.chkConnectOnStartup.Location = New System.Drawing.Point(24, 176)
		Me.chkConnectOnStartup.TabIndex = 21
		Me.ToolTip1.SetToolTip(Me.chkConnectOnStartup, "Automatically connect when the bot starts up.")
		Me.chkConnectOnStartup.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkConnectOnStartup.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkConnectOnStartup.CausesValidation = True
		Me.chkConnectOnStartup.Enabled = True
		Me.chkConnectOnStartup.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkConnectOnStartup.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkConnectOnStartup.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkConnectOnStartup.TabStop = True
		Me.chkConnectOnStartup.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkConnectOnStartup.Visible = True
		Me.chkConnectOnStartup.Name = "chkConnectOnStartup"
		Me.txtEmail.AutoSize = False
		Me.txtEmail.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtEmail.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtEmail.ForeColor = System.Drawing.Color.White
		Me.txtEmail.Size = New System.Drawing.Size(201, 19)
		Me.txtEmail.Location = New System.Drawing.Point(208, 192)
		Me.txtEmail.TabIndex = 23
		Me.ToolTip1.SetToolTip(Me.txtEmail, "Created accounts will automatically be registered with this address. Leave blank to prompt each time.")
		Me.txtEmail.AcceptsReturn = True
		Me.txtEmail.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtEmail.CausesValidation = True
		Me.txtEmail.Enabled = True
		Me.txtEmail.HideSelection = True
		Me.txtEmail.ReadOnly = False
		Me.txtEmail.Maxlength = 0
		Me.txtEmail.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtEmail.MultiLine = False
		Me.txtEmail.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtEmail.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtEmail.TabStop = True
		Me.txtEmail.Visible = True
		Me.txtEmail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtEmail.Name = "txtEmail"
		Me.cboBNLSServer.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.cboBNLSServer.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboBNLSServer.ForeColor = System.Drawing.Color.White
		Me.cboBNLSServer.Size = New System.Drawing.Size(249, 21)
		Me.cboBNLSServer.Location = New System.Drawing.Point(168, 80)
		Me.cboBNLSServer.TabIndex = 20
		Me.cboBNLSServer.Text = "cboBNLSServer"
		Me.cboBNLSServer.CausesValidation = True
		Me.cboBNLSServer.Enabled = True
		Me.cboBNLSServer.IntegralHeight = True
		Me.cboBNLSServer.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboBNLSServer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboBNLSServer.Sorted = False
		Me.cboBNLSServer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboBNLSServer.TabStop = True
		Me.cboBNLSServer.Visible = True
		Me.cboBNLSServer.Name = "cboBNLSServer"
		Me.txtReconDelay.AutoSize = False
		Me.txtReconDelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me.txtReconDelay.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtReconDelay.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtReconDelay.ForeColor = System.Drawing.Color.White
		Me.txtReconDelay.Size = New System.Drawing.Size(41, 19)
		Me.txtReconDelay.Location = New System.Drawing.Point(112, 200)
		Me.txtReconDelay.Maxlength = 15
		Me.txtReconDelay.TabIndex = 22
		Me.txtReconDelay.Text = "1000"
		Me.txtReconDelay.AcceptsReturn = True
		Me.txtReconDelay.CausesValidation = True
		Me.txtReconDelay.Enabled = True
		Me.txtReconDelay.HideSelection = True
		Me.txtReconDelay.ReadOnly = False
		Me.txtReconDelay.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtReconDelay.MultiLine = False
		Me.txtReconDelay.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtReconDelay.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtReconDelay.TabStop = True
		Me.txtReconDelay.Visible = True
		Me.txtReconDelay.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtReconDelay.Name = "txtReconDelay"
		Me.optSocks5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optSocks5.BackColor = System.Drawing.Color.Black
		Me.optSocks5.Text = "SOCKS5"
		Me.optSocks5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optSocks5.ForeColor = System.Drawing.Color.White
		Me.optSocks5.Size = New System.Drawing.Size(49, 17)
		Me.optSocks5.Location = New System.Drawing.Point(360, 232)
		Me.optSocks5.Appearance = System.Windows.Forms.Appearance.Button
		Me.optSocks5.TabIndex = 30
		Me.optSocks5.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optSocks5.CausesValidation = True
		Me.optSocks5.Enabled = True
		Me.optSocks5.Cursor = System.Windows.Forms.Cursors.Default
		Me.optSocks5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optSocks5.TabStop = True
		Me.optSocks5.Checked = False
		Me.optSocks5.Visible = True
		Me.optSocks5.Name = "optSocks5"
		Me.optSocks4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.optSocks4.BackColor = System.Drawing.Color.Black
		Me.optSocks4.Text = "SOCKS4"
		Me.optSocks4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optSocks4.ForeColor = System.Drawing.Color.White
		Me.optSocks4.Size = New System.Drawing.Size(49, 17)
		Me.optSocks4.Location = New System.Drawing.Point(304, 232)
		Me.optSocks4.Appearance = System.Windows.Forms.Appearance.Button
		Me.optSocks4.TabIndex = 29
		Me.optSocks4.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optSocks4.CausesValidation = True
		Me.optSocks4.Enabled = True
		Me.optSocks4.Cursor = System.Windows.Forms.Cursors.Default
		Me.optSocks4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optSocks4.TabStop = True
		Me.optSocks4.Checked = False
		Me.optSocks4.Visible = True
		Me.optSocks4.Name = "optSocks4"
		Me.txtProxyPort.AutoSize = False
		Me.txtProxyPort.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtProxyPort.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtProxyPort.ForeColor = System.Drawing.Color.White
		Me.txtProxyPort.Size = New System.Drawing.Size(73, 19)
		Me.txtProxyPort.Location = New System.Drawing.Point(264, 280)
		Me.txtProxyPort.Maxlength = 5
		Me.txtProxyPort.TabIndex = 28
		Me.txtProxyPort.AcceptsReturn = True
		Me.txtProxyPort.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtProxyPort.CausesValidation = True
		Me.txtProxyPort.Enabled = True
		Me.txtProxyPort.HideSelection = True
		Me.txtProxyPort.ReadOnly = False
		Me.txtProxyPort.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtProxyPort.MultiLine = False
		Me.txtProxyPort.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtProxyPort.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtProxyPort.TabStop = True
		Me.txtProxyPort.Visible = True
		Me.txtProxyPort.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtProxyPort.Name = "txtProxyPort"
		Me.txtProxyIP.AutoSize = False
		Me.txtProxyIP.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtProxyIP.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtProxyIP.ForeColor = System.Drawing.Color.White
		Me.txtProxyIP.Size = New System.Drawing.Size(145, 19)
		Me.txtProxyIP.Location = New System.Drawing.Point(264, 256)
		Me.txtProxyIP.Maxlength = 15
		Me.txtProxyIP.TabIndex = 27
		Me.txtProxyIP.AcceptsReturn = True
		Me.txtProxyIP.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtProxyIP.CausesValidation = True
		Me.txtProxyIP.Enabled = True
		Me.txtProxyIP.HideSelection = True
		Me.txtProxyIP.ReadOnly = False
		Me.txtProxyIP.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtProxyIP.MultiLine = False
		Me.txtProxyIP.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtProxyIP.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtProxyIP.TabStop = True
		Me.txtProxyIP.Visible = True
		Me.txtProxyIP.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtProxyIP.Name = "txtProxyIP"
		Me.chkUseProxies.BackColor = System.Drawing.Color.Black
		Me.chkUseProxies.Text = "Use a proxy"
		Me.chkUseProxies.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkUseProxies.ForeColor = System.Drawing.Color.White
		Me.chkUseProxies.Size = New System.Drawing.Size(81, 17)
		Me.chkUseProxies.Location = New System.Drawing.Point(200, 232)
		Me.chkUseProxies.TabIndex = 26
		Me.ToolTip1.SetToolTip(Me.chkUseProxies, "Routes your Battle.net and/or BNLS connection through a SOCKS4 or 5 proxy.")
		Me.chkUseProxies.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkUseProxies.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkUseProxies.CausesValidation = True
		Me.chkUseProxies.Enabled = True
		Me.chkUseProxies.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkUseProxies.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkUseProxies.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkUseProxies.TabStop = True
		Me.chkUseProxies.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkUseProxies.Visible = True
		Me.chkUseProxies.Name = "chkUseProxies"
		Me.chkUDP.BackColor = System.Drawing.Color.Black
		Me.chkUDP.Text = "Use Lag Plug"
		Me.chkUDP.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkUDP.ForeColor = System.Drawing.Color.White
		Me.chkUDP.Size = New System.Drawing.Size(81, 17)
		Me.chkUDP.Location = New System.Drawing.Point(32, 280)
		Me.chkUDP.TabIndex = 25
		Me.ToolTip1.SetToolTip(Me.chkUDP, "Sets whether or not you have a lag plug when you sign on. If you don't know what this is, leave it off.")
		Me.chkUDP.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkUDP.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkUDP.CausesValidation = True
		Me.chkUDP.Enabled = True
		Me.chkUDP.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkUDP.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkUDP.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkUDP.TabStop = True
		Me.chkUDP.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkUDP.Visible = True
		Me.chkUDP.Name = "chkUDP"
		Me.cboSpoof.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.cboSpoof.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboSpoof.ForeColor = System.Drawing.Color.White
		Me.cboSpoof.Size = New System.Drawing.Size(145, 21)
		Me.cboSpoof.Location = New System.Drawing.Point(24, 256)
		Me.cboSpoof.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboSpoof.TabIndex = 24
		Me.cboSpoof.CausesValidation = True
		Me.cboSpoof.Enabled = True
		Me.cboSpoof.IntegralHeight = True
		Me.cboSpoof.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboSpoof.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboSpoof.Sorted = False
		Me.cboSpoof.TabStop = True
		Me.cboSpoof.Visible = True
		Me.cboSpoof.Name = "cboSpoof"
		Me.cboConnMethod.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.cboConnMethod.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboConnMethod.ForeColor = System.Drawing.Color.White
		Me.cboConnMethod.Size = New System.Drawing.Size(289, 21)
		Me.cboConnMethod.Location = New System.Drawing.Point(128, 56)
		Me.cboConnMethod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboConnMethod.TabIndex = 19
		Me.cboConnMethod.CausesValidation = True
		Me.cboConnMethod.Enabled = True
		Me.cboConnMethod.IntegralHeight = True
		Me.cboConnMethod.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboConnMethod.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboConnMethod.Sorted = False
		Me.cboConnMethod.TabStop = True
		Me.cboConnMethod.Visible = True
		Me.cboConnMethod.Name = "cboConnMethod"
		Me.Line7.BorderColor = System.Drawing.Color.White
		Me.Line7.X1 = 176
		Me.Line7.X2 = 16
		Me.Line7.Y1 = 211
		Me.Line7.Y2 = 211
		Me.Line7.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line7.BorderWidth = 1
		Me.Line7.Visible = True
		Me.Line7.Name = "Line7"
		Me._lbl5_13.BackColor = System.Drawing.Color.Black
		Me._lbl5_13.Text = "(for account registration)"
		Me._lbl5_13.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lbl5_13.ForeColor = System.Drawing.Color.White
		Me._lbl5_13.Size = New System.Drawing.Size(121, 17)
		Me._lbl5_13.Location = New System.Drawing.Point(288, 176)
		Me._lbl5_13.TabIndex = 195
		Me._lbl5_13.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lbl5_13.Enabled = True
		Me._lbl5_13.Cursor = System.Windows.Forms.Cursors.Default
		Me._lbl5_13.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lbl5_13.UseMnemonic = True
		Me._lbl5_13.Visible = True
		Me._lbl5_13.AutoSize = False
		Me._lbl5_13.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lbl5_13.Name = "_lbl5_13"
		Me._lbl5_1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._lbl5_1.BackColor = System.Drawing.Color.Black
		Me._lbl5_1.Text = "Email Address"
		Me._lbl5_1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lbl5_1.ForeColor = System.Drawing.Color.White
		Me._lbl5_1.Size = New System.Drawing.Size(73, 17)
		Me._lbl5_1.Location = New System.Drawing.Point(200, 176)
		Me._lbl5_1.TabIndex = 194
		Me._lbl5_1.Enabled = True
		Me._lbl5_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lbl5_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lbl5_1.UseMnemonic = True
		Me._lbl5_1.Visible = True
		Me._lbl5_1.AutoSize = False
		Me._lbl5_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lbl5_1.Name = "_lbl5_1"
		Me.Line6.BorderColor = System.Drawing.Color.White
		Me.Line6.X1 = 192
		Me.Line6.X2 = 424
		Me.Line6.Y1 = 211
		Me.Line6.Y2 = 211
		Me.Line6.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line6.BorderWidth = 1
		Me.Line6.Visible = True
		Me.Line6.Name = "Line6"
		Me._lbl5_12.BackColor = System.Drawing.Color.Black
		Me._lbl5_12.Text = "BNLS server, if applicable:"
		Me._lbl5_12.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lbl5_12.ForeColor = System.Drawing.Color.White
		Me._lbl5_12.Size = New System.Drawing.Size(137, 17)
		Me._lbl5_12.Location = New System.Drawing.Point(24, 80)
		Me._lbl5_12.TabIndex = 192
		Me._lbl5_12.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lbl5_12.Enabled = True
		Me._lbl5_12.Cursor = System.Windows.Forms.Cursors.Default
		Me._lbl5_12.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lbl5_12.UseMnemonic = True
		Me._lbl5_12.Visible = True
		Me._lbl5_12.AutoSize = False
		Me._lbl5_12.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lbl5_12.Name = "_lbl5_12"
		Me._lbl5_11.BackColor = System.Drawing.Color.Black
		Me._lbl5_11.Text = "ms"
		Me._lbl5_11.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lbl5_11.ForeColor = System.Drawing.Color.White
		Me._lbl5_11.Size = New System.Drawing.Size(17, 17)
		Me._lbl5_11.Location = New System.Drawing.Point(160, 200)
		Me._lbl5_11.TabIndex = 189
		Me._lbl5_11.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lbl5_11.Enabled = True
		Me._lbl5_11.Cursor = System.Windows.Forms.Cursors.Default
		Me._lbl5_11.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lbl5_11.UseMnemonic = True
		Me._lbl5_11.Visible = True
		Me._lbl5_11.AutoSize = False
		Me._lbl5_11.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lbl5_11.Name = "_lbl5_11"
		Me._lbl5_10.BackColor = System.Drawing.Color.Black
		Me._lbl5_10.Text = "Reconnect delay"
		Me._lbl5_10.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lbl5_10.ForeColor = System.Drawing.Color.White
		Me._lbl5_10.Size = New System.Drawing.Size(81, 17)
		Me._lbl5_10.Location = New System.Drawing.Point(24, 200)
		Me._lbl5_10.TabIndex = 188
		Me._lbl5_10.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lbl5_10.Enabled = True
		Me._lbl5_10.Cursor = System.Windows.Forms.Cursors.Default
		Me._lbl5_10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lbl5_10.UseMnemonic = True
		Me._lbl5_10.Visible = True
		Me._lbl5_10.AutoSize = False
		Me._lbl5_10.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lbl5_10.Name = "_lbl5_10"
		Me._lbl5_9.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._lbl5_9.BackColor = System.Drawing.Color.Black
		Me._lbl5_9.Text = "Port"
		Me._lbl5_9.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lbl5_9.ForeColor = System.Drawing.Color.White
		Me._lbl5_9.Size = New System.Drawing.Size(57, 17)
		Me._lbl5_9.Location = New System.Drawing.Point(200, 280)
		Me._lbl5_9.TabIndex = 184
		Me._lbl5_9.Enabled = True
		Me._lbl5_9.Cursor = System.Windows.Forms.Cursors.Default
		Me._lbl5_9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lbl5_9.UseMnemonic = True
		Me._lbl5_9.Visible = True
		Me._lbl5_9.AutoSize = False
		Me._lbl5_9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lbl5_9.Name = "_lbl5_9"
		Me._lbl5_8.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._lbl5_8.BackColor = System.Drawing.Color.Black
		Me._lbl5_8.Text = "IP address"
		Me._lbl5_8.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lbl5_8.ForeColor = System.Drawing.Color.White
		Me._lbl5_8.Size = New System.Drawing.Size(57, 17)
		Me._lbl5_8.Location = New System.Drawing.Point(200, 256)
		Me._lbl5_8.TabIndex = 183
		Me._lbl5_8.Enabled = True
		Me._lbl5_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._lbl5_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lbl5_8.UseMnemonic = True
		Me._lbl5_8.Visible = True
		Me._lbl5_8.AutoSize = False
		Me._lbl5_8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lbl5_8.Name = "_lbl5_8"
		Me.Line4.BorderColor = System.Drawing.Color.White
		Me.Line4.X1 = 184
		Me.Line4.X2 = 184
		Me.Line4.Y1 = 155
		Me.Line4.Y2 = 299
		Me.Line4.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line4.BorderWidth = 1
		Me.Line4.Visible = True
		Me.Line4.Name = "Line4"
		Me._lbl5_0.BackColor = System.Drawing.Color.Black
		Me._lbl5_0.Text = "Ping spoofing"
		Me._lbl5_0.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lbl5_0.ForeColor = System.Drawing.Color.White
		Me._lbl5_0.Size = New System.Drawing.Size(65, 17)
		Me._lbl5_0.Location = New System.Drawing.Point(16, 240)
		Me._lbl5_0.TabIndex = 148
		Me._lbl5_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lbl5_0.Enabled = True
		Me._lbl5_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lbl5_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lbl5_0.UseMnemonic = True
		Me._lbl5_0.Visible = True
		Me._lbl5_0.AutoSize = False
		Me._lbl5_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lbl5_0.Name = "_lbl5_0"
		Me.lblHashPath.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblHashPath.BackColor = System.Drawing.SystemColors.ControlText
		Me.lblHashPath.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblHashPath.ForeColor = System.Drawing.Color.White
		Me.lblHashPath.Size = New System.Drawing.Size(401, 25)
		Me.lblHashPath.Location = New System.Drawing.Point(24, 128)
		Me.lblHashPath.TabIndex = 147
		Me.lblHashPath.Enabled = True
		Me.lblHashPath.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblHashPath.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblHashPath.UseMnemonic = True
		Me.lblHashPath.Visible = True
		Me.lblHashPath.AutoSize = False
		Me.lblHashPath.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblHashPath.Name = "lblHashPath"
		Me._lbl5_2.BackColor = System.Drawing.Color.Black
		Me._lbl5_2.Text = "Connection method:"
		Me._lbl5_2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lbl5_2.ForeColor = System.Drawing.Color.White
		Me._lbl5_2.Size = New System.Drawing.Size(105, 17)
		Me._lbl5_2.Location = New System.Drawing.Point(24, 56)
		Me._lbl5_2.TabIndex = 146
		Me._lbl5_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lbl5_2.Enabled = True
		Me._lbl5_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lbl5_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lbl5_2.UseMnemonic = True
		Me._lbl5_2.Visible = True
		Me._lbl5_2.AutoSize = False
		Me._lbl5_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lbl5_2.Name = "_lbl5_2"
		Me._lbl5_3.BackColor = System.Drawing.Color.Black
		Me._lbl5_3.Text = "Local hashing is supported for all game clients. Your current hash file path is:"
		Me._lbl5_3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lbl5_3.ForeColor = System.Drawing.Color.White
		Me._lbl5_3.Size = New System.Drawing.Size(393, 17)
		Me._lbl5_3.Location = New System.Drawing.Point(24, 112)
		Me._lbl5_3.TabIndex = 145
		Me._lbl5_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lbl5_3.Enabled = True
		Me._lbl5_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lbl5_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lbl5_3.UseMnemonic = True
		Me._lbl5_3.Visible = True
		Me._lbl5_3.AutoSize = False
		Me._lbl5_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lbl5_3.Name = "_lbl5_3"
		Me._Line1_2.BorderColor = System.Drawing.Color.White
		Me._Line1_2.X1 = 24
		Me._Line1_2.X2 = 416
		Me._Line1_2.Y1 = 27
		Me._Line1_2.Y2 = 27
		Me._Line1_2.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line1_2.BorderWidth = 1
		Me._Line1_2.Visible = True
		Me._Line1_2.Name = "_Line1_2"
		Me._Label1_11.BackColor = System.Drawing.Color.Black
		Me._Label1_11.Text = "Advanced connection settings"
		Me._Label1_11.Font = New System.Drawing.Font("Tahoma", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_11.ForeColor = System.Drawing.Color.White
		Me._Label1_11.Size = New System.Drawing.Size(321, 25)
		Me._Label1_11.Location = New System.Drawing.Point(24, 16)
		Me._Label1_11.TabIndex = 143
		Me._Label1_11.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_11.Enabled = True
		Me._Label1_11.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_11.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_11.UseMnemonic = True
		Me._Label1_11.Visible = True
		Me._Label1_11.AutoSize = False
		Me._Label1_11.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_11.Name = "_Label1_11"
		Me._fraPanel_5.BackColor = System.Drawing.Color.Black
		Me._fraPanel_5.ForeColor = System.Drawing.Color.White
		Me._fraPanel_5.Size = New System.Drawing.Size(441, 321)
		Me._fraPanel_5.Location = New System.Drawing.Point(200, 0)
		Me._fraPanel_5.TabIndex = 123
		Me._fraPanel_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraPanel_5.Enabled = True
		Me._fraPanel_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraPanel_5.Visible = True
		Me._fraPanel_5.Padding = New System.Windows.Forms.Padding(0)
		Me._fraPanel_5.Name = "_fraPanel_5"
		Me.chkWhisperGreet.BackColor = System.Drawing.Color.Black
		Me.chkWhisperGreet.Text = "Whisper the greet message"
		Me.chkWhisperGreet.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkWhisperGreet.ForeColor = System.Drawing.Color.White
		Me.chkWhisperGreet.Size = New System.Drawing.Size(177, 17)
		Me.chkWhisperGreet.Location = New System.Drawing.Point(24, 80)
		Me.chkWhisperGreet.TabIndex = 95
		Me.chkWhisperGreet.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkWhisperGreet.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkWhisperGreet.CausesValidation = True
		Me.chkWhisperGreet.Enabled = True
		Me.chkWhisperGreet.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkWhisperGreet.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkWhisperGreet.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkWhisperGreet.TabStop = True
		Me.chkWhisperGreet.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkWhisperGreet.Visible = True
		Me.chkWhisperGreet.Name = "chkWhisperGreet"
		Me.txtGreetMsg.AutoSize = False
		Me.txtGreetMsg.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtGreetMsg.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtGreetMsg.ForeColor = System.Drawing.Color.White
		Me.txtGreetMsg.Size = New System.Drawing.Size(393, 19)
		Me.txtGreetMsg.Location = New System.Drawing.Point(24, 120)
		Me.txtGreetMsg.Maxlength = 200
		Me.txtGreetMsg.TabIndex = 96
		Me.txtGreetMsg.Text = "Welcome to %c, %0! I am a %v."
		Me.txtGreetMsg.AcceptsReturn = True
		Me.txtGreetMsg.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtGreetMsg.CausesValidation = True
		Me.txtGreetMsg.Enabled = True
		Me.txtGreetMsg.HideSelection = True
		Me.txtGreetMsg.ReadOnly = False
		Me.txtGreetMsg.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtGreetMsg.MultiLine = False
		Me.txtGreetMsg.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtGreetMsg.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtGreetMsg.TabStop = True
		Me.txtGreetMsg.Visible = True
		Me.txtGreetMsg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtGreetMsg.Name = "txtGreetMsg"
		Me.chkGreetMsg.BackColor = System.Drawing.Color.Black
		Me.chkGreetMsg.Text = "Greet users who join the channel"
		Me.chkGreetMsg.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkGreetMsg.ForeColor = System.Drawing.Color.White
		Me.chkGreetMsg.Size = New System.Drawing.Size(201, 17)
		Me.chkGreetMsg.Location = New System.Drawing.Point(24, 56)
		Me.chkGreetMsg.TabIndex = 94
		Me.chkGreetMsg.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkGreetMsg.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkGreetMsg.CausesValidation = True
		Me.chkGreetMsg.Enabled = True
		Me.chkGreetMsg.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkGreetMsg.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkGreetMsg.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkGreetMsg.TabStop = True
		Me.chkGreetMsg.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkGreetMsg.Visible = True
		Me.chkGreetMsg.Name = "chkGreetMsg"
		Me.lblGreetVars.BackColor = System.Drawing.Color.Black
		Me.lblGreetVars.Text = "greet variable container label"
		Me.lblGreetVars.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblGreetVars.ForeColor = System.Drawing.Color.White
		Me.lblGreetVars.Size = New System.Drawing.Size(393, 145)
		Me.lblGreetVars.Location = New System.Drawing.Point(24, 152)
		Me.lblGreetVars.TabIndex = 175
		Me.lblGreetVars.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblGreetVars.Enabled = True
		Me.lblGreetVars.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblGreetVars.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblGreetVars.UseMnemonic = True
		Me.lblGreetVars.Visible = True
		Me.lblGreetVars.AutoSize = False
		Me.lblGreetVars.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblGreetVars.Name = "lblGreetVars"
		Me._lblIdle_2.BackColor = System.Drawing.Color.Black
		Me._lblIdle_2.Text = "Greet Message"
		Me._lblIdle_2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblIdle_2.ForeColor = System.Drawing.Color.White
		Me._lblIdle_2.Size = New System.Drawing.Size(129, 17)
		Me._lblIdle_2.Location = New System.Drawing.Point(24, 104)
		Me._lblIdle_2.TabIndex = 174
		Me._lblIdle_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblIdle_2.Enabled = True
		Me._lblIdle_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblIdle_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblIdle_2.UseMnemonic = True
		Me._lblIdle_2.Visible = True
		Me._lblIdle_2.AutoSize = False
		Me._lblIdle_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblIdle_2.Name = "_lblIdle_2"
		Me._Line1_6.BorderColor = System.Drawing.Color.White
		Me._Line1_6.X1 = 24
		Me._Line1_6.X2 = 416
		Me._Line1_6.Y1 = 27
		Me._Line1_6.Y2 = 27
		Me._Line1_6.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line1_6.BorderWidth = 1
		Me._Line1_6.Visible = True
		Me._Line1_6.Name = "_Line1_6"
		Me._Label1_16.BackColor = System.Drawing.Color.Black
		Me._Label1_16.Text = "Greet message settings"
		Me._Label1_16.Font = New System.Drawing.Font("Tahoma", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_16.ForeColor = System.Drawing.Color.White
		Me._Label1_16.Size = New System.Drawing.Size(321, 25)
		Me._Label1_16.Location = New System.Drawing.Point(24, 16)
		Me._Label1_16.TabIndex = 173
		Me._Label1_16.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_16.Enabled = True
		Me._Label1_16.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_16.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_16.UseMnemonic = True
		Me._Label1_16.Visible = True
		Me._Label1_16.AutoSize = False
		Me._Label1_16.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_16.Name = "_Label1_16"
		Me._fraPanel_9.BackColor = System.Drawing.Color.Black
		Me._fraPanel_9.ForeColor = System.Drawing.Color.White
		Me._fraPanel_9.Size = New System.Drawing.Size(441, 321)
		Me._fraPanel_9.Location = New System.Drawing.Point(200, 0)
		Me._fraPanel_9.TabIndex = 196
		Me._fraPanel_9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraPanel_9.Enabled = True
		Me._fraPanel_9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraPanel_9.Visible = True
		Me._fraPanel_9.Padding = New System.Windows.Forms.Padding(0)
		Me._fraPanel_9.Name = "_fraPanel_9"
		Me.chkLogDBActions.BackColor = System.Drawing.Color.Black
		Me.chkLogDBActions.Text = "Log database changes"
		Me.chkLogDBActions.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkLogDBActions.ForeColor = System.Drawing.Color.White
		Me.chkLogDBActions.Size = New System.Drawing.Size(153, 17)
		Me.chkLogDBActions.Location = New System.Drawing.Point(32, 152)
		Me.chkLogDBActions.TabIndex = 90
		Me.ToolTip1.SetToolTip(Me.chkLogDBActions, "Any actions that change the database will be logged in the log folder in a file called 'database.txt'.")
		Me.chkLogDBActions.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLogDBActions.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLogDBActions.CausesValidation = True
		Me.chkLogDBActions.Enabled = True
		Me.chkLogDBActions.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLogDBActions.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLogDBActions.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLogDBActions.TabStop = True
		Me.chkLogDBActions.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLogDBActions.Visible = True
		Me.chkLogDBActions.Name = "chkLogDBActions"
		Me.chkLogAllCommands.BackColor = System.Drawing.Color.Black
		Me.chkLogAllCommands.Text = "Log all commands"
		Me.chkLogAllCommands.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkLogAllCommands.ForeColor = System.Drawing.Color.White
		Me.chkLogAllCommands.Size = New System.Drawing.Size(153, 17)
		Me.chkLogAllCommands.Location = New System.Drawing.Point(32, 174)
		Me.chkLogAllCommands.TabIndex = 91
		Me.ToolTip1.SetToolTip(Me.chkLogAllCommands, "Any commands issued to the bot will be logged in a file in the bot's Logs folder called 'commandlog.txt'.")
		Me.chkLogAllCommands.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLogAllCommands.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLogAllCommands.CausesValidation = True
		Me.chkLogAllCommands.Enabled = True
		Me.chkLogAllCommands.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLogAllCommands.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLogAllCommands.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLogAllCommands.TabStop = True
		Me.chkLogAllCommands.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLogAllCommands.Visible = True
		Me.chkLogAllCommands.Name = "chkLogAllCommands"
		Me.txtMaxLogSize.AutoSize = False
		Me.txtMaxLogSize.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtMaxLogSize.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtMaxLogSize.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtMaxLogSize.ForeColor = System.Drawing.Color.White
		Me.txtMaxLogSize.Size = New System.Drawing.Size(49, 19)
		Me.txtMaxLogSize.Location = New System.Drawing.Point(136, 288)
		Me.txtMaxLogSize.TabIndex = 93
		Me.txtMaxLogSize.Text = "0"
		Me.txtMaxLogSize.AcceptsReturn = True
		Me.txtMaxLogSize.CausesValidation = True
		Me.txtMaxLogSize.Enabled = True
		Me.txtMaxLogSize.HideSelection = True
		Me.txtMaxLogSize.ReadOnly = False
		Me.txtMaxLogSize.Maxlength = 0
		Me.txtMaxLogSize.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMaxLogSize.MultiLine = False
		Me.txtMaxLogSize.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMaxLogSize.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMaxLogSize.TabStop = True
		Me.txtMaxLogSize.Visible = True
		Me.txtMaxLogSize.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMaxLogSize.Name = "txtMaxLogSize"
		Me.cboLogging.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.cboLogging.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboLogging.ForeColor = System.Drawing.Color.White
		Me.cboLogging.Size = New System.Drawing.Size(393, 21)
		Me.cboLogging.Location = New System.Drawing.Point(24, 112)
		Me.cboLogging.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboLogging.TabIndex = 89
		Me.cboLogging.CausesValidation = True
		Me.cboLogging.Enabled = True
		Me.cboLogging.IntegralHeight = True
		Me.cboLogging.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboLogging.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboLogging.Sorted = False
		Me.cboLogging.TabStop = True
		Me.cboLogging.Visible = True
		Me.cboLogging.Name = "cboLogging"
		Me.txtMaxBackLogSize.AutoSize = False
		Me.txtMaxBackLogSize.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
		Me.txtMaxBackLogSize.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtMaxBackLogSize.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtMaxBackLogSize.ForeColor = System.Drawing.Color.White
		Me.txtMaxBackLogSize.Size = New System.Drawing.Size(49, 19)
		Me.txtMaxBackLogSize.Location = New System.Drawing.Point(136, 264)
		Me.txtMaxBackLogSize.TabIndex = 92
		Me.txtMaxBackLogSize.Text = "10000"
		Me.ToolTip1.SetToolTip(Me.txtMaxBackLogSize, "Backlog is the text shown in the bot's chat window. Text is cleared from the top as new text is added. This must be done to keep the bot running smoothly.")
		Me.txtMaxBackLogSize.AcceptsReturn = True
		Me.txtMaxBackLogSize.CausesValidation = True
		Me.txtMaxBackLogSize.Enabled = True
		Me.txtMaxBackLogSize.HideSelection = True
		Me.txtMaxBackLogSize.ReadOnly = False
		Me.txtMaxBackLogSize.Maxlength = 0
		Me.txtMaxBackLogSize.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMaxBackLogSize.MultiLine = False
		Me.txtMaxBackLogSize.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMaxBackLogSize.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMaxBackLogSize.TabStop = True
		Me.txtMaxBackLogSize.Visible = True
		Me.txtMaxBackLogSize.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMaxBackLogSize.Name = "txtMaxBackLogSize"
		Me.lbl9.BackColor = System.Drawing.Color.Black
		Me.lbl9.Text = "Size limits (set 0 for unlimited)"
		Me.lbl9.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lbl9.ForeColor = System.Drawing.Color.White
		Me.lbl9.Size = New System.Drawing.Size(241, 17)
		Me.lbl9.Location = New System.Drawing.Point(24, 224)
		Me.lbl9.TabIndex = 204
		Me.lbl9.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lbl9.Enabled = True
		Me.lbl9.Cursor = System.Windows.Forms.Cursors.Default
		Me.lbl9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lbl9.UseMnemonic = True
		Me.lbl9.Visible = True
		Me.lbl9.AutoSize = False
		Me.lbl9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lbl9.Name = "lbl9"
		Me.Line8.BorderColor = System.Drawing.Color.White
		Me.Line8.X1 = 24
		Me.Line8.X2 = 248
		Me.Line8.Y1 = 235
		Me.Line8.Y2 = 235
		Me.Line8.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line8.BorderWidth = 1
		Me.Line8.Visible = True
		Me.Line8.Name = "Line8"
		Me._lbl5_5.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lbl5_5.BackColor = System.Drawing.Color.Black
		Me._lbl5_5.Text = "Log files are used to keep a record of events witnessed by your bot. They can also be used to track user database changes and command use. A separate file is created for each day."
		Me._lbl5_5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lbl5_5.ForeColor = System.Drawing.Color.White
		Me._lbl5_5.Size = New System.Drawing.Size(377, 41)
		Me._lbl5_5.Location = New System.Drawing.Point(32, 48)
		Me._lbl5_5.TabIndex = 203
		Me._lbl5_5.Enabled = True
		Me._lbl5_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._lbl5_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lbl5_5.UseMnemonic = True
		Me._lbl5_5.Visible = True
		Me._lbl5_5.AutoSize = False
		Me._lbl5_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lbl5_5.Name = "_lbl5_5"
		Me._lbl5_4.BackColor = System.Drawing.Color.Black
		Me._lbl5_4.Text = "Text logging"
		Me._lbl5_4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lbl5_4.ForeColor = System.Drawing.Color.White
		Me._lbl5_4.Size = New System.Drawing.Size(169, 17)
		Me._lbl5_4.Location = New System.Drawing.Point(32, 96)
		Me._lbl5_4.TabIndex = 202
		Me._lbl5_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lbl5_4.Enabled = True
		Me._lbl5_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._lbl5_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lbl5_4.UseMnemonic = True
		Me._lbl5_4.Visible = True
		Me._lbl5_4.AutoSize = False
		Me._lbl5_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lbl5_4.Name = "_lbl5_4"
		Me._lbl5_6.BackColor = System.Drawing.Color.Black
		Me._lbl5_6.Text = "Maximum logfile size"
		Me._lbl5_6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lbl5_6.ForeColor = System.Drawing.Color.White
		Me._lbl5_6.Size = New System.Drawing.Size(105, 17)
		Me._lbl5_6.Location = New System.Drawing.Point(24, 288)
		Me._lbl5_6.TabIndex = 201
		Me._lbl5_6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lbl5_6.Enabled = True
		Me._lbl5_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._lbl5_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lbl5_6.UseMnemonic = True
		Me._lbl5_6.Visible = True
		Me._lbl5_6.AutoSize = False
		Me._lbl5_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lbl5_6.Name = "_lbl5_6"
		Me._lbl5_7.BackColor = System.Drawing.Color.Black
		Me._lbl5_7.Text = "  megabytes"
		Me._lbl5_7.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lbl5_7.ForeColor = System.Drawing.Color.White
		Me._lbl5_7.Size = New System.Drawing.Size(65, 17)
		Me._lbl5_7.Location = New System.Drawing.Point(186, 288)
		Me._lbl5_7.TabIndex = 200
		Me._lbl5_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lbl5_7.Enabled = True
		Me._lbl5_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._lbl5_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lbl5_7.UseMnemonic = True
		Me._lbl5_7.Visible = True
		Me._lbl5_7.AutoSize = False
		Me._lbl5_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lbl5_7.Name = "_lbl5_7"
		Me.lblBacklogSize.BackColor = System.Drawing.Color.Black
		Me.lblBacklogSize.Text = "  bytes"
		Me.lblBacklogSize.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblBacklogSize.ForeColor = System.Drawing.Color.White
		Me.lblBacklogSize.Size = New System.Drawing.Size(65, 17)
		Me.lblBacklogSize.Location = New System.Drawing.Point(186, 264)
		Me.lblBacklogSize.TabIndex = 199
		Me.lblBacklogSize.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblBacklogSize.Enabled = True
		Me.lblBacklogSize.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblBacklogSize.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblBacklogSize.UseMnemonic = True
		Me.lblBacklogSize.Visible = True
		Me.lblBacklogSize.AutoSize = False
		Me.lblBacklogSize.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblBacklogSize.Name = "lblBacklogSize"
		Me.lblBacklog.BackColor = System.Drawing.Color.Black
		Me.lblBacklog.Text = "Maximum backlog size"
		Me.lblBacklog.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblBacklog.ForeColor = System.Drawing.Color.White
		Me.lblBacklog.Size = New System.Drawing.Size(105, 17)
		Me.lblBacklog.Location = New System.Drawing.Point(24, 264)
		Me.lblBacklog.TabIndex = 198
		Me.lblBacklog.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblBacklog.Enabled = True
		Me.lblBacklog.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblBacklog.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblBacklog.UseMnemonic = True
		Me.lblBacklog.Visible = True
		Me.lblBacklog.AutoSize = False
		Me.lblBacklog.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblBacklog.Name = "lblBacklog"
		Me._Label1_20.BackColor = System.Drawing.Color.Transparent
		Me._Label1_20.Text = "Logging settings"
		Me._Label1_20.Font = New System.Drawing.Font("Tahoma", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_20.ForeColor = System.Drawing.Color.White
		Me._Label1_20.Size = New System.Drawing.Size(217, 25)
		Me._Label1_20.Location = New System.Drawing.Point(24, 16)
		Me._Label1_20.TabIndex = 197
		Me._Label1_20.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_20.Enabled = True
		Me._Label1_20.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_20.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_20.UseMnemonic = True
		Me._Label1_20.Visible = True
		Me._Label1_20.AutoSize = False
		Me._Label1_20.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_20.Name = "_Label1_20"
		Me._Line1_10.BorderColor = System.Drawing.Color.White
		Me._Line1_10.X1 = 24
		Me._Line1_10.X2 = 416
		Me._Line1_10.Y1 = 27
		Me._Line1_10.Y2 = 27
		Me._Line1_10.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line1_10.BorderWidth = 1
		Me._Line1_10.Visible = True
		Me._Line1_10.Name = "_Line1_10"
		Me._fraPanel_8.BackColor = System.Drawing.Color.Black
		Me._fraPanel_8.ForeColor = System.Drawing.Color.White
		Me._fraPanel_8.Size = New System.Drawing.Size(441, 321)
		Me._fraPanel_8.Location = New System.Drawing.Point(200, 0)
		Me._fraPanel_8.TabIndex = 126
		Me._fraPanel_8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraPanel_8.Enabled = True
		Me._fraPanel_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraPanel_8.Visible = True
		Me._fraPanel_8.Padding = New System.Windows.Forms.Padding(0)
		Me._fraPanel_8.Name = "_fraPanel_8"
		Me.lblSplash.BackColor = System.Drawing.Color.Black
		Me.lblSplash.Text = "Splash message container label."
		Me.lblSplash.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblSplash.ForeColor = System.Drawing.Color.White
		Me.lblSplash.Size = New System.Drawing.Size(401, 209)
		Me.lblSplash.Location = New System.Drawing.Point(24, 64)
		Me.lblSplash.TabIndex = 128
		Me.lblSplash.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblSplash.Enabled = True
		Me.lblSplash.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblSplash.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSplash.UseMnemonic = True
		Me.lblSplash.Visible = True
		Me.lblSplash.AutoSize = False
		Me.lblSplash.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblSplash.Name = "lblSplash"
		Me._Line1_0.BorderColor = System.Drawing.Color.White
		Me._Line1_0.X1 = 24
		Me._Line1_0.X2 = 416
		Me._Line1_0.Y1 = 35
		Me._Line1_0.Y2 = 35
		Me._Line1_0.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line1_0.BorderWidth = 1
		Me._Line1_0.Visible = True
		Me._Line1_0.Name = "_Line1_0"
		Me._Label1_0.BackColor = System.Drawing.Color.Black
		Me._Label1_0.Text = "Welcome to &StealthBot"
		Me._Label1_0.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_0.ForeColor = System.Drawing.Color.White
		Me._Label1_0.Size = New System.Drawing.Size(217, 25)
		Me._Label1_0.Location = New System.Drawing.Point(24, 16)
		Me._Label1_0.TabIndex = 127
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_0.Enabled = True
		Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_0.UseMnemonic = True
		Me._Label1_0.Visible = True
		Me._Label1_0.AutoSize = False
		Me._Label1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_0.Name = "_Label1_0"
		Me.Controls.Add(cboProfile)
		Me.Controls.Add(cmdWebsite)
		Me.Controls.Add(cmdReadme)
		Me.Controls.Add(cmdStepByStep)
		Me.Controls.Add(cmdSave)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(tvw)
		Me.Controls.Add(_fraPanel_0)
		Me.Controls.Add(_fraPanel_2)
		Me.Controls.Add(_fraPanel_3)
		Me.Controls.Add(_fraPanel_4)
		Me.Controls.Add(_fraPanel_6)
		Me.Controls.Add(_fraPanel_7)
		Me.Controls.Add(_fraPanel_1)
		Me.Controls.Add(_fraPanel_5)
		Me.Controls.Add(_fraPanel_9)
		Me.Controls.Add(_fraPanel_8)
		Me._fraPanel_0.Controls.Add(txtCDKey)
		Me._fraPanel_0.Controls.Add(chkSHR)
		Me._fraPanel_0.Controls.Add(chkJPN)
		Me._fraPanel_0.Controls.Add(optDRTL)
		Me._fraPanel_0.Controls.Add(optW3XP)
		Me._fraPanel_0.Controls.Add(optSTAR)
		Me._fraPanel_0.Controls.Add(optSEXP)
		Me._fraPanel_0.Controls.Add(optD2DV)
		Me._fraPanel_0.Controls.Add(optD2XP)
		Me._fraPanel_0.Controls.Add(optW2BN)
		Me._fraPanel_0.Controls.Add(optWAR3)
		Me._fraPanel_0.Controls.Add(txtUsername)
		Me._fraPanel_0.Controls.Add(txtPassword)
		Me._fraPanel_0.Controls.Add(txtHomeChan)
		Me._fraPanel_0.Controls.Add(cboServer)
		Me._fraPanel_0.Controls.Add(txtExpKey)
		Me._fraPanel_0.Controls.Add(chkUseRealm)
		Me._fraPanel_0.Controls.Add(chkSpawn)
		Me._fraPanel_0.Controls.Add(lblAddCurrentKey)
		Me._fraPanel_0.Controls.Add(lblManageKeys)
		Me._fraPanel_0.Controls.Add(_Label1_10)
		Me.ShapeContainer1.Shapes.Add(_Line3_2)
		Me.ShapeContainer1.Shapes.Add(_Line3_1)
		Me.ShapeContainer1.Shapes.Add(_Line3_0)
		Me.ShapeContainer1.Shapes.Add(Line2)
		Me._fraPanel_0.Controls.Add(_Label1_9)
		Me.ShapeContainer1.Shapes.Add(_Line1_1)
		Me._fraPanel_0.Controls.Add(_Label1_6)
		Me._fraPanel_0.Controls.Add(_Label1_4)
		Me._fraPanel_0.Controls.Add(_Label1_1)
		Me._fraPanel_0.Controls.Add(_Label1_2)
		Me._fraPanel_0.Controls.Add(_Label1_3)
		Me._fraPanel_0.Controls.Add(_Label1_5)
		Me._fraPanel_0.Controls.Add(_Label1_7)
		Me._fraPanel_0.Controls.Add(ShapeContainer1)
		Me._fraPanel_2.Controls.Add(chkMinimizeOnStartup)
		Me._fraPanel_2.Controls.Add(chkURLDetect)
		Me._fraPanel_2.Controls.Add(chkDisablePrefix)
		Me._fraPanel_2.Controls.Add(chkDisableSuffix)
		Me._fraPanel_2.Controls.Add(chkShowUserGameStatsIcons)
		Me._fraPanel_2.Controls.Add(chkShowUserFlagsIcons)
		Me._fraPanel_2.Controls.Add(chkNoColoring)
		Me._fraPanel_2.Controls.Add(chkNoAutocomplete)
		Me._fraPanel_2.Controls.Add(chkNoTray)
		Me._fraPanel_2.Controls.Add(chkFlash)
		Me._fraPanel_2.Controls.Add(chkUTF8)
		Me._fraPanel_2.Controls.Add(cboTimestamp)
		Me._fraPanel_2.Controls.Add(chkSplash)
		Me._fraPanel_2.Controls.Add(chkFilter)
		Me._fraPanel_2.Controls.Add(chkJoinLeaves)
		Me.ShapeContainer2.Shapes.Add(Line12)
		Me.ShapeContainer2.Shapes.Add(Line11)
		Me.ShapeContainer2.Shapes.Add(Line10)
		Me.ShapeContainer2.Shapes.Add(Line9)
		Me._fraPanel_2.Controls.Add(_Label8_8)
		Me.ShapeContainer2.Shapes.Add(_Line1_3)
		Me._fraPanel_2.Controls.Add(_Label1_12)
		Me._fraPanel_2.Controls.Add(ShapeContainer2)
		Me._fraPanel_3.Controls.Add(cmdSaveColor)
		Me._fraPanel_3.Controls.Add(cboColorList)
		Me._fraPanel_3.Controls.Add(txtValue)
		Me._fraPanel_3.Controls.Add(cmdColorPicker)
		Me._fraPanel_3.Controls.Add(txtR)
		Me._fraPanel_3.Controls.Add(txtG)
		Me._fraPanel_3.Controls.Add(txtB)
		Me._fraPanel_3.Controls.Add(cmdGetRGB)
		Me._fraPanel_3.Controls.Add(cmdImport)
		Me._fraPanel_3.Controls.Add(cmdDefaults)
		Me._fraPanel_3.Controls.Add(cmdExport)
		Me._fraPanel_3.Controls.Add(txtHTML)
		Me._fraPanel_3.Controls.Add(cmdHTMLGen)
		Me._fraPanel_3.Controls.Add(txtChanFont)
		Me._fraPanel_3.Controls.Add(txtChatFont)
		Me._fraPanel_3.Controls.Add(txtChanSize)
		Me._fraPanel_3.Controls.Add(txtChatSize)
		Me._fraPanel_3.Controls.Add(lblCurrentValue)
		Me._fraPanel_3.Controls.Add(lblEg)
		Me._fraPanel_3.Controls.Add(_Label1_14)
		Me._fraPanel_3.Controls.Add(Label3)
		Me._fraPanel_3.Controls.Add(Label4)
		Me._fraPanel_3.Controls.Add(Label5)
		Me._fraPanel_3.Controls.Add(Label6)
		Me._fraPanel_3.Controls.Add(Label7)
		Me._fraPanel_3.Controls.Add(_Label8_5)
		Me._fraPanel_3.Controls.Add(_Label8_0)
		Me._fraPanel_3.Controls.Add(_Label8_2)
		Me._fraPanel_3.Controls.Add(_Label8_7)
		Me._fraPanel_3.Controls.Add(_Label8_1)
		Me._fraPanel_3.Controls.Add(_Label8_3)
		Me._fraPanel_3.Controls.Add(_Label8_4)
		Me.ShapeContainer3.Shapes.Add(_Line1_4)
		Me._fraPanel_3.Controls.Add(_Label1_13)
		Me._fraPanel_3.Controls.Add(lblColorStatus)
		Me._fraPanel_3.Controls.Add(ShapeContainer3)
		Me._fraPanel_4.Controls.Add(chkPhraseKick)
		Me._fraPanel_4.Controls.Add(chkQuietKick)
		Me._fraPanel_4.Controls.Add(chkPingBan)
		Me._fraPanel_4.Controls.Add(txtPingLevel)
		Me._fraPanel_4.Controls.Add(chkBanEvasion)
		Me._fraPanel_4.Controls.Add(chkIdleKick)
		Me._fraPanel_4.Controls.Add(chkPeonbans)
		Me._fraPanel_4.Controls.Add(txtLevelBanMsg)
		Me._fraPanel_4.Controls.Add(_chkCBan_5)
		Me._fraPanel_4.Controls.Add(_chkCBan_3)
		Me._fraPanel_4.Controls.Add(chkPhrasebans)
		Me._fraPanel_4.Controls.Add(chkIPBans)
		Me._fraPanel_4.Controls.Add(_chkCBan_0)
		Me._fraPanel_4.Controls.Add(_chkCBan_1)
		Me._fraPanel_4.Controls.Add(_chkCBan_2)
		Me._fraPanel_4.Controls.Add(_chkCBan_6)
		Me._fraPanel_4.Controls.Add(_chkCBan_4)
		Me._fraPanel_4.Controls.Add(chkQuiet)
		Me._fraPanel_4.Controls.Add(txtProtectMsg)
		Me._fraPanel_4.Controls.Add(chkProtect)
		Me._fraPanel_4.Controls.Add(chkKOY)
		Me._fraPanel_4.Controls.Add(txtBanW3)
		Me._fraPanel_4.Controls.Add(chkPlugban)
		Me._fraPanel_4.Controls.Add(txtBanD2)
		Me._fraPanel_4.Controls.Add(chkIdlebans)
		Me._fraPanel_4.Controls.Add(txtIdleBanDelay)
		Me._fraPanel_4.Controls.Add(_lblMod_7)
		Me.ShapeContainer4.Shapes.Add(Line16)
		Me.ShapeContainer4.Shapes.Add(Line15)
		Me.ShapeContainer4.Shapes.Add(Line14)
		Me.ShapeContainer4.Shapes.Add(Line13)
		Me._fraPanel_4.Controls.Add(_lblMod_0)
		Me._fraPanel_4.Controls.Add(_lblMod_2)
		Me._fraPanel_4.Controls.Add(_lblMod_3)
		Me._fraPanel_4.Controls.Add(_lblMod_5)
		Me._fraPanel_4.Controls.Add(_lblMod_1)
		Me._fraPanel_4.Controls.Add(_lblMod_4)
		Me._fraPanel_4.Controls.Add(_lblMod_6)
		Me.ShapeContainer4.Shapes.Add(_Line1_5)
		Me._fraPanel_4.Controls.Add(_Label1_15)
		Me._fraPanel_4.Controls.Add(ShapeContainer4)
		Me._fraPanel_6.Controls.Add(optMsg)
		Me._fraPanel_6.Controls.Add(optQuote)
		Me._fraPanel_6.Controls.Add(optUptime)
		Me._fraPanel_6.Controls.Add(optMP3)
		Me._fraPanel_6.Controls.Add(chkIdles)
		Me._fraPanel_6.Controls.Add(txtIdleMsg)
		Me._fraPanel_6.Controls.Add(txtIdleWait)
		Me.ShapeContainer5.Shapes.Add(_Line5_4)
		Me.ShapeContainer5.Shapes.Add(_Line5_3)
		Me.ShapeContainer5.Shapes.Add(_Line5_2)
		Me.ShapeContainer5.Shapes.Add(_Line5_1)
		Me.ShapeContainer5.Shapes.Add(_Line5_0)
		Me._fraPanel_6.Controls.Add(_lblIdle_3)
		Me._fraPanel_6.Controls.Add(_lblIdle_0)
		Me._fraPanel_6.Controls.Add(_lblIdle_1)
		Me.ShapeContainer5.Shapes.Add(_Line1_7)
		Me._fraPanel_6.Controls.Add(lblIdleVars)
		Me._fraPanel_6.Controls.Add(_Label1_17)
		Me._fraPanel_6.Controls.Add(ShapeContainer5)
		Me._fraPanel_7.Controls.Add(txtOwner)
		Me._fraPanel_7.Controls.Add(txtTrigger)
		Me._fraPanel_7.Controls.Add(chkD2Naming)
		Me._fraPanel_7.Controls.Add(_optNaming_3)
		Me._fraPanel_7.Controls.Add(_optNaming_2)
		Me._fraPanel_7.Controls.Add(_optNaming_1)
		Me._fraPanel_7.Controls.Add(_optNaming_0)
		Me._fraPanel_7.Controls.Add(chkShowOffline)
		Me._fraPanel_7.Controls.Add(txtBackupChan)
		Me._fraPanel_7.Controls.Add(chkAllowMP3)
		Me._fraPanel_7.Controls.Add(chkWhisperCmds)
		Me._fraPanel_7.Controls.Add(chkPAmp)
		Me._fraPanel_7.Controls.Add(chkMail)
		Me._fraPanel_7.Controls.Add(chkBackup)
		Me._fraPanel_7.Controls.Add(_Label1_19)
		Me._fraPanel_7.Controls.Add(_Label1_8)
		Me._fraPanel_7.Controls.Add(_Label8_6)
		Me._fraPanel_7.Controls.Add(_Label8_10)
		Me.ShapeContainer6.Shapes.Add(_Line1_9)
		Me.ShapeContainer6.Shapes.Add(_Line1_8)
		Me._fraPanel_7.Controls.Add(_Label1_18)
		Me._fraPanel_7.Controls.Add(ShapeContainer6)
		Me._fraPanel_1.Controls.Add(chkConnectOnStartup)
		Me._fraPanel_1.Controls.Add(txtEmail)
		Me._fraPanel_1.Controls.Add(cboBNLSServer)
		Me._fraPanel_1.Controls.Add(txtReconDelay)
		Me._fraPanel_1.Controls.Add(optSocks5)
		Me._fraPanel_1.Controls.Add(optSocks4)
		Me._fraPanel_1.Controls.Add(txtProxyPort)
		Me._fraPanel_1.Controls.Add(txtProxyIP)
		Me._fraPanel_1.Controls.Add(chkUseProxies)
		Me._fraPanel_1.Controls.Add(chkUDP)
		Me._fraPanel_1.Controls.Add(cboSpoof)
		Me._fraPanel_1.Controls.Add(cboConnMethod)
		Me.ShapeContainer7.Shapes.Add(Line7)
		Me._fraPanel_1.Controls.Add(_lbl5_13)
		Me._fraPanel_1.Controls.Add(_lbl5_1)
		Me.ShapeContainer7.Shapes.Add(Line6)
		Me._fraPanel_1.Controls.Add(_lbl5_12)
		Me._fraPanel_1.Controls.Add(_lbl5_11)
		Me._fraPanel_1.Controls.Add(_lbl5_10)
		Me._fraPanel_1.Controls.Add(_lbl5_9)
		Me._fraPanel_1.Controls.Add(_lbl5_8)
		Me.ShapeContainer7.Shapes.Add(Line4)
		Me._fraPanel_1.Controls.Add(_lbl5_0)
		Me._fraPanel_1.Controls.Add(lblHashPath)
		Me._fraPanel_1.Controls.Add(_lbl5_2)
		Me._fraPanel_1.Controls.Add(_lbl5_3)
		Me.ShapeContainer7.Shapes.Add(_Line1_2)
		Me._fraPanel_1.Controls.Add(_Label1_11)
		Me._fraPanel_1.Controls.Add(ShapeContainer7)
		Me._fraPanel_5.Controls.Add(chkWhisperGreet)
		Me._fraPanel_5.Controls.Add(txtGreetMsg)
		Me._fraPanel_5.Controls.Add(chkGreetMsg)
		Me._fraPanel_5.Controls.Add(lblGreetVars)
		Me._fraPanel_5.Controls.Add(_lblIdle_2)
		Me.ShapeContainer8.Shapes.Add(_Line1_6)
		Me._fraPanel_5.Controls.Add(_Label1_16)
		Me._fraPanel_5.Controls.Add(ShapeContainer8)
		Me._fraPanel_9.Controls.Add(chkLogDBActions)
		Me._fraPanel_9.Controls.Add(chkLogAllCommands)
		Me._fraPanel_9.Controls.Add(txtMaxLogSize)
		Me._fraPanel_9.Controls.Add(cboLogging)
		Me._fraPanel_9.Controls.Add(txtMaxBackLogSize)
		Me._fraPanel_9.Controls.Add(lbl9)
		Me.ShapeContainer9.Shapes.Add(Line8)
		Me._fraPanel_9.Controls.Add(_lbl5_5)
		Me._fraPanel_9.Controls.Add(_lbl5_4)
		Me._fraPanel_9.Controls.Add(_lbl5_6)
		Me._fraPanel_9.Controls.Add(_lbl5_7)
		Me._fraPanel_9.Controls.Add(lblBacklogSize)
		Me._fraPanel_9.Controls.Add(lblBacklog)
		Me._fraPanel_9.Controls.Add(_Label1_20)
		Me.ShapeContainer9.Shapes.Add(_Line1_10)
		Me._fraPanel_9.Controls.Add(ShapeContainer9)
		Me._fraPanel_8.Controls.Add(lblSplash)
		Me.ShapeContainer10.Shapes.Add(_Line1_0)
		Me._fraPanel_8.Controls.Add(_Label1_0)
		Me._fraPanel_8.Controls.Add(ShapeContainer10)
		Me.Label1.SetIndex(_Label1_10, CType(10, Short))
		Me.Label1.SetIndex(_Label1_9, CType(9, Short))
		Me.Label1.SetIndex(_Label1_6, CType(6, Short))
		Me.Label1.SetIndex(_Label1_4, CType(4, Short))
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.Label1.SetIndex(_Label1_2, CType(2, Short))
		Me.Label1.SetIndex(_Label1_3, CType(3, Short))
		Me.Label1.SetIndex(_Label1_5, CType(5, Short))
		Me.Label1.SetIndex(_Label1_7, CType(7, Short))
		Me.Label1.SetIndex(_Label1_12, CType(12, Short))
		Me.Label1.SetIndex(_Label1_14, CType(14, Short))
		Me.Label1.SetIndex(_Label1_13, CType(13, Short))
		Me.Label1.SetIndex(_Label1_15, CType(15, Short))
		Me.Label1.SetIndex(_Label1_17, CType(17, Short))
		Me.Label1.SetIndex(_Label1_19, CType(19, Short))
		Me.Label1.SetIndex(_Label1_8, CType(8, Short))
		Me.Label1.SetIndex(_Label1_18, CType(18, Short))
		Me.Label1.SetIndex(_Label1_11, CType(11, Short))
		Me.Label1.SetIndex(_Label1_16, CType(16, Short))
		Me.Label1.SetIndex(_Label1_20, CType(20, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.Label8.SetIndex(_Label8_8, CType(8, Short))
		Me.Label8.SetIndex(_Label8_5, CType(5, Short))
		Me.Label8.SetIndex(_Label8_0, CType(0, Short))
		Me.Label8.SetIndex(_Label8_2, CType(2, Short))
		Me.Label8.SetIndex(_Label8_7, CType(7, Short))
		Me.Label8.SetIndex(_Label8_1, CType(1, Short))
		Me.Label8.SetIndex(_Label8_3, CType(3, Short))
		Me.Label8.SetIndex(_Label8_4, CType(4, Short))
		Me.Label8.SetIndex(_Label8_6, CType(6, Short))
		Me.Label8.SetIndex(_Label8_10, CType(10, Short))
		Me.Line1.SetIndex(_Line1_1, CType(1, Short))
		Me.Line1.SetIndex(_Line1_3, CType(3, Short))
		Me.Line1.SetIndex(_Line1_4, CType(4, Short))
		Me.Line1.SetIndex(_Line1_5, CType(5, Short))
		Me.Line1.SetIndex(_Line1_7, CType(7, Short))
		Me.Line1.SetIndex(_Line1_9, CType(9, Short))
		Me.Line1.SetIndex(_Line1_8, CType(8, Short))
		Me.Line1.SetIndex(_Line1_2, CType(2, Short))
		Me.Line1.SetIndex(_Line1_6, CType(6, Short))
		Me.Line1.SetIndex(_Line1_10, CType(10, Short))
		Me.Line1.SetIndex(_Line1_0, CType(0, Short))
		Me.Line3.SetIndex(_Line3_2, CType(2, Short))
		Me.Line3.SetIndex(_Line3_1, CType(1, Short))
		Me.Line3.SetIndex(_Line3_0, CType(0, Short))
		Me.Line5.SetIndex(_Line5_4, CType(4, Short))
		Me.Line5.SetIndex(_Line5_3, CType(3, Short))
		Me.Line5.SetIndex(_Line5_2, CType(2, Short))
		Me.Line5.SetIndex(_Line5_1, CType(1, Short))
		Me.Line5.SetIndex(_Line5_0, CType(0, Short))
		Me.chkCBan.SetIndex(_chkCBan_5, CType(5, Short))
		Me.chkCBan.SetIndex(_chkCBan_3, CType(3, Short))
		Me.chkCBan.SetIndex(_chkCBan_0, CType(0, Short))
		Me.chkCBan.SetIndex(_chkCBan_1, CType(1, Short))
		Me.chkCBan.SetIndex(_chkCBan_2, CType(2, Short))
		Me.chkCBan.SetIndex(_chkCBan_6, CType(6, Short))
		Me.chkCBan.SetIndex(_chkCBan_4, CType(4, Short))
		Me.fraPanel.SetIndex(_fraPanel_0, CType(0, Short))
		Me.fraPanel.SetIndex(_fraPanel_2, CType(2, Short))
		Me.fraPanel.SetIndex(_fraPanel_3, CType(3, Short))
		Me.fraPanel.SetIndex(_fraPanel_4, CType(4, Short))
		Me.fraPanel.SetIndex(_fraPanel_6, CType(6, Short))
		Me.fraPanel.SetIndex(_fraPanel_7, CType(7, Short))
		Me.fraPanel.SetIndex(_fraPanel_1, CType(1, Short))
		Me.fraPanel.SetIndex(_fraPanel_5, CType(5, Short))
		Me.fraPanel.SetIndex(_fraPanel_9, CType(9, Short))
		Me.fraPanel.SetIndex(_fraPanel_8, CType(8, Short))
		Me.lbl5.SetIndex(_lbl5_13, CType(13, Short))
		Me.lbl5.SetIndex(_lbl5_1, CType(1, Short))
		Me.lbl5.SetIndex(_lbl5_12, CType(12, Short))
		Me.lbl5.SetIndex(_lbl5_11, CType(11, Short))
		Me.lbl5.SetIndex(_lbl5_10, CType(10, Short))
		Me.lbl5.SetIndex(_lbl5_9, CType(9, Short))
		Me.lbl5.SetIndex(_lbl5_8, CType(8, Short))
		Me.lbl5.SetIndex(_lbl5_0, CType(0, Short))
		Me.lbl5.SetIndex(_lbl5_2, CType(2, Short))
		Me.lbl5.SetIndex(_lbl5_3, CType(3, Short))
		Me.lbl5.SetIndex(_lbl5_5, CType(5, Short))
		Me.lbl5.SetIndex(_lbl5_4, CType(4, Short))
		Me.lbl5.SetIndex(_lbl5_6, CType(6, Short))
		Me.lbl5.SetIndex(_lbl5_7, CType(7, Short))
		Me.lblIdle.SetIndex(_lblIdle_3, CType(3, Short))
		Me.lblIdle.SetIndex(_lblIdle_0, CType(0, Short))
		Me.lblIdle.SetIndex(_lblIdle_1, CType(1, Short))
		Me.lblIdle.SetIndex(_lblIdle_2, CType(2, Short))
		Me.lblMod.SetIndex(_lblMod_7, CType(7, Short))
		Me.lblMod.SetIndex(_lblMod_0, CType(0, Short))
		Me.lblMod.SetIndex(_lblMod_2, CType(2, Short))
		Me.lblMod.SetIndex(_lblMod_3, CType(3, Short))
		Me.lblMod.SetIndex(_lblMod_5, CType(5, Short))
		Me.lblMod.SetIndex(_lblMod_1, CType(1, Short))
		Me.lblMod.SetIndex(_lblMod_4, CType(4, Short))
		Me.lblMod.SetIndex(_lblMod_6, CType(6, Short))
		Me.optNaming.SetIndex(_optNaming_3, CType(3, Short))
		Me.optNaming.SetIndex(_optNaming_2, CType(2, Short))
		Me.optNaming.SetIndex(_optNaming_1, CType(1, Short))
		Me.optNaming.SetIndex(_optNaming_0, CType(0, Short))
		CType(Me.optNaming, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblMod, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblIdle, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lbl5, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.fraPanel, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.chkCBan, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Line5, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Line3, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label8, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.tvw, System.ComponentModel.ISupportInitialize).EndInit()
		Me._fraPanel_0.ResumeLayout(False)
		Me._fraPanel_2.ResumeLayout(False)
		Me._fraPanel_3.ResumeLayout(False)
		Me._fraPanel_4.ResumeLayout(False)
		Me._fraPanel_6.ResumeLayout(False)
		Me._fraPanel_7.ResumeLayout(False)
		Me._fraPanel_1.ResumeLayout(False)
		Me._fraPanel_5.ResumeLayout(False)
		Me._fraPanel_9.ResumeLayout(False)
		Me._fraPanel_8.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class