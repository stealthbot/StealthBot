<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmChat
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
	Public WithEvents mnuConnect2 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDisconnect2 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSepT As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuUsers As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCommandManager As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuScriptManager As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSepTabcd As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuGetNews As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuUpdateVerbytes As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSepZ As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuRealmSwitch As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuIgnoreInvites As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSep1 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuHomeChannel As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuLastChannel As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuQCDash As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuQCHeader As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuCustomChannels_0 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuCustomChannels_1 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuCustomChannels_2 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuCustomChannels_3 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuCustomChannels_4 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuCustomChannels_5 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuCustomChannels_6 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuCustomChannels_7 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuCustomChannels_8 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCustomChannelAdd As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCustomChannelEdit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPCDash As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuPCHeader As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuPublicChannels_0 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuQCTop As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSep5 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuEditCaught As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuOpenBotFolder As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSepA As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuClearedTxt As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuWhisperCleared As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFiles As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuToolsMenuWarning As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuRepairDataFiles As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuRepairVerbytes As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuRepairCleanMail As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPacketLog As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSettingsRepair As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSep2 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuBot As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSetup As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuUTF8 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuLog0 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuLog1 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuLog2 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuLogging As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSep4 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuProfile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFilters As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuCatchPhrases As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSep6 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuReload As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSetTop As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuConnect As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDisconnect As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuToggle As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHideBans As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuLock As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuToggleFilters As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSep7 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuToggleWWUse As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSP3 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuToggleShowOutgoing As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHideWhispersInrtbChat As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSepC As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuClear As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuClearWW As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSepD As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuFlash As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuDisableVoidView As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuRecordWindowPos As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuWindow As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuTrayCaption As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuRestore As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuTraySep As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuTrayExit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuTray As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopWhisper As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopCopy As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopAddLeft As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopAddToFList As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopInvite As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopShitlist As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopSafelist As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopSep1 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuPopUserlistWhois As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopWhois As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopStats As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopProfile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopWebProfile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopSep2 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuPopKick As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopBan As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopSquelch As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopUnsquelch As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopDes As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopup As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuReloadScripts As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuOpenScriptFolder As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuScriptingDash_0 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuScripting As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuAbout As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelpReadme As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelpWebsite As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuTerms As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuChangeLog As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuItalic As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuBold As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuUnderline As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuListViewButton_0 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuListViewButton_1 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents _mnuListViewButton_2 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuShortcuts As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanWhisper As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanCopy As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanAddLeft As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanAddToFList As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanSep1 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuPopClanUserlistWhois As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanWhois As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanStatsWAR3 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanStatsW3XP As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanStats As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanProfile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanWebProfileWAR3 As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanWebProfileW3XP As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanWebProfile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanSep2 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuPopClanPromote As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanDemote As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanRemove As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanLeave As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanMakeChief As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanDisband As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopClanList As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopFLWhisper As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopFLCopy As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopFLAddLeft As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopFLInvite As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopFLSep1 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuPopFLUserlistWhois As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopFLWhois As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopFLStats As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopFLProfile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopFLWebProfile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopFLSep2 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuPopFLPromote As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopFLDemote As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopFLRemove As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopFLSep3 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuPopFLRefresh As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuPopFList As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents _lvChannel_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
	Public WithEvents _lvChannel_ColumnHeader_2 As System.Windows.Forms.ColumnHeader
	Public WithEvents _lvChannel_ColumnHeader_3 As System.Windows.Forms.ColumnHeader
	Public WithEvents lvChannel As System.Windows.Forms.ListView
	Public WithEvents tmrAccountLock As System.Windows.Forms.Timer
	Public WithEvents SControl As MSScriptControl.ScriptControl
	Public WithEvents SCRestricted As MSScriptControl.ScriptControl
	Public WithEvents _tmrScriptLong_0 As System.Windows.Forms.Timer
	Public WithEvents _itcScript_0 As AxInetCtlsObjects.AxInet
	Public WithEvents _sckScript_0 As AxMSWinsockLib.AxWinsock
	Public WithEvents ChatQueueTimer As System.Windows.Forms.Timer
	Public WithEvents cacheTimer As System.Windows.Forms.Timer
	Public WithEvents imlClan As System.Windows.Forms.ImageList
	Public WithEvents imlIcons As System.Windows.Forms.ImageList
	Public WithEvents _tmrScript_0 As System.Windows.Forms.Timer
	Public WithEvents _tmrSilentChannel_1 As System.Windows.Forms.Timer
	Public WithEvents _tmrSilentChannel_0 As System.Windows.Forms.Timer
	Public WithEvents cboSend As System.Windows.Forms.ComboBox
	Public WithEvents txtPost As System.Windows.Forms.TextBox
	Public WithEvents txtPre As System.Windows.Forms.TextBox
	Public WithEvents _ListviewTabs_TabPage0 As System.Windows.Forms.TabPage
	Public WithEvents _ListviewTabs_TabPage1 As System.Windows.Forms.TabPage
	Public WithEvents _ListviewTabs_TabPage2 As System.Windows.Forms.TabPage
	Public WithEvents ListviewTabs As System.Windows.Forms.TabControl
	Public WithEvents tmrIdleTimer As System.Windows.Forms.Timer
	Public WithEvents cmdShowHide As System.Windows.Forms.Button
	Public WithEvents sckMCP As AxMSWinsockLib.AxWinsock
	Public WithEvents scTimer As System.Windows.Forms.Timer
	Public WithEvents INet As AxInetCtlsObjects.AxInet
	Public WithEvents sckBNLS As AxMSWinsockLib.AxWinsock
	Public WithEvents sckBNet As AxMSWinsockLib.AxWinsock
	Public WithEvents UpTimer As System.Windows.Forms.Timer
	Public WithEvents _lvClanList_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
	Public WithEvents _lvClanList_ColumnHeader_2 As System.Windows.Forms.ColumnHeader
	Public WithEvents _lvClanList_ColumnHeader_3 As System.Windows.Forms.ColumnHeader
	Public WithEvents _lvClanList_ColumnHeader_4 As System.Windows.Forms.ColumnHeader
	Public WithEvents lvClanList As System.Windows.Forms.ListView
	Public WithEvents _lvFriendList_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
	Public WithEvents _lvFriendList_ColumnHeader_2 As System.Windows.Forms.ColumnHeader
	Public WithEvents lvFriendList As System.Windows.Forms.ListView
	Public WithEvents rtbWhispers As System.Windows.Forms.RichTextBox
	Public WithEvents rtbChat As System.Windows.Forms.RichTextBox
	Public WithEvents lblCurrentChannel As System.Windows.Forms.Label
	Public WithEvents itcScript As AxInetArray
	Public WithEvents mnuCustomChannels As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
	Public WithEvents mnuListViewButton As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
	Public WithEvents mnuPublicChannels As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
	Public WithEvents mnuScriptingDash As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
	Public WithEvents sckScript As AxWinsockArray
	Public WithEvents tmrScript As Microsoft.VisualBasic.Compatibility.VB6.TimerArray
	Public WithEvents tmrScriptLong As Microsoft.VisualBasic.Compatibility.VB6.TimerArray
	Public WithEvents tmrSilentChannel As Microsoft.VisualBasic.Compatibility.VB6.TimerArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmChat))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.mnuBot = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuConnect2 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuDisconnect2 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSepT = New System.Windows.Forms.ToolStripSeparator
		Me.mnuUsers = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCommandManager = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuScriptManager = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSepTabcd = New System.Windows.Forms.ToolStripSeparator
		Me.mnuGetNews = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuUpdateVerbytes = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSepZ = New System.Windows.Forms.ToolStripSeparator
		Me.mnuRealmSwitch = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuIgnoreInvites = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSep1 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuQCTop = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuHomeChannel = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuLastChannel = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuQCDash = New System.Windows.Forms.ToolStripSeparator
		Me.mnuQCHeader = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuCustomChannels_0 = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuCustomChannels_1 = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuCustomChannels_2 = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuCustomChannels_3 = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuCustomChannels_4 = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuCustomChannels_5 = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuCustomChannels_6 = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuCustomChannels_7 = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuCustomChannels_8 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCustomChannelAdd = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCustomChannelEdit = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPCDash = New System.Windows.Forms.ToolStripSeparator
		Me.mnuPCHeader = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuPublicChannels_0 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSep5 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuEditCaught = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuFiles = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuOpenBotFolder = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSepA = New System.Windows.Forms.ToolStripSeparator
		Me.mnuClearedTxt = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuWhisperCleared = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSettingsRepair = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuToolsMenuWarning = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuRepairDataFiles = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuRepairVerbytes = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuRepairCleanMail = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPacketLog = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSep2 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSetTop = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSetup = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuUTF8 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuLogging = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuLog0 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuLog1 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuLog2 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSep4 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuProfile = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuFilters = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCatchPhrases = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSep6 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuReload = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuConnect = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuDisconnect = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuWindow = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuToggle = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuHideBans = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuLock = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuToggleFilters = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSep7 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuToggleWWUse = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSP3 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuToggleShowOutgoing = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuHideWhispersInrtbChat = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSepC = New System.Windows.Forms.ToolStripSeparator
		Me.mnuClear = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuClearWW = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSepD = New System.Windows.Forms.ToolStripSeparator
		Me.mnuFlash = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuDisableVoidView = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuRecordWindowPos = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuTray = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuTrayCaption = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuRestore = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuTraySep = New System.Windows.Forms.ToolStripSeparator
		Me.mnuTrayExit = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopup = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopWhisper = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopCopy = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopAddLeft = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopAddToFList = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopInvite = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopShitlist = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopSafelist = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopSep1 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuPopUserlistWhois = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopWhois = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopStats = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopProfile = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopWebProfile = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopSep2 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuPopKick = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopBan = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopSquelch = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopUnsquelch = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopDes = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuScripting = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuReloadScripts = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuOpenScriptFolder = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuScriptingDash_0 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuAbout = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuHelpReadme = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuHelpWebsite = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuTerms = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuChangeLog = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuShortcuts = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuItalic = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuBold = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuUnderline = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuListViewButton_0 = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuListViewButton_1 = New System.Windows.Forms.ToolStripMenuItem
		Me._mnuListViewButton_2 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanList = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanWhisper = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanCopy = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanAddLeft = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanAddToFList = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanSep1 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuPopClanUserlistWhois = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanWhois = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanStats = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanStatsWAR3 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanStatsW3XP = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanProfile = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanWebProfile = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanWebProfileWAR3 = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanWebProfileW3XP = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanSep2 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuPopClanPromote = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanDemote = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanRemove = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanLeave = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanMakeChief = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopClanDisband = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopFList = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopFLWhisper = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopFLCopy = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopFLAddLeft = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopFLInvite = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopFLSep1 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuPopFLUserlistWhois = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopFLWhois = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopFLStats = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopFLProfile = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopFLWebProfile = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopFLSep2 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuPopFLPromote = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopFLDemote = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopFLRemove = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuPopFLSep3 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuPopFLRefresh = New System.Windows.Forms.ToolStripMenuItem
		Me.lvChannel = New System.Windows.Forms.ListView
		Me._lvChannel_ColumnHeader_1 = New System.Windows.Forms.ColumnHeader
		Me._lvChannel_ColumnHeader_2 = New System.Windows.Forms.ColumnHeader
		Me._lvChannel_ColumnHeader_3 = New System.Windows.Forms.ColumnHeader
		Me.tmrAccountLock = New System.Windows.Forms.Timer(components)
		Me.SControl = New MSScriptControl.ScriptControl
		Me.SCRestricted = New MSScriptControl.ScriptControl
		Me._tmrScriptLong_0 = New System.Windows.Forms.Timer(components)
		Me._itcScript_0 = New AxInetCtlsObjects.AxInet
		Me._sckScript_0 = New AxMSWinsockLib.AxWinsock
		Me.ChatQueueTimer = New System.Windows.Forms.Timer(components)
		Me.cacheTimer = New System.Windows.Forms.Timer(components)
		Me.imlClan = New System.Windows.Forms.ImageList
		Me.imlIcons = New System.Windows.Forms.ImageList
		Me._tmrScript_0 = New System.Windows.Forms.Timer(components)
		Me._tmrSilentChannel_1 = New System.Windows.Forms.Timer(components)
		Me._tmrSilentChannel_0 = New System.Windows.Forms.Timer(components)
		Me.cboSend = New System.Windows.Forms.ComboBox
		Me.txtPost = New System.Windows.Forms.TextBox
		Me.txtPre = New System.Windows.Forms.TextBox
		Me.ListviewTabs = New System.Windows.Forms.TabControl
		Me._ListviewTabs_TabPage0 = New System.Windows.Forms.TabPage
		Me._ListviewTabs_TabPage1 = New System.Windows.Forms.TabPage
		Me._ListviewTabs_TabPage2 = New System.Windows.Forms.TabPage
		Me.tmrIdleTimer = New System.Windows.Forms.Timer(components)
		Me.cmdShowHide = New System.Windows.Forms.Button
		Me.sckMCP = New AxMSWinsockLib.AxWinsock
		Me.scTimer = New System.Windows.Forms.Timer(components)
		Me.INet = New AxInetCtlsObjects.AxInet
		Me.sckBNLS = New AxMSWinsockLib.AxWinsock
		Me.sckBNet = New AxMSWinsockLib.AxWinsock
		Me.UpTimer = New System.Windows.Forms.Timer(components)
		Me.lvClanList = New System.Windows.Forms.ListView
		Me._lvClanList_ColumnHeader_1 = New System.Windows.Forms.ColumnHeader
		Me._lvClanList_ColumnHeader_2 = New System.Windows.Forms.ColumnHeader
		Me._lvClanList_ColumnHeader_3 = New System.Windows.Forms.ColumnHeader
		Me._lvClanList_ColumnHeader_4 = New System.Windows.Forms.ColumnHeader
		Me.lvFriendList = New System.Windows.Forms.ListView
		Me._lvFriendList_ColumnHeader_1 = New System.Windows.Forms.ColumnHeader
		Me._lvFriendList_ColumnHeader_2 = New System.Windows.Forms.ColumnHeader
		Me.rtbWhispers = New System.Windows.Forms.RichTextBox
		Me.rtbChat = New System.Windows.Forms.RichTextBox
		Me.lblCurrentChannel = New System.Windows.Forms.Label
		Me.itcScript = New AxInetArray(components)
		Me.mnuCustomChannels = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(components)
		Me.mnuListViewButton = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(components)
		Me.mnuPublicChannels = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(components)
		Me.mnuScriptingDash = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(components)
		Me.sckScript = New AxWinsockArray(components)
		Me.tmrScript = New Microsoft.VisualBasic.Compatibility.VB6.TimerArray(components)
		Me.tmrScriptLong = New Microsoft.VisualBasic.Compatibility.VB6.TimerArray(components)
		Me.tmrSilentChannel = New Microsoft.VisualBasic.Compatibility.VB6.TimerArray(components)
		Me.MainMenu1.SuspendLayout()
		Me.lvChannel.SuspendLayout()
		Me.ListviewTabs.SuspendLayout()
		Me.lvClanList.SuspendLayout()
		Me.lvFriendList.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.SControl, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SCRestricted, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._itcScript_0, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._sckScript_0, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.sckMCP, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.INet, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.sckBNLS, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.sckBNet, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.itcScript, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mnuCustomChannels, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mnuListViewButton, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mnuPublicChannels, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mnuScriptingDash, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.sckScript, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.tmrScript, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.tmrScriptLong, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.tmrSilentChannel, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.Color.Black
		Me.Text = ":: StealthBot &version :: Disconnected ::"
		Me.ClientSize = New System.Drawing.Size(760, 555)
		Me.Location = New System.Drawing.Point(15, 58)
		Me.ForeColor = System.Drawing.Color.Black
		Me.Icon = CType(resources.GetObject("frmChat.Icon"), System.Drawing.Icon)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Name = "frmChat"
		Me.mnuBot.Name = "mnuBot"
		Me.mnuBot.Text = "&Bot"
		Me.mnuBot.Checked = False
		Me.mnuBot.Enabled = True
		Me.mnuBot.Visible = True
		Me.mnuConnect2.Name = "mnuConnect2"
		Me.mnuConnect2.Text = "&Connect"
		Me.mnuConnect2.Checked = False
		Me.mnuConnect2.Enabled = True
		Me.mnuConnect2.Visible = True
		Me.mnuDisconnect2.Name = "mnuDisconnect2"
		Me.mnuDisconnect2.Text = "&Disconnect"
		Me.mnuDisconnect2.Checked = False
		Me.mnuDisconnect2.Enabled = True
		Me.mnuDisconnect2.Visible = True
		Me.mnuSepT.Enabled = True
		Me.mnuSepT.Visible = True
		Me.mnuSepT.Name = "mnuSepT"
		Me.mnuUsers.Name = "mnuUsers"
		Me.mnuUsers.Text = "&User Database Manager..."
		Me.mnuUsers.Checked = False
		Me.mnuUsers.Enabled = True
		Me.mnuUsers.Visible = True
		Me.mnuCommandManager.Name = "mnuCommandManager"
		Me.mnuCommandManager.Text = "Command &Manager..."
		Me.mnuCommandManager.Checked = False
		Me.mnuCommandManager.Enabled = True
		Me.mnuCommandManager.Visible = True
		Me.mnuScriptManager.Name = "mnuScriptManager"
		Me.mnuScriptManager.Text = "&Script Settings Manager..."
		Me.mnuScriptManager.Visible = False
		Me.mnuScriptManager.Checked = False
		Me.mnuScriptManager.Enabled = True
		Me.mnuSepTabcd.Enabled = True
		Me.mnuSepTabcd.Visible = True
		Me.mnuSepTabcd.Name = "mnuSepTabcd"
		Me.mnuGetNews.Name = "mnuGetNews"
		Me.mnuGetNews.Text = "Get &News and Check for Updates"
		Me.mnuGetNews.Checked = False
		Me.mnuGetNews.Enabled = True
		Me.mnuGetNews.Visible = True
		Me.mnuUpdateVerbytes.Name = "mnuUpdateVerbytes"
		Me.mnuUpdateVerbytes.Text = "Update &Version Bytes"
		Me.mnuUpdateVerbytes.Checked = False
		Me.mnuUpdateVerbytes.Enabled = True
		Me.mnuUpdateVerbytes.Visible = True
		Me.mnuSepZ.Visible = False
		Me.mnuSepZ.Enabled = True
		Me.mnuSepZ.Name = "mnuSepZ"
		Me.mnuRealmSwitch.Name = "mnuRealmSwitch"
		Me.mnuRealmSwitch.Text = "Switch &Realm Character..."
		Me.mnuRealmSwitch.Visible = False
		Me.mnuRealmSwitch.Checked = False
		Me.mnuRealmSwitch.Enabled = True
		Me.mnuIgnoreInvites.Name = "mnuIgnoreInvites"
		Me.mnuIgnoreInvites.Text = "&Ignore Clan Invitations"
		Me.mnuIgnoreInvites.Visible = False
		Me.mnuIgnoreInvites.Checked = False
		Me.mnuIgnoreInvites.Enabled = True
		Me.mnuSep1.Enabled = True
		Me.mnuSep1.Visible = True
		Me.mnuSep1.Name = "mnuSep1"
		Me.mnuQCTop.Name = "mnuQCTop"
		Me.mnuQCTop.Text = "Channel &List"
		Me.mnuQCTop.Checked = False
		Me.mnuQCTop.Enabled = True
		Me.mnuQCTop.Visible = True
		Me.mnuHomeChannel.Name = "mnuHomeChannel"
		Me.mnuHomeChannel.Text = "&Bot Home"
		Me.mnuHomeChannel.Visible = False
		Me.mnuHomeChannel.Checked = False
		Me.mnuHomeChannel.Enabled = True
		Me.mnuLastChannel.Name = "mnuLastChannel"
		Me.mnuLastChannel.Text = "&Last Channel"
		Me.mnuLastChannel.Visible = False
		Me.mnuLastChannel.Checked = False
		Me.mnuLastChannel.Enabled = True
		Me.mnuQCDash.Visible = False
		Me.mnuQCDash.Enabled = True
		Me.mnuQCDash.Name = "mnuQCDash"
		Me.mnuQCHeader.Name = "mnuQCHeader"
		Me.mnuQCHeader.Text = "- QuickChannels -"
		Me.mnuQCHeader.Enabled = False
		Me.mnuQCHeader.Checked = False
		Me.mnuQCHeader.Visible = True
		Me._mnuCustomChannels_0.Name = "_mnuCustomChannels_0"
		Me._mnuCustomChannels_0.Text = ""
		Me._mnuCustomChannels_0.ShortcutKeys = CType(System.Windows.Forms.Keys.F1, System.Windows.Forms.Keys)
		Me._mnuCustomChannels_0.Checked = False
		Me._mnuCustomChannels_0.Enabled = True
		Me._mnuCustomChannels_0.Visible = True
		Me._mnuCustomChannels_1.Name = "_mnuCustomChannels_1"
		Me._mnuCustomChannels_1.Text = ""
		Me._mnuCustomChannels_1.ShortcutKeys = CType(System.Windows.Forms.Keys.F2, System.Windows.Forms.Keys)
		Me._mnuCustomChannels_1.Checked = False
		Me._mnuCustomChannels_1.Enabled = True
		Me._mnuCustomChannels_1.Visible = True
		Me._mnuCustomChannels_2.Name = "_mnuCustomChannels_2"
		Me._mnuCustomChannels_2.Text = ""
		Me._mnuCustomChannels_2.ShortcutKeys = CType(System.Windows.Forms.Keys.F3, System.Windows.Forms.Keys)
		Me._mnuCustomChannels_2.Checked = False
		Me._mnuCustomChannels_2.Enabled = True
		Me._mnuCustomChannels_2.Visible = True
		Me._mnuCustomChannels_3.Name = "_mnuCustomChannels_3"
		Me._mnuCustomChannels_3.Text = ""
		Me._mnuCustomChannels_3.ShortcutKeys = CType(System.Windows.Forms.Keys.F4, System.Windows.Forms.Keys)
		Me._mnuCustomChannels_3.Checked = False
		Me._mnuCustomChannels_3.Enabled = True
		Me._mnuCustomChannels_3.Visible = True
		Me._mnuCustomChannels_4.Name = "_mnuCustomChannels_4"
		Me._mnuCustomChannels_4.Text = ""
		Me._mnuCustomChannels_4.ShortcutKeys = CType(System.Windows.Forms.Keys.F5, System.Windows.Forms.Keys)
		Me._mnuCustomChannels_4.Checked = False
		Me._mnuCustomChannels_4.Enabled = True
		Me._mnuCustomChannels_4.Visible = True
		Me._mnuCustomChannels_5.Name = "_mnuCustomChannels_5"
		Me._mnuCustomChannels_5.Text = ""
		Me._mnuCustomChannels_5.ShortcutKeys = CType(System.Windows.Forms.Keys.F6, System.Windows.Forms.Keys)
		Me._mnuCustomChannels_5.Checked = False
		Me._mnuCustomChannels_5.Enabled = True
		Me._mnuCustomChannels_5.Visible = True
		Me._mnuCustomChannels_6.Name = "_mnuCustomChannels_6"
		Me._mnuCustomChannels_6.Text = ""
		Me._mnuCustomChannels_6.ShortcutKeys = CType(System.Windows.Forms.Keys.F7, System.Windows.Forms.Keys)
		Me._mnuCustomChannels_6.Checked = False
		Me._mnuCustomChannels_6.Enabled = True
		Me._mnuCustomChannels_6.Visible = True
		Me._mnuCustomChannels_7.Name = "_mnuCustomChannels_7"
		Me._mnuCustomChannels_7.Text = ""
		Me._mnuCustomChannels_7.ShortcutKeys = CType(System.Windows.Forms.Keys.F8, System.Windows.Forms.Keys)
		Me._mnuCustomChannels_7.Checked = False
		Me._mnuCustomChannels_7.Enabled = True
		Me._mnuCustomChannels_7.Visible = True
		Me._mnuCustomChannels_8.Name = "_mnuCustomChannels_8"
		Me._mnuCustomChannels_8.Text = ""
		Me._mnuCustomChannels_8.ShortcutKeys = CType(System.Windows.Forms.Keys.F9, System.Windows.Forms.Keys)
		Me._mnuCustomChannels_8.Checked = False
		Me._mnuCustomChannels_8.Enabled = True
		Me._mnuCustomChannels_8.Visible = True
		Me.mnuCustomChannelAdd.Name = "mnuCustomChannelAdd"
		Me.mnuCustomChannelAdd.Text = "&Add QuickChannel"
		Me.mnuCustomChannelAdd.Visible = False
		Me.mnuCustomChannelAdd.Checked = False
		Me.mnuCustomChannelAdd.Enabled = True
		Me.mnuCustomChannelEdit.Name = "mnuCustomChannelEdit"
		Me.mnuCustomChannelEdit.Text = "&Edit QuickChannels..."
		Me.mnuCustomChannelEdit.Checked = False
		Me.mnuCustomChannelEdit.Enabled = True
		Me.mnuCustomChannelEdit.Visible = True
		Me.mnuPCDash.Enabled = True
		Me.mnuPCDash.Visible = True
		Me.mnuPCDash.Name = "mnuPCDash"
		Me.mnuPCHeader.Name = "mnuPCHeader"
		Me.mnuPCHeader.Text = "- Public Channels -"
		Me.mnuPCHeader.Enabled = False
		Me.mnuPCHeader.Checked = False
		Me.mnuPCHeader.Visible = True
		Me._mnuPublicChannels_0.Name = "_mnuPublicChannels_0"
		Me._mnuPublicChannels_0.Text = ""
		Me._mnuPublicChannels_0.Checked = False
		Me._mnuPublicChannels_0.Enabled = True
		Me._mnuPublicChannels_0.Visible = True
		Me.mnuSep5.Enabled = True
		Me.mnuSep5.Visible = True
		Me.mnuSep5.Name = "mnuSep5"
		Me.mnuEditCaught.Name = "mnuEditCaught"
		Me.mnuEditCaught.Text = "View Caught P&hrases..."
		Me.mnuEditCaught.Checked = False
		Me.mnuEditCaught.Enabled = True
		Me.mnuEditCaught.Visible = True
		Me.mnuFiles.Name = "mnuFiles"
		Me.mnuFiles.Text = "View &Files"
		Me.mnuFiles.Checked = False
		Me.mnuFiles.Enabled = True
		Me.mnuFiles.Visible = True
		Me.mnuOpenBotFolder.Name = "mnuOpenBotFolder"
		Me.mnuOpenBotFolder.Text = "Open Bot &Folder"
		Me.mnuOpenBotFolder.Checked = False
		Me.mnuOpenBotFolder.Enabled = True
		Me.mnuOpenBotFolder.Visible = True
		Me.mnuSepA.Enabled = True
		Me.mnuSepA.Visible = True
		Me.mnuSepA.Name = "mnuSepA"
		Me.mnuClearedTxt.Name = "mnuClearedTxt"
		Me.mnuClearedTxt.Text = "Current Text &Log"
		Me.mnuClearedTxt.Checked = False
		Me.mnuClearedTxt.Enabled = True
		Me.mnuClearedTxt.Visible = True
		Me.mnuWhisperCleared.Name = "mnuWhisperCleared"
		Me.mnuWhisperCleared.Text = "&Whisper Window Text Log"
		Me.mnuWhisperCleared.Checked = False
		Me.mnuWhisperCleared.Enabled = True
		Me.mnuWhisperCleared.Visible = True
		Me.mnuSettingsRepair.Name = "mnuSettingsRepair"
		Me.mnuSettingsRepair.Text = "&Tools"
		Me.mnuSettingsRepair.Checked = False
		Me.mnuSettingsRepair.Enabled = True
		Me.mnuSettingsRepair.Visible = True
		Me.mnuToolsMenuWarning.Name = "mnuToolsMenuWarning"
		Me.mnuToolsMenuWarning.Text = "- Use Carefully -"
		Me.mnuToolsMenuWarning.Enabled = False
		Me.mnuToolsMenuWarning.Checked = False
		Me.mnuToolsMenuWarning.Visible = True
		Me.mnuRepairDataFiles.Name = "mnuRepairDataFiles"
		Me.mnuRepairDataFiles.Text = "Delete &Data Files"
		Me.mnuRepairDataFiles.Checked = False
		Me.mnuRepairDataFiles.Enabled = True
		Me.mnuRepairDataFiles.Visible = True
		Me.mnuRepairVerbytes.Name = "mnuRepairVerbytes"
		Me.mnuRepairVerbytes.Text = "Restore Default &Version Bytes"
		Me.mnuRepairVerbytes.Checked = False
		Me.mnuRepairVerbytes.Enabled = True
		Me.mnuRepairVerbytes.Visible = True
		Me.mnuRepairCleanMail.Name = "mnuRepairCleanMail"
		Me.mnuRepairCleanMail.Text = "Clean Up &Mail Database"
		Me.mnuRepairCleanMail.Checked = False
		Me.mnuRepairCleanMail.Enabled = True
		Me.mnuRepairCleanMail.Visible = True
		Me.mnuPacketLog.Name = "mnuPacketLog"
		Me.mnuPacketLog.Text = "Log StealthBot &Packet Traffic"
		Me.mnuPacketLog.Checked = False
		Me.mnuPacketLog.Enabled = True
		Me.mnuPacketLog.Visible = True
		Me.mnuSep2.Enabled = True
		Me.mnuSep2.Visible = True
		Me.mnuSep2.Name = "mnuSep2"
		Me.mnuExit.Name = "mnuExit"
		Me.mnuExit.Text = "E&xit"
		Me.mnuExit.Checked = False
		Me.mnuExit.Enabled = True
		Me.mnuExit.Visible = True
		Me.mnuSetTop.Name = "mnuSetTop"
		Me.mnuSetTop.Text = "&Settings"
		Me.mnuSetTop.Checked = False
		Me.mnuSetTop.Enabled = True
		Me.mnuSetTop.Visible = True
		Me.mnuSetup.Name = "mnuSetup"
		Me.mnuSetup.Text = "&Bot Settings..."
		Me.mnuSetup.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.P, System.Windows.Forms.Keys)
		Me.mnuSetup.Checked = False
		Me.mnuSetup.Enabled = True
		Me.mnuSetup.Visible = True
		Me.mnuUTF8.Name = "mnuUTF8"
		Me.mnuUTF8.Text = "Use &UTF-8 in Chat"
		Me.mnuUTF8.Checked = False
		Me.mnuUTF8.Enabled = True
		Me.mnuUTF8.Visible = True
		Me.mnuLogging.Name = "mnuLogging"
		Me.mnuLogging.Text = "&Logging Settings"
		Me.mnuLogging.Checked = False
		Me.mnuLogging.Enabled = True
		Me.mnuLogging.Visible = True
		Me.mnuLog0.Name = "mnuLog0"
		Me.mnuLog0.Text = "Full Text Logging"
		Me.mnuLog0.Checked = False
		Me.mnuLog0.Enabled = True
		Me.mnuLog0.Visible = True
		Me.mnuLog1.Name = "mnuLog1"
		Me.mnuLog1.Text = "Temporary Logging"
		Me.mnuLog1.Checked = False
		Me.mnuLog1.Enabled = True
		Me.mnuLog1.Visible = True
		Me.mnuLog2.Name = "mnuLog2"
		Me.mnuLog2.Text = "No Logging"
		Me.mnuLog2.Checked = False
		Me.mnuLog2.Enabled = True
		Me.mnuLog2.Visible = True
		Me.mnuSep4.Enabled = True
		Me.mnuSep4.Visible = True
		Me.mnuSep4.Name = "mnuSep4"
		Me.mnuProfile.Name = "mnuProfile"
		Me.mnuProfile.Text = "Edit &Profile..."
		Me.mnuProfile.Checked = False
		Me.mnuProfile.Enabled = True
		Me.mnuProfile.Visible = True
		Me.mnuFilters.Name = "mnuFilters"
		Me.mnuFilters.Text = "&Edit Chat Filters..."
		Me.mnuFilters.Checked = False
		Me.mnuFilters.Enabled = True
		Me.mnuFilters.Visible = True
		Me.mnuCatchPhrases.Name = "mnuCatchPhrases"
		Me.mnuCatchPhrases.Text = "Edit &Catch Phrases..."
		Me.mnuCatchPhrases.Checked = False
		Me.mnuCatchPhrases.Enabled = True
		Me.mnuCatchPhrases.Visible = True
		Me.mnuSep6.Enabled = True
		Me.mnuSep6.Visible = True
		Me.mnuSep6.Name = "mnuSep6"
		Me.mnuReload.Name = "mnuReload"
		Me.mnuReload.Text = "&Reload Config"
		Me.mnuReload.Checked = False
		Me.mnuReload.Enabled = True
		Me.mnuReload.Visible = True
		Me.mnuConnect.Name = "mnuConnect"
		Me.mnuConnect.Text = "&Connect"
		Me.mnuConnect.Checked = False
		Me.mnuConnect.Enabled = True
		Me.mnuConnect.Visible = True
		Me.mnuDisconnect.Name = "mnuDisconnect"
		Me.mnuDisconnect.Text = "&Disconnect"
		Me.mnuDisconnect.Checked = False
		Me.mnuDisconnect.Enabled = True
		Me.mnuDisconnect.Visible = True
		Me.mnuWindow.Name = "mnuWindow"
		Me.mnuWindow.Text = "&Window"
		Me.mnuWindow.Checked = False
		Me.mnuWindow.Enabled = True
		Me.mnuWindow.Visible = True
		Me.mnuToggle.Name = "mnuToggle"
		Me.mnuToggle.Text = "Hide &Join/Leave Messages"
		Me.mnuToggle.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.J, System.Windows.Forms.Keys)
		Me.mnuToggle.Checked = False
		Me.mnuToggle.Enabled = True
		Me.mnuToggle.Visible = True
		Me.mnuHideBans.Name = "mnuHideBans"
		Me.mnuHideBans.Text = "Hide &Ban Messages"
		Me.mnuHideBans.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.H, System.Windows.Forms.Keys)
		Me.mnuHideBans.Checked = False
		Me.mnuHideBans.Enabled = True
		Me.mnuHideBans.Visible = True
		Me.mnuLock.Name = "mnuLock"
		Me.mnuLock.Text = "Loc&k Chat Window"
		Me.mnuLock.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.L, System.Windows.Forms.Keys)
		Me.mnuLock.Checked = False
		Me.mnuLock.Enabled = True
		Me.mnuLock.Visible = True
		Me.mnuToggleFilters.Name = "mnuToggleFilters"
		Me.mnuToggleFilters.Text = "Use Chat &Filtering"
		Me.mnuToggleFilters.Checked = True
		Me.mnuToggleFilters.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.F, System.Windows.Forms.Keys)
		Me.mnuToggleFilters.Enabled = True
		Me.mnuToggleFilters.Visible = True
		Me.mnuSep7.Enabled = True
		Me.mnuSep7.Visible = True
		Me.mnuSep7.Name = "mnuSep7"
		Me.mnuToggleWWUse.Name = "mnuToggleWWUse"
		Me.mnuToggleWWUse.Text = "Use Individual &Whisper Windows"
		Me.mnuToggleWWUse.Checked = False
		Me.mnuToggleWWUse.Enabled = True
		Me.mnuToggleWWUse.Visible = True
		Me.mnuSP3.Enabled = True
		Me.mnuSP3.Visible = True
		Me.mnuSP3.Name = "mnuSP3"
		Me.mnuToggleShowOutgoing.Name = "mnuToggleShowOutgoing"
		Me.mnuToggleShowOutgoing.Text = "Show &Outgoing Whispers in Whisper Box"
		Me.mnuToggleShowOutgoing.Checked = False
		Me.mnuToggleShowOutgoing.Enabled = True
		Me.mnuToggleShowOutgoing.Visible = True
		Me.mnuHideWhispersInrtbChat.Name = "mnuHideWhispersInrtbChat"
		Me.mnuHideWhispersInrtbChat.Text = "&Hide Whispers in Main Window"
		Me.mnuHideWhispersInrtbChat.Checked = False
		Me.mnuHideWhispersInrtbChat.Enabled = True
		Me.mnuHideWhispersInrtbChat.Visible = True
		Me.mnuSepC.Enabled = True
		Me.mnuSepC.Visible = True
		Me.mnuSepC.Name = "mnuSepC"
		Me.mnuClear.Name = "mnuClear"
		Me.mnuClear.Text = "&Clear Chat Window"
		Me.mnuClear.ShortcutKeys = CType(System.Windows.Forms.Keys.Shift or System.Windows.Forms.Keys.Delete, System.Windows.Forms.Keys)
		Me.mnuClear.Checked = False
		Me.mnuClear.Enabled = True
		Me.mnuClear.Visible = True
		Me.mnuClearWW.Name = "mnuClearWW"
		Me.mnuClearWW.Text = "Cl&ear Whisper Window"
		Me.mnuClearWW.Checked = False
		Me.mnuClearWW.Enabled = True
		Me.mnuClearWW.Visible = True
		Me.mnuSepD.Enabled = True
		Me.mnuSepD.Visible = True
		Me.mnuSepD.Name = "mnuSepD"
		Me.mnuFlash.Name = "mnuFlash"
		Me.mnuFlash.Text = "Fl&ash Window on Events"
		Me.mnuFlash.Checked = False
		Me.mnuFlash.Enabled = True
		Me.mnuFlash.Visible = True
		Me.mnuDisableVoidView.Name = "mnuDisableVoidView"
		Me.mnuDisableVoidView.Text = "Disable &Silent Channel View"
		Me.mnuDisableVoidView.Checked = False
		Me.mnuDisableVoidView.Enabled = True
		Me.mnuDisableVoidView.Visible = True
		Me.mnuRecordWindowPos.Name = "mnuRecordWindowPos"
		Me.mnuRecordWindowPos.Text = "&Record Current Position"
		Me.mnuRecordWindowPos.Checked = False
		Me.mnuRecordWindowPos.Enabled = True
		Me.mnuRecordWindowPos.Visible = True
		Me.mnuTray.Name = "mnuTray"
		Me.mnuTray.Text = "systray"
		Me.mnuTray.Visible = False
		Me.mnuTray.Checked = False
		Me.mnuTray.Enabled = True
		Me.mnuTrayCaption.Name = "mnuTrayCaption"
		Me.mnuTrayCaption.Text = ""
		Me.mnuTrayCaption.Checked = False
		Me.mnuTrayCaption.Enabled = True
		Me.mnuTrayCaption.Visible = True
		Me.mnuRestore.Name = "mnuRestore"
		Me.mnuRestore.Text = "&Restore"
		Me.mnuRestore.Checked = False
		Me.mnuRestore.Enabled = True
		Me.mnuRestore.Visible = True
		Me.mnuTraySep.Enabled = True
		Me.mnuTraySep.Visible = True
		Me.mnuTraySep.Name = "mnuTraySep"
		Me.mnuTrayExit.Name = "mnuTrayExit"
		Me.mnuTrayExit.Text = "E&xit"
		Me.mnuTrayExit.Checked = False
		Me.mnuTrayExit.Enabled = True
		Me.mnuTrayExit.Visible = True
		Me.mnuPopup.Name = "mnuPopup"
		Me.mnuPopup.Text = "popupmenu"
		Me.mnuPopup.Visible = False
		Me.mnuPopup.Checked = False
		Me.mnuPopup.Enabled = True
		Me.mnuPopWhisper.Name = "mnuPopWhisper"
		Me.mnuPopWhisper.Text = "W&hisper"
		Me.mnuPopWhisper.Checked = False
		Me.mnuPopWhisper.Enabled = True
		Me.mnuPopWhisper.Visible = True
		Me.mnuPopCopy.Name = "mnuPopCopy"
		Me.mnuPopCopy.Text = "&Copy Name to Clipboard"
		Me.mnuPopCopy.Checked = False
		Me.mnuPopCopy.Enabled = True
		Me.mnuPopCopy.Visible = True
		Me.mnuPopAddLeft.Name = "mnuPopAddLeft"
		Me.mnuPopAddLeft.Text = "Add to &Left Send Box"
		Me.mnuPopAddLeft.Checked = False
		Me.mnuPopAddLeft.Enabled = True
		Me.mnuPopAddLeft.Visible = True
		Me.mnuPopAddToFList.Name = "mnuPopAddToFList"
		Me.mnuPopAddToFList.Text = "Add to &Friends List"
		Me.mnuPopAddToFList.Checked = False
		Me.mnuPopAddToFList.Enabled = True
		Me.mnuPopAddToFList.Visible = True
		Me.mnuPopInvite.Name = "mnuPopInvite"
		Me.mnuPopInvite.Text = "&Invite to Warcraft III Clan"
		Me.mnuPopInvite.Checked = False
		Me.mnuPopInvite.Enabled = True
		Me.mnuPopInvite.Visible = True
		Me.mnuPopShitlist.Name = "mnuPopShitlist"
		Me.mnuPopShitlist.Text = "Shi&tlist"
		Me.mnuPopShitlist.Checked = False
		Me.mnuPopShitlist.Enabled = True
		Me.mnuPopShitlist.Visible = True
		Me.mnuPopSafelist.Name = "mnuPopSafelist"
		Me.mnuPopSafelist.Text = "S&afelist"
		Me.mnuPopSafelist.Checked = False
		Me.mnuPopSafelist.Enabled = True
		Me.mnuPopSafelist.Visible = True
		Me.mnuPopSep1.Enabled = True
		Me.mnuPopSep1.Visible = True
		Me.mnuPopSep1.Name = "mnuPopSep1"
		Me.mnuPopUserlistWhois.Name = "mnuPopUserlistWhois"
		Me.mnuPopUserlistWhois.Text = "&Userlist Whois"
		Me.mnuPopUserlistWhois.Checked = False
		Me.mnuPopUserlistWhois.Enabled = True
		Me.mnuPopUserlistWhois.Visible = True
		Me.mnuPopWhois.Name = "mnuPopWhois"
		Me.mnuPopWhois.Text = "Battle.net &Whois"
		Me.mnuPopWhois.Checked = False
		Me.mnuPopWhois.Enabled = True
		Me.mnuPopWhois.Visible = True
		Me.mnuPopStats.Name = "mnuPopStats"
		Me.mnuPopStats.Text = "Battle.net &Stats"
		Me.mnuPopStats.Checked = False
		Me.mnuPopStats.Enabled = True
		Me.mnuPopStats.Visible = True
		Me.mnuPopProfile.Name = "mnuPopProfile"
		Me.mnuPopProfile.Text = "Battle.net &Profile"
		Me.mnuPopProfile.Checked = False
		Me.mnuPopProfile.Enabled = True
		Me.mnuPopProfile.Visible = True
		Me.mnuPopWebProfile.Name = "mnuPopWebProfile"
		Me.mnuPopWebProfile.Text = "W&eb Profile"
		Me.mnuPopWebProfile.Checked = False
		Me.mnuPopWebProfile.Enabled = True
		Me.mnuPopWebProfile.Visible = True
		Me.mnuPopSep2.Enabled = True
		Me.mnuPopSep2.Visible = True
		Me.mnuPopSep2.Name = "mnuPopSep2"
		Me.mnuPopKick.Name = "mnuPopKick"
		Me.mnuPopKick.Text = "&Kick"
		Me.mnuPopKick.Checked = False
		Me.mnuPopKick.Enabled = True
		Me.mnuPopKick.Visible = True
		Me.mnuPopBan.Name = "mnuPopBan"
		Me.mnuPopBan.Text = "&Ban"
		Me.mnuPopBan.Checked = False
		Me.mnuPopBan.Enabled = True
		Me.mnuPopBan.Visible = True
		Me.mnuPopSquelch.Name = "mnuPopSquelch"
		Me.mnuPopSquelch.Text = "S&quelch"
		Me.mnuPopSquelch.Checked = False
		Me.mnuPopSquelch.Enabled = True
		Me.mnuPopSquelch.Visible = True
		Me.mnuPopUnsquelch.Name = "mnuPopUnsquelch"
		Me.mnuPopUnsquelch.Text = "U&nsquelch"
		Me.mnuPopUnsquelch.Checked = False
		Me.mnuPopUnsquelch.Enabled = True
		Me.mnuPopUnsquelch.Visible = True
		Me.mnuPopDes.Name = "mnuPopDes"
		Me.mnuPopDes.Text = "&Designate"
		Me.mnuPopDes.Checked = False
		Me.mnuPopDes.Enabled = True
		Me.mnuPopDes.Visible = True
		Me.mnuScripting.Name = "mnuScripting"
		Me.mnuScripting.Text = "Sc&ripting"
		Me.mnuScripting.Checked = False
		Me.mnuScripting.Enabled = True
		Me.mnuScripting.Visible = True
		Me.mnuReloadScripts.Name = "mnuReloadScripts"
		Me.mnuReloadScripts.Text = "Reload Scripts"
		Me.mnuReloadScripts.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.R, System.Windows.Forms.Keys)
		Me.mnuReloadScripts.Checked = False
		Me.mnuReloadScripts.Enabled = True
		Me.mnuReloadScripts.Visible = True
		Me.mnuOpenScriptFolder.Name = "mnuOpenScriptFolder"
		Me.mnuOpenScriptFolder.Text = "Open Script Folder"
		Me.mnuOpenScriptFolder.Checked = False
		Me.mnuOpenScriptFolder.Enabled = True
		Me.mnuOpenScriptFolder.Visible = True
		Me._mnuScriptingDash_0.Visible = False
		Me._mnuScriptingDash_0.Enabled = True
		Me._mnuScriptingDash_0.Name = "_mnuScriptingDash_0"
		Me.mnuHelp.Name = "mnuHelp"
		Me.mnuHelp.Text = "&Help"
		Me.mnuHelp.Checked = False
		Me.mnuHelp.Enabled = True
		Me.mnuHelp.Visible = True
		Me.mnuAbout.Name = "mnuAbout"
		Me.mnuAbout.Text = "&About..."
		Me.mnuAbout.Checked = False
		Me.mnuAbout.Enabled = True
		Me.mnuAbout.Visible = True
		Me.mnuHelpReadme.Name = "mnuHelpReadme"
		Me.mnuHelpReadme.Text = "&Wiki"
		Me.mnuHelpReadme.Checked = False
		Me.mnuHelpReadme.Enabled = True
		Me.mnuHelpReadme.Visible = True
		Me.mnuHelpWebsite.Name = "mnuHelpWebsite"
		Me.mnuHelpWebsite.Text = "&Forum"
		Me.mnuHelpWebsite.Checked = False
		Me.mnuHelpWebsite.Enabled = True
		Me.mnuHelpWebsite.Visible = True
		Me.mnuTerms.Name = "mnuTerms"
		Me.mnuTerms.Text = "&End-User License Agreement"
		Me.mnuTerms.Checked = False
		Me.mnuTerms.Enabled = True
		Me.mnuTerms.Visible = True
		Me.mnuChangeLog.Name = "mnuChangeLog"
		Me.mnuChangeLog.Text = "&Change Log"
		Me.mnuChangeLog.Checked = False
		Me.mnuChangeLog.Enabled = True
		Me.mnuChangeLog.Visible = True
		Me.mnuShortcuts.Name = "mnuShortcuts"
		Me.mnuShortcuts.Text = "invisibleMenu"
		Me.mnuShortcuts.Visible = False
		Me.mnuShortcuts.Checked = False
		Me.mnuShortcuts.Enabled = True
		Me.mnuItalic.Name = "mnuItalic"
		Me.mnuItalic.Text = "Italic"
		Me.mnuItalic.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.I, System.Windows.Forms.Keys)
		Me.mnuItalic.Checked = False
		Me.mnuItalic.Enabled = True
		Me.mnuItalic.Visible = True
		Me.mnuBold.Name = "mnuBold"
		Me.mnuBold.Text = "Bold"
		Me.mnuBold.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.B, System.Windows.Forms.Keys)
		Me.mnuBold.Checked = False
		Me.mnuBold.Enabled = True
		Me.mnuBold.Visible = True
		Me.mnuUnderline.Name = "mnuUnderline"
		Me.mnuUnderline.Text = "Underline"
		Me.mnuUnderline.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.U, System.Windows.Forms.Keys)
		Me.mnuUnderline.Checked = False
		Me.mnuUnderline.Enabled = True
		Me.mnuUnderline.Visible = True
		Me._mnuListViewButton_0.Name = "_mnuListViewButton_0"
		Me._mnuListViewButton_0.Text = "invisible CTRL+A Channel"
		Me._mnuListViewButton_0.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.A, System.Windows.Forms.Keys)
		Me._mnuListViewButton_0.Visible = False
		Me._mnuListViewButton_0.Checked = False
		Me._mnuListViewButton_0.Enabled = True
		Me._mnuListViewButton_1.Name = "_mnuListViewButton_1"
		Me._mnuListViewButton_1.Text = "invisible CTRL+S Friends"
		Me._mnuListViewButton_1.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.S, System.Windows.Forms.Keys)
		Me._mnuListViewButton_1.Visible = False
		Me._mnuListViewButton_1.Checked = False
		Me._mnuListViewButton_1.Enabled = True
		Me._mnuListViewButton_2.Name = "_mnuListViewButton_2"
		Me._mnuListViewButton_2.Text = "invisible CTRL+D Clan"
		Me._mnuListViewButton_2.ShortcutKeys = CType(System.Windows.Forms.Keys.Control or System.Windows.Forms.Keys.D, System.Windows.Forms.Keys)
		Me._mnuListViewButton_2.Visible = False
		Me._mnuListViewButton_2.Checked = False
		Me._mnuListViewButton_2.Enabled = True
		Me.mnuPopClanList.Name = "mnuPopClanList"
		Me.mnuPopClanList.Text = "clanlist popup menu"
		Me.mnuPopClanList.Visible = False
		Me.mnuPopClanList.Checked = False
		Me.mnuPopClanList.Enabled = True
		Me.mnuPopClanWhisper.Name = "mnuPopClanWhisper"
		Me.mnuPopClanWhisper.Text = "W&hisper"
		Me.mnuPopClanWhisper.Checked = False
		Me.mnuPopClanWhisper.Enabled = True
		Me.mnuPopClanWhisper.Visible = True
		Me.mnuPopClanCopy.Name = "mnuPopClanCopy"
		Me.mnuPopClanCopy.Text = "&Copy Name to Clipboard"
		Me.mnuPopClanCopy.Checked = False
		Me.mnuPopClanCopy.Enabled = True
		Me.mnuPopClanCopy.Visible = True
		Me.mnuPopClanAddLeft.Name = "mnuPopClanAddLeft"
		Me.mnuPopClanAddLeft.Text = "Add to &Left Send Box"
		Me.mnuPopClanAddLeft.Checked = False
		Me.mnuPopClanAddLeft.Enabled = True
		Me.mnuPopClanAddLeft.Visible = True
		Me.mnuPopClanAddToFList.Name = "mnuPopClanAddToFList"
		Me.mnuPopClanAddToFList.Text = "Add to &Friends List"
		Me.mnuPopClanAddToFList.Checked = False
		Me.mnuPopClanAddToFList.Enabled = True
		Me.mnuPopClanAddToFList.Visible = True
		Me.mnuPopClanSep1.Enabled = True
		Me.mnuPopClanSep1.Visible = True
		Me.mnuPopClanSep1.Name = "mnuPopClanSep1"
		Me.mnuPopClanUserlistWhois.Name = "mnuPopClanUserlistWhois"
		Me.mnuPopClanUserlistWhois.Text = "&Userlist Whois"
		Me.mnuPopClanUserlistWhois.Checked = False
		Me.mnuPopClanUserlistWhois.Enabled = True
		Me.mnuPopClanUserlistWhois.Visible = True
		Me.mnuPopClanWhois.Name = "mnuPopClanWhois"
		Me.mnuPopClanWhois.Text = "Battle.net &Whois"
		Me.mnuPopClanWhois.Checked = False
		Me.mnuPopClanWhois.Enabled = True
		Me.mnuPopClanWhois.Visible = True
		Me.mnuPopClanStats.Name = "mnuPopClanStats"
		Me.mnuPopClanStats.Text = "Battle.net &Stats"
		Me.mnuPopClanStats.Checked = False
		Me.mnuPopClanStats.Enabled = True
		Me.mnuPopClanStats.Visible = True
		Me.mnuPopClanStatsWAR3.Name = "mnuPopClanStatsWAR3"
		Me.mnuPopClanStatsWAR3.Text = "&Reign of Chaos"
		Me.mnuPopClanStatsWAR3.Checked = False
		Me.mnuPopClanStatsWAR3.Enabled = True
		Me.mnuPopClanStatsWAR3.Visible = True
		Me.mnuPopClanStatsW3XP.Name = "mnuPopClanStatsW3XP"
		Me.mnuPopClanStatsW3XP.Text = "The &Frozen Throne"
		Me.mnuPopClanStatsW3XP.Checked = False
		Me.mnuPopClanStatsW3XP.Enabled = True
		Me.mnuPopClanStatsW3XP.Visible = True
		Me.mnuPopClanProfile.Name = "mnuPopClanProfile"
		Me.mnuPopClanProfile.Text = "&Battle.net Profile"
		Me.mnuPopClanProfile.Checked = False
		Me.mnuPopClanProfile.Enabled = True
		Me.mnuPopClanProfile.Visible = True
		Me.mnuPopClanWebProfile.Name = "mnuPopClanWebProfile"
		Me.mnuPopClanWebProfile.Text = "W&eb Profile"
		Me.mnuPopClanWebProfile.Checked = False
		Me.mnuPopClanWebProfile.Enabled = True
		Me.mnuPopClanWebProfile.Visible = True
		Me.mnuPopClanWebProfileWAR3.Name = "mnuPopClanWebProfileWAR3"
		Me.mnuPopClanWebProfileWAR3.Text = "&Reign of Chaos"
		Me.mnuPopClanWebProfileWAR3.Checked = False
		Me.mnuPopClanWebProfileWAR3.Enabled = True
		Me.mnuPopClanWebProfileWAR3.Visible = True
		Me.mnuPopClanWebProfileW3XP.Name = "mnuPopClanWebProfileW3XP"
		Me.mnuPopClanWebProfileW3XP.Text = "The &Frozen Throne"
		Me.mnuPopClanWebProfileW3XP.Checked = False
		Me.mnuPopClanWebProfileW3XP.Enabled = True
		Me.mnuPopClanWebProfileW3XP.Visible = True
		Me.mnuPopClanSep2.Enabled = True
		Me.mnuPopClanSep2.Visible = True
		Me.mnuPopClanSep2.Name = "mnuPopClanSep2"
		Me.mnuPopClanPromote.Name = "mnuPopClanPromote"
		Me.mnuPopClanPromote.Text = "&Promote"
		Me.mnuPopClanPromote.Checked = False
		Me.mnuPopClanPromote.Enabled = True
		Me.mnuPopClanPromote.Visible = True
		Me.mnuPopClanDemote.Name = "mnuPopClanDemote"
		Me.mnuPopClanDemote.Text = "&Demote"
		Me.mnuPopClanDemote.Checked = False
		Me.mnuPopClanDemote.Enabled = True
		Me.mnuPopClanDemote.Visible = True
		Me.mnuPopClanRemove.Name = "mnuPopClanRemove"
		Me.mnuPopClanRemove.Text = "Remove from Clan"
		Me.mnuPopClanRemove.Checked = False
		Me.mnuPopClanRemove.Enabled = True
		Me.mnuPopClanRemove.Visible = True
		Me.mnuPopClanLeave.Name = "mnuPopClanLeave"
		Me.mnuPopClanLeave.Text = "Leave Clan"
		Me.mnuPopClanLeave.Visible = False
		Me.mnuPopClanLeave.Checked = False
		Me.mnuPopClanLeave.Enabled = True
		Me.mnuPopClanMakeChief.Name = "mnuPopClanMakeChief"
		Me.mnuPopClanMakeChief.Text = "Make Chieftain"
		Me.mnuPopClanMakeChief.Checked = False
		Me.mnuPopClanMakeChief.Enabled = True
		Me.mnuPopClanMakeChief.Visible = True
		Me.mnuPopClanDisband.Name = "mnuPopClanDisband"
		Me.mnuPopClanDisband.Text = "Disband Clan"
		Me.mnuPopClanDisband.Checked = False
		Me.mnuPopClanDisband.Enabled = True
		Me.mnuPopClanDisband.Visible = True
		Me.mnuPopFList.Name = "mnuPopFList"
		Me.mnuPopFList.Text = "flistpopup"
		Me.mnuPopFList.Visible = False
		Me.mnuPopFList.Checked = False
		Me.mnuPopFList.Enabled = True
		Me.mnuPopFLWhisper.Name = "mnuPopFLWhisper"
		Me.mnuPopFLWhisper.Text = "W&hisper"
		Me.mnuPopFLWhisper.Checked = False
		Me.mnuPopFLWhisper.Enabled = True
		Me.mnuPopFLWhisper.Visible = True
		Me.mnuPopFLCopy.Name = "mnuPopFLCopy"
		Me.mnuPopFLCopy.Text = "&Copy Name to Clipboard"
		Me.mnuPopFLCopy.Checked = False
		Me.mnuPopFLCopy.Enabled = True
		Me.mnuPopFLCopy.Visible = True
		Me.mnuPopFLAddLeft.Name = "mnuPopFLAddLeft"
		Me.mnuPopFLAddLeft.Text = "Add to &Left Send Box"
		Me.mnuPopFLAddLeft.Checked = False
		Me.mnuPopFLAddLeft.Enabled = True
		Me.mnuPopFLAddLeft.Visible = True
		Me.mnuPopFLInvite.Name = "mnuPopFLInvite"
		Me.mnuPopFLInvite.Text = "&Invite to Warcraft III Clan"
		Me.mnuPopFLInvite.Checked = False
		Me.mnuPopFLInvite.Enabled = True
		Me.mnuPopFLInvite.Visible = True
		Me.mnuPopFLSep1.Enabled = True
		Me.mnuPopFLSep1.Visible = True
		Me.mnuPopFLSep1.Name = "mnuPopFLSep1"
		Me.mnuPopFLUserlistWhois.Name = "mnuPopFLUserlistWhois"
		Me.mnuPopFLUserlistWhois.Text = "&Userlist Whois"
		Me.mnuPopFLUserlistWhois.Checked = False
		Me.mnuPopFLUserlistWhois.Enabled = True
		Me.mnuPopFLUserlistWhois.Visible = True
		Me.mnuPopFLWhois.Name = "mnuPopFLWhois"
		Me.mnuPopFLWhois.Text = "Battle.net &Whois"
		Me.mnuPopFLWhois.Checked = False
		Me.mnuPopFLWhois.Enabled = True
		Me.mnuPopFLWhois.Visible = True
		Me.mnuPopFLStats.Name = "mnuPopFLStats"
		Me.mnuPopFLStats.Text = "Battle.net &Stats"
		Me.mnuPopFLStats.Checked = False
		Me.mnuPopFLStats.Enabled = True
		Me.mnuPopFLStats.Visible = True
		Me.mnuPopFLProfile.Name = "mnuPopFLProfile"
		Me.mnuPopFLProfile.Text = "&Battle.net Profile"
		Me.mnuPopFLProfile.Checked = False
		Me.mnuPopFLProfile.Enabled = True
		Me.mnuPopFLProfile.Visible = True
		Me.mnuPopFLWebProfile.Name = "mnuPopFLWebProfile"
		Me.mnuPopFLWebProfile.Text = "W&eb Profile"
		Me.mnuPopFLWebProfile.Checked = False
		Me.mnuPopFLWebProfile.Enabled = True
		Me.mnuPopFLWebProfile.Visible = True
		Me.mnuPopFLSep2.Enabled = True
		Me.mnuPopFLSep2.Visible = True
		Me.mnuPopFLSep2.Name = "mnuPopFLSep2"
		Me.mnuPopFLPromote.Name = "mnuPopFLPromote"
		Me.mnuPopFLPromote.Text = "&Promote"
		Me.mnuPopFLPromote.Checked = False
		Me.mnuPopFLPromote.Enabled = True
		Me.mnuPopFLPromote.Visible = True
		Me.mnuPopFLDemote.Name = "mnuPopFLDemote"
		Me.mnuPopFLDemote.Text = "&Demote"
		Me.mnuPopFLDemote.Checked = False
		Me.mnuPopFLDemote.Enabled = True
		Me.mnuPopFLDemote.Visible = True
		Me.mnuPopFLRemove.Name = "mnuPopFLRemove"
		Me.mnuPopFLRemove.Text = "&Remove"
		Me.mnuPopFLRemove.Checked = False
		Me.mnuPopFLRemove.Enabled = True
		Me.mnuPopFLRemove.Visible = True
		Me.mnuPopFLSep3.Enabled = True
		Me.mnuPopFLSep3.Visible = True
		Me.mnuPopFLSep3.Name = "mnuPopFLSep3"
		Me.mnuPopFLRefresh.Name = "mnuPopFLRefresh"
		Me.mnuPopFLRefresh.Text = "Refresh and Reorder"
		Me.mnuPopFLRefresh.Checked = False
		Me.mnuPopFLRefresh.Enabled = True
		Me.mnuPopFLRefresh.Visible = True
		Me.lvChannel.Size = New System.Drawing.Size(247, 425)
		Me.lvChannel.Location = New System.Drawing.Point(592, 40)
		Me.lvChannel.TabIndex = 1
		Me.lvChannel.TabStop = 0
		Me.lvChannel.View = System.Windows.Forms.View.Details
		Me.lvChannel.LabelEdit = False
		Me.lvChannel.LabelWrap = False
		Me.lvChannel.HideSelection = True
		Me.lvChannel.SmallImageList = imlIcons
		Me.lvChannel.ForeColor = System.Drawing.Color.FromARGB(0, 204, 153)
		Me.lvChannel.BackColor = System.Drawing.Color.Black
		Me.lvChannel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lvChannel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lvChannel.Name = "lvChannel"
		Me._lvChannel_ColumnHeader_1.Width = 277
		Me._lvChannel_ColumnHeader_2.Width = 84
		Me._lvChannel_ColumnHeader_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._lvChannel_ColumnHeader_3.Width = 39
		Me.tmrAccountLock.Enabled = False
		Me.tmrAccountLock.Interval = 30000
		SControl.OcxState = CType(resources.GetObject("SControl.OcxState"), System.Windows.Forms.AxHost.State)
		Me.SControl.Location = New System.Drawing.Point(8, 72)
		Me.SControl.Name = "SControl"
		SCRestricted.OcxState = CType(resources.GetObject("SCRestricted.OcxState"), System.Windows.Forms.AxHost.State)
		Me.SCRestricted.Location = New System.Drawing.Point(392, 248)
		Me.SCRestricted.Name = "SCRestricted"
		Me._tmrScriptLong_0.Enabled = False
		Me._tmrScriptLong_0.Interval = 1
		_itcScript_0.OcxState = CType(resources.GetObject("_itcScript_0.OcxState"), System.Windows.Forms.AxHost.State)
		Me._itcScript_0.Location = New System.Drawing.Point(8, 32)
		Me._itcScript_0.Name = "_itcScript_0"
		_sckScript_0.OcxState = CType(resources.GetObject("_sckScript_0.OcxState"), System.Windows.Forms.AxHost.State)
		Me._sckScript_0.Location = New System.Drawing.Point(48, 32)
		Me._sckScript_0.Name = "_sckScript_0"
		Me.ChatQueueTimer.Enabled = False
		Me.ChatQueueTimer.Interval = 500
		Me.cacheTimer.Enabled = False
		Me.cacheTimer.Interval = 2500
		Me.imlClan.ImageSize = New System.Drawing.Size(37, 23)
		Me.imlClan.ImageStream = CType(resources.GetObject("imlClan.ImageStream"), System.Windows.Forms.ImageListStreamer)
		Me.imlClan.Images.SetKeyName(0, "")
		Me.imlClan.Images.SetKeyName(1, "")
		Me.imlClan.Images.SetKeyName(2, "")
		Me.imlClan.Images.SetKeyName(3, "")
		Me.imlClan.Images.SetKeyName(4, "")
		Me.imlClan.Images.SetKeyName(5, "")
		Me.imlClan.Images.SetKeyName(6, "")
		Me.imlIcons.ImageSize = New System.Drawing.Size(28, 18)
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
		Me.imlIcons.Images.SetKeyName(13, "")
		Me.imlIcons.Images.SetKeyName(14, "")
		Me.imlIcons.Images.SetKeyName(15, "")
		Me.imlIcons.Images.SetKeyName(16, "")
		Me.imlIcons.Images.SetKeyName(17, "")
		Me.imlIcons.Images.SetKeyName(18, "")
		Me.imlIcons.Images.SetKeyName(19, "")
		Me.imlIcons.Images.SetKeyName(20, "")
		Me.imlIcons.Images.SetKeyName(21, "")
		Me.imlIcons.Images.SetKeyName(22, "")
		Me.imlIcons.Images.SetKeyName(23, "")
		Me.imlIcons.Images.SetKeyName(24, "")
		Me.imlIcons.Images.SetKeyName(25, "")
		Me.imlIcons.Images.SetKeyName(26, "")
		Me.imlIcons.Images.SetKeyName(27, "")
		Me.imlIcons.Images.SetKeyName(28, "")
		Me.imlIcons.Images.SetKeyName(29, "")
		Me.imlIcons.Images.SetKeyName(30, "")
		Me.imlIcons.Images.SetKeyName(31, "")
		Me.imlIcons.Images.SetKeyName(32, "")
		Me.imlIcons.Images.SetKeyName(33, "")
		Me.imlIcons.Images.SetKeyName(34, "")
		Me.imlIcons.Images.SetKeyName(35, "")
		Me.imlIcons.Images.SetKeyName(36, "")
		Me.imlIcons.Images.SetKeyName(37, "")
		Me.imlIcons.Images.SetKeyName(38, "")
		Me.imlIcons.Images.SetKeyName(39, "")
		Me.imlIcons.Images.SetKeyName(40, "")
		Me.imlIcons.Images.SetKeyName(41, "")
		Me.imlIcons.Images.SetKeyName(42, "")
		Me.imlIcons.Images.SetKeyName(43, "")
		Me.imlIcons.Images.SetKeyName(44, "")
		Me.imlIcons.Images.SetKeyName(45, "")
		Me.imlIcons.Images.SetKeyName(46, "")
		Me.imlIcons.Images.SetKeyName(47, "")
		Me.imlIcons.Images.SetKeyName(48, "")
		Me.imlIcons.Images.SetKeyName(49, "")
		Me.imlIcons.Images.SetKeyName(50, "")
		Me.imlIcons.Images.SetKeyName(51, "")
		Me.imlIcons.Images.SetKeyName(52, "")
		Me.imlIcons.Images.SetKeyName(53, "")
		Me.imlIcons.Images.SetKeyName(54, "")
		Me.imlIcons.Images.SetKeyName(55, "")
		Me.imlIcons.Images.SetKeyName(56, "")
		Me.imlIcons.Images.SetKeyName(57, "")
		Me.imlIcons.Images.SetKeyName(58, "")
		Me.imlIcons.Images.SetKeyName(59, "")
		Me.imlIcons.Images.SetKeyName(60, "")
		Me.imlIcons.Images.SetKeyName(61, "")
		Me.imlIcons.Images.SetKeyName(62, "")
		Me.imlIcons.Images.SetKeyName(63, "")
		Me.imlIcons.Images.SetKeyName(64, "")
		Me.imlIcons.Images.SetKeyName(65, "")
		Me.imlIcons.Images.SetKeyName(66, "")
		Me.imlIcons.Images.SetKeyName(67, "")
		Me.imlIcons.Images.SetKeyName(68, "")
		Me.imlIcons.Images.SetKeyName(69, "")
		Me.imlIcons.Images.SetKeyName(70, "")
		Me.imlIcons.Images.SetKeyName(71, "")
		Me.imlIcons.Images.SetKeyName(72, "")
		Me.imlIcons.Images.SetKeyName(73, "")
		Me.imlIcons.Images.SetKeyName(74, "")
		Me.imlIcons.Images.SetKeyName(75, "")
		Me.imlIcons.Images.SetKeyName(76, "")
		Me.imlIcons.Images.SetKeyName(77, "")
		Me.imlIcons.Images.SetKeyName(78, "")
		Me.imlIcons.Images.SetKeyName(79, "")
		Me.imlIcons.Images.SetKeyName(80, "")
		Me.imlIcons.Images.SetKeyName(81, "")
		Me.imlIcons.Images.SetKeyName(82, "")
		Me.imlIcons.Images.SetKeyName(83, "")
		Me.imlIcons.Images.SetKeyName(84, "")
		Me.imlIcons.Images.SetKeyName(85, "")
		Me.imlIcons.Images.SetKeyName(86, "")
		Me.imlIcons.Images.SetKeyName(87, "")
		Me.imlIcons.Images.SetKeyName(88, "")
		Me.imlIcons.Images.SetKeyName(89, "")
		Me.imlIcons.Images.SetKeyName(90, "")
		Me.imlIcons.Images.SetKeyName(91, "")
		Me.imlIcons.Images.SetKeyName(92, "")
		Me.imlIcons.Images.SetKeyName(93, "")
		Me.imlIcons.Images.SetKeyName(94, "")
		Me.imlIcons.Images.SetKeyName(95, "")
		Me.imlIcons.Images.SetKeyName(96, "")
		Me.imlIcons.Images.SetKeyName(97, "")
		Me.imlIcons.Images.SetKeyName(98, "")
		Me.imlIcons.Images.SetKeyName(99, "")
		Me.imlIcons.Images.SetKeyName(100, "")
		Me.imlIcons.Images.SetKeyName(101, "")
		Me.imlIcons.Images.SetKeyName(102, "")
		Me.imlIcons.Images.SetKeyName(103, "")
		Me.imlIcons.Images.SetKeyName(104, "")
		Me.imlIcons.Images.SetKeyName(105, "")
		Me.imlIcons.Images.SetKeyName(106, "")
		Me.imlIcons.Images.SetKeyName(107, "")
		Me.imlIcons.Images.SetKeyName(108, "")
		Me.imlIcons.Images.SetKeyName(109, "")
		Me.imlIcons.Images.SetKeyName(110, "")
		Me.imlIcons.Images.SetKeyName(111, "")
		Me.imlIcons.Images.SetKeyName(112, "")
		Me.imlIcons.Images.SetKeyName(113, "")
		Me.imlIcons.Images.SetKeyName(114, "")
		Me.imlIcons.Images.SetKeyName(115, "")
		Me.imlIcons.Images.SetKeyName(116, "")
		Me.imlIcons.Images.SetKeyName(117, "")
		Me.imlIcons.Images.SetKeyName(118, "")
		Me.imlIcons.Images.SetKeyName(119, "")
		Me.imlIcons.Images.SetKeyName(120, "")
		Me.imlIcons.Images.SetKeyName(121, "")
		Me.imlIcons.Images.SetKeyName(122, "")
		Me.imlIcons.Images.SetKeyName(123, "")
		Me.imlIcons.Images.SetKeyName(124, "")
		Me.imlIcons.Images.SetKeyName(125, "")
		Me.imlIcons.Images.SetKeyName(126, "")
		Me.imlIcons.Images.SetKeyName(127, "")
		Me.imlIcons.Images.SetKeyName(128, "")
		Me.imlIcons.Images.SetKeyName(129, "")
		Me.imlIcons.Images.SetKeyName(130, "")
		Me.imlIcons.Images.SetKeyName(131, "")
		Me.imlIcons.Images.SetKeyName(132, "")
		Me.imlIcons.Images.SetKeyName(133, "")
		Me.imlIcons.Images.SetKeyName(134, "")
		Me.imlIcons.Images.SetKeyName(135, "")
		Me.imlIcons.Images.SetKeyName(136, "")
		Me.imlIcons.Images.SetKeyName(137, "")
		Me.imlIcons.Images.SetKeyName(138, "")
		Me.imlIcons.Images.SetKeyName(139, "")
		Me.imlIcons.Images.SetKeyName(140, "")
		Me.imlIcons.Images.SetKeyName(141, "")
		Me.imlIcons.Images.SetKeyName(142, "")
		Me.imlIcons.Images.SetKeyName(143, "")
		Me.imlIcons.Images.SetKeyName(144, "")
		Me.imlIcons.Images.SetKeyName(145, "")
		Me.imlIcons.Images.SetKeyName(146, "")
		Me.imlIcons.Images.SetKeyName(147, "")
		Me.imlIcons.Images.SetKeyName(148, "")
		Me.imlIcons.Images.SetKeyName(149, "")
		Me.imlIcons.Images.SetKeyName(150, "")
		Me.imlIcons.Images.SetKeyName(151, "")
		Me.imlIcons.Images.SetKeyName(152, "")
		Me.imlIcons.Images.SetKeyName(153, "")
		Me.imlIcons.Images.SetKeyName(154, "")
		Me.imlIcons.Images.SetKeyName(155, "")
		Me.imlIcons.Images.SetKeyName(156, "")
		Me.imlIcons.Images.SetKeyName(157, "")
		Me._tmrScript_0.Enabled = False
		Me._tmrScript_0.Interval = 1
		Me._tmrSilentChannel_1.Enabled = False
		Me._tmrSilentChannel_1.Interval = 30000
		Me._tmrSilentChannel_0.Enabled = False
		Me._tmrSilentChannel_0.Interval = 500
		Me.cboSend.BackColor = System.Drawing.Color.Black
		Me.cboSend.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboSend.ForeColor = System.Drawing.Color.White
		Me.cboSend.Size = New System.Drawing.Size(513, 21)
		Me.cboSend.Location = New System.Drawing.Point(40, 464)
		Me.cboSend.TabIndex = 7
		Me.cboSend.CausesValidation = True
		Me.cboSend.Enabled = True
		Me.cboSend.IntegralHeight = True
		Me.cboSend.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboSend.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboSend.Sorted = False
		Me.cboSend.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboSend.TabStop = True
		Me.cboSend.Visible = True
		Me.cboSend.Name = "cboSend"
		Me.txtPost.AutoSize = False
		Me.txtPost.BackColor = System.Drawing.Color.Black
		Me.txtPost.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPost.ForeColor = System.Drawing.Color.White
		Me.txtPost.Size = New System.Drawing.Size(41, 21)
		Me.txtPost.Location = New System.Drawing.Point(552, 464)
		Me.txtPost.TabIndex = 9
		Me.txtPost.AcceptsReturn = True
		Me.txtPost.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPost.CausesValidation = True
		Me.txtPost.Enabled = True
		Me.txtPost.HideSelection = True
		Me.txtPost.ReadOnly = False
		Me.txtPost.Maxlength = 0
		Me.txtPost.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPost.MultiLine = False
		Me.txtPost.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPost.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPost.TabStop = True
		Me.txtPost.Visible = True
		Me.txtPost.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPost.Name = "txtPost"
		Me.txtPre.AutoSize = False
		Me.txtPre.BackColor = System.Drawing.Color.Black
		Me.txtPre.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPre.ForeColor = System.Drawing.Color.White
		Me.txtPre.Size = New System.Drawing.Size(41, 21)
		Me.txtPre.Location = New System.Drawing.Point(0, 464)
		Me.txtPre.TabIndex = 8
		Me.txtPre.AcceptsReturn = True
		Me.txtPre.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPre.CausesValidation = True
		Me.txtPre.Enabled = True
		Me.txtPre.HideSelection = True
		Me.txtPre.ReadOnly = False
		Me.txtPre.Maxlength = 0
		Me.txtPre.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPre.MultiLine = False
		Me.txtPre.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPre.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPre.TabStop = True
		Me.txtPre.Visible = True
		Me.txtPre.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPre.Name = "txtPre"
		Me.ListviewTabs.Size = New System.Drawing.Size(247, 25)
		Me.ListviewTabs.Location = New System.Drawing.Point(592, 464)
		Me.ListviewTabs.TabIndex = 0
		Me.ListviewTabs.Alignment = System.Windows.Forms.TabAlignment.Bottom
		Me.ListviewTabs.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
		Me.ListviewTabs.SelectedIndex = 1
		Me.ListviewTabs.ItemSize = New System.Drawing.Size(42, 18)
		Me.ListviewTabs.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ListviewTabs.Name = "ListviewTabs"
		Me._ListviewTabs_TabPage0.Text = "Channel  "
		Me._ListviewTabs_TabPage1.Text = "Friends  "
		Me._ListviewTabs_TabPage2.Text = "Clan  "
		Me.tmrIdleTimer.Interval = 1000
		Me.tmrIdleTimer.Enabled = True
		Me.cmdShowHide.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdShowHide.Text = " ^^^^"
		Me.cmdShowHide.Font = New System.Drawing.Font("Tahoma", 5.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdShowHide.Size = New System.Drawing.Size(17, 113)
		Me.cmdShowHide.Location = New System.Drawing.Point(824, 464)
		Me.cmdShowHide.TabIndex = 3
		Me.cmdShowHide.TabStop = False
		Me.cmdShowHide.BackColor = System.Drawing.SystemColors.Control
		Me.cmdShowHide.CausesValidation = True
		Me.cmdShowHide.Enabled = True
		Me.cmdShowHide.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdShowHide.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdShowHide.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdShowHide.Name = "cmdShowHide"
		sckMCP.OcxState = CType(resources.GetObject("sckMCP.OcxState"), System.Windows.Forms.AxHost.State)
		Me.sckMCP.Location = New System.Drawing.Point(352, 208)
		Me.sckMCP.Name = "sckMCP"
		Me.scTimer.Enabled = False
		Me.scTimer.Interval = 1
		INet.OcxState = CType(resources.GetObject("INet.OcxState"), System.Windows.Forms.AxHost.State)
		Me.INet.Location = New System.Drawing.Point(352, 248)
		Me.INet.Name = "INet"
		sckBNLS.OcxState = CType(resources.GetObject("sckBNLS.OcxState"), System.Windows.Forms.AxHost.State)
		Me.sckBNLS.Location = New System.Drawing.Point(384, 208)
		Me.sckBNLS.Name = "sckBNLS"
		sckBNet.OcxState = CType(resources.GetObject("sckBNet.OcxState"), System.Windows.Forms.AxHost.State)
		Me.sckBNet.Location = New System.Drawing.Point(416, 208)
		Me.sckBNet.Name = "sckBNet"
		Me.UpTimer.Enabled = False
		Me.UpTimer.Interval = 1
		Me.lvClanList.Size = New System.Drawing.Size(247, 425)
		Me.lvClanList.Location = New System.Drawing.Point(592, 40)
		Me.lvClanList.TabIndex = 5
		Me.lvClanList.TabStop = 0
		Me.lvClanList.View = System.Windows.Forms.View.Details
		Me.lvClanList.LabelEdit = False
		Me.lvClanList.LabelWrap = False
		Me.lvClanList.HideSelection = True
		Me.lvClanList.SmallImageList = imlIcons
		Me.lvClanList.ForeColor = System.Drawing.Color.FromARGB(0, 204, 153)
		Me.lvClanList.BackColor = System.Drawing.Color.Black
		Me.lvClanList.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lvClanList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lvClanList.Name = "lvClanList"
		Me._lvClanList_ColumnHeader_1.Width = 271
		Me._lvClanList_ColumnHeader_2.Width = 170
		Me._lvClanList_ColumnHeader_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._lvClanList_ColumnHeader_3.Width = 170
		Me._lvClanList_ColumnHeader_4.Width = 6
		Me.lvFriendList.Size = New System.Drawing.Size(247, 425)
		Me.lvFriendList.Location = New System.Drawing.Point(592, 40)
		Me.lvFriendList.TabIndex = 2
		Me.lvFriendList.TabStop = 0
		Me.lvFriendList.View = System.Windows.Forms.View.Details
		Me.lvFriendList.LabelEdit = False
		Me.lvFriendList.LabelWrap = False
		Me.lvFriendList.HideSelection = True
		Me.lvFriendList.LargeImageList = imlIcons
		Me.lvFriendList.SmallImageList = imlIcons
		Me.lvFriendList.ForeColor = System.Drawing.Color.FromARGB(0, 204, 153)
		Me.lvFriendList.BackColor = System.Drawing.Color.Black
		Me.lvFriendList.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lvFriendList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lvFriendList.Name = "lvFriendList"
		Me._lvFriendList_ColumnHeader_1.Width = 271
		Me._lvFriendList_ColumnHeader_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
		Me._lvFriendList_ColumnHeader_2.Width = 6
		Me.rtbWhispers.Size = New System.Drawing.Size(825, 113)
		Me.rtbWhispers.Location = New System.Drawing.Point(0, 488)
		Me.rtbWhispers.TabIndex = 4
		Me.rtbWhispers.TabStop = 0
		Me.rtbWhispers.BackColor = System.Drawing.Color.Black
		Me.rtbWhispers.ReadOnly = True
		Me.rtbWhispers.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
		Me.rtbWhispers.RTF = resources.GetString("rtbWhispers.TextRTF")
		Me.rtbWhispers.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.rtbWhispers.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.rtbWhispers.Name = "rtbWhispers"
		Me.rtbChat.Size = New System.Drawing.Size(593, 441)
		Me.rtbChat.Location = New System.Drawing.Point(0, 24)
		Me.rtbChat.TabIndex = 6
		Me.rtbChat.TabStop = 0
		Me.rtbChat.BackColor = System.Drawing.Color.Black
		Me.rtbChat.ReadOnly = True
		Me.rtbChat.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
		Me.rtbChat.RTF = resources.GetString("rtbChat.TextRTF")
		Me.rtbChat.Font = New System.Drawing.Font("Tahoma", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.rtbChat.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.rtbChat.Name = "rtbChat"
		Me.lblCurrentChannel.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblCurrentChannel.BackColor = System.Drawing.Color.FromARGB(0, 51, 204)
		Me.lblCurrentChannel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCurrentChannel.Size = New System.Drawing.Size(247, 17)
		Me.lblCurrentChannel.Location = New System.Drawing.Point(592, 24)
		Me.lblCurrentChannel.TabIndex = 10
		Me.lblCurrentChannel.Enabled = True
		Me.lblCurrentChannel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCurrentChannel.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCurrentChannel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCurrentChannel.UseMnemonic = True
		Me.lblCurrentChannel.Visible = True
		Me.lblCurrentChannel.AutoSize = False
		Me.lblCurrentChannel.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCurrentChannel.Name = "lblCurrentChannel"
		Me.Controls.Add(lvChannel)
		Me.Controls.Add(SControl)
		Me.Controls.Add(SCRestricted)
		Me.Controls.Add(_itcScript_0)
		Me.Controls.Add(_sckScript_0)
		Me.Controls.Add(cboSend)
		Me.Controls.Add(txtPost)
		Me.Controls.Add(txtPre)
		Me.Controls.Add(ListviewTabs)
		Me.Controls.Add(cmdShowHide)
		Me.Controls.Add(sckMCP)
		Me.Controls.Add(INet)
		Me.Controls.Add(sckBNLS)
		Me.Controls.Add(sckBNet)
		Me.Controls.Add(lvClanList)
		Me.Controls.Add(lvFriendList)
		Me.Controls.Add(rtbWhispers)
		Me.Controls.Add(rtbChat)
		Me.Controls.Add(lblCurrentChannel)
		Me.lvChannel.Columns.Add(_lvChannel_ColumnHeader_1)
		Me.lvChannel.Columns.Add(_lvChannel_ColumnHeader_2)
		Me.lvChannel.Columns.Add(_lvChannel_ColumnHeader_3)
		Me.ListviewTabs.Controls.Add(_ListviewTabs_TabPage0)
		Me.ListviewTabs.Controls.Add(_ListviewTabs_TabPage1)
		Me.ListviewTabs.Controls.Add(_ListviewTabs_TabPage2)
		Me.lvClanList.Columns.Add(_lvClanList_ColumnHeader_1)
		Me.lvClanList.Columns.Add(_lvClanList_ColumnHeader_2)
		Me.lvClanList.Columns.Add(_lvClanList_ColumnHeader_3)
		Me.lvClanList.Columns.Add(_lvClanList_ColumnHeader_4)
		Me.lvFriendList.Columns.Add(_lvFriendList_ColumnHeader_1)
		Me.lvFriendList.Columns.Add(_lvFriendList_ColumnHeader_2)
		Me.itcScript.SetIndex(_itcScript_0, CType(0, Short))
		Me.mnuCustomChannels.SetIndex(_mnuCustomChannels_0, CType(0, Short))
		Me.mnuCustomChannels.SetIndex(_mnuCustomChannels_1, CType(1, Short))
		Me.mnuCustomChannels.SetIndex(_mnuCustomChannels_2, CType(2, Short))
		Me.mnuCustomChannels.SetIndex(_mnuCustomChannels_3, CType(3, Short))
		Me.mnuCustomChannels.SetIndex(_mnuCustomChannels_4, CType(4, Short))
		Me.mnuCustomChannels.SetIndex(_mnuCustomChannels_5, CType(5, Short))
		Me.mnuCustomChannels.SetIndex(_mnuCustomChannels_6, CType(6, Short))
		Me.mnuCustomChannels.SetIndex(_mnuCustomChannels_7, CType(7, Short))
		Me.mnuCustomChannels.SetIndex(_mnuCustomChannels_8, CType(8, Short))
		Me.mnuListViewButton.SetIndex(_mnuListViewButton_0, CType(0, Short))
		Me.mnuListViewButton.SetIndex(_mnuListViewButton_1, CType(1, Short))
		Me.mnuListViewButton.SetIndex(_mnuListViewButton_2, CType(2, Short))
		Me.mnuPublicChannels.SetIndex(_mnuPublicChannels_0, CType(0, Short))
		Me.mnuScriptingDash.SetIndex(_mnuScriptingDash_0, CType(0, Short))
		Me.sckScript.SetIndex(_sckScript_0, CType(0, Short))
		Me.tmrScript.SetIndex(_tmrScript_0, CType(0, Short))
		Me.tmrScriptLong.SetIndex(_tmrScriptLong_0, CType(0, Short))
		Me.tmrSilentChannel.SetIndex(_tmrSilentChannel_1, CType(1, Short))
		Me.tmrSilentChannel.SetIndex(_tmrSilentChannel_0, CType(0, Short))
		CType(Me.tmrSilentChannel, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.tmrScriptLong, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.tmrScript, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.sckScript, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mnuScriptingDash, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mnuPublicChannels, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mnuListViewButton, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mnuCustomChannels, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.itcScript, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.sckBNet, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.sckBNLS, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.INet, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.sckMCP, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._sckScript_0, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._itcScript_0, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SCRestricted, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SControl, System.ComponentModel.ISupportInitialize).EndInit()
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuBot, Me.mnuSetTop, Me.mnuConnect, Me.mnuDisconnect, Me.mnuWindow, Me.mnuTray, Me.mnuPopup, Me.mnuScripting, Me.mnuHelp, Me.mnuShortcuts, Me.mnuPopClanList, Me.mnuPopFList})
		mnuBot.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuConnect2, Me.mnuDisconnect2, Me.mnuSepT, Me.mnuUsers, Me.mnuCommandManager, Me.mnuScriptManager, Me.mnuSepTabcd, Me.mnuGetNews, Me.mnuUpdateVerbytes, Me.mnuSepZ, Me.mnuRealmSwitch, Me.mnuIgnoreInvites, Me.mnuSep1, Me.mnuQCTop, Me.mnuSep5, Me.mnuEditCaught, Me.mnuFiles, Me.mnuSettingsRepair, Me.mnuSep2, Me.mnuExit})
		mnuQCTop.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuHomeChannel, Me.mnuLastChannel, Me.mnuQCDash, Me.mnuQCHeader, Me._mnuCustomChannels_0, Me._mnuCustomChannels_1, Me._mnuCustomChannels_2, Me._mnuCustomChannels_3, Me._mnuCustomChannels_4, Me._mnuCustomChannels_5, Me._mnuCustomChannels_6, Me._mnuCustomChannels_7, Me._mnuCustomChannels_8, Me.mnuCustomChannelAdd, Me.mnuCustomChannelEdit, Me.mnuPCDash, Me.mnuPCHeader, Me._mnuPublicChannels_0})
		mnuFiles.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuOpenBotFolder, Me.mnuSepA, Me.mnuClearedTxt, Me.mnuWhisperCleared})
		mnuSettingsRepair.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuToolsMenuWarning, Me.mnuRepairDataFiles, Me.mnuRepairVerbytes, Me.mnuRepairCleanMail, Me.mnuPacketLog})
		mnuSetTop.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuSetup, Me.mnuUTF8, Me.mnuLogging, Me.mnuSep4, Me.mnuProfile, Me.mnuFilters, Me.mnuCatchPhrases, Me.mnuSep6, Me.mnuReload})
		mnuLogging.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuLog0, Me.mnuLog1, Me.mnuLog2})
		mnuWindow.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuToggle, Me.mnuHideBans, Me.mnuLock, Me.mnuToggleFilters, Me.mnuSep7, Me.mnuToggleWWUse, Me.mnuSP3, Me.mnuToggleShowOutgoing, Me.mnuHideWhispersInrtbChat, Me.mnuSepC, Me.mnuClear, Me.mnuClearWW, Me.mnuSepD, Me.mnuFlash, Me.mnuDisableVoidView, Me.mnuRecordWindowPos})
		mnuTray.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuTrayCaption, Me.mnuRestore, Me.mnuTraySep, Me.mnuTrayExit})
		mnuPopup.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuPopWhisper, Me.mnuPopCopy, Me.mnuPopAddLeft, Me.mnuPopAddToFList, Me.mnuPopInvite, Me.mnuPopShitlist, Me.mnuPopSafelist, Me.mnuPopSep1, Me.mnuPopUserlistWhois, Me.mnuPopWhois, Me.mnuPopStats, Me.mnuPopProfile, Me.mnuPopWebProfile, Me.mnuPopSep2, Me.mnuPopKick, Me.mnuPopBan, Me.mnuPopSquelch, Me.mnuPopUnsquelch, Me.mnuPopDes})
		mnuScripting.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuReloadScripts, Me.mnuOpenScriptFolder, Me._mnuScriptingDash_0})
		mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuAbout, Me.mnuHelpReadme, Me.mnuHelpWebsite, Me.mnuTerms, Me.mnuChangeLog})
		mnuShortcuts.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuItalic, Me.mnuBold, Me.mnuUnderline, Me._mnuListViewButton_0, Me._mnuListViewButton_1, Me._mnuListViewButton_2})
		mnuPopClanList.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuPopClanWhisper, Me.mnuPopClanCopy, Me.mnuPopClanAddLeft, Me.mnuPopClanAddToFList, Me.mnuPopClanSep1, Me.mnuPopClanUserlistWhois, Me.mnuPopClanWhois, Me.mnuPopClanStats, Me.mnuPopClanProfile, Me.mnuPopClanWebProfile, Me.mnuPopClanSep2, Me.mnuPopClanPromote, Me.mnuPopClanDemote, Me.mnuPopClanRemove, Me.mnuPopClanLeave, Me.mnuPopClanMakeChief, Me.mnuPopClanDisband})
		mnuPopClanStats.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuPopClanStatsWAR3, Me.mnuPopClanStatsW3XP})
		mnuPopClanWebProfile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuPopClanWebProfileWAR3, Me.mnuPopClanWebProfileW3XP})
		mnuPopFList.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuPopFLWhisper, Me.mnuPopFLCopy, Me.mnuPopFLAddLeft, Me.mnuPopFLInvite, Me.mnuPopFLSep1, Me.mnuPopFLUserlistWhois, Me.mnuPopFLWhois, Me.mnuPopFLStats, Me.mnuPopFLProfile, Me.mnuPopFLWebProfile, Me.mnuPopFLSep2, Me.mnuPopFLPromote, Me.mnuPopFLDemote, Me.mnuPopFLRemove, Me.mnuPopFLSep3, Me.mnuPopFLRefresh})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.lvChannel.ResumeLayout(False)
		Me.ListviewTabs.ResumeLayout(False)
		Me.lvClanList.ResumeLayout(False)
		Me.lvFriendList.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class