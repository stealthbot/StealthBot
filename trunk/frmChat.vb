Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmChat
	Inherits System.Windows.Forms.Form
	'StealthBot 11/5/02-Present
	'Source Code Version: 2.7RC2
	
	'REVISION #1337!
	
	'Classes
	Public WithEvents ClanHandler As clsClanPacketHandler
	Public WithEvents FriendListHandler As clsFriendlistHandler
	
	'Variables
	Private m_lCurItemIndex As Integer
	Private MultiLinePaste As Boolean
	Private doAuth As Boolean
	Private AUTH_CHECKED As Boolean
	
	'Forms
	Public SettingsForm As frmSettings
	Public ChatNoScroll As Boolean
	
	Private Const HKEY_LOCAL_MACHINE As Integer = &H80000002
	Private Const REG_SZ As Short = 1
	Private Const WM_USER As Short = 1024
	Private Const CB_LIMITTEXT As Integer = &H141
	Private Const SB_BOTTOM As Short = 7
	Private Const EM_SCROLL As Integer = &HB5
	
	Private Structure sockaddr_in
		Dim sin_family As Short
		Dim sin_port As Short
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(4),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=4)> Public sin_addr() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(8),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=8)> Public sin_zero() As Char
	End Structure
	
	Private Const SB_INET_UNSET As String = vbNullString
	Private Const SB_INET_NEWS1 As String = "SBNEWS_AUTOCONNECT"
	Private Const SB_INET_NEWS As String = "SBNEWS"
	Private Const SB_INET_BNLS1 As String = "BNLSFINDER"
	Private Const SB_INET_BNLS2 As String = "BNLSFINDER_DEFAULT"
	Private Const SB_INET_VBYTE As String = "VERBYTE"
	Private Const SB_INET_BETA As String = "AUTHBETA"
	
	' LET IT BEGIN
	Private Sub frmChat_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim s As String
		Dim f As Short
		Dim L As Integer
		Dim FrmSplashInUse As Boolean
		Dim strBeta As String
		Dim sStr() As String
		
		' COMPILER FLAGS
#If (BETA = 1) Then
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CVERSION = StringFormat("StealthBot Beta v{0}.{1} - Build {2}", My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Revision)
#Else
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression Else did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		CVERSION = StringFormat("StealthBot v{0}.{1}.{2}", App.Major, App.Minor, App.REVISION)
#End If
		
#If (COMPILE_DEBUG = 1) Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression (COMPILE_DEBUG = 1) did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		CVERSION = StringFormat("{0} - DEBUG", CVERSION)
#End If
		
#If (COMPILE_CRC = 1) Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression (COMPILE_CRC = 1) did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		Dim crc As New clsCRC32
		If (Not crc.ValidateExecutable) Then
		MsgBox GetHexProtectionMessage, vbOKOnly + vbCritical
		Call Form_Unload(0)
		Unload frmChat
		Exit Sub
		End If
		Set crc = Nothing
#End If
		
#If (COMPILE_DEBUG = 0) Then
		HookWindowProc(Me.Handle.ToInt32)
#End If
		
		SendMessage(Me.cboSend.Handle.ToInt32, CB_LIMITTEXT, 0, 0)
		
		colWhisperWindows = New Collection
		colLastSeen = New Collection
		GErrorHandler = New clsErrorHandler
		BotVars = New clsBotVars
		
		'UPGRADE_WARNING: Couldn't resolve default property of object SetCommandLine(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sStr = SetCommandLine(VB.Command())
		
		' EVERYTHING ELSE
		rtbWhispers.Visible = False 'default
		rtbWhispersVisible = False
		
		'Set dictTimerInterval = New Dictionary
		'Set dictTimerEnabled = New Dictionary
		'Set dictTimerCount = New Dictionary
		'dictTimerInterval.CompareMode = TextCompare
		'dictTimerEnabled.CompareMode = TextCompare
		'dictTimerCount.CompareMode = TextCompare
		
		With mnuTrayCaption
			.Text = CVERSION
			.Enabled = False
		End With
		
		mail = True
		f = FreeFile
		
		With rtbChat
			.Font = VB6.FontChangeSize(.Font, 8)
			.SelectionTabs(0) = 15 * VB6.TwipsPerPixelX
			.SelectionHangingIndent = VB6.PixelsToTwipsX(.SelectionTabs(0))
		End With
		
		With rtbWhispers
			.Font = VB6.FontChangeSize(.Font, 8)
			.SelectionTabs(0) = 15 * VB6.TwipsPerPixelX
			.SelectionHangingIndent = VB6.PixelsToTwipsX(.SelectionTabs(0))
		End With
		
		lvChannel.View = System.Windows.Forms.View.Details
		lvChannel.LargeImageList = imlIcons
		lvClanList.View = System.Windows.Forms.View.Details
		lvClanList.LargeImageList = imlIcons
		
		ReDim Phrases(0)
		'UPGRADE_NOTE: Catch was upgraded to Catch_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		ReDim Catch_Renamed(0)
		ReDim gBans(0)
		ReDim gOutFilters(0)
		ReDim gFilters(0)
		
		Call BuildProductInfo()
		
		Config = New clsConfig
		Config.Load(GetConfigFilePath())
		
		' SPLASH SCREEN
		If Config.ShowSplashScreen Then
			frmSplash.Show()
			FrmSplashInUse = True
		End If
		
		If Config.ShowWhisperBox Then
			If Not rtbWhispersVisible Then Call cmdShowHide_Click(cmdShowHide, New System.EventArgs())
		Else
			If rtbWhispersVisible Then Call cmdShowHide_Click(cmdShowHide, New System.EventArgs())
		End If
		
		If Config.PositionHeight > 0 Then
			L = (IIf(CInt(Config.PositionHeight) < 200, 200, CInt(Config.PositionHeight)) * VB6.TwipsPerPixelY)
			
			If (rtbWhispersVisible) Then
				L = L - (VB6.PixelsToTwipsY(rtbWhispers.Height) / VB6.TwipsPerPixelY)
			End If
			
			Me.Height = VB6.TwipsToPixelsY(L)
		End If
		
		If Config.PositionWidth > 0 Then
			Me.Width = VB6.TwipsToPixelsX((IIf(CInt(Config.PositionWidth) < 300, 300, CInt(Config.PositionWidth)) * VB6.TwipsPerPixelX))
		End If
		
		'Set window position
		Me.Left = VB6.TwipsToPixelsX(CInt(Config.PositionLeft) * VB6.TwipsPerPixelX)
		Me.Top = VB6.TwipsToPixelsY(CInt(Config.PositionTop) * VB6.TwipsPerPixelY)
		
		'Make sure the window is on the screen
		If Config.EnforceScreenBounds Then
			If Config.MonitorCount <> GetMonitorCount Then
				If (VB6.PixelsToTwipsX(Me.Left) > (VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width))) Then
					Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)))
				End If
				
				If (VB6.PixelsToTwipsY(Me.Top) > (VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height))) Then
					Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)))
				End If
				
				Config.MonitorCount = GetMonitorCount
			End If
		End If
		
		'Support for recording maxmized position. - FrOzeN
		If Config.IsMaximized Then
			Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		End If
		
		ClanHandler = New clsClanPacketHandler
		FriendListHandler = New clsFriendlistHandler
		ListToolTip = New clsCTooltip
		
		Call ReloadConfig()
		
		Call frmChat_Resize(Me, New System.EventArgs())
		
		Call GetColorLists()
		Call InitListviewTabs()
		Call DisableListviewTabs()
		
		ListviewTabs.SelectedIndex = 0
		
		With ListToolTip
			.Style = clsCTooltip.ttStyleEnum.TTStandard
			.Icon = clsCTooltip.ttIconType.TTNoIcon
			.DelayTime = 100
		End With
		
		Call ClearChannel()
		
		With lvClanList
			.View = System.Windows.Forms.View.Details
			.SmallImageList = imlClan
			'UPGRADE_WARNING: Lower bound of collection lvClanList.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			.Columns.Item(1).Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(.Width) \ 4) * 3 - 150)
			'UPGRADE_WARNING: Lower bound of collection lvClanList.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			.Columns.Item(2).Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(.Width) \ 4 + 200)
			'UPGRADE_WARNING: Lower bound of collection lvClanList.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			.Columns.Item(3).Width = 0
		End With
		
		Me.KeyPreview = True
		SetTitle("Disconnected")
		
		Me.UpdateTrayTooltip()
		
		Me.Show()
		Me.Refresh()
		'UPGRADE_ISSUE: Form property frmChat.AutoRedraw was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        Me.DoubleBuffered = True
		
		AddChat(RTBColors.ConsoleText, "-> Welcome to " & CVERSION & ", by Stealth.")
		AddChat(RTBColors.ConsoleText, "-> If you enjoy StealthBot, consider supporting its development at http://donate.stealthbot.net")
		
		Dim x As Short
		For x = LBound(sStr) To UBound(sStr)
            If (Len(sStr(x)) > 0) Then
                AddChat(RTBColors.InformationText, sStr(x))
            End If
		Next x
		
		On Error Resume Next
		
		VoteDuration = -1
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If (Len(Dir(GetConfigFilePath())) = 0) Then
            AddChat(RTBColors.ServerInfoText, "If you're new to bots, start by choosing 'Bot Settings' " & "under the 'Settings' menu above.")
            AddChat(RTBColors.ServerInfoText, "For more help, click the 'Step-By-Step Configuration' " & "button inside Settings.")
            AddChat(RTBColors.ServerInfoText, "For more information and a list of commands, see the " & "Readme by clicking 'Readme' under the 'Help' menu.")
            AddChat(RTBColors.ServerInfoText, "Please note that any usage of this program is subject to " & "the terms of the End-User License Agreement available at http://eula.stealthbot.net.")
        End If
		
		
		Randomize()
		
		ID_TASKBARICON = (Rnd() * 100)
		
		TASKBARCREATED_MSGID = RegisterWindowMessage("TaskbarCreated")
		
		cboSend.Focus()
		
		LoadQuickChannels()
		InitScriptControl(SControl)
		
		On Error Resume Next
		'News call and scripting events
		
		LoadScripts()
		'InitMenus
		InitScripts()
		
		If FrmSplashInUse Then frmSplash.Activate()
		
		If Not MDebug("debug") Then
			mnuRecordWindowPos.Visible = False
		End If
		
		'Update the config if it's an old version
		If Config.Version < CDbl(CONFIG_VERSION) Then
			Call Config.Save()
		End If
		
#If COMPILE_DEBUG = 0 Then
		If Config.MinimizeOnStartup Then
			Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
			Call frmChat_Resize(Me, New System.EventArgs())
		End If
#End If
		
		If Not Config.DisableNews Then
			Call RequestINetPage(GetNewsURL(), SB_INET_NEWS1, True)
		ElseIf Config.AutoConnect Then 
			Call DoConnect()
		End If
		
		'Now loads scripts when the bot opens, instead of after connecting. - FrOzeN
		'RunInAll "Event_Load"
		
		'Dim I As Integer
		'Dim tmp As String
		'Dim str As String
		
		'str = "flood"
		
		'For I = 1 To Len(str)
		'   tmp = tmp & Hex(Asc(Mid(str, I, 1)))
		'Next I
		
		'    BotVars.UseProxy = True
		'    BotVars.ProxyIP = "213.210.194.139"
		'    BotVars.ProxyPort = 1080
		'BotVars.ProxyIsSocks5 = True
		
		'UPGRADE_WARNING: Lower bound of collection lvFriendList.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		lvFriendList.Columns.Item(2).Width = VB6.TwipsToPixelsX(imlIcons.ImageSize.Width)
		'UPGRADE_WARNING: Lower bound of collection lvClanList.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		lvClanList.Columns.Item(2).Width = VB6.TwipsToPixelsX(imlClan.ImageSize.Width)
	End Sub
	
	Public Sub cacheTimer_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cacheTimer.Tick
		' this code updated 7/23/05 in Chihuahua, Chihuahua, MX
		Dim strArray() As String
		Dim ret As String
		Dim lPos As Integer
		Dim y As String
		Dim c, n As Short
		If (Caching) Then ' time to retrieve stored information and squelch or ban a channel
			
			Caching = False
			
			' Changed 08-18-09 - Hdx - Uses the new Channel cache function, Eventually to beremoved to script
			'ret = CacheChannelList(vbNullString, 0, Y)
			ret = CacheChannelList(modCommandsOps.CacheChanneListEnum.enRetrieve, y)
			
			lPos = InStr(1, ret, ", ", CompareMethod.Binary)
			
			If (lPos) Then
				strArray = Split(ret, ", ")
			Else
				ReDim Preserve strArray(0)
				strArray(0) = ret
			End If
			
			For c = 0 To UBound(strArray)
				' [CHANNELOP]  -  [*CHANNELOP]  -  [CHARACTER@USEast (*CHANNELOP)]
				If StrComp(UCase(strArray(c)), strArray(c), CompareMethod.Binary) = 0 Then
					If VB.Left(strArray(c), 1) = "[" And VB.Right(strArray(c), 1) = "]" Then
						strArray(c) = Mid(strArray(c), 2, Len(strArray(c)) - 2)
					End If
				End If
				
				strArray(c) = ConvertUsername(CleanUsername(strArray(c)))
				
				'AddChat vbRed, strArray(C)
				
				If Len(strArray(c)) > 1 Then
					If InStr(y, "ban") Then
						If (g_Channel.Self.IsOperator) Then
							Ban(strArray(c), AutoModSafelistValue - 1, 0)
						End If
					Else
						If (GetSafelist(strArray(c)) = False) Then
							AddQ("/squelch " & strArray(c))
						End If
					End If
				End If
			Next c
		End If
		
		cacheTimer.Enabled = False
	End Sub
	
	Private Sub ChatQueueTimer_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ChatQueueTimer.Tick
		modChatQueue.ChatQueueTimerProc()
	End Sub
	
	Private Sub frmChat_GotFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.GotFocus
		On Error GoTo ERROR_HANDLER
		
		If (cboSendHadFocus) Then
			cboSend.Focus()
		End If
		
		Exit Sub
		
ERROR_HANDLER: 
		AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in Form_GotFocus().")
		
		Exit Sub
	End Sub
	
	' asynchronous INet
	Private Function RequestINetPage(ByVal URL As String, ByVal Request As String, ByVal CancelStillExecuting As Boolean) As Boolean
		On Error GoTo ERROR_HANDLER
		
		Dim ret As String
		With INet
			If .StillExecuting Then
				If CancelStillExecuting Then
					'UPGRADE_ISSUE: VBControlExtender property INet.Cancel was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					.Cancel()
				Else
					RequestINetPage = False
					
					Exit Function
				End If
			End If
			
			.RequestTimeout = 5
			.Tag = Request
            .URL = URL
            .Execute()
			
			RequestINetPage = True
		End With
		
		Exit Function
		
ERROR_HANDLER: 
		AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in RequestINetPage().")
		
		RequestINetPage = False
		
		Exit Function
	End Function
	
	' asynchronous INet response
	Private Sub INet_StateChanged(ByVal eventSender As System.Object, ByVal eventArgs As AxInetCtlsObjects.DInetEvents_StateChangedEvent) Handles INet.StateChanged
		Dim strData As String
		Dim Buffer As String
		
		Select Case eventArgs.State
			Case InetCtlsObjects.StateConstants.icResponseCompleted, InetCtlsObjects.StateConstants.icError
				If INet.ResponseCode >= 1000 Then
					Buffer = "INet Error #" & INet.ResponseCode & ": " & INet.ResponseInfo
				ElseIf INet.ResponseCode <> 0 Then 
					Buffer = "HTTP Error " & INet.ResponseCode & " " & INet.ResponseInfo
				Else
					Do 
						'UPGRADE_WARNING: Couldn't resolve default property of object INet.GetChunk(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						strData = INet.GetChunk(1024, InetCtlsObjects.DataTypeConstants.icString)
						If Len(strData) = 0 Then Exit Do
						Buffer = Buffer & strData
					Loop 
					
                    If Len(Buffer) = 0 Then
                        Buffer = "Empty response"
                    End If
				End If
				
				Select Case INet.Tag
					Case SB_INET_NEWS
						Call HandleNews(Buffer, INet.ResponseCode)
					Case SB_INET_NEWS1
						Call HandleNews(Buffer, INet.ResponseCode)
						If Config.AutoConnect And Not g_Connected Then
							Call DoConnect()
						End If
					Case SB_INET_VBYTE
						Call HandleUpdateVerbyte(Buffer, INet.ResponseCode)
					Case SB_INET_BNLS1
						Call HandleFindBNLSServerListResult(Buffer, INet.ResponseCode, True)
					Case SB_INET_BNLS2
						Call HandleFindBNLSServerListResult(Buffer, INet.ResponseCode, False)
				End Select
				
				INet.Tag = SB_INET_UNSET
				'UPGRADE_ISSUE: VBControlExtender property INet.Cancel was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				INet.Cancel()
		End Select
	End Sub
	
	Private Sub HandleUpdateVerbyte(ByVal Buffer As String, ByVal ResponseCode As Integer)
		On Error Resume Next
		
		Dim ary() As String
		Dim i As Short
		
		Dim keys(3) As String
		If INet.ResponseCode <> 0 Then
			AddChat(RTBColors.ErrorMessageText, Buffer & ". Error retrieving version bytes from http://www.stealthbot.net. Please visit it for instructions.")
		ElseIf Len(Buffer) <> 11 Then 
			AddChat(RTBColors.ErrorMessageText, "Format not understood. Error retrieving version bytes from http://www.stealthbot.net. Please visit it for instructions.")
		Else
			'W2 SC D2 W3
			
			keys(0) = "W2"
			keys(1) = "SC"
			keys(2) = "D2"
			keys(3) = "W3"
			
			ary = Split(Buffer, " ")
			
			For i = 0 To 3
				Config.SetVersionByte(keys(i), CInt(Val("&H" & ary(i))))
			Next i
			Config.SetVersionByte("D2X", CInt(Val("&H" & ary(2))))
			
			Call Config.Save()
			
			AddChat(RTBColors.SuccessText, "Your config.ini file has been loaded with current version bytes.")
		End If
	End Sub
	
	
	'RTB ADDCHAT SUBROUTINE - originally written by Grok[vL] - modified to support
	'                         logging and timestamps and color decoding
	' Updated 7/23/05 to remove many bulky calls to Len()
	' Updated 9/01/05 to remove the changes made on 7/23/05 *smack forehead*
	' Updated 9/25/05-10/25/05 to add HTML logging
	' Updated 1/3/06 to remove HTML logging
	' Updated 8/4/06 to add scrollbar locking (thanks FrOzeN)
	' Updated 11/8/06 to log incoming text immediately
	' Updated 4/17/07 to not flash the desktop when the scrollbar is held up
	' Updated 8/07/07 with greater precision
	'UPGRADE_WARNING: ParamArray saElements was changed from ByRef to ByVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93C6A0DC-8C99-429A-8696-35FC4DCEFCCC"'
	Sub AddChat(ParamArray ByVal saElements() As Object)
		Dim arr() As Object
		Dim i As Short
		
		arr = VB6.CopyArray(saElements)
		Call DisplayRichText((Me.rtbChat), arr)
	End Sub
	
	'UPGRADE_WARNING: ParamArray saElements was changed from ByRef to ByVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93C6A0DC-8C99-429A-8696-35FC4DCEFCCC"'
	Sub AddWhisper(ParamArray ByVal saElements() As Object)
		Dim arr() As Object
		
		arr = VB6.CopyArray(saElements)
		
		Call DisplayRichText((Me.rtbWhispers), arr)
		Exit Sub
		
		
		Dim s As String
		Dim L As Integer
		Dim i As Short
		
		If Not BotVars.LockChat Then
			'If ((BotVars.MaxBacklogSize) And (Len(rtbWhispers.text) >= BotVars.MaxBacklogSize)) Then
			If BotVars.Logging < 2 Then
				FileClose(1)
				FileOpen(1, (StringFormat("{0}{1}-WHISPERS.txt", GetFolderPath("Logs"), VB6.Format(Today, "YYYY-MM-DD"))), OpenMode.Append)
			End If
			
			With rtbWhispers
				.Visible = False
				.SelectionStart = 0
				.SelectionLength = InStr(1, .Text, vbLf, CompareMethod.Binary)
				If BotVars.Logging < 2 Then PrintLine(1, VB.Left(vbCrLf, -2 * CInt((i + 1) = UBound(saElements))))
				.SelectedText = vbNullString
				.Visible = True
			End With
			
			FileClose(1)
			'End If
			
			Select Case BotVars.TSSetting
				Case 0 : s = " [" & TimeOfDay & "] "
				Case 1 : s = " [" & VB6.Format(TimeOfDay, "HH:MM:SS") & "] "
				Case 2 : s = " [" & VB6.Format(TimeOfDay, "HH:MM:SS") & "." & GetCurrentMS & "] "
				Case 3 : s = vbNullString
			End Select
			
			With rtbWhispers
				.SelectionStart = Len(.Text)
				.SelectionLength = 0
				.SelectionColor = System.Drawing.ColorTranslator.FromOle(RTBColors.TimeStamps)
				If .SelectionFont.Bold = True Then .Font = VB6.FontChangeBold(.SelectionFont, False)
				If .SelectionFont.Italic = True Then .SelectionFont = VB6.FontChangeItalic(.SelectionFont, False)
				.SelectedText = s
				.SelectionStart = Len(.Text)
			End With
			
			For i = LBound(saElements) To UBound(saElements) Step 2
				'UPGRADE_WARNING: Couldn't resolve default property of object saElements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If InStr(1, saElements(i), Chr(0), CompareMethod.Binary) > 0 Then KillNull(saElements(i))
				
				If Len(saElements(i + 1)) > 0 Then
					With rtbWhispers
						.SelectionStart = Len(.Text)
						L = .SelectionStart
						.SelectionLength = 0
						'UPGRADE_WARNING: Couldn't resolve default property of object saElements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.SelectionColor = System.Drawing.ColorTranslator.FromOle(saElements(i))
						'UPGRADE_WARNING: Couldn't resolve default property of object saElements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.SelectedText = saElements(i + 1) & VB.Left(vbCrLf, -2 * CInt((i + 1) = UBound(saElements)))
						.SelectionStart = Len(.Text)
					End With
				End If
			Next i
			
			Call ColorModify(rtbWhispers, L)
		End If
	End Sub
	
	
	'BNLS EVENTS
	Sub Event_BNetConnected()
		If (BotVars.UseProxy) Then
			AddChat(RTBColors.SuccessText, "[PROXY] Connected!")
		Else
			AddChat(RTBColors.SuccessText, "[BNCS] Connected!")
		End If
		
		Call SetNagelStatus(sckBNet.SocketHandle, False)
	End Sub
	
	Sub Event_BNetConnecting()
		If BotVars.UseProxy Then
			AddChat(RTBColors.InformationText, "[PROXY] Connecting to the Battle.net server at " & BotVars.Server & "...")
		Else
			AddChat(RTBColors.InformationText, "[BNCS] Connecting to the Battle.net server at " & BotVars.Server & "...")
		End If
	End Sub
	
	Sub Event_BNetDisconnected()
		'UPGRADE_WARNING: Timer property tmrIdleTimer.Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
		tmrIdleTimer.Interval = 0
		'UPGRADE_WARNING: Timer property UpTimer.Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
		UpTimer.Interval = 0
		BotVars.JoinWatch = 0
		
		AddChat(RTBColors.ErrorMessageText, IIf(BotVars.UseProxy And BotVars.ProxyStatus <> modEnum.enuProxyStatus.psOnline, "[PROXY] ", "[BNCS] ") & "Disconnected.")
		
		DoDisconnect((1))
		
		SetTitle("Disconnected")
		
		UpdateTrayTooltip()
		
		g_Online = False
		
		Call ClearChannel()
		
		UpdateProxyStatus(modEnum.enuProxyStatus.psNotConnected)
		'AddChat RTBColors.ErrorMessageText, "[BNCS] Attempting to reconnect, please wait..."
		'AddChat RTBColors.SuccessText, "Connection initialized."
		
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If sckBNet.CtlState <> 0 Then sckBNet.Close()
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If sckBNLS.CtlState <> 0 Then sckBNLS.Close()
		
		BNLSAuthorized = False
		
		'If Not UserCancelledConnect Then
		'    ReconnectTimerID = SetTimer(0, 0, BotVars.ReconnectDelay, _
		''        AddressOf Reconnect_TimerProc)
		'End If
	End Sub
	
	Sub Event_BNetError(ByRef ErrorNumber As Short, ByRef Description As String)
		Dim s As String
		
		If BotVars.UseProxy And BotVars.ProxyStatus <> modEnum.enuProxyStatus.psOnline Then
			s = "[PROXY] "
		Else
			s = "[BNCS] "
		End If
		
		AddChat(RTBColors.ErrorMessageText, s & ErrorNumber & " -- " & Description)
		AddChat(RTBColors.ErrorMessageText, s & "Disconnected.")
		
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If (sckBNet.CtlState <> 0) Then
			Call sckBNet.Close()
		End If
		
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If (sckBNLS.CtlState <> 0) Then
			Call sckBNLS.Close()
		End If
		
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If (sckMCP.CtlState <> 0) Then
			Call sckMCP.Close()
		End If
		
		g_Connected = False
		
		UserCancelledConnect = False
		
		DoDisconnect(1, True)
		
		SetTitle("Disconnected")
		
		Me.UpdateTrayTooltip()
		
		Call ClearChannel()
		lvClanList.Items.Clear()
		lvFriendList.Items.Clear()
		
		lblCurrentChannel.Text = GetChannelString
		
		' NOV 18 04 Change here should fix the attention-grabbing on errors
		'If Me.WindowState <> vbMinimized Then cboSend.SetFocus
		
		
		If DisplayError(ErrorNumber, IIf(BotVars.UseProxy And BotVars.ProxyStatus <> modEnum.enuProxyStatus.psOnline, 2, 1), modEnum.enuErrorSources.BNET) = True Then
			AddChat(RTBColors.ErrorMessageText, IIf(BotVars.UseProxy And BotVars.ProxyStatus <> modEnum.enuProxyStatus.psOnline, "[PROXY] ", "[BNCS] ") & "Attempting to reconnect in " & (BotVars.ReconnectDelay / 1000) & IIf(((BotVars.ReconnectDelay / 1000) > 1), " seconds", " second") & "...")
			
			UserCancelledConnect = False 'this should fix the beta reconnect problems
			
			'ReconnectTimerID = SetTimer(0, 0, BotVars.ReconnectDelay, _
			''    AddressOf Reconnect_TimerProc)
			
			'ExReconTicks = 0
			'ExReconMinutes = BotVars.ReconnectDelay / 1000
			'ExReconnectTimerID = SetTimer(0, ExReconnectTimerID, _
			''    1000, AddressOf ExtendedReconnect_TimerProc)
		End If
	End Sub
	
	Sub Event_BNLSAuthEvent(ByRef Success As Boolean)
		If Success = True Then
			AddChat(RTBColors.SuccessText, "[BNLS] Authorized!")
		Else
			AddChat(RTBColors.ErrorMessageText, "[BNLS] Authorization failed! Please download the latest version of StealthBot from http://www.stealthbot.net.")
			Call DoDisconnect()
		End If
	End Sub
	
	Sub Event_BNLSConnected()
		AddChat(RTBColors.SuccessText, "[BNLS] Connected!")
		
		Call SetNagelStatus(sckBNLS.SocketHandle, False)
	End Sub
	
	Sub Event_BNLSConnecting()
		AddChat(RTBColors.InformationText, "[BNLS] Connecting to the BNLS server at " & BotVars.BNLSServer & "...")
	End Sub
	
	Sub Event_BNLSDataError(ByRef Message As Byte)
		If Message = 0 Then
			AddChat(RTBColors.ErrorMessageText, "[BNLS] Your CD-Key was rejected. It may be invalid. Try connecting again.")
		ElseIf Message = 1 Then 
			AddChat(RTBColors.ErrorMessageText, "[BNLS] Error! Your CD-Key is bad.")
		ElseIf Message = 2 Then 
			AddChat(RTBColors.ErrorMessageText, "[BNLS] Error! BNLS has failed CheckRevision. Please check your bot's settings and try again.")
			AddChat(RTBColors.ErrorMessageText, "[BNLS] Product: " & StrReverse(BotVars.Product) & ".")
		ElseIf Message = 3 Then 
			AddChat(RTBColors.ErrorMessageText, "[BNLS] Error! Bad NLS revision.")
		End If
	End Sub
	
	Private Sub Event_BNLSError(ByRef ErrorNumber As Short, ByRef Description As String)
		If Not HandleBnlsError("[BNLS] Error " & ErrorNumber & ": " & Description) Then
			' if we aren't using the finder display the error
			DisplayError(ErrorNumber, 0, modEnum.enuErrorSources.BNLS)
		End If
	End Sub
	
	
	' this function will return whether we are going to use the finder
	Public Function HandleBnlsError(ByVal ErrorMessage As String) As Boolean
		HandleBnlsError = False
		
		sckBNet.Close()
		
		' Is the BNLS server finder enabled?
		If Config.BNLSFinder Then
			Call RotateBnlsServer()
		Else
			AddChat(RTBColors.ErrorMessageText, ErrorMessage)
			
			UserCancelledConnect = False
			DoDisconnect(1, True)
		End If
		
		' return the BotVars
		HandleBnlsError = Config.BNLSFinder
	End Function
	
	' Moves the connection to the next available BNLS server
	Public Sub RotateBnlsServer()
		'Close the current BNLS connection
		sckBNLS.Close()
		
		'Notify user the current BNLS server failed
		AddChat(RTBColors.ErrorMessageText, "[BNLS] Connection to " & BotVars.BNLSServer & " failed.")
		
		'Notify user other BNLS servers are being located
		AddChat(RTBColors.InformationText, "[BNLS] Locating other BNLS servers...")
		
		Call FindBNLSServer()
	End Sub
	
	Public Sub HandleFindBNLSServerListResult(ByVal strReturn As String, ByVal Result As Short, ByVal ConfigListSource As Boolean)
		' convert to LF
		strReturn = Replace(strReturn, vbCr, vbLf)
		strReturn = Replace(strReturn, vbLf & vbLf, vbLf)
		
		If (INet.ResponseCode <> 0) Or (VB.Right(strReturn, 1) <> vbLf) Then
			If ConfigListSource Then
				If Not RequestINetPage(BNLS_DEFAULT_SOURCE, SB_INET_BNLS2, True) Then
					Call HandleFindBNLSServerListResult("INet is busy", -1, False)
				End If
			Else
				AddChat(RTBColors.ErrorMessageText, "[BNLS] " & strReturn & ". Unable to use BNLS server finder.")
				AddChat(RTBColors.ErrorMessageText, "[BNLS] An error occured while trying to locate an alternative BNLS server.")
				AddChat(RTBColors.ErrorMessageText, "[BNLS]   You may not be connected to the internet or may be having DNS resolution issues.")
				AddChat(RTBColors.ErrorMessageText, "[BNLS]   Visit http://www.stealthbot.net/ and check the Technical Support forum for more information.")
				DoDisconnect()
				
				' ensure that we update our listing on following connection(s)
				BNLSFinderGotList = False
				
				' ensure checker starts at 0 again on following connection(s)
				BNLSFinderIndex = 0
			End If
			
			Exit Sub
		Else
			' Split the page up into an array of servers.
			BNLSFinderEntries = Split(strReturn, vbLf)
		End If
		
		'Mark GotBNLSList as True so it's no longer downloaded for each attempt
		BNLSFinderGotList = True
		
		Call FindBNLSServerEntry()
	End Sub
	
	'Locates alternative BNLS servers for the bot to use if the current one fails
	Public Sub FindBNLSServer()
		'Error handler
		On Error GoTo ERROR_HANDLER
		
		BNLSFinderIndex = BNLSFinderIndex + 1
		
		'Check if the BNLS list has been downloaded
		If (BNLSFinderGotList = False) Then
			'Reset the counter
			BNLSFinderIndex = 0
			
			' store first bnls server used so that we can avoid connecting to it again
			BNLSFinderLatest = BotVars.BNLSServer
			
			'Get the servers as a list from http://stealthbot.net/p/bnls.php
            If (Len(Config.BNLSFinderSource) > 0) Then
                If Not RequestINetPage(Config.BNLSFinderSource, SB_INET_BNLS1, True) Then
                    Call HandleFindBNLSServerListResult("INet is busy", -1, False)
                End If
            Else
                If Not RequestINetPage(BNLS_DEFAULT_SOURCE, SB_INET_BNLS2, True) Then
                    Call HandleFindBNLSServerListResult("INet is busy", -1, False)
                End If
            End If
			
			Exit Sub
		End If
		
		Call FindBNLSServerEntry()
		
		Exit Sub
		
ERROR_HANDLER: 
		
		'Display the error message to the user
		If Err.Number = ERROR_FINDBNLSSERVER Then
			AddChat(RTBColors.ErrorMessageText, "[BNLS] " & Err.Description)
			AddChat(RTBColors.ErrorMessageText, "[BNLS]   Visit http://www.stealthbot.net/ and check the Technical Support forum for more information.")
			DoDisconnect()
			
			' ensure that we update our listing on following connection(s)
			BNLSFinderGotList = False
			
			' ensure checker starts at 0 again on following connection(s)
			BNLSFinderIndex = 0
			
		Else
			Resume Next
		End If
		
		Exit Sub
	End Sub
	
	Sub FindBNLSServerEntry()
		If BNLSFinderIndex > UBound(BNLSFinderEntries) Then
			'All BNLS servers have been tried and failed
			Err.Raise(ERROR_FINDBNLSSERVER,  , "All the BNLS servers have failed.")
		End If
		
		' keep increasing counter until we find a server that is valid and isn't the same as the first one
        Do While (StrComp(BNLSFinderEntries(BNLSFinderIndex), BNLSFinderLatest, CompareMethod.Text) = 0) Or (Len(BNLSFinderEntries(BNLSFinderIndex)) = 0)
            BNLSFinderIndex = BNLSFinderIndex + 1

            If BNLSFinderIndex > UBound(BNLSFinderEntries) Then
                'All BNLS servers have been tried and failed
                Err.Raise(ERROR_FINDBNLSSERVER, , "All the BNLS servers have failed.")
                Exit Do
            End If
        Loop
		
		BotVars.BNLSServer = BNLSFinderEntries(BNLSFinderIndex)
		
		ConnectBNLS()
	End Sub
	
	' Updated 8/8/07 to support new prefix/suffix box feature
	'UPGRADE_WARNING: Event frmChat.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Sub frmChat_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error Resume Next
		
		Dim lblHeight As Short
		Static WasMaximized As Boolean
		Static DoMaximize As Boolean
		
		If Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then
			If Not BotVars.NoTray Then
#If Not COMPILE_DEBUG = 1 Then
				Me.Hide()
				
				With nid
                    .cbSize = Len(nid)
					.hWnd = Me.Handle.ToInt32
					.uId = ID_TASKBARICON
					.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
					.uCallBackMessage = WM_ICONNOTIFY
					'UPGRADE_ISSUE: Picture property Icon.handle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					.hIcon = Me.Icon.Handle
					.szTip = GenerateTooltip()
				End With
				
				Shell_NotifyIcon(NIM_ADD, nid)
#End If
			End If
		Else
			Shell_NotifyIcon(NIM_DELETE, nid)
			cboSend.Focus()
			
			If txtPre.Visible Then
				txtPre.Height = cboSend.Height
				txtPre.Width = txtPost.Width
			End If
			
			If txtPost.Visible Then
				txtPost.Height = cboSend.Height
			End If
			
			'sizing + positioning
			
			'Added IsWindowsVista() call within an IIf() statement.
			'This shrinks the size of the entire layout by a further 80 twips.
			'This will act as a fix for Vista's screen cut off issues.
			'   - FrOzeN
			'This issue only occured under Aero, so the fix breaks the GUI for classic themes
			'on vista. This should be fixed now. ~Pyro
			With lvChannel
				'rtbChat.Width = Me.Width - .Width - IIf(g_OSVersion.IsWindowsVista, 200, 120)
				
				rtbChat.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.ClientRectangle.Width) - VB6.PixelsToTwipsX(.Width))
				
				'    .Width = (Me.Width / 4) - 120 'magic number?
				'    If .Width > (.ColumnHeaders.Item(1).Width + 700) Then
				'        .Width = .ColumnHeaders.Item(1).Width + 700
				'
				'        rtbChat.Width = Me.Width - .Width - 120
				'    Else
				'        rtbChat.Width = ((Me.Width / 4) * 3)
				'    End If
				'
				'    .ColumnHeaders.Item(1).Width = (.Width / 3) * 2.5
			End With
			
			lblCurrentChannel.Width = lvChannel.Width
			lvFriendList.Width = lvChannel.Width
			lvClanList.Width = lvChannel.Width
			cboSend.Width = rtbChat.Width
			
			With cmdShowHide
				If rtbWhispersVisible Then
					'Debug.Print "-> " & rtbWhispers.Height
					.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(rtbWhispers.Height) + 285)
					.Text = CAP_HIDE
					ToolTip1.SetToolTip(cmdShowHide, TIP_HIDE)
				Else
					.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(txtPre.Height) - VB6.TwipsPerPixelY)
					.Text = CAP_SHOW
					ToolTip1.SetToolTip(cmdShowHide, TIP_SHOW)
				End If
				
				.BringToFront()
			End With
			
			rtbWhispers.Visible = rtbWhispersVisible
			
			'height is based on rtbchat.height + cmdshowhide.height
			If rtbWhispersVisible Then
				rtbChat.Height = VB6.TwipsToPixelsY(((VB6.PixelsToTwipsY(Me.ClientRectangle.Height) / VB6.TwipsPerPixelY) - (VB6.PixelsToTwipsY(txtPre.Height) / VB6.TwipsPerPixelY) - (VB6.PixelsToTwipsY(rtbWhispers.Height) / VB6.TwipsPerPixelY)) * (VB6.TwipsPerPixelY))
				rtbWhispers.SetBounds(rtbChat.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(cboSend.Top) + VB6.PixelsToTwipsY(cboSend.Height)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			Else
				rtbChat.Height = VB6.TwipsToPixelsY(((VB6.PixelsToTwipsY(Me.ClientRectangle.Height) / VB6.TwipsPerPixelY) - (VB6.PixelsToTwipsY(txtPre.Height) / VB6.TwipsPerPixelY)) * (VB6.TwipsPerPixelY))
			End If
			
			lvChannel.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(rtbChat.Left) + VB6.PixelsToTwipsX(rtbChat.Width)), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lblCurrentChannel.Top) + VB6.PixelsToTwipsY(lblCurrentChannel.Height)), lvChannel.Width, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(rtbChat.Height) - VB6.PixelsToTwipsY(lblCurrentChannel.Height)))
			lvFriendList.SetBounds(lvChannel.Left, lvChannel.Top, lvChannel.Width, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(rtbChat.Height) - VB6.PixelsToTwipsY(lblCurrentChannel.Height)))
			lvClanList.SetBounds(lvChannel.Left, lvChannel.Top, lvChannel.Width, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(rtbChat.Height) - VB6.PixelsToTwipsY(lblCurrentChannel.Height)))
			lblCurrentChannel.SetBounds(lvChannel.Left, rtbChat.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			
			If txtPre.Visible Then
				txtPre.SetBounds(rtbChat.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(rtbChat.Top) + VB6.PixelsToTwipsY(rtbChat.Height) + (VB6.TwipsPerPixelY / 3)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				cboSend.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(txtPre.Left) + VB6.PixelsToTwipsX(txtPre.Width)), txtPre.Top, VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(rtbChat.Width) - VB6.PixelsToTwipsX(txtPre.Width)), 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y Or Windows.Forms.BoundsSpecified.Width)
			Else
				cboSend.SetBounds(rtbChat.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(rtbChat.Top) + VB6.PixelsToTwipsY(rtbChat.Height) + (VB6.TwipsPerPixelY / 3)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			End If
			
			If txtPost.Visible Then
				cboSend.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cboSend.Width) - VB6.PixelsToTwipsX(txtPost.Width))
				txtPost.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cboSend.Left) + VB6.PixelsToTwipsX(cboSend.Width)), cboSend.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			End If
			
			lvChannel.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(rtbChat.Height) - VB6.PixelsToTwipsY(lblCurrentChannel.Height))
			lvFriendList.Height = lvChannel.Height
			lvClanList.Height = lvChannel.Height
			
			'Minus 80 twips from rtbWhispers.Width if using Vista to fix width issue
			'the issue is not with Vista, but with Aero.
			With rtbWhispers
				If .Visible Then
					.SetBounds(rtbChat.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(cboSend.Top) + VB6.PixelsToTwipsY(cboSend.Height)), VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(Me.ClientRectangle.Width) - VB6.PixelsToTwipsX(cmdShowHide.Width) - VB6.TwipsPerPixelX)), 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y Or Windows.Forms.BoundsSpecified.Width)
				End If
			End With
			
			ListviewTabs.Height = cboSend.Height
			ListviewTabs.SetBounds(lvChannel.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(cboSend.Top) + VB6.TwipsPerPixelY), VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvChannel.Width) - VB6.PixelsToTwipsX(cmdShowHide.Width) - VB6.TwipsPerPixelX), cboSend.Height) '+ 2 * Screen.TwipsPerPixelY
			
			If rtbWhispersVisible Then
				cmdShowHide.SetBounds(VB6.TwipsToPixelsX((((VB6.PixelsToTwipsX(rtbWhispers.Left) + VB6.PixelsToTwipsX(rtbWhispers.Width)) / VB6.TwipsPerPixelX) + 1) * VB6.TwipsPerPixelX), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lvChannel.Top) + VB6.PixelsToTwipsY(lvChannel.Height) + VB6.TwipsPerPixelY), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			Else
				cmdShowHide.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(ListviewTabs.Left) + VB6.PixelsToTwipsX(ListviewTabs.Width)), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lvChannel.Top) + VB6.PixelsToTwipsY(lvChannel.Height)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			End If
			
			With lvClanList
				'UPGRADE_WARNING: Lower bound of collection lvClanList.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				.Columns.Item(1).Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(.Width) \ 4) * 3 - 150)
				'UPGRADE_WARNING: Lower bound of collection lvClanList.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				.Columns.Item(2).Width = VB6.TwipsToPixelsX(imlClan.ImageSize.Width) '.Width \ 4 + 200
				'UPGRADE_WARNING: Lower bound of collection lvClanList.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				.Columns.Item(3).Width = 0
			End With
			
			With lvFriendList
				'UPGRADE_WARNING: Lower bound of collection lvFriendList.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				.Columns.Item(1).Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(.Width) \ 4) * 3)
				'UPGRADE_WARNING: Lower bound of collection lvFriendList.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				.Columns.Item(2).Width = VB6.TwipsToPixelsX(imlIcons.ImageSize.Width) '.Width \ 4 + 200
			End With
		End If
		
		If Me.WindowState = System.Windows.Forms.FormWindowState.Maximized Then
			WasMaximized = True
			Call RecordWindowPosition(True)
		ElseIf Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then 
			If WasMaximized Then
				WasMaximized = False
				DoMaximize = True
			End If
		Else
			WasMaximized = False
			
			If DoMaximize Then
				DoMaximize = False
				
				Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
			End If
		End If
		
		Call rtbChat.Refresh()
		
		Exit Sub
		
ERROR_HANDLER: 
		AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in Form_Resize().")
	End Sub
	
	Function GenerateTooltip() As String
		GenerateTooltip = New String(vbNullChar, 64)
        GenerateTooltip = IIf(Len(GetCurrentUsername) > 0, GetCurrentUsername, "offline") & " @ " & BotVars.Server & " (" & StrReverse(BotVars.Product) & ")" & Chr(0)
	End Function
	
	Sub UpdateTrayTooltip()
		On Error Resume Next
		
		If Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then
			With nid
                .cbSize = Len(nid)
				.hWnd = Me.Handle.ToInt32
				.uId = ID_TASKBARICON
				.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
				.uCallBackMessage = WM_ICONNOTIFY
				'UPGRADE_ISSUE: Picture property Icon.handle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				.hIcon = Me.Icon.Handle
				.szTip = GenerateTooltip()
			End With
			
			Shell_NotifyIcon(NIM_MODIFY, nid)
		End If
	End Sub
	
	Private Sub ClanHandler_CandidateList(ByVal Status As Byte, ByRef Users As System.Array) Handles ClanHandler.CandidateList
		Dim i As Integer
		
		'Valid Status codes:
		'   0x00: Successfully found candidate(s)
		'   0x01: Clan tag already taken
		'   0x08: Already in clan
		'   0x0a: Invalid clan tag specified
		
		If MDebug("debug") Then
			AddChat(RTBColors.ErrorMessageText, "CandidateList received. Status code [0x" & Hex(Status) & "].")
			If UBound(Users) > -1 Then
				AddChat(RTBColors.InformationText, "Potential clan members:")
				
				For i = 0 To UBound(Users)
					AddChat(RTBColors.InformationText, Users(i))
				Next i
			End If
		End If
		
		RunInAll("Event_ClanCandidateList", Status, New Object(){ConvertStringArray(Users)})
	End Sub
	
	Private Sub ClanHandler_MemberLeaves(ByVal Member As String) Handles ClanHandler.MemberLeaves
		AddChat(RTBColors.InformationText, "[CLAN] " & Member & " has left the clan.")
		
		Dim x As System.Windows.Forms.ListViewItem
		Dim pos As Short
		
		pos = g_Clan.GetUserIndexEx(Member)
		
		If (pos > 0) Then
			g_Clan.Members.Remove(pos)
		End If
		
		Member = ConvertUsername(Member)
		
		x = lvClanList.FindItemWithText(Member)
		
		If (Not (x Is Nothing)) Then
			lvClanList.Items.RemoveAt(x.Index)
			
			lvClanList.Refresh()
			
			'UPGRADE_NOTE: Object x may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			x = Nothing
		End If
		
		On Error Resume Next
		
		RunInAll("Event_ClanMemberLeaves", Member)
	End Sub
	
	Private Sub ClanHandler_RemovedFromClan(ByVal Status As Byte) Handles ClanHandler.RemovedFromClan
		If Status = 1 Then
			If (AwaitingSelfRemoval = 0) Then
				g_Clan = New clsClanObj
				
				Clan.isUsed = False
				
				ListviewTabs.TabPages.Item(LVW_BUTTON_CLAN).Enabled = False
				lvClanList.Items.Clear()
				ListviewTabs.SelectedIndex = LVW_BUTTON_CHANNEL
				Call ListviewTabs_SelectedIndexChanged(ListviewTabs, New System.EventArgs())
				
				AddChat(RTBColors.ErrorMessageText, "[CLAN] You have been removed from the clan, or it has been disbanded.")
				
				On Error Resume Next
				RunInAll("Event_BotRemovedFromClan")
			End If
			
			On Error Resume Next
			RunInAll("Event_BotRemovedFromClan")
		End If
	End Sub
	
	Private Sub ClanHandler_MyRankChange(ByVal NewRank As Byte) Handles ClanHandler.MyRankChange
		If (g_Clan.Self.Rank < NewRank) Then
			AddChat(RTBColors.SuccessText, "[CLAN] You have been promoted. Your new rank is ", RTBColors.InformationText, GetRank(NewRank), RTBColors.SuccessText, ".")
		ElseIf (g_Clan.Self.Rank > NewRank) Then 
			AddChat(RTBColors.SuccessText, "[CLAN] You have been demoted. Your new rank is ", RTBColors.InformationText, GetRank(NewRank), RTBColors.SuccessText, ".")
		Else
			AddChat(RTBColors.SuccessText, "[CLAN] Your new rank is ", RTBColors.InformationText, GetRank(NewRank), RTBColors.SuccessText, ".")
		End If
		
		g_Clan.Self.Rank = NewRank
		
		On Error Resume Next
		
		RunInAll("Event_BotClanRankChanged", NewRank)
	End Sub
	
	Private Sub ClanHandler_ClanInfo(ByVal ClanTag As String, ByVal RawClanTag As String, ByVal Rank As Byte) Handles ClanHandler.ClanInfo
		g_Clan = New clsClanObj
		
		With Clan
			.Name = ClanTag
			.DWName = RawClanTag
			.MyRank = Rank
			.isUsed = True
		End With
		
		With g_Clan
			.Name = ClanTag
		End With
		
		Call InitListviewTabs()
		
		'If g_Clan.Self.Rank = 0 Then g_Clan.Self.Rank = 1
		On Error Resume Next
		
		ClanTag = KillNull(ClanTag)
		
		BotVars.Clan = ClanTag
		
		If AwaitingClanMembership = 1 Then
			AddChat(RTBColors.SuccessText, "[CLAN] You are now a member of ", RTBColors.InformationText, "Clan " & ClanTag, RTBColors.SuccessText, "!")
			AwaitingClanMembership = 0
			
			RunInAll("Event_BotJoinedClan", ClanTag)
		Else
			AddChat(RTBColors.SuccessText, "[CLAN] You are a ", RTBColors.InformationText, GetRank(Rank), RTBColors.SuccessText, " in ", RTBColors.InformationText, "Clan " & ClanTag, RTBColors.SuccessText, ".")
			
			RunInAll("Event_BotClanInfo", ClanTag, Rank)
		End If
		
		RequestClanList()
		RequestClanMOTD()
		
		'frmChat.ClanHandler.RequestClanMotd 1

        'Me.ListviewTabs_Click(0)
	End Sub
	
	Private Sub ClanHandler_ClanInvitation(ByVal Token As String, ByVal ClanTag As String, ByVal RawClanTag As String, ByVal ClanName As String, ByVal InvitedBy As String, ByVal NewClan As Boolean) Handles ClanHandler.ClanInvitation
		If Not mnuIgnoreInvites.Checked And IsW3 Then
			Clan.Token = Token
			Clan.DWName = RawClanTag
			Clan.Creator = InvitedBy
			Clan.Name = ClanName
			If NewClan Then Clan.isNew = 1
			
			With RTBColors
				AddChat(.SuccessText, "[CLAN] ", .InformationText, ConvertUsername(InvitedBy), .SuccessText, " has invited you to join a clan: ", .InformationText, ClanName, .SuccessText, " [", .InformationText, ClanTag, .SuccessText, "]")
			End With
			
			frmClanInvite.Show()
		End If
		
		RunInAll("Event_ClanInvitation", Token, ClanTag, RawClanTag, ClanName, InvitedBy, NewClan)
	End Sub
	
	Private Sub ClanHandler_ClanMemberList(ByRef Members As System.Array) Handles ClanHandler.ClanMemberList
		Dim ClanMember As clsClanMemberObj
		Dim i As Integer
		
		If AwaitingClanList = 1 Then
			For i = 0 To UBound(Members) Step 4
				ClanMember = New clsClanMemberObj
				
				With ClanMember
					.Name = Members(i)
					.Rank = Val(Members(i + 1))
					.Status = Val(Members(i + 2))
					.Location = Members(i + 3)
				End With
				
				g_Clan.Members.Add(ClanMember)
				If ((Len(Members(i)) > 0) And (UBound(Members) >= i + 1)) Then
					AddClanMember(ClanMember.DisplayName, Val(Members(i + 1)), Val(Members(i + 2)))
					
					On Error Resume Next
					
					RunInAll("Event_ClanMemberList", ClanMember.DisplayName, Val(Members(i + 1)), Val(Members(i + 2)))
				End If
			Next i
		End If
		
		lblCurrentChannel.Text = GetChannelString()
		
        'Me.ListviewTabs_Click(0)
	End Sub
	
	'UPGRADE_NOTE: Location was upgraded to Location_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub ClanHandler_ClanMemberUpdate(ByVal Username As String, ByVal Rank As Byte, ByVal IsOnline As Byte, ByVal Location_Renamed As String) Handles ClanHandler.ClanMemberUpdate
		Dim x As System.Windows.Forms.ListViewItem
		Dim pos As Short
		
		pos = g_Clan.GetUserIndexEx(Username)
		
		Dim ClanMember As clsClanMemberObj
		If (pos > 0) Then
			With g_Clan.Members.Item(pos)
				'UPGRADE_WARNING: Couldn't resolve default property of object g_Clan.Members().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Name = Username
				'UPGRADE_WARNING: Couldn't resolve default property of object g_Clan.Members().Rank. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Rank = Rank
				'UPGRADE_WARNING: Couldn't resolve default property of object g_Clan.Members().Status. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Status = IsOnline
				'UPGRADE_WARNING: Couldn't resolve default property of object g_Clan.Members().Location. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Location = Location_Renamed
			End With
		Else
			
			ClanMember = New clsClanMemberObj
			
			With ClanMember
				.Name = Username
				.Rank = Rank
				.Status = IsOnline
				.Location = Location_Renamed
			End With
			
			g_Clan.Members.Add(ClanMember)
		End If
		
		Username = ConvertUsername(Username)
		
		x = lvClanList.FindItemWithText(Username)
		
		If StrComp(Username, CurrentUsername, CompareMethod.Text) = 0 Then
			g_Clan.Self.Rank = IIf(Rank = 0, Rank + 1, Rank)
			AwaitingClanInfo = 1
		End If
		
		If AwaitingClanInfo = 1 Then
			AwaitingClanInfo = 0
			AddChat(RTBColors.SuccessText, "[CLAN] Member update: ", RTBColors.InformationText, Username, RTBColors.SuccessText, " is now a " & GetRank(Rank) & ".")
		End If
		
		If Not (x Is Nothing) Then
			lvClanList.Items.RemoveAt(x.Index)
			'UPGRADE_NOTE: Object x may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			x = Nothing
		End If
		
		AddClanMember(Username, CShort(Rank), CShort(IsOnline))
		
		On Error Resume Next
		RunInAll("Event_ClanMemberUpdate", Username, Rank, IsOnline)
	End Sub
	
	Private Sub ClanHandler_ClanMOTD(ByVal cookie As Integer, ByVal Message As String) Handles ClanHandler.ClanMOTD
		g_Clan.MOTD = Message
		
		If (cookie = 1) Then
			PassedClanMotdCheck = True
			
			'If (g_Clan.MOTD <> vbNullString) Then
			'    frmChat.AddChat vbBlue, Message
			'End If
		End If
		
		On Error Resume Next
		
		RunInAll("Event_ClanMOTD", Message)
	End Sub
	
	Private Sub ClanHandler_DemoteUserReply(ByVal Success As Boolean) Handles ClanHandler.DemoteUserReply
		If Success Then
			AddChat(RTBColors.SuccessText, "[CLAN] User demoted successfully.")
		Else
			AddChat(RTBColors.ErrorMessageText, "[CLAN] User demotion failed.")
		End If
		
		lblCurrentChannel.Text = GetChannelString
		
		RunInAll("Event_ClanDemoteUserReply", Success)
	End Sub
	
	Private Sub ClanHandler_DisbandClanReply(ByVal Success As Boolean) Handles ClanHandler.DisbandClanReply
		If MDebug("debug") Then
			AddChat(RTBColors.ConsoleText, "DisbandClanReply: " & Success)
		End If
		RunInAll("Event_ClanDisbandReply", Success)
	End Sub
	
	Private Sub ClanHandler_InviteUserReply(ByVal Status As Byte) Handles ClanHandler.InviteUserReply
		'0x00: Invitation accepted
		'0x04: Invitation declined
		'0x05: Failed to invite user
		'0x09: Clan is full
		
		Select Case Status
			Case 0, 1 : AddChat(RTBColors.SuccessText, "[CLAN] The invitation was accepted.")
			Case 4 : AddChat(RTBColors.ErrorMessageText, "[CLAN] The invitation was declined.")
			Case 5 : AddChat(RTBColors.ErrorMessageText, "[CLAN] The invitation failed.")
			Case 9 : AddChat(RTBColors.ErrorMessageText, "[CLAN] The invitation failed: Your clan is full.")
			Case Else : AddChat(RTBColors.ErrorMessageText, "[CLAN] Unknown invitation status: " & Status)
		End Select
		RunInAll("Event_ClanInviteUserReply", Status)
	End Sub
	
	Private Sub ClanHandler_PromoteUserReply(ByVal Success As Boolean) Handles ClanHandler.PromoteUserReply
		If Success Then
			AddChat(RTBColors.SuccessText, "[CLAN] User promoted successfully.")
		Else
			AddChat(RTBColors.ErrorMessageText, "[CLAN] User promotion failed.")
		End If
		
		lblCurrentChannel.Text = GetChannelString
		RunInAll("Event_ClanPromoteUserReply", Success)
	End Sub
	
	Private Sub ClanHandler_RemoveUserReply(ByVal Result As Byte) Handles ClanHandler.RemoveUserReply
		
		'    0x00: Successfully removed user from clan
		'    0x02: Too soon to remove user
		'    0x03: Not enough members to remove this user
		'    0x07: Not authorized to remove the user
		'    0x08: User is not in your clan
		
		'Debug.Print "Removed successfully!"
		
		Select Case Result
			Case 0
				If AwaitingSelfRemoval = 1 Then
					AwaitingSelfRemoval = 0
					Clan.isUsed = False
					
					ListviewTabs.TabPages.Item(LVW_BUTTON_CLAN).Enabled = False
					lvClanList.Items.Clear()
					
					ListviewTabs.SelectedIndex = LVW_BUTTON_CHANNEL
					Call ListviewTabs_SelectedIndexChanged(ListviewTabs, New System.EventArgs())
					
					g_Clan = New clsClanObj
					
					AddChat(RTBColors.SuccessText, "[CLAN] You have successfully left the clan.")
				Else
					AddChat(RTBColors.SuccessText, "[CLAN] User removed successfully.")
					
					RequestClanList()
				End If
				
			Case 2
				AddChat(RTBColors.ErrorMessageText, "[CLAN] That user is currently on probation.")
				
			Case 3
				AddChat(RTBColors.ErrorMessageText, "[CLAN] There are not enough members for you to remove that user.")
				
			Case 7
				AddChat(RTBColors.ErrorMessageText, "[CLAN] You are not authorized to remove that user.")
				
			Case 8
				AddChat(RTBColors.ErrorMessageText, "[CLAN] You are not allowed to remove that user.")
				
			Case Else
				AddChat(RTBColors.InformationText, "[CLAN] 0x78 Response code: 0x" & Hex(Result))
				AddChat(RTBColors.InformationText, "[CLAN] You failed to remove that user from the clan.")
		End Select
		
		lblCurrentChannel.Text = GetChannelString
		RunInAll("Event_ClanRemoveUserReply", Result)
	End Sub
	
	Private Sub ClanHandler_UnknownClanEvent(ByVal PacketID As Byte, ByVal Data As String) Handles ClanHandler.UnknownClanEvent
		If MDebug("debug") Then
			Me.AddChat(RTBColors.ErrorMessageText, "[CLAN] Unknown clan event [0x" & Hex(PacketID) & "]. Data is as follows:")
			Me.AddChat(RTBColors.ErrorMessageText, Data)
		End If
	End Sub
	
	Public Function GetLogFilePath() As String
		
		Dim Path As String
		Dim f As Short
		
		f = FreeFile
		
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Path = StringFormat("{0}{1}.txt", GetFolderPath("Logs"), VB6.Format(Today, "YYYY-MM-DD"))
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If (Dir(Path) = vbNullString) Then
			FileOpen(f, Path, OpenMode.Output)
			FileClose(1)
		End If
		
		GetLogFilePath = Path
		
	End Function
	
	Sub frmChat_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Dim Key As String
		Dim L As Integer
		
		'Me.WindowState = vbNormal
		'Me.Show
		
		'UserCancelledConnect = False
		
		'Cancel = 1
		
		'scTimer.Enabled = False
		'SControl.Reset
		
		'UPGRADE_WARNING: Untranslated statement in Form_Unload. Please check source code.
		
		AddChat(RTBColors.ErrorMessageText, "Shutting down...")
		
		If Config.FileExists Then
			If Me.WindowState <> System.Windows.Forms.FormWindowState.Minimized Then
				Call RecordWindowPosition(CBool(Me.WindowState = System.Windows.Forms.FormWindowState.Maximized))
			End If
			
			Call Config.Save()
		End If
		
		'With frmChat.INet
		'    If .StillExecuting Then .Cancel
		'End With
		
		Call DoDisconnect(1)
		
		Shell_NotifyIcon(NIM_DELETE, nid)
		
		On Error Resume Next
		
		RunInAll("Event_LoggedOff")
		RunInAll("Event_Close")
		RunInAll("Event_Shutdown")
		
		DestroyObjs()
		
		On Error GoTo 0
		
		If ExReconnectTimerID > 0 Then
			KillTimer(0, ExReconnectTimerID)
		End If
		
		'    If AttemptedNewVerbyte Then
		'        AttemptedNewVerbyte = False
		'        l = CLng(Val("&H" & ReadCFG("Main", Key & "VerByte")))
		'        WriteINI "Main", Key & "VerByte", Hex(l - 1)
		'    End If
		
		Call modWarden.WardenCleanup(WardenInstance)
		
		'Call ChatQueue_Terminate
		
		DisableURLDetect(Me.rtbChat.Handle.ToInt32)
		UnhookWindowProc(Me.Handle.ToInt32)
		
		'Call SharedScriptSupport.Dispose 'Explicit call the Class_Terminate sub in the ScriptSupportClass to destroy all the forms. - FrOzeN
		
		'DeconstructSettings
		'DeconstructMonitor
		DestroyAllWWs()
		
		'UPGRADE_NOTE: Object g_Logger may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		g_Logger = Nothing
		'UPGRADE_NOTE: Object BotVars may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		BotVars = Nothing
		'UPGRADE_NOTE: Object ClanHandler may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ClanHandler = Nothing
		'UPGRADE_NOTE: Object ListToolTip may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ListToolTip = Nothing
		'UPGRADE_NOTE: Object GErrorHandler may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		GErrorHandler = Nothing
		'UPGRADE_NOTE: Object FriendListHandler may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		FriendListHandler = Nothing
		'UPGRADE_NOTE: Object colWhisperWindows may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		colWhisperWindows = Nothing
		'UPGRADE_NOTE: Object colLastSeen may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		colLastSeen = Nothing
		'UPGRADE_NOTE: Object SharedScriptSupport may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		SharedScriptSupport = Nothing
		'UPGRADE_NOTE: Object ds may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ds = Nothing
		
		'Set dictTimerInterval = Nothing
		'Set dictTimerCount = Nothing
		'Set dictTimerEnabled = Nothing
		
		' Updated to match current form list 2009-02-09 Andy
		frmAbout.Close()
		frmCatch.Close()
		frmCommands.Close()
		frmClanInvite.Close()
		frmCustomInputBox.Close()
		frmDBType.Close()
		frmEMailReg.Close()
		frmFilters.Close()
		frmDBGameSelection.Close()
		frmDBNameEntry.Close()
		frmDBManager.Close()
		frmManageKeys.Close()
		'Unload frmMonitor
		frmProfile.Close()
		'Unload frmProfileManager
		frmQuickChannel.Close()
		frmRealm.Close()
		'Unload frmScriptUI
		frmScript.Close()
		frmSettings.Close()
		frmSplash.Close()
		'Unload frmUserManager
		frmWhisperWindow.Close()
		'Unload frmWriteProfile
		
		' Added this instead of End to try and fix some system tray crashes 2009-0211-andy
		'  It was used in some capacity before since the API was already declared
		'   in modAPI...
		' added preprocessor check; the bot was ending the VB6 IDE's process too! - ribose
		' if it was compiled with the debugger, we don't allow minimizing to tray anyway
#If Not COMPILE_DEBUG = 1 Then
		Call ExitProcess(0)
#Else
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression Else did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		End
#End If
	End Sub
	
	
	Public Sub AddFriend(ByVal Username As String, ByVal Product As String, ByRef IsOnline As Boolean)
		Dim i, OnlineIcon As Short
		Dim f As System.Windows.Forms.ListViewItem
		Const ICOFFLINE As Short = MONITOR_OFFLINE
		Const ICONLINE As Short = MONITOR_ONLINE
		
		If IsOnline Then OnlineIcon = ICONLINE Else OnlineIcon = ICOFFLINE
		
		'Everybody Else
		Select Case Product
			Case Is = PRODUCT_STAR
				i = ICSTAR
			Case Is = PRODUCT_SEXP
				i = ICSEXP
			Case Is = PRODUCT_D2DV
				i = ICD2DV
			Case Is = PRODUCT_D2XP
				i = ICD2XP
			Case Is = PRODUCT_W2BN
				i = ICW2BN
			Case Is = PRODUCT_WAR3
				i = ICWAR3
			Case Is = PRODUCT_W3XP
				i = ICWAR3X
			Case Is = PRODUCT_CHAT
				i = ICCHAT
			Case Is = PRODUCT_DRTL
				i = ICDIABLO
			Case Is = PRODUCT_DSHR
				i = ICDIABLOSW
			Case Is = PRODUCT_JSTR
				i = ICJSTR
			Case Is = PRODUCT_SSHR
				i = ICSCSW
			Case Else
				i = ICUNKNOWN
		End Select
		
		f = lvFriendList.FindItemWithText(Username)
		
		If (f Is Nothing) Then
			With lvFriendList.Items
                .Add(Username, i)
                .Item(.Count).SubItems.Add(IsOnline)
			End With
		Else
            f.ImageIndex = i
            f.SubItems.Item(1).Text = IsOnline

            f = Nothing
		End If
		
		lblCurrentChannel.Text = GetChannelString()
		
        'Me.ListviewTabs_Click(0)
	End Sub
	
	'UPGRADE_NOTE: Location was upgraded to Location_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub FriendListHandler_FriendAdded(ByVal Username As String, ByVal Product As String, ByVal Location_Renamed As Byte, ByVal Status As Byte, ByVal Channel As String) Handles FriendListHandler.FriendAdded
		AddFriend(Username, Product, (Location_Renamed > 0))
		lblCurrentChannel.Text = GetChannelString
	End Sub
	
	Private Sub FriendListHandler_FriendListReceived(ByVal FriendCount As Byte) Handles FriendListHandler.FriendListReceived
		lvFriendList.Items.Clear()
	End Sub
	
	'UPGRADE_NOTE: Location was upgraded to Location_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub FriendListHandler_FriendListEntry(ByVal Username As String, ByVal Product As String, ByVal Channel As String, ByVal Status As Byte, ByVal Location_Renamed As Byte) Handles FriendListHandler.FriendListEntry
		AddFriend(Username, Product, (Location_Renamed > 0))
		lblCurrentChannel.Text = GetChannelString
	End Sub
	
	Private Sub FriendListHandler_FriendMoved() Handles FriendListHandler.FriendMoved
		Call FriendListHandler.RequestFriendsList(PBuffer)
	End Sub
	
	Private Sub FriendListHandler_FriendRemoved(ByVal Username As String) Handles FriendListHandler.FriendRemoved
		Dim flItem As System.Windows.Forms.ListViewItem
		
		flItem = lvFriendList.FindItemWithText(Username)
		
		If (Not (flItem Is Nothing)) Then
			lvFriendList.Items.RemoveAt(flItem.Index)
			
			'UPGRADE_NOTE: Object flItem may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			flItem = Nothing
		End If
		
		lblCurrentChannel.Text = GetChannelString
	End Sub
	
	Private Sub FriendListHandler_FriendUpdate(ByVal Username As String, ByVal FLIndex As Byte) Handles FriendListHandler.FriendUpdate
		On Error GoTo ERROR_HANDLER
		
		Dim x As System.Windows.Forms.ListViewItem
		Dim i As Short
		Const ICONLINE As Short = MONITOR_ONLINE
		Const ICOFFLINE As Short = MONITOR_OFFLINE
		
		x = lvFriendList.FindItemWithText(Username)
		
		If Not (x Is Nothing) Then
			With g_Friends.Item(FLIndex)
				'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends.Item(FLIndex).LocationID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case .LocationID
					Case FRL_OFFLINE
                        x.ImageIndex = ICUNKNOWN
						
                        x.SubItems.Item(1).Text = "False"
						
					Case Else
						'UPGRADE_WARNING: Lower bound of collection x.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
						'UPGRADE_ISSUE: MSComctlLib.ListSubItem property ListSubItems.Item.ReportIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                        If x.SubItems.Item(1).Text = "False" Then
                            'Friend is now online - notify user?
                        End If
						
						'UPGRADE_WARNING: Lower bound of collection x.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
						'UPGRADE_ISSUE: MSComctlLib.ListSubItem property ListSubItems.Item.ReportIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                        x.SubItems.Item(1).Text = "True"
						
						'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends.Item(FLIndex).Game. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Select Case .Game
							Case Is = PRODUCT_STAR : i = ICSTAR
							Case Is = PRODUCT_SEXP : i = ICSEXP
							Case Is = PRODUCT_D2DV : i = ICD2DV
							Case Is = PRODUCT_D2XP : i = ICD2XP
							Case Is = PRODUCT_W2BN : i = ICW2BN
							Case Is = PRODUCT_WAR3 : i = ICWAR3
							Case Is = PRODUCT_W3XP : i = ICWAR3X
							Case Is = PRODUCT_CHAT : i = ICCHAT
							Case Is = PRODUCT_DRTL : i = ICDIABLO
							Case Is = PRODUCT_DSHR : i = ICDIABLOSW
							Case Is = PRODUCT_JSTR : i = ICJSTR
							Case Is = PRODUCT_SSHR : i = ICSCSW
							Case Else : i = ICUNKNOWN
						End Select
						
						'UPGRADE_ISSUE: MSComctlLib.ListItem property x.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                        x.ImageIndex = i
				End Select
			End With
			
		End If
		
		'UPGRADE_NOTE: Object x may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		x = Nothing
		
		Exit Sub
		
ERROR_HANDLER: 
		AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in FriendUpdate().")
		
		Exit Sub
	End Sub
	
	Private Sub lblCurrentChannel_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblCurrentChannel.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)

        mnuQCTop.ShowDropDown()
	End Sub
	
	Public Sub ListviewTabs_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ListviewTabs.SelectedIndexChanged
		Static PreviousTab As Short = ListviewTabs.SelectedIndex()
		Dim CurrentTab As Short
		
		CurrentTab = ListviewTabs.SelectedIndex
		
		Select Case CurrentTab
			Case LVW_BUTTON_CHANNEL ' = 0 = Channel button clicked
				ToolTip1.SetToolTip(lblCurrentChannel, "Currently in " & g_Channel.SType() & " channel " & g_Channel.Name & " (" & g_Channel.Users.Count() & ")")
				
				lvChannel.BringToFront()
				
			Case LVW_BUTTON_FRIENDS ' = 1 = Friends button clicked
				ToolTip1.SetToolTip(lblCurrentChannel, "Currently viewing " & g_Friends.Count() & " friends")
				
				lvFriendList.BringToFront()
				
			Case LVW_BUTTON_CLAN ' = 2 = Clan button clicked
				ToolTip1.SetToolTip(lblCurrentChannel, "Currently viewing " & g_Clan.Members.Count() & " members of clan " & Clan.Name)
				
				lvClanList.BringToFront()
				
		End Select
		
		lvChannel.HideSelection = True
		lvFriendList.HideSelection = True
		lvClanList.HideSelection = True
		
		lblCurrentChannel.Text = GetChannelString
		PreviousTab = ListviewTabs.SelectedIndex()
	End Sub
	
	' This procedure relies on code in RecordcboSendSelInfo() that sets global variables
	'  cboSendSelLength and cboSendSelStart
	' These two properties are zeroed out as the control loses focus and inaccessible
	'  (zeroed) at both access time in this method AND in the _LostFocus sub
	Private Sub lvChannel_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lvChannel.DoubleClick
		Dim Value As String
		
		Value = GetSelectedUser
        If Len(cboSend.Text) = 0 Then Value = Value & BotVars.AutoCompletePostfix
		Value = Value & Space(1)
		
		If (Len(Value) > 0) Then
			With cboSend
				.SelectionStart = cboSendSelStart
				.SelectionLength = cboSendSelLength
				.SelectedText = Value
				
				cboSendSelStart = cboSendSelStart + Len(Value)
				cboSendSelLength = 0
				
				.Focus()
			End With
		End If
	End Sub
	
	Private Sub lvChannel_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles lvChannel.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Const S_ALT As Short = 4
		
		Dim sStart As Short
		If (KeyCode = 93) Then
			lvChannel_MouseUp(lvChannel, New System.Windows.Forms.MouseEventArgs(2 * &H100000, 0, 0, 0, 0))
		ElseIf (KeyCode = KEY_ALTN And Shift = S_ALT) Then 
			
			With lvChannel
				If Not (.FocusedItem Is Nothing) Then
					cboSend.SelectionStart = Len(cboSend.Text)
					cboSend.SelectedText = .FocusedItem.Text
					
					KeyCode = 0
					Shift = 0
					
					Exit Sub
				End If
			End With
		End If
	End Sub
	
	Private Sub lvFriendList_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles lvFriendList.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Const S_ALT As Short = 4
		
		Dim sStart As Short
		If (KeyCode = 93) Then
			lvFriendList_MouseUp(lvFriendList, New System.Windows.Forms.MouseEventArgs(2 * &H100000, 0, 0, 0, 0))
		ElseIf (KeyCode = KEY_ALTN And Shift = S_ALT) Then 
			
			With lvFriendList
				If Not (.FocusedItem Is Nothing) Then
					cboSend.SelectionStart = Len(cboSend.Text)
					cboSend.SelectedText = .FocusedItem.Text
					
					KeyCode = 0
					Shift = 0
					
					Exit Sub
				End If
			End With
		End If
	End Sub
	
	Private Sub lvClanList_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles lvClanList.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Const S_ALT As Short = 4
		
		Dim sStart As Short
		If (KeyCode = 93) Then
			lvClanList_MouseUp(lvClanList, New System.Windows.Forms.MouseEventArgs(2 * &H100000, 0, 0, 0, 0))
		ElseIf (KeyCode = KEY_ALTN And Shift = S_ALT) Then 
			
			With lvClanList
				If Not (.FocusedItem Is Nothing) Then
					cboSend.SelectionStart = Len(cboSend.Text)
					cboSend.SelectedText = .FocusedItem.Text
					
					KeyCode = 0
					Shift = 0
					
					Exit Sub
				End If
			End With
		End If
	End Sub
	
	Private Sub lvFriendList_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lvFriendList.DoubleClick
		Dim Value As String
		
		Value = GetFriendsSelectedUser
        If Len(cboSend.Text) = 0 Then Value = Value & BotVars.AutoCompletePostfix
		Value = Value & Space(1)
		
		If (Len(Value) > 0) Then
			With cboSend
				.SelectionStart = cboSendSelStart
				.SelectionLength = cboSendSelLength
				.SelectedText = Value
				
				cboSendSelStart = cboSendSelStart + Len(Value)
				cboSendSelLength = 0
				
				.Focus()
			End With
		End If
	End Sub
	
	Private Sub lvClanList_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lvClanList.DoubleClick
		Dim Value As String
		
		Value = GetClanSelectedUser
        If Len(cboSend.Text) = 0 Then Value = Value & BotVars.AutoCompletePostfix
		Value = Value & Space(1)
		
		If (Len(Value) > 0) Then
			With cboSend
				.SelectionStart = cboSendSelStart
				.SelectionLength = cboSendSelLength
				.SelectedText = Value
				
				cboSendSelStart = cboSendSelStart + Len(Value)
				cboSendSelLength = 0
				
				.Focus()
			End With
		End If
	End Sub
	
	Private Sub lvChannel_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lvChannel.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim aInx As Short
		Dim sProd As New VB6.FixedLengthString(4)
		Dim HasOps As Boolean
		Dim UserIsW3 As Boolean
		Dim UserHasStats As Boolean
		
		If (lvChannel.FocusedItem Is Nothing) Then
			Exit Sub
		End If
		
		If Button = VB6.MouseButtonConstants.RightButton Then
			If Not (lvChannel.FocusedItem Is Nothing) Then
				aInx = g_Channel.GetUserIndex(GetSelectedUser)
				
				If aInx > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().Game. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sProd.Value = g_Channel.Users.Item(aInx).Game
					UserIsW3 = (sProd.Value = PRODUCT_W3XP Or sProd.Value = PRODUCT_WAR3)
					Select Case sProd.Value
						Case PRODUCT_STAR, PRODUCT_SEXP, PRODUCT_W2BN, PRODUCT_WAR3, PRODUCT_W3XP, PRODUCT_JSTR, PRODUCT_SSHR
							UserHasStats = True
					End Select
					HasOps = g_Channel.Self.IsOperator()
					
					'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().Clan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    mnuPopInvite.Enabled = UserIsW3 And Len(g_Channel.Users.Item(aInx).Clan) = 0 And InStr(1, GetSelectedUser, "#") = 0 And g_Clan.Self.Rank >= 3
					
					mnuPopStats.Enabled = UserHasStats
					mnuPopWebProfile.Enabled = UserIsW3
					mnuPopKick.Enabled = HasOps
					mnuPopDes.Enabled = HasOps
					mnuPopBan.Enabled = HasOps
				End If
			End If
			
			mnuPopup.Tag = lvChannel.FocusedItem.Text 'Record which user is selected at time of right-clicking. - FrOzeN
			
            mnuPopup.ShowDropDown()
		End If
	End Sub
	
	Private Sub lvFriendList_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lvFriendList.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim aInx As Short
		Dim sProd As New VB6.FixedLengthString(4)
		Dim bIsOn As Boolean
		Dim bIsMutual As Boolean
		Dim UserIsW3 As Boolean
		Dim UserHasStats As Boolean
		Dim SelfHasStats As Boolean
		
		If (lvFriendList.FocusedItem Is Nothing) Then
			Exit Sub
		End If
		
		If Button = VB6.MouseButtonConstants.RightButton Then
			If Not (lvFriendList.FocusedItem Is Nothing) Then
				aInx = lvFriendList.FocusedItem.Index
				
				If aInx > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends().Game. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sProd.Value = g_Friends.Item(aInx).Game
					'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends().IsOnline. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					bIsOn = g_Friends.Item(aInx).IsOnline
					'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends().IsMutual. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					bIsMutual = g_Friends.Item(aInx).IsMutual
					UserIsW3 = (sProd.Value = PRODUCT_W3XP Or sProd.Value = PRODUCT_WAR3)
					Select Case sProd.Value
						Case PRODUCT_STAR, PRODUCT_SEXP, PRODUCT_W2BN, PRODUCT_WAR3, PRODUCT_W3XP, PRODUCT_JSTR, PRODUCT_SSHR
							UserHasStats = True
					End Select
					Select Case StrReverse(BotVars.Product)
						Case PRODUCT_STAR, PRODUCT_SEXP, PRODUCT_W2BN, PRODUCT_WAR3, PRODUCT_W3XP, PRODUCT_JSTR, PRODUCT_SSHR
							SelfHasStats = True
					End Select
					
					mnuPopFLWhisper.Enabled = bIsOn
					mnuPopFLInvite.Enabled = UserIsW3 And bIsMutual And g_Clan.Self.Rank >= 3
					
					mnuPopFLStats.Enabled = UserHasStats Or SelfHasStats
					mnuPopFLWebProfile.Enabled = UserIsW3
				End If
			End If
			
			mnuPopFList.Tag = lvFriendList.FocusedItem.Text 'Record which user is selected at time of right-clicking. - FrOzeN
			
            mnuPopFList.ShowDropDown()
		End If
	End Sub
	
	Private Sub lvChannel_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lvChannel.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim lvhti As LVHITTESTINFO
		Dim lItemIndex As Integer
		Dim sOutBuf As String
		Dim sTemp As String
		Dim UserAccess As udtGetAccessResponse
		Dim Clan As String
		
		lvhti.pt.x = x / VB6.TwipsPerPixelX
		lvhti.pt.y = y / VB6.TwipsPerPixelY
		'UPGRADE_WARNING: Couldn't resolve default property of object lvhti. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lItemIndex = SendMessageAny(lvChannel.Handle.ToInt32, LVM_HITTEST, -1, lvhti) + 1
		
		If m_lCurItemIndex <> lItemIndex Then
			m_lCurItemIndex = lItemIndex
			
			If m_lCurItemIndex = 0 Then ' no item under the mouse pointer
				ListToolTip.Destroy()
			Else
				'UserAccess = GetCumulativeAccess(lvChannel.ListItems(m_lCurItemIndex).text, "USER")
				
				'UPGRADE_WARNING: Lower bound of collection lvChannel.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				ListToolTip.Title = "Information for " & lvChannel.Items.Item(m_lCurItemIndex).Text
				
				'If (UserAccess.Name <> vbNullString) Then
				'    sTemp = sTemp & "["
				'
				'    If (UserAccess.Access > 0) Then
				'        sTemp = sTemp & "rank: " & UserAccess.Access
				'    End If
				'
				'    If ((UserAccess.Flags <> "%") And (UserAccess.Flags <> vbNullString)) Then
				'        If (UserAccess.Access > 0) Then
				'            sTemp = sTemp & ", "
				'        End If
				'
				'        sTemp = sTemp & "flags: " & UserAccess.Flags
				'    End If
				'
				'    sTemp = sTemp & "]" & vbCrLf
				'End If
				
				
				'UPGRADE_WARNING: Lower bound of collection lvChannel.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				lItemIndex = g_Channel.GetUserIndex(lvChannel.Items.Item(m_lCurItemIndex).Text)
				
				If (lItemIndex > 0) Then
					With g_Channel.Users.Item(lItemIndex)
						'ParseStatstring .Statstring, sOutBuf, Clan
						
						'sTemp = sTemp & vbCrLf
						'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().Ping. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sTemp = sTemp & "Ping at login: " & .Ping & "ms" & vbCrLf
						'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().Flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sTemp = sTemp & "Flags: " & FlagDescription(.Flags, True) & vbCrLf
						sTemp = sTemp & vbCrLf
						'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().Stats. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sTemp = sTemp & .Stats.ToString
						
						ListToolTip.TipText = sTemp
						
					End With
					
					Call ListToolTip.Create(lvChannel.Handle.ToInt32, CInt(x), CInt(y))
				End If
			End If
		End If
	End Sub
	
	Private Sub lvFriendList_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lvFriendList.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim lvhti As LVHITTESTINFO
		Dim lItemIndex As Integer
		
		lvhti.pt.x = x / VB6.TwipsPerPixelX
		lvhti.pt.y = y / VB6.TwipsPerPixelY
		'UPGRADE_WARNING: Couldn't resolve default property of object lvhti. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lItemIndex = SendMessageAny(lvFriendList.Handle.ToInt32, LVM_HITTEST, 0, lvhti) + 1
		
		Dim sTemp As String
		If m_lCurItemIndex <> lItemIndex Then
			m_lCurItemIndex = lItemIndex
			
			If m_lCurItemIndex = 0 Then ' no item under the mouse pointer
				ListToolTip.Destroy()
			Else
				'UPGRADE_WARNING: Lower bound of collection lvFriendList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				ListToolTip.Title = "Information for " & lvFriendList.Items.Item(m_lCurItemIndex).Text
				
				
				If ((lItemIndex > 0) And (g_Friends.Count() > 0)) Then
					'UPGRADE_WARNING: Lower bound of collection lvFriendList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					lItemIndex = FriendListHandler.UsernameToFLIndex(lvFriendList.Items.Item(m_lCurItemIndex).Text)
					
					With g_Friends.Item(lItemIndex)
						'                    Private Const FRL_OFFLINE& = &H0
						'                    Private Const FRL_NOTINCHAT& = &H1
						'                    Private Const FRL_INCHAT& = &H2
						'                    Private Const FRL_PUBLICGAME& = &H3
						'                    Private Const FRL_PRIVATEGAME& = &H5
						'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends.Item(lItemIndex).IsOnline. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If .IsOnline Then
							'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends.Item().Game. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sTemp = sTemp & "Using " & ProductCodeToFullName(.Game) & " "
						End If
						
						'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends.Item(lItemIndex).LocationID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Select Case .LocationID
							Case FRL_OFFLINE
								sTemp = sTemp & "This person is offline."
							Case FRL_NOTINCHAT
								sTemp = sTemp & "in limbo. (not yet in chat)"
							Case FRL_INCHAT
								sTemp = sTemp & "in a Battle.net channel."
							Case FRL_PUBLICGAME
								sTemp = sTemp & "in a public game."
							Case FRL_PRIVATEGAME
								sTemp = sTemp & "in a private game."
						End Select
						
						'                    Private Const FRS_NONE& = &H0
						'                    Private Const FRS_MUTUAL& = &H1
						'                    Private Const FRS_DND& = &H2
						'                    Private Const FRS_AWAY& = &H4
						
						'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends.Item(lItemIndex).Status. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (.Status And FRS_MUTUAL) = FRS_MUTUAL Then
							sTemp = sTemp & vbCrLf & "Mutual friend"
							
							'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends.Item(lItemIndex).LocationID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Select Case (.LocationID)
								Case FRL_INCHAT
									'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends.Item().Location. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sTemp = sTemp & ", in channel " & .Location & "."
								Case FRL_PRIVATEGAME
									'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends.Item().Location. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sTemp = sTemp & ", in game " & .Location & "."
								Case Else
									sTemp = sTemp & "."
							End Select
						End If
						
						'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends.Item(lItemIndex).LocationID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (.LocationID = FRL_PUBLICGAME) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends.Item().Location. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sTemp = sTemp & vbCrLf & "Currently in the public game '" & .Location & "'."
						End If
						
						ListToolTip.TipText = sTemp
					End With
					
					Call ListToolTip.Create(lvFriendList.Handle.ToInt32, CInt(x), CInt(y))
				End If
			End If
		End If
	End Sub
	
	Private Sub lvClanList_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lvClanList.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim lvhti As LVHITTESTINFO
		Dim lItemIndex As Short
		Dim ClanMember As clsClanMemberObj
		
		lvhti.pt.x = x / VB6.TwipsPerPixelX
		lvhti.pt.y = y / VB6.TwipsPerPixelY
		'UPGRADE_WARNING: Couldn't resolve default property of object lvhti. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lItemIndex = SendMessageAny(lvClanList.Handle.ToInt32, LVM_HITTEST, 0, lvhti) + 1
		
		Dim sTemp As String
		If m_lCurItemIndex <> lItemIndex Then
			m_lCurItemIndex = lItemIndex
			
			If m_lCurItemIndex = 0 Then ' no item under the mouse pointer
				ListToolTip.Destroy()
			Else
				'UPGRADE_WARNING: Lower bound of collection lvClanList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				ListToolTip.Title = "Information for " & lvClanList.Items.Item(m_lCurItemIndex).Text
				
				
				If ((lItemIndex > 0) And (g_Clan.Members.Count() > 0)) Then
					'UPGRADE_WARNING: Lower bound of collection lvClanList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					ClanMember = g_Clan.GetMember(lvClanList.Items.Item(m_lCurItemIndex).Text)
					
					If (Not ClanMember Is Nothing) Then
						With ClanMember
							If (.Rank = 4) Then
								sTemp = sTemp & "The "
							Else
								sTemp = sTemp & "This "
							End If
							
							sTemp = sTemp & .RankName & " is currently "
							
							If (.IsOnline) Then
								sTemp = sTemp & "online"
								
                                If (Len(.Location) > 0) Then
                                    sTemp = sTemp & " in " & .Location
                                End If
							Else
								sTemp = sTemp & "offline"
							End If
							
							sTemp = sTemp & "."
							
							ListToolTip.TipText = sTemp
						End With
						'UPGRADE_NOTE: Object ClanMember may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						ClanMember = Nothing
					End If
					
					Call ListToolTip.Create(lvClanList.Handle.ToInt32, CInt(x), CInt(y))
				End If
			End If
		End If
	End Sub
	
	Public Sub mnuCatchPhrases_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCatchPhrases.Click
		frmCatch.Show()
	End Sub
	
	Public Sub mnuChangeLog_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuChangeLog.Click
		ShellOpenURL("http://www.stealthbot.net/wiki/changelog", "the StealthBot Changelog")
	End Sub
	
	Public Sub mnuOpenScriptFolder_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOpenScriptFolder.Click
		Dim sPath As String
		sPath = GetFolderPath("Scripts")
		
		' Does the script folder exist?
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If (Len(Dir(sPath, FileAttribute.Directory)) > 0) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Shell(StringFormat("explorer.exe {0}", sPath), AppWinStyle.NormalFocus)
        Else
            ' Try and create it
            MkDir(sPath)

            ' Did it work?
            'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            If (Len(Dir(sPath, FileAttribute.Directory)) > 0) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Shell(StringFormat("explorer.exe {0}", sPath), AppWinStyle.NormalFocus)
            Else
                Call Me.AddChat(RTBColors.ErrorMessageText, "Your script folder does not exist, and could not be created.")
                Call Me.AddChat(RTBColors.ErrorMessageText, "Script folder path: " & sPath)
                Exit Sub
            End If
        End If
	End Sub
	
	Public Sub mnuPopClanAddLeft_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanAddLeft.Click
		On Error Resume Next
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		
		If txtPre.Enabled Then 'fix for topic 25290 -a
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			txtPre.Text = StringFormat("/w {0} ", GetClanSelectedUser)
			
			cboSend.Focus()
			cboSend.SelectionStart = Len(cboSend.Text)
		End If
	End Sub
	
	Public Sub mnuPopClanAddToFList_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanAddToFList.Click
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		
		If Not (lvClanList.FocusedItem Is Nothing) Then
			AddQ("/f a " & CleanUsername(GetClanSelectedUser), (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
		End If
	End Sub
	
	Public Sub mnuPopClanCopy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanCopy.Click
		On Error Resume Next
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		My.Computer.Clipboard.Clear()
		
		My.Computer.Clipboard.SetText(GetClanSelectedUser)
	End Sub
	
	Public Sub mnuPopClanDemote_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanDemote.Click
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		If MsgBox("Are you sure you want to demote " & GetClanSelectedUser & "?", MsgBoxStyle.YesNo, "StealthBot") = MsgBoxResult.Yes Then
			
			With PBuffer
				.InsertDWord(&H1)
				.InsertNTString(GetClanSelectedUser)
				'UPGRADE_WARNING: Lower bound of collection lvClanList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_ISSUE: MSComctlLib.ListItem property lvClanList.ListItems.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                .InsertByte(lvClanList.Items.Item(lvClanList.FocusedItem.Index).ImageIndex - 1)
				.SendPacket(SID_CLANRANKCHANGE)
			End With
			
			AwaitingClanInfo = 1
			
		End If
	End Sub
	
	Public Sub mnuPopClanDisband_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanDisband.Click
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		If MsgBox("Are you sure you want to disband this clan?", MsgBoxStyle.YesNo Or MsgBoxStyle.Critical, "StealthBot") = MsgBoxResult.Yes Then
			With PBuffer
				.InsertDWord(&H1) '//cookie
				.SendPacket(SID_CLANDISBAND)
			End With
			
			AwaitingSelfRemoval = 1
		End If
	End Sub
	
	Public Sub mnuPopClanLeave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanLeave.Click
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		If MsgBox("Are you sure you want to leave the clan?", MsgBoxStyle.YesNo, "StealthBot") = MsgBoxResult.Yes Then
			With PBuffer
				.InsertDWord(&H1) '//cookie
				.InsertNTString(GetCurrentUsername)
				.SendPacket(SID_CLANREMOVEMEMBER)
			End With
			
			AwaitingSelfRemoval = 1
		End If
	End Sub
	
	Public Sub mnuPopClanMakeChief_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanMakeChief.Click
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		If MsgBox("Are you sure you want to make " & GetClanSelectedUser & " the new Chieftain?", MsgBoxStyle.YesNo Or MsgBoxStyle.Critical, "StealthBot") = MsgBoxResult.Yes Then
			With PBuffer
				.InsertDWord(&H1) '//cookie
				.InsertNTString(GetClanSelectedUser)
				.SendPacket(SID_CLANMAKECHIEFTAIN)
			End With
			
			AwaitingClanInfo = 1
		End If
	End Sub
	
	Public Sub mnuPopClanProfile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanProfile.Click
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		RequestProfile(GetClanSelectedUser)
		
		frmProfile.PrepareForProfile(GetClanSelectedUser, False)
	End Sub
	
	Public Sub mnuPopClanPromote_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanPromote.Click
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		If MsgBox("Are you sure you want to promote " & GetClanSelectedUser & "?", MsgBoxStyle.YesNo, "StealthBot") = MsgBoxResult.Yes Then
			With PBuffer
				.InsertDWord(&H3)
				.InsertNTString(GetClanSelectedUser)
				'UPGRADE_WARNING: Lower bound of collection lvClanList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_ISSUE: MSComctlLib.ListItem property lvClanList.ListItems.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                .InsertByte(lvClanList.Items.Item(lvClanList.FocusedItem.Index).ImageIndex + 1)
				.SendPacket(SID_CLANRANKCHANGE)
			End With
			
			AwaitingClanInfo = 1
		End If
	End Sub
	
	Public Sub mnuPopClanRemove_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanRemove.Click
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		Dim L As Integer
		L = TimeSinceLastRemoval
		
		If L < 30 Then
			AddChat(RTBColors.ErrorMessageText, "You must wait " & 30 - L & " more seconds before you " & "can remove another user from your clan.")
		Else
			If MsgBox("Are you sure you want to remove this user from the clan?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "StealthBot") = MsgBoxResult.Yes Then
				
				With PBuffer
					If lvClanList.FocusedItem.Index > 0 Then
						.InsertDWord(1) 'lvClanList.ListItems(lvClanList.SelectedItem.Index).SmallIcon
						.InsertNTString(GetClanSelectedUser)
						.SendPacket(SID_CLANREMOVEMEMBER)
					End If
					
					AwaitingClanInfo = 1
				End With
				
				LastRemoval = GetTickCount
			End If
		End If
	End Sub
	
	Public Sub mnuPopClanStatsWAR3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanStatsWAR3.Click
		Dim sProd As String
		
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		sProd = PRODUCT_WAR3
		
		If (StrComp(sProd, StrReverse(BotVars.Product), CompareMethod.Binary) = 0) Then
			sProd = vbNullString
		Else
			sProd = Space(1) & sProd
		End If
		
		AddQ("/stats " & CleanUsername(GetClanSelectedUser) & sProd, (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
	End Sub
	
	Public Sub mnuPopClanStatsW3XP_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanStatsW3XP.Click
		Dim sProd As String
		
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		sProd = PRODUCT_W3XP
		
		If (StrComp(sProd, StrReverse(BotVars.Product), CompareMethod.Binary) = 0) Then
			sProd = vbNullString
		Else
			sProd = Space(1) & sProd
		End If
		
		AddQ("/stats " & CleanUsername(GetClanSelectedUser) & sProd, (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
	End Sub
	
	Public Sub mnuPopClanUserlistWhois_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanUserlistWhois.Click
		On Error Resume Next
		
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		Dim temp As udtGetAccessResponse
		Dim s As String
		
		s = GetClanSelectedUser
		
		'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		temp = GetAccess(s)
		
		With RTBColors
			If temp.Rank > -1 Then
				If temp.Rank > 0 Then
					If temp.Flags <> vbNullString Then
						AddChat(.ConsoleText, "Found user " & s & ", with rank " & temp.Rank & " and flags " & temp.Flags & ".")
					Else
						AddChat(.ConsoleText, "Found user " & s & ", with rank " & temp.Rank & ".")
					End If
				Else
					If temp.Flags <> vbNullString Then
						AddChat(.ConsoleText, "Found user " & s & ", with flags " & temp.Flags & ".")
					Else
						AddChat(.ConsoleText, "User not found.")
					End If
				End If
			Else
				AddChat(.ConsoleText, "User not found.")
			End If
		End With
	End Sub
	
	Public Sub mnuPopClanWebProfileWAR3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanWebProfileWAR3.Click
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		GetW3LadderProfile(GetClanSelectedUser, modEnum.enuWebProfileTypes.WAR3)
	End Sub
	
	Public Sub mnuPopClanWebProfileW3XP_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanWebProfileW3XP.Click
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		GetW3LadderProfile(GetClanSelectedUser, modEnum.enuWebProfileTypes.W3XP)
	End Sub
	
	Public Sub mnuPopClanWhisper_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanWhisper.Click
		On Error Resume Next
		
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		If cboSend.Text <> vbNullString Then
			AddQ("/w " & CleanUsername(GetClanSelectedUser, True) & Space(1) & cboSend.Text, (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
			
			cboSend.Items.Insert(0, cboSend.Text)
			cboSend.Text = vbNullString
			cboSend.Focus()
		End If
	End Sub
	
	Public Sub mnuPopClanWhois_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopClanWhois.Click
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		If Not (lvClanList.FocusedItem Is Nothing) Then
			AddQ("/whois " & lvClanList.FocusedItem.Text, (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
		End If
	End Sub
	
	Public Sub mnuPopFLAddLeft_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopFLAddLeft.Click
		On Error Resume Next
		If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		Dim Index As Short
		Dim s As String
		
		If txtPre.Enabled Then 'fix for topic 25290 -a
			s = vbNullString
			If Dii Then s = "*"
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = StringFormat("/w {0}{1} ", s, GetFriendsSelectedUser)
			txtPre.Text = s
			
			cboSend.Focus()
			cboSend.SelectionStart = Len(cboSend.Text)
		End If
	End Sub
	
	Public Sub mnuPopFLCopy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopFLCopy.Click
		On Error Resume Next
		If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		My.Computer.Clipboard.Clear()
		
		My.Computer.Clipboard.SetText(GetFriendsSelectedUser)
	End Sub
	
	'Will move the selected user one spot down on the friends list.
	Public Sub mnuPopFLDemote_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopFLDemote.Click
		If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		If Not (lvFriendList.FocusedItem Is Nothing) Then
			With lvFriendList.FocusedItem
				If (.Index < lvFriendList.Items.Count) Then
					AddQ("/f d " & GetFriendsSelectedUser, (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
					'MoveFriend .index, .index + 1
				End If
			End With
		End If
	End Sub
	
	Public Sub mnuPopFLInvite_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopFLInvite.Click
		If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		Dim sPlayer As String
		
		If Not lvFriendList.FocusedItem Is Nothing Then
			sPlayer = GetFriendsSelectedUser
		End If
		
        If Len(sPlayer) > 0 Then
            If g_Clan.Self.Rank >= 3 Then
                InviteToClan((ReverseConvertUsernameGateway(sPlayer)))
                AddChat(RTBColors.InformationText, "[CLAN] Invitation sent to " & GetFriendsSelectedUser() & ", awaiting reply.")
            End If
        End If
	End Sub
	
	Public Sub mnuPopFLProfile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopFLProfile.Click
		If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		If Not lvFriendList.FocusedItem Is Nothing Then
			RequestProfile(CleanUsername(lvFriendList.FocusedItem.Text))
			
			frmProfile.PrepareForProfile(CleanUsername(lvFriendList.FocusedItem.Text), False)
		End If
	End Sub
	
	'Will move the selected user one spot up on the friends list.
	Public Sub mnuPopFLPromote_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopFLPromote.Click
		If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		If Not (lvFriendList.FocusedItem Is Nothing) Then
			With lvFriendList.FocusedItem
				If (.Index > 1) Then
					AddQ("/f p " & GetFriendsSelectedUser, (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
					'MoveFriend .index, .index - 1
				End If
			End With
		End If
	End Sub
	
	Public Sub mnuPopFLRefresh_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopFLRefresh.Click
		lvFriendList.Items.Clear()
		Call FriendListHandler.RequestFriendsList(PBuffer)
	End Sub
	
	Public Sub mnuPopFLRemove_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopFLRemove.Click
		If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		If Not (lvFriendList.FocusedItem Is Nothing) Then
			AddQ("/f r " & GetFriendsSelectedUser, (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
		End If
	End Sub
	
	Public Sub mnuPopFLStats_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopFLStats.Click
		If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		Dim aInx As Short
		Dim sProd As String
		
		aInx = lvFriendList.FocusedItem.Index
		'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends().Game. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sProd = g_Friends.Item(aInx).Game
		Select Case sProd
			Case PRODUCT_STAR, PRODUCT_SEXP, PRODUCT_W2BN, PRODUCT_WAR3, PRODUCT_W3XP, PRODUCT_JSTR, PRODUCT_SSHR
				' get stats for user on their current product
			Case Else
				' their current product does not have stats, or they are offline
				Select Case StrReverse(BotVars.Product)
					Case PRODUCT_STAR, PRODUCT_SEXP, PRODUCT_W2BN, PRODUCT_WAR3, PRODUCT_W3XP, PRODUCT_JSTR, PRODUCT_SSHR
						' get stats for user on the bot's product
						sProd = StrReverse(BotVars.Product)
					Case Else
						' unspecified product
						AddChat(RTBColors.ConsoleText, "This bot and the specified friend are not on a game that stores stats viewable via the Battle.net /stats command. " & "Type /stats " & CleanUsername(GetFriendsSelectedUser) & " <desired product code> to get this user's stats for another game.")
						Exit Sub
				End Select
		End Select
		
		If (StrComp(sProd, StrReverse(BotVars.Product), CompareMethod.Binary) = 0) Then
			sProd = vbNullString
		Else
			sProd = Space(1) & sProd
		End If
		
		AddQ("/stats " & CleanUsername(GetFriendsSelectedUser) & sProd, (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
	End Sub
	
	Public Sub mnuPopFLUserlistWhois_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopFLUserlistWhois.Click
		On Error Resume Next
		
		If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		Dim temp As udtGetAccessResponse
		Dim s As String
		
		s = GetFriendsSelectedUser
		
		'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		temp = GetAccess(s)
		
		With RTBColors
			If temp.Rank > -1 Then
				If temp.Rank > 0 Then
					If temp.Flags <> vbNullString Then
						AddChat(.ConsoleText, "Found user " & s & ", with rank " & temp.Rank & " and flags " & temp.Flags & ".")
					Else
						AddChat(.ConsoleText, "Found user " & s & ", with rank " & temp.Rank & ".")
					End If
				Else
					If temp.Flags <> vbNullString Then
						AddChat(.ConsoleText, "Found user " & s & ", with flags " & temp.Flags & ".")
					Else
						AddChat(.ConsoleText, "User not found.")
					End If
				End If
			Else
				AddChat(.ConsoleText, "User not found.")
			End If
		End With
	End Sub
	
	Public Sub mnuPopFLWebProfile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopFLWebProfile.Click
		If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		Dim aInx As Short
		Dim sProd As String
		Dim webProd As modEnum.enuWebProfileTypes
		
		aInx = lvFriendList.FocusedItem.Index
		'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends().Game. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sProd = g_Friends.Item(aInx).Game
		Select Case sProd
			' get web profile for user on their current product
			Case PRODUCT_WAR3
				webProd = modEnum.enuWebProfileTypes.WAR3
			Case PRODUCT_W3XP
				webProd = modEnum.enuWebProfileTypes.W3XP
			Case Else
				Select Case StrReverse(BotVars.Product)
					' get web profile for user on the bot's product
					Case PRODUCT_WAR3
						webProd = modEnum.enuWebProfileTypes.WAR3
					Case PRODUCT_W3XP
						webProd = modEnum.enuWebProfileTypes.W3XP
					Case Else
						' their current product does not have stats, or they are offline
						AddChat(RTBColors.ConsoleText, "The specified friend must be online to decide which web profile to view.")
						Exit Sub
				End Select
		End Select
		
		GetW3LadderProfile(CleanUsername(GetFriendsSelectedUser), webProd)
	End Sub
	
	Public Sub mnuPopFLWhisper_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopFLWhisper.Click
		Dim Value As String
		
		If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		Value = cboSend.Text
		
        If Len(Value) > 0 Then
            Value = "/w " & CleanUsername(GetFriendsSelectedUser, True) & Space(1) & Value

            AddQ(Value, (modQueueObj.PRIORITY.CONSOLE_MESSAGE))

            cboSend.Items.Insert(0, Value)
            cboSend.Text = vbNullString

            On Error Resume Next
            cboSend.Focus()
        End If
	End Sub
	
	Public Sub mnuPopInvite_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopInvite.Click
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		Dim sPlayer As String
		
		If Not lvChannel.FocusedItem Is Nothing Then
			sPlayer = GetSelectedUser
		End If
		
        If Len(sPlayer) > 0 Then
            If g_Clan.Self.Rank >= 3 Then
                InviteToClan((ReverseConvertUsernameGateway(sPlayer)))
                AddChat(RTBColors.InformationText, "[CLAN] Invitation sent to " & GetSelectedUser() & ", awaiting reply.")
            End If
        End If
	End Sub
	
	Public Sub mnuPopProfile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopProfile.Click
		On Error Resume Next
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		Dim sUser As String
		sUser = StripAccountNumber(CleanUsername(GetSelectedUser))
		
		RequestProfile(sUser)
		
		frmProfile.PrepareForProfile(sUser, False)
	End Sub
	
	Public Sub mnuPopStats_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopStats.Click
		Dim aInx As Short
		Dim sProd As String
		
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		
		aInx = g_Channel.GetUserIndex(GetSelectedUser)
		'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().Game. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sProd = g_Channel.Users.Item(aInx).Game
		Select Case sProd
			Case PRODUCT_STAR, PRODUCT_SEXP, PRODUCT_W2BN, PRODUCT_WAR3, PRODUCT_W3XP, PRODUCT_JSTR, PRODUCT_SSHR
				' get stats for user on their current product
			Case Else
				' unspecified product
				AddChat(RTBColors.ConsoleText, "The specified user is not on a game that stores stats viewable via the Battle.net /stats command. " & "Type /stats " & CleanUsername(GetSelectedUser) & " <desired product code> to get this user's stats for another game.")
				Exit Sub
		End Select
		
		If (StrComp(sProd, StrReverse(BotVars.Product), CompareMethod.Binary) = 0) Then
			sProd = vbNullString
		Else
			sProd = Space(1) & sProd
		End If
		
		AddQ("/stats " & StripAccountNumber(CleanUsername(GetSelectedUser)) & sProd, (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
	End Sub
	
	Public Sub mnuPopUserlistWhois_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopUserlistWhois.Click
		On Error Resume Next
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		
		Dim temp As udtGetAccessResponse
		Dim s As String
		
		s = GetSelectedUser
		
		'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		temp = GetAccess(s)
		
		With RTBColors
			If temp.Rank > -1 Then
				If temp.Rank > 0 Then
					If temp.Flags <> vbNullString Then
						AddChat(.ConsoleText, "Found user " & s & ", with rank " & temp.Rank & " and flags " & temp.Flags & ".")
					Else
						AddChat(.ConsoleText, "Found user " & s & ", with rank " & temp.Rank & ".")
					End If
				Else
					If temp.Flags <> vbNullString Then
						AddChat(.ConsoleText, "Found user " & s & ", with flags " & temp.Flags & ".")
					Else
						AddChat(.ConsoleText, "User not found.")
					End If
				End If
			Else
				AddChat(.ConsoleText, "User not found.")
			End If
		End With
	End Sub
	
	Public Sub mnuHomeChannel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHomeChannel.Click
        If (Len(Config.HomeChannel) = 0) Then
            ' do product home join instead
            Call DoChannelJoinProductHome()
        Else
            ' go home
            Call FullJoin((Config.HomeChannel), 2)
        End If
	End Sub
	
	Public Sub mnuLastChannel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLastChannel.Click
        If (Len(BotVars.LastChannel) = 0) Then
            ' do product home join instead
            Call DoChannelJoinProductHome()
        Else
            ' go to last
            Call FullJoin((BotVars.LastChannel), 2)
        End If
	End Sub
	
	Public Sub mnuRealmSwitch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRealmSwitch.Click
		If Dii Then
			If ds.MCPHandler Is Nothing Then
				ds.MCPHandler = New clsMCPHandler
				
				SEND_SID_QUERYREALMS2()
			Else
				Call ds.MCPHandler.HandleQueryRealmServersResponse()
			End If
		End If
	End Sub
	
	Public Sub mnuPublicChannels_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPublicChannels.Click
		Dim Index As Short = mnuPublicChannels.GetIndex(eventSender)
		' some public channels are redirects
		'If (StrComp(mnuPublicChannels(Index).Caption, g_Channel.Name, vbTextCompare) = 0) Then
		'    Exit Sub
		'End If
		
		If Not BotVars.PublicChannels Is Nothing Then
			If BotVars.PublicChannels.Count() > Index Then
				Select Case Config.AutoCreateChannels
					Case "ALERT", "NEVER"
						'UPGRADE_WARNING: Couldn't resolve default property of object BotVars.PublicChannels.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call FullJoin(BotVars.PublicChannels.Item(Index + 1), 0)
					Case Else ' "ALWAYS"
						'UPGRADE_WARNING: Couldn't resolve default property of object BotVars.PublicChannels.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Call FullJoin(BotVars.PublicChannels.Item(Index + 1), 2)
				End Select
			End If
			'AddQ "/join " & PublicChannels.Item(Index + 1), PRIORITY.CONSOLE_MESSAGE
		End If
	End Sub
	
	Public Sub mnuCustomChannels_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCustomChannels.Click
		Dim Index As Short = mnuCustomChannels.GetIndex(eventSender)
		If (StrComp(QC(Index + 1), g_Channel.Name, CompareMethod.Text) = 0) Then
			Exit Sub
		End If
		
		Select Case Config.AutoCreateChannels
			Case "ALERT", "NEVER"
				Call FullJoin(QC(Index + 1), 0)
			Case Else ' "ALWAYS"
				Call FullJoin(QC(Index + 1), 2)
		End Select
		
		'AddQ "/join " & QC(Index + 1), PRIORITY.CONSOLE_MESSAGE
	End Sub
	
	Public Sub mnuCommandManager_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCommandManager.Click
		'frmCommands.Show vbModal, Me
		frmCommands.Show()
	End Sub
	
	Public Sub mnuConnect2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuConnect2.Click
		Call DoConnect()
	End Sub
	
	Public Sub mnuDisableVoidView_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDisableVoidView.Click
		mnuDisableVoidView.Checked = Not (mnuDisableVoidView.Checked)
		Config.VoidView = Not CBool(mnuDisableVoidView.Checked)
		Call Config.Save()
		
		If Config.VoidView And g_Channel.IsSilent Then
			tmrSilentChannel(1).Enabled = Config.VoidView
			AddQ("/unignore " & GetCurrentUsername)
		End If
	End Sub
	
	Public Sub mnuDisconnect2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDisconnect2.Click
		Dim Key As String
		Dim L As Integer
		Key = GetProductKey()
		
		'    If AttemptedNewVerbyte Then
		'        AttemptedNewVerbyte = False
		'        l = CLng(Val("&H" & ReadCFG("Main", Key & "VerByte")))
		'        WriteINI "Main", Key & "VerByte", Hex(l - 1)
		'    End If
		
		GErrorHandler.Reset_Renamed()
		Call DoDisconnect()
	End Sub
	
	Public Sub mnuEditCaught_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditCaught.Click
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Dir(GetFilePath(FILE_CAUGHT_PHRASES)) = vbNullString Then
			MsgBox("The bot has not caught any phrases yet.")
			Exit Sub
		Else
			ShellOpenURL(GetFilePath(FILE_CAUGHT_PHRASES),  , False)
		End If
	End Sub
	
	Public Sub mnuFlash_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFlash.Click
		mnuFlash.Checked = (Not mnuFlash.Checked)
		Config.FlashOnEvents = CBool(mnuFlash.Checked)
		Call Config.Save()
	End Sub
	
	'Moves a person in the friends list view.
	Private Sub MoveFriend(ByRef startPos As Short, ByRef endPos As Short)
		With lvFriendList.Items
			If (startPos > endPos) Then
				'UPGRADE_WARNING: Lower bound of collection lvFriendList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_ISSUE: MSComctlLib.ListItem property lvFriendList.ListItems.Item.SmallIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_WARNING: Lower bound of collection lvFriendList.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
                .Insert(endPos, .Item(startPos).Text, .Item(startPos).ImageIndex)
				'UPGRADE_WARNING: Lower bound of collection lvFriendList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_WARNING: Lower bound of collection lvFriendList.ListItems.Item().ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_ISSUE: MSComctlLib.ListSubItem property lvFriendList.ListItems.Item.ListSubItems.Item.ReportIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: MSComctlLib.ListSubItems method lvFriendList.ListItems.Item.ListSubItems.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                .Item(endPos).SubItems.Add(.Item(startPos + 1).SubItems.Item(1).Text)
				.RemoveAt(startPos + 1)
			Else
				'UPGRADE_WARNING: Lower bound of collection lvFriendList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_ISSUE: MSComctlLib.ListItem property lvFriendList.ListItems.Item.Icon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_WARNING: Lower bound of collection lvFriendList.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_WARNING: SmallIcon parameter was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B392B480-6E26-492E-84C3-16C21D8D0807"'
                .Insert(endPos + 1, .Item(startPos))
				'UPGRADE_WARNING: Lower bound of collection lvFriendList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_WARNING: Lower bound of collection lvFriendList.ListItems.Item().ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_ISSUE: MSComctlLib.ListSubItem property lvFriendList.ListItems.Item.ListSubItems.Item.ReportIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: MSComctlLib.ListSubItems method lvFriendList.ListItems.Item.ListSubItems.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                .Item(endPos + 1).SubItems.Add(.Item(startPos).SubItems.Item(1).Text)
				.RemoveAt(startPos)
			End If
		End With
	End Sub
	
	Public Sub mnuGetNews_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuGetNews.Click
		If Not RequestINetPage(GetNewsURL(), SB_INET_NEWS, False) Then
			Call HandleNews("INet is busy", -1)
		End If
	End Sub
	
	Public Sub mnuHelpReadme_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHelpReadme.Click
		OpenReadme()
	End Sub
	
	Public Sub mnuHelpWebsite_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHelpWebsite.Click
		ShellOpenURL("http://www.stealthbot.net", "the StealthBot Forum")
	End Sub
	
	Public Sub mnuHideBans_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHideBans.Click
		mnuHideBans.Checked = (Not mnuHideBans.Checked)
		Config.HideBanMessages = CBool(mnuHideBans.Checked)
		Call Config.Save()
		AddChat(RTBColors.InformationText, "Ban messages " & IIf(mnuHideBans.Checked, "disabled", "enabled") & ".")
	End Sub
	
	Public Sub mnuHideWhispersInrtbChat_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHideWhispersInrtbChat.Click
		mnuHideWhispersInrtbChat.Checked = (Not mnuHideWhispersInrtbChat.Checked)
		Config.HideWhispersInMain = CBool(mnuHideWhispersInrtbChat.Checked)
		Call Config.Save()
	End Sub
	
	Public Sub mnuIgnoreInvites_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuIgnoreInvites.Click
		If mnuIgnoreInvites.Checked Then
			mnuIgnoreInvites.Checked = False
		Else
			mnuIgnoreInvites.Checked = True
		End If
		
		Config.IgnoreClanInvites = CBool(mnuIgnoreInvites.Checked)
		Call Config.Save()
	End Sub
	
	Public Sub mnuLog0_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLog0.Click
		BotVars.Logging = 2
		Config.LoggingMode = BotVars.Logging
		Call Config.Save()
		
		AddChat(RTBColors.InformationText, "Full text logging enabled.")
		mnuLog1.Checked = False
		mnuLog0.Checked = True
		mnuLog2.Checked = False
		'mnuLog3.Checked = False
		
		'MakeLoggingDirectory
	End Sub
	
	Public Sub mnuLog1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLog1.Click
		BotVars.Logging = 1
		Config.LoggingMode = BotVars.Logging
		Call Config.Save()
		
		AddChat(RTBColors.InformationText, "Partial text logging enabled.")
		mnuLog1.Checked = True
		mnuLog0.Checked = False
		mnuLog2.Checked = False
		'mnuLog3.Checked = False
		
		'MakeLoggingDirectory
	End Sub
	
	Public Sub mnuLog2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLog2.Click
		BotVars.Logging = 0
		Config.LoggingMode = BotVars.Logging
		Call Config.Save()
		
		AddChat(RTBColors.InformationText, "Logging disabled.")
		mnuLog1.Checked = False
		mnuLog0.Checked = False
		mnuLog2.Checked = True
		'mnuLog3.Checked = False
	End Sub
	
	Public Sub mnuOpenBotFolder_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOpenBotFolder.Click
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Shell(StringFormat("explorer.exe {0}", CurDir()), AppWinStyle.NormalFocus)
	End Sub
	
	
	Public Sub mnuPacketLog_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPacketLog.Click
		Dim f As Short
		
		If mnuPacketLog.Checked Then
			' turning this feature off
			AddChat(RTBColors.SuccessText, "StealthBot packet traffic will no longer be logged.")
		Else
			' turning it on
			AddChat(RTBColors.SuccessText, "StealthBot packet traffic will be logged in the bot's folder, in a file named " & VB6.Format(Today, "yyyy-MM-dd") & "-PacketLog.txt.")
			AddChat(RTBColors.SuccessText, "--")
			AddChat(RTBColors.SuccessText, "Log packets at your own risk! Please read the note below:")
			AddChat(RTBColors.ErrorMessageText, "*** CAUTION: THIS LOG MAY CONTAIN PRIVATE INFORMATION.")
			AddChat(RTBColors.ErrorMessageText, "*** CAUTION: DO NOT DISTRIBUTE it in public posts on StealthBot.net or on any other website!")
			AddChat(RTBColors.ErrorMessageText, "*** CAUTION: Only produce a packet log if you're specifically instructed to by")
			AddChat(RTBColors.ErrorMessageText, "*** CAUTION: a StealthBot.net tech, or you know what you're doing!")
			AddChat(RTBColors.SuccessText, "If you wish to stop logging packets, uncheck the menu item or restart your bot.")
			AddChat(RTBColors.SuccessText, "This feature only logs StealthBot traffic. It is not a system-wide packet capture utility.")
		End If
		
		mnuPacketLog.Checked = Not mnuPacketLog.Checked
		LogPacketTraffic = mnuPacketLog.Checked
	End Sub
	
	Public Sub mnuPopAddLeft_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopAddLeft.Click
		On Error Resume Next
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		Dim Index As Short
		Dim s As String
		
		If txtPre.Enabled Then 'fix for topic 25290 -a
			s = vbNullString
			If Dii Then s = "*"
			'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = StringFormat("/w {0}{1} ", s, GetSelectedUser)
			txtPre.Text = s
			
			cboSend.Focus()
			cboSend.SelectionStart = Len(cboSend.Text)
		End If
	End Sub
	
	Public Sub mnuPopAddToFList_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopAddToFList.Click
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		
		If Not (lvChannel.FocusedItem Is Nothing) Then
			AddQ("/f a " & StripAccountNumber(CleanUsername(GetSelectedUser)), (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
		End If
	End Sub
	
	Public Sub mnuPopDes_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopDes.Click
		On Error Resume Next
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		
		AddQ("/designate " & CleanUsername(GetSelectedUser, True), (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
	End Sub
	
	Public Sub mnuPopFLWhois_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopFLWhois.Click
		If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
		
		If Not (lvFriendList.FocusedItem Is Nothing) Then
			AddQ("/whois " & lvFriendList.FocusedItem.Text, (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
		End If
	End Sub
	
	Public Sub mnuPopSafelist_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopSafelist.Click
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		
		Dim gAcc As udtGetAccessResponse
		Dim toSafe As String
		
		On Error Resume Next
		
		toSafe = GetSelectedUser
		
		gAcc.Rank = 1000
		
		Call ProcessCommand(GetCurrentUsername, "/safeadd " & toSafe, True, False)
	End Sub
	
	Public Sub mnuPopShitlist_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopShitlist.Click
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		
		Dim gAcc As udtGetAccessResponse
		Dim toBan As String
		
		On Error Resume Next
		
		toBan = GetSelectedUser
		
		gAcc.Rank = 1000
		
		Call ProcessCommand(GetCurrentUsername, "/shitadd " & toBan, True, False)
	End Sub
	
	Public Sub mnuPopSquelch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopSquelch.Click
		On Error Resume Next
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		
		AddQ("/squelch " & GetSelectedUser, (modQueueObj.PRIORITY.CONSOLE_MESSAGE), CStr(modQueueObj.PRIORITY.CONSOLE_MESSAGE))
	End Sub
	
	
	Public Sub mnuPopUnsquelch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopUnsquelch.Click
		'On Error Resume Next
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		
		AddQ("/unsquelch " & GetSelectedUser, (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
	End Sub
	
	Public Sub mnuPopWhisper_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopWhisper.Click
		On Error Resume Next
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		
		If cboSend.Text <> vbNullString Then
			AddQ("/w " & CleanUsername(GetSelectedUser, True) & Space(1) & cboSend.Text, (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
			
			cboSend.Items.Insert(0, cboSend.Text)
			cboSend.Text = vbNullString
			cboSend.Focus()
		End If
	End Sub
	
	Public Sub mnuClear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuClear.Click
		Call ClearChatScreen(3)
	End Sub
	
	Public Sub mnuClearWW_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuClearWW.Click
		Call ClearChatScreen(2)
	End Sub
	
	' clear the chat screen:
	' 1 - clear chat window only
	' 2 - clear whisper window only
	' 3 (default) - clear chat and whisper
	Sub ClearChatScreen(Optional ByVal ClearOption As Short = 3)
		' if they passed 0 (False), change to 3
		If ClearOption = 0 Then ClearOption = 3
		' if they passed -1 (True), change to 1 (old behavior: DoNotClearWhispers)
		If ClearOption = -1 Then ClearOption = 1
		' check for 2 (or 3) and clear whispers
		If ClearOption And 2 Then
			rtbWhispers.Text = vbNullString
			' add cleared message
			AddWhisper(RTBColors.ConsoleText, ">> Whisper window cleared.")
		End If
		
		' check for 1 (or 3) and clear chats
		If ClearOption And 1 Then
			rtbChat.Text = vbNullString
			rtbChatLength = 0
			' add a sensical cleared message
			If ClearOption And 2 Then
				AddChat(RTBColors.InformationText, "Chat window cleared.")
			Else
				AddChat(RTBColors.InformationText, "Chat and whisper windows cleared.")
			End If
		End If
		
		' set focus to send box
		On Error Resume Next
		cboSend.Focus()
	End Sub
	
	Public Sub mnuPopWhois_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopWhois.Click
		On Error Resume Next
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		
		AddQ("/whois " & CleanUsername(GetSelectedUser, True), (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
	End Sub
	
	Public Sub mnuPopWebProfile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopWebProfile.Click
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		
		Dim aInx As Short
		Dim sProd As String
		Dim webProd As modEnum.enuWebProfileTypes
		
		aInx = lvChannel.FocusedItem.Index
		'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().Game. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sProd = g_Channel.Users.Item(aInx).Game
		Select Case sProd
			' get web profile for user on their current product
			Case PRODUCT_WAR3
				webProd = modEnum.enuWebProfileTypes.WAR3
			Case PRODUCT_W3XP
				webProd = modEnum.enuWebProfileTypes.W3XP
			Case Else
				' their current product does not have a web profile
				AddChat(RTBColors.ConsoleText, "The specified user must be on WarCraft III to view their web profile.")
				Exit Sub
		End Select
		
		GetW3LadderProfile(StripAccountNumber(CleanUsername(GetSelectedUser)), webProd)
	End Sub
	
	Public Sub mnuClearedTxt_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuClearedTxt.Click
		Dim sPath As String
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sPath = StringFormat("{0}{1}.txt", g_Logger.LogPath, VB6.Format(Today, "yyyy-MM-dd"))
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Len(Dir(sPath)) = 0 Then
            AddChat(RTBColors.ErrorMessageText, "The log file for today is empty.")
        Else
            ShellOpenURL(sPath, , False)
        End If
	End Sub
	
	Public Sub mnuRecordWindowPos_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRecordWindowPos.Click
		RecordWindowPosition()
	End Sub
	
	Public Sub mnuRepairCleanMail_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRepairCleanMail.Click
		CleanUpMailFile()
		Me.AddChat(RTBColors.SuccessText, "Delivered and invalid pieces of mail have been removed from your mail.dat file.")
	End Sub
	
	Public Sub mnuRepairDataFiles_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRepairDataFiles.Click
		If MsgBox("Are you sure? This action will delete your mail.dat (Bot mail database) file.", MsgBoxStyle.YesNo, "Repair data files") = MsgBoxResult.Yes Then
			On Error Resume Next
			Kill(GetFilePath(FILE_MAILDB))
			AddChat(RTBColors.SuccessText, "The bot's DAT data files have been removed.")
		End If
	End Sub
	
	Public Sub mnuRepairVerbytes_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRepairVerbytes.Click
		Dim Index As Short
		
		For Index = 0 To UBound(ProductList)
			If ProductList(Index).BNLS_ID > 0 Then
				Config.SetVersionByte(ProductList(Index).ShortCode, GetVerByte(ProductList(Index).Code, 1))
			End If
		Next 
		
		Call Config.Save()
		
		Me.AddChat(RTBColors.SuccessText, "The version bytes stored in config.ini have been restored to their defaults.")
	End Sub
	
	Private Sub mnuScripts_Click()
		'do nothing
	End Sub
	
	Public Sub mnuToggleShowOutgoing_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuToggleShowOutgoing.Click
		mnuToggleShowOutgoing.Checked = (Not mnuToggleShowOutgoing.Checked)
		Config.ShowOutgoingWhispers = CBool(mnuToggleShowOutgoing.Checked)
		Call Config.Save()
	End Sub
	
	Public Sub mnuToggleWWUse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuToggleWWUse.Click
		mnuToggleWWUse.Checked = (Not mnuToggleWWUse.Checked)
		
		Config.WhisperWindows = CBool(mnuToggleWWUse.Checked)
		Call Config.Save()
		
		If Not mnuToggleWWUse.Checked Then
			DestroyAllWWs()
		End If
	End Sub
	
	Public Sub mnuUpdateVerbytes_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuUpdateVerbytes.Click
		If Not RequestINetPage(VERBYTE_SOURCE, SB_INET_VBYTE, False) Then
			Call HandleUpdateVerbyte("INet is busy", -1)
		End If
	End Sub
	
	Public Sub mnuWhisperCleared_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuWhisperCleared.Click
		Dim sPath As String
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sPath = StringFormat("{0}{1}-WHISPERS.txt", g_Logger.LogPath, VB6.Format(Today, "yyyy-MM-dd"))
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Len(Dir(sPath)) = 0 Then
            AddChat(RTBColors.ErrorMessageText, "The whisper log file for today is empty.")
        Else
            ShellOpenURL(sPath, , False)
        End If
	End Sub
	
	Private Sub mnuEditUsers_Click()
		ShellOpenURL(GetFilePath(FILE_USERDB),  , False)
	End Sub
	
	Public Sub mnuReloadScripts_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReloadScripts.Click
		
		On Error GoTo ERROR_HANDLER
		
		'RunInAll "Event_LoggedOff"
		RunInAll("Event_Close")
		
		InitScriptControl(SControl)
		LoadScripts()
		
		InitScripts()
		
		Exit Sub
		
ERROR_HANDLER: 
		
		' Cannot call this method while the script is executing
		If (Err.Number = -2147467259) Then
			Me.AddChat(RTBColors.ErrorMessageText, "Error: Script is still executing.")
			
			Exit Sub
		End If
		
		Me.AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in mnuReloadScripts_Click().")
		
	End Sub
	
	Public Sub mnuSetTop_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSetTop.Click
		mnuLog0.Checked = False
		mnuLog1.Checked = False
		mnuLog2.Checked = False
		
		Select Case BotVars.Logging
			Case 2 : mnuLog0.Checked = True
			Case 1 : mnuLog1.Checked = True
			Case 0 : mnuLog2.Checked = True
		End Select
	End Sub
	
	Public Sub mnuTerms_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuTerms.Click
		ShellOpenURL("http://eula.stealthbot.net", "the StealthBot EULA")
	End Sub
	
	Public Sub mnuFilters_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFilters.Click
		frmFilters.Show()
	End Sub
	
	Public Sub mnuPopCopy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopCopy.Click
		On Error Resume Next
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		
		My.Computer.Clipboard.Clear()
		
		My.Computer.Clipboard.SetText(GetSelectedUser)
	End Sub
	
	Public Sub mnuProfile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuProfile.Click
		frmProfile.PrepareForProfile(vbNullString, True)
	End Sub
	
	Public Sub mnuCustomChannelAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCustomChannelAdd.Click
		Dim i As Short
		
        If Len(g_Channel.Name) > 0 Then

            For i = LBound(QC) To UBound(QC)
                If Len(Trim(QC(i))) = 0 Then
                    QC(i) = g_Channel.Name
                    PrepareQuickChannelMenu()

                    Exit Sub
                End If
            Next i

        End If
	End Sub
	
	Public Sub mnuCustomChannelEdit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCustomChannelEdit.Click
		frmQuickChannel.Show()
	End Sub
	
	Public Sub mnuReload_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuReload.Click
		On Error Resume Next
		Call ReloadConfig(1)
		AddChat(RTBColors.SuccessText, "Configuration file loaded.")
	End Sub
	
	Public Sub mnuUTF8_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuUTF8.Click
		If mnuUTF8.Checked Then
			mnuUTF8.Checked = False
			AddChat(RTBColors.ConsoleText, "Messages will no longer be UTF-8-decoded.")
		Else
			mnuUTF8.Checked = True
			AddChat(RTBColors.ConsoleText, "Messages will now be UTF-8-decoded.")
		End If
		
		Config.UseUTF8 = CBool(mnuUTF8.Checked)
		Call Config.Save()
	End Sub
	
	Private Sub rtbChat_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles rtbChat.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		If (Shift = VB6.ShiftConstants.CtrlMask) And ((KeyCode = System.Windows.Forms.Keys.L) Or (KeyCode = System.Windows.Forms.Keys.E) Or (KeyCode = System.Windows.Forms.Keys.R)) Then
			'Call Ctrl+L and Ctrl+R keyboard shortcuts as they code to automatically handle them will be canceled out below
			Select Case KeyCode
				Case System.Windows.Forms.Keys.L
					Call mnuLock_Click(mnuLock, New System.EventArgs())
				Case System.Windows.Forms.Keys.R
					Call mnuReloadScripts_Click(mnuReloadScripts, New System.EventArgs())
			End Select
			
			'Disable Ctrl+L, Ctrl+E, and Ctrl+R
			KeyCode = 0
		ElseIf (Shift = VB6.ShiftConstants.ShiftMask) And (KeyCode = System.Windows.Forms.Keys.Delete) Then 
			'Call Shift+DEL keyboard shortcut since it doens't work with RTB focus.
			Call mnuClear_Click(mnuClear, New System.EventArgs())
		End If
	End Sub
	
	Private Sub rtbChat_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles rtbChat.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If (KeyAscii < 32) Then
			GoTo EventExitSub
		End If
		
		cboSend.Focus()
		cboSend.SelectedText = Chr(KeyAscii)
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub rtbWhispers_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles rtbWhispers.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		If (Shift = VB6.ShiftConstants.CtrlMask) And ((KeyCode = System.Windows.Forms.Keys.L) Or (KeyCode = System.Windows.Forms.Keys.E) Or (KeyCode = System.Windows.Forms.Keys.R)) Then
			'Call Ctrl+L and Ctrl+R keyboard shortcuts as they code to automatically handle them will be canceled out below
			Select Case KeyCode
				Case System.Windows.Forms.Keys.L
					Call mnuLock_Click(mnuLock, New System.EventArgs())
				Case System.Windows.Forms.Keys.R
					Call mnuReloadScripts_Click(mnuReloadScripts, New System.EventArgs())
			End Select
			
			'Disable Ctrl+L, Ctrl+E, and Ctrl+R
			KeyCode = 0
		ElseIf (Shift = VB6.ShiftConstants.ShiftMask) And (KeyCode = System.Windows.Forms.Keys.Delete) Then 
			'Call Shift+DEL keyboard shortcut since it doens't work with RTB focus.
			Call mnuClear_Click(mnuClear, New System.EventArgs())
		End If
	End Sub
	
	Private Sub rtbWhispers_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles rtbWhispers.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If (KeyAscii < 32) Then
			GoTo EventExitSub
		End If
		
		cboSend.Focus()
		cboSend.SelectedText = Chr(KeyAscii)
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub rtbChat_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles rtbChat.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		On Error Resume Next
		
		If Button = 1 And Len(rtbChat.SelectedText) > 0 Then
			If Not BotVars.NoRTBAutomaticCopy Then
				My.Computer.Clipboard.Clear()
				My.Computer.Clipboard.SetText(rtbChat.SelectedText)
			End If
		End If
	End Sub
	
	Private Sub rtbWhispers_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles rtbWhispers.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		On Error Resume Next
		
		If Button = 1 And Len(rtbWhispers.SelectedText) > 0 Then
			My.Computer.Clipboard.Clear()
			My.Computer.Clipboard.SetText(rtbWhispers.SelectedText)
		End If
	End Sub
	
	Public Sub mnuToggleFilters_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuToggleFilters.Click
		mnuToggleFilters.Checked = (Not (mnuToggleFilters.Checked))
		
		If Filters Then
			Filters = False
			AddChat(RTBColors.InformationText, "Chat filtering disabled.")
		Else
			Filters = True
			AddChat(RTBColors.InformationText, "Chat filtering enabled.")
		End If
		
		Config.ChatFilters = Filters
		Call Config.Save()
	End Sub
	
	Public Sub mnuConnect_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuConnect.Click
		GErrorHandler.Reset_Renamed()
		Call DoConnect()
	End Sub
	
	Public Sub mnuPopKick_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopKick.Click
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		
		If MyFlags = 2 Or MyFlags = 18 Then
			AddQ("/kick " & CleanUsername(GetSelectedUser, True), (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
		End If
	End Sub
	
	Public Sub mnuPopBan_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPopBan.Click
		If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
		
		If MyFlags = 2 Or MyFlags = 18 Then
			AddQ("/ban " & CleanUsername(GetSelectedUser, True), (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
		End If
	End Sub
	
	Public Sub mnuTrayExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuTrayExit.Click
		Dim Result As MsgBoxResult
		
		Result = MsgBox("Are you sure you want to quit?", MsgBoxStyle.YesNo, "StealthBot")
		
		If (Result = MsgBoxResult.Yes) Then
			'frmChat.Show
			
			'UnhookWindowProc
			'RESTORE FORM
			'Call NewWindowProc(frmChat.hWnd, 0&, ID_TASKBARICON, WM_LBUTTONDOWN)
			'Call Form_Unload(0)
			Me.Close()
		End If
	End Sub
	
	Public Sub mnuRestore_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRestore.Click
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Show()
	End Sub
	
	Public Sub mnuLock_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLock.Click
		mnuLock.Checked = (Not (mnuLock.Checked))
		
		If BotVars.LockChat = False Then
			AddChat(RTBColors.InformationText, "Chat window locked.")
			AddChat(RTBColors.ErrorMessageText, "NO MESSAGES WHATSOEVER WILL BE DISPLAYED UNTIL YOU UNLOCK THE WINDOW.")
			AddChat(RTBColors.ErrorMessageText, "To return to normal mode, press CTRL+L or use the toggle under the Window menu.")
			BotVars.LockChat = True
		Else
			BotVars.LockChat = False
			AddChat(RTBColors.SuccessText, "Chat window unlocked.")
		End If
	End Sub
	
	Public Sub mnuDisconnect_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDisconnect.Click
		Dim Key As String
		Dim L As Integer
		Key = GetProductKey()
		
		'    If AttemptedNewVerbyte Then
		'        AttemptedNewVerbyte = False
		'        l = CLng(Val("&H" & ReadCFG("Main", Key & "VerByte")))
		'        WriteINI "Main", Key & "VerByte", Hex(l - 1)
		'    End If
		
		GErrorHandler.Reset_Renamed()
		Call DoDisconnect()
	End Sub
	
	Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click
		'Call Form_Unload(0)
		Me.Close()
	End Sub
	
	Public Sub mnuSetup_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSetup.Click
		If (SettingsForm Is Nothing) Then
			SettingsForm = New frmSettings
		End If
		
		SettingsForm.Show()
	End Sub
	
	Public Sub mnuAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAbout.Click
		frmAbout.Show()
	End Sub
	
	Public Sub mnuToggle_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuToggle.Click
		mnuToggle.Checked = (Not (mnuToggle.Checked))
		
		If JoinMessagesOff = False Then
			AddChat(RTBColors.InformationText, "Join/Leave messages disabled.")
			JoinMessagesOff = True
		Else
			AddChat(RTBColors.InformationText, "Join/Leave messages enabled.")
			JoinMessagesOff = False
		End If
		
		Config.ShowJoinLeaves = Not JoinMessagesOff
		Call Config.Save()
	End Sub
	
	Public Sub mnuUsers_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuUsers.Click
		frmDBManager.Show()
	End Sub
	
	Private Sub mnuScript_Click(ByRef Index As Short)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Menu", Index)
		RunInSingle(obj.SCModule, obj.ObjName & "_Click")
	End Sub
	
	Private Sub sckScript_ConnectEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sckScript.ConnectEvent
		Dim Index As Short = sckScript.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Winsock", Index)
		RunInSingle(obj.SCModule, obj.ObjName & "_Connect")
	End Sub
	
	Private Sub sckScript_ConnectionRequest(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_ConnectionRequestEvent) Handles sckScript.ConnectionRequest
		Dim Index As Short = sckScript.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Winsock", Index)
		RunInSingle(obj.SCModule, obj.ObjName & "_ConnectionRequest", eventArgs.requestID)
	End Sub
	
	Private Sub sckScript_DataArrival(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_DataArrivalEvent) Handles sckScript.DataArrival
		Dim Index As Short = sckScript.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Winsock", Index)
		RunInSingle(obj.SCModule, obj.ObjName & "_DataArrival", eventArgs.bytesTotal)
	End Sub
	
	Private Sub sckScript_SendComplete(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sckScript.SendComplete
		Dim Index As Short = sckScript.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Winsock", Index)
		RunInSingle(obj.SCModule, obj.ObjName & "_SendComplete")
	End Sub
	
	Private Sub sckScript_SendProgress(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_SendProgressEvent) Handles sckScript.SendProgress
		Dim Index As Short = sckScript.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Winsock", Index)
		RunInSingle(obj.SCModule, obj.ObjName & "_SendProgress", eventArgs.bytesSent, eventArgs.bytesRemaining)
	End Sub
	
	Private Sub sckScript_CloseEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sckScript.CloseEvent
		Dim Index As Short = sckScript.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Winsock", Index)
		RunInSingle(obj.SCModule, obj.ObjName & "_Close")
	End Sub
	
	Private Sub sckScript_Error(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_ErrorEvent) Handles sckScript.Error
		Dim Index As Short = sckScript.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Winsock", Index)
		RunInSingle(obj.SCModule, obj.ObjName & "_Error", eventArgs.Number, eventArgs.Description, eventArgs.sCode, eventArgs.source, eventArgs.HelpFile, eventArgs.HelpContext, eventArgs.CancelDisplay)
	End Sub
	
	Private Sub itcScript_StateChanged(ByVal eventSender As System.Object, ByVal eventArgs As AxInetCtlsObjects.DInetEvents_StateChangedEvent) Handles itcScript.StateChanged
		Dim Index As Short = itcScript.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Inet", Index)
		RunInSingle(obj.SCModule, obj.ObjName & "_StateChanged", eventArgs.State)
	End Sub
	
	Private Sub tmrAccountLock_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrAccountLock.Tick
		tmrAccountLock.Enabled = False
		
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If (Not sckBNet.CtlState = MSWinsockLib.StateConstants.sckConnected) Then 'g_online is set to true AFTER we login... makes this moot, changed to socket being connected.
			Exit Sub
		End If
		
		AddChat(RTBColors.ErrorMessageText, "[BNCS] Your account appears to be locked, likely due to an excessive number of " & "invalid logins.  Please try connecting again in 15-20 minutes.")
		
		DoDisconnect()
	End Sub
	
	Private Sub tmrScript_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrScript.Tick
		Dim Index As Short = tmrScript.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("Timer", Index)
		RunInSingle(obj.SCModule, obj.ObjName & "_Timer")
	End Sub
	
	Private Sub tmrScriptLong_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrScriptLong.Tick
		Dim Index As Short = tmrScriptLong.GetIndex(eventSender)
		On Error Resume Next
		
		Dim obj As scObj
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj = GetScriptObjByIndex("LongTimer", Index)
		'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.Counter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		obj.obj.Counter = (obj.obj.Counter + 1)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.Interval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.Counter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (obj.obj.Counter >= obj.obj.Interval) Then
			RunInSingle(obj.SCModule, obj.ObjName & "_Timer")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object obj.obj.Counter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			obj.obj.Counter = 0
		End If
	End Sub
	
	Private Sub txtPre_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPre.Enter
		Call cboSend_Enter(cboSend, New System.EventArgs())
	End Sub
	
	Private Sub txtPost_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPost.Enter
		Call cboSend_Enter(cboSend, New System.EventArgs())
	End Sub
	
	Private Sub cboSend_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboSend.Enter
		On Error Resume Next
		
		Dim i As Short
		cboSend.SelectionStart = cboSendSelStart
		cboSend.SelectionLength = cboSendSelLength
		
		If (BotVars.NoAutocompletion = False) Then
			'UPGRADE_WARNING: Controls method Controls.Count has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			For i = 0 To (Controls.Count() - 1)
				'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				If (TypeOf CType(Controls(i), Object) Is System.Windows.Forms.ListView) Or (TypeOf CType(Controls(i), Object) Is System.Windows.Forms.TabControl) Or (TypeOf CType(Controls(i), Object) Is System.Windows.Forms.RichTextBox) Or (TypeOf CType(Controls(i), Object) Is System.Windows.Forms.TextBox) Then
					If (CType(Controls(i), Object).TabStop = False) Then
						CType(Controls(i), Object).Tag = "False"
					End If
					
					CType(Controls(i), Object).TabStop = False
				End If
			Next i
		End If
		
		cboSendHadFocus = True
	End Sub
	
	Private Sub txtPre_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPre.Leave
		Call cboSend_Leave(cboSend, New System.EventArgs())
	End Sub
	
	Private Sub txtPost_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPost.Leave
		Call cboSend_Leave(cboSend, New System.EventArgs())
	End Sub
	
	Private Sub cboSend_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboSend.Leave
		On Error Resume Next
		
		Dim i As Short
		
		If (BotVars.NoAutocompletion = False) Then
			'UPGRADE_WARNING: Controls method Controls.Count has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			For i = 0 To (Controls.Count() - 1)
				'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				If (TypeOf CType(Controls(i), Object) Is System.Windows.Forms.ListView) Or (TypeOf CType(Controls(i), Object) Is AxMSComctlLib.AxTabStrip) Or (TypeOf CType(Controls(i), Object) Is System.Windows.Forms.RichTextBox) Or (TypeOf CType(Controls(i), Object) Is System.Windows.Forms.TextBox) Then
					
					If (CType(Controls(i), Object).Tag <> "False") Then
						CType(Controls(i), Object).TabStop = True
					End If
				End If
			Next i
		End If
		
		cboSendHadFocus = False
	End Sub
	
	'UPGRADE_WARNING: Event cboSend.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboSend_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboSend.SelectedIndexChanged
		RecordcboSendSelInfo()
	End Sub
	
	'UPGRADE_WARNING: Event cboSend.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cboSend.Change was upgraded to cboSend.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cboSend_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboSend.TextChanged
		'Debug.Print cboSendSelLength & "\" & cboSendSelStart
		RecordcboSendSelInfo()
	End Sub
	
	Private Sub cboSend_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cboSend.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		RecordcboSendSelInfo()
	End Sub
	
	Private Sub cboSend_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cboSend.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		On Error GoTo ERROR_HANDLER
		
		Static ACBuffer As String
		Static ACWordStart As Integer
		Static ACWordEnd As Integer
		Static ACUserPLen As Integer
		Static ACUserIndex As Short
		
		Dim AccessAmount As udtGetAccessResponse
		
		Dim Vetoed As Boolean '
		
		Dim i As Short
		Dim j As Short
		Dim SavedSelPos As Integer
		Dim SavedSelLen As Integer
		Dim CurrentTab As Short
		Dim CurrentList As System.Windows.Forms.ListView
		Dim CurrentUser As Integer
		Dim Value As String
		
		SavedSelPos = cboSend.SelectionStart
		SavedSelLen = cboSend.SelectionLength
		
		CurrentTab = ListviewTabs.SelectedIndex
		Select Case CurrentTab
			Case LVW_BUTTON_FRIENDS : CurrentList = lvFriendList
			Case LVW_BUTTON_CLAN : CurrentList = lvClanList
			Case Else : CurrentList = lvChannel
		End Select
		
		If (Not (CurrentList.FocusedItem Is Nothing)) Then
			CurrentUser = CurrentList.FocusedItem.Index
		End If
		
		Dim x() As String
		Dim n As Integer
		Dim ACList As Collection
		Dim ACUser As String
		Dim ACText As String
		Dim DoRunCommands As Boolean
		Dim NoProcs() As String
		Dim StartOutfilterPos As Integer
		Dim Args() As String
		Dim UserArg As String
		Dim CommandList As String
		Select Case (KeyCode)
			Case System.Windows.Forms.Keys.Space
				With cboSend
                    If (Len(LastWhisper) > 0) Then
                        If (Len(.Text) >= 2) Then
                            If StrComp(VB.Left(.Text, 2), "/r", CompareMethod.Text) = 0 Then
                                .SelectionStart = 0
                                .SelectionLength = Len(.Text)
                                .SelectedText = "/w " & CleanUsername(LastWhisper, True)
                                .SelectionStart = Len(.Text)
                            End If
                        End If

                        If (Len(.Text) >= 6) Then
                            If StrComp(VB.Left(.Text, 6), "/reply", CompareMethod.Text) = 0 Then
                                .SelectionStart = 0
                                .SelectionLength = Len(.Text)
                                .SelectedText = "/w " & CleanUsername(LastWhisper, True)
                                .SelectionStart = Len(.Text)
                            End If
                        End If
                    End If
					
                    If (Len(LastWhisperTo) > 0) Then
                        If (Len(.Text) >= 3) Then
                            If StrComp(VB.Left(.Text, 3), "/rw", CompareMethod.Text) = 0 Then
                                .SelectionStart = 0
                                .SelectionLength = Len(.Text)

                                If StrComp(LastWhisperTo, "%f%") = 0 Then
                                    .SelectedText = "/f m"
                                Else
                                    .SelectedText = "/w " & CleanUsername(LastWhisperTo, True)
                                End If

                                .SelectionStart = Len(.Text)
                            End If
                        End If
                    End If
				End With
				
			Case System.Windows.Forms.Keys.PageDown 'ALT + PAGEDOWN
				If Shift = VB6.ShiftConstants.AltMask Then
					With CurrentList
						If CurrentUser > 0 And CurrentUser < .Items.Count Then
							.HideSelection = False
							'UPGRADE_WARNING: Lower bound of collection CurrentList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
							With .Items.Item(CurrentUser + 1)
								'UPGRADE_WARNING: Lower bound of collection CurrentList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
								.Selected = True
								'UPGRADE_WARNING: Lower bound of collection CurrentList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
								'UPGRADE_WARNING: MSComctlLib.IListItem method ListItems.Item.EnsureVisible has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
								.EnsureVisible()
							End With
						End If
					End With
					
					cboSend.Focus()
					cboSend.SelectionStart = SavedSelPos
					cboSend.SelectionLength = SavedSelLen
					Exit Sub
				End If
				
			Case System.Windows.Forms.Keys.PageUp 'ALT + PAGEUP
				If Shift = VB6.ShiftConstants.AltMask Then
					With CurrentList
						If CurrentUser > 1 Then
							.HideSelection = False
							'UPGRADE_WARNING: Lower bound of collection CurrentList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
							With .Items.Item(CurrentUser - 1)
								'UPGRADE_WARNING: Lower bound of collection CurrentList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
								.Selected = True
								'UPGRADE_WARNING: Lower bound of collection CurrentList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
								'UPGRADE_WARNING: MSComctlLib.IListItem method ListItems.Item.EnsureVisible has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
								.EnsureVisible()
							End With
						End If
					End With
					
					cboSend.Focus()
					cboSend.SelectionStart = SavedSelPos
					cboSend.SelectionLength = SavedSelLen
					Exit Sub
				End If
				
			Case System.Windows.Forms.Keys.Home 'ALT+HOME
				If Shift = VB6.ShiftConstants.AltMask Then
					With CurrentList
						If .Items.Count > 0 Then
							.HideSelection = False
							'UPGRADE_WARNING: Lower bound of collection CurrentList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
							With .Items.Item(1)
								'UPGRADE_WARNING: Lower bound of collection CurrentList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
								.Selected = True
								'UPGRADE_WARNING: Lower bound of collection CurrentList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
								'UPGRADE_WARNING: MSComctlLib.IListItem method ListItems.Item.EnsureVisible has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
								.EnsureVisible()
							End With
						End If
					End With
					
					cboSend.Focus()
					cboSend.SelectionStart = SavedSelPos
					cboSend.SelectionLength = SavedSelLen
					Exit Sub
				End If
				
			Case System.Windows.Forms.Keys.End 'ALT+END
				If Shift = VB6.ShiftConstants.AltMask Then
					With CurrentList
						If .Items.Count > 0 Then
							.HideSelection = False
							'UPGRADE_WARNING: Lower bound of collection CurrentList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
							With .Items.Item(.Items.Count)
								'UPGRADE_WARNING: Lower bound of collection CurrentList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
								.Selected = True
								'UPGRADE_WARNING: Lower bound of collection CurrentList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
								'UPGRADE_WARNING: MSComctlLib.IListItem method ListItems.Item.EnsureVisible has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
								.EnsureVisible()
							End With
						End If
					End With
					
					cboSend.Focus()
					cboSend.SelectionStart = SavedSelPos
					cboSend.SelectionLength = SavedSelLen
					Exit Sub
				End If
				
			Case System.Windows.Forms.Keys.N, System.Windows.Forms.Keys.Insert 'ALT + N or ALT + INSERT
				If (Shift = VB6.ShiftConstants.AltMask) Then
					If (Not (CurrentList.FocusedItem Is Nothing)) Then
						Value = CurrentList.FocusedItem.Text
						If cboSend.SelectionStart = 0 Then Value = Value & BotVars.AutoCompletePostfix
						Value = Value & Space(1)
						
						cboSend.SelectedText = Value
						cboSend.SelectionStart = cboSend.SelectionStart + Len(Value)
						Exit Sub
					End If
				End If
				
			Case System.Windows.Forms.Keys.V 'PASTE
				
				If (IsScrolling(rtbChat)) Then
					LockWindowUpdate(rtbChat.Handle.ToInt32)
					
					SendMessage(rtbChat.Handle.ToInt32, EM_SCROLL, SB_BOTTOM, &H0)
					
					LockWindowUpdate(&H0)
				End If
				
				If (Shift = VB6.ShiftConstants.CtrlMask) Then
					On Error Resume Next
					
					If (InStr(1, My.Computer.Clipboard.GetText, Chr(13), CompareMethod.Text) <> 0) Then
						x = Split(My.Computer.Clipboard.GetText, Chr(10))
						
						If UBound(x) > 0 Then
							For n = LBound(x) To UBound(x)
								x(n) = Replace(x(n), Chr(13), vbNullString)
								
								If (x(n) <> vbNullString) Then
									If (n <> LBound(x)) Then
										AddQ(txtPre.Text & x(n) & txtPost.Text, (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
										
										cboSend.Items.Insert(0, txtPre.Text & x(n) & txtPost.Text)
									Else
										AddQ(txtPre.Text & cboSend.Text & x(n) & txtPost.Text, (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
										
										cboSend.Items.Insert(0, txtPre.Text & cboSend.Text & x(n) & txtPost.Text)
									End If
								End If
							Next n
							
							cboSend.Text = vbNullString
							
							MultiLinePaste = True
						End If
					End If
				End If
				
			Case System.Windows.Forms.Keys.A
				If (Shift = VB6.ShiftConstants.CtrlMask) Then
					If (CurrentTab <> LVW_BUTTON_CHANNEL) Then
						ListviewTabs.SelectedIndex = LVW_BUTTON_CHANNEL
						Call ListviewTabs_SelectedIndexChanged(ListviewTabs, New System.EventArgs())
					Else
						cboSend.SelectionStart = 0
						cboSend.SelectionLength = Len(cboSend.Text)
					End If
				End If
				
			Case System.Windows.Forms.Keys.S
				If (Shift = VB6.ShiftConstants.CtrlMask) Then
					If (CurrentTab <> LVW_BUTTON_FRIENDS) And (ListviewTabs.TabPages.Item(LVW_BUTTON_FRIENDS).Enabled) Then
						ListviewTabs.SelectedIndex = LVW_BUTTON_FRIENDS
						Call ListviewTabs_SelectedIndexChanged(ListviewTabs, New System.EventArgs())
					End If
				End If
				
			Case System.Windows.Forms.Keys.D
				If (Shift = VB6.ShiftConstants.CtrlMask) Then
					If (CurrentTab <> LVW_BUTTON_CLAN) And (ListviewTabs.TabPages.Item(LVW_BUTTON_CLAN).Enabled) Then
						ListviewTabs.SelectedIndex = LVW_BUTTON_CLAN
						Call ListviewTabs_SelectedIndexChanged(ListviewTabs, New System.EventArgs())
					End If
				End If
				
			Case System.Windows.Forms.Keys.B
				If (Shift = VB6.ShiftConstants.CtrlMask) Then
					cboSend.SelectedText = "ÿcb"
				End If
				
				'Case vbKeyJ
				'    If (Shift = vbCtrlMask) Then
				'        Call mnuToggle_Click
				'    End If
				
			Case System.Windows.Forms.Keys.U
				If (Shift = VB6.ShiftConstants.CtrlMask) Then
					cboSend.SelectedText = "ÿcu"
				End If
				
			Case System.Windows.Forms.Keys.I
				If (Shift = VB6.ShiftConstants.CtrlMask) Then
					cboSend.SelectedText = "ÿci"
				End If
				
			Case System.Windows.Forms.Keys.Delete
				ACUserIndex = 0
				
			Case System.Windows.Forms.Keys.Tab
				
				If (Shift <> 0) Then
					Call cboSend_Leave(cboSend, New System.EventArgs())
					
					If (txtPre.Visible = True) Then
						Call txtPre.Focus()
					Else
						Call ListviewTabs.Focus()
						Call ListviewTabs_SelectedIndexChanged(ListviewTabs, New System.EventArgs())
					End If
				Else
					
					With cboSend
						' check if auto-complete active
						If ACUserIndex = 0 Then
							' reset the static variables and auto-complete the current word
							ACBuffer = .Text
							
							If .SelectionStart = 0 Then
								ACWordStart = 1
							Else
								ACWordStart = InStrRev(ACBuffer, Space(1), .SelectionStart, CompareMethod.Binary) + 1
							End If
							ACWordEnd = InStr(.SelectionStart + 1, ACBuffer, Space(1), CompareMethod.Binary)
							If ACWordEnd = 0 Then ACWordEnd = Len(ACBuffer) + 1
							
							ACUserPLen = ACWordEnd - ACWordStart
							If ACUserPLen < 0 Then ACUserPLen = 0
							
							'AddChat vbWhite, ACWordEnd - ACWordStart
							'AddChat vbWhite, Mid$(ACBuffer, ACWordStart, ACUserPLen)
							
							ACUserIndex = 1
						Else
							' advance last auto-complete
							ACUserIndex = ACUserIndex + 1
						End If
						
						' repopulate autocomplete candidates
						' sort as we get them
						ACList = New Collection
						'If ACUserPLen > 0 Then
						For i = 1 To CurrentList.Items.Count
							'UPGRADE_WARNING: Lower bound of collection CurrentList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
							If StrComp(VB.Left(CurrentList.Items.Item(i).Text, ACUserPLen), Mid(ACBuffer, ACWordStart, ACUserPLen), CompareMethod.Text) = 0 Then
								For j = 1 To ACList.Count()
									'UPGRADE_WARNING: Couldn't resolve default property of object ACList.Item(j). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Lower bound of collection CurrentList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
									If StrComp(CurrentList.Items.Item(i).Text, ACList.Item(j), CompareMethod.Text) < 0 Then
										Exit For
									End If
								Next j
								
								'AddChat vbYellow, j & " - " & CurrentList.ListItems.Item(i).Text
								
								' add at found position
								If j - 1 = ACList.Count() Then
									'UPGRADE_WARNING: Lower bound of collection CurrentList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
									ACList.Add(CurrentList.Items.Item(i).Text)
								Else
									'UPGRADE_WARNING: Lower bound of collection CurrentList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
									ACList.Add(CurrentList.Items.Item(i).Text,  , j)
								End If
							End If
						Next i
						'End If
						
						' set the user to the next candidate
						ACUser = vbNullString
						If ACList.Count() > 0 Then
							If ACList.Count() < ACUserIndex Then
								ACUserIndex = 1
							End If
							
							'UPGRADE_WARNING: Couldn't resolve default property of object ACList.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							ACUser = ACList.Item(ACUserIndex)
						End If
						'UPGRADE_NOTE: Object ACList may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						ACList = Nothing
						
						' save to the text box
                        If (Len(ACUser) > 0) Then
                            If (ACWordStart > 1) Then
                                ACText = VB.Left(ACBuffer, ACWordStart - 1)
                            ElseIf (ACWordStart = 1) Then
                                ' only includes the postfix for the first word
                                ACUser = ACUser & BotVars.AutoCompletePostfix
                            End If

                            ACText = ACText & ACUser & Space(1)

                            SavedSelPos = Len(ACText)

                            If (ACWordEnd > 0) Then
                                ACText = ACText & Mid(ACBuffer, ACWordEnd + 1)
                            End If

                            .Text = ACText
                            .SelectionStart = SavedSelPos
                        Else
                            ACUserIndex = 0
                        End If
					End With
				End If
				
			Case System.Windows.Forms.Keys.Return
				
				If (IsScrolling(rtbChat)) Then
					LockWindowUpdate(rtbChat.Handle.ToInt32)
					
					SendMessage(rtbChat.Handle.ToInt32, EM_SCROLL, SB_BOTTOM, &H0)
					
					LockWindowUpdate(&H0)
				End If
				
				DoRunCommands = True
				
				Select Case (Shift)
					Case VB6.ShiftConstants.ShiftMask 'CTRL+ENTER - rewhisper
                        If Len(cboSend.Text) > 0 Then
                            Value = "/w " & LastWhisperTo & Space(1)
                            DoRunCommands = False
                        End If
						
					Case VB6.ShiftConstants.ShiftMask Or VB6.ShiftConstants.CtrlMask 'CTRL+SHIFT+ENTER - reply
                        If Len(cboSend.Text) > 0 Then
                            Value = "/w " & LastWhisper & Space(1)
                            DoRunCommands = False
                        End If
						
					Case Else 'normal ENTER - old rules apply
                        If Len(cboSend.Text) > 0 Then
                            Value = vbNullString
                        End If
				End Select
				
				' prefix box
				If txtPre.Visible Then Value = Value & txtPre.Text
				' sendbox
				Value = Value & cboSend.Text
				' suffix box
				If txtPost.Visible Then Value = Value & txtPost.Text
				
				If DoRunCommands Then
					On Error Resume Next
					DoRunCommands = Not RunInAll("Event_PressedEnter", Value)
					On Error GoTo 0
				End If
				
				If (VB.Left(Value, 6) = "/tell ") Then
					Value = "/w " & Mid(Value, 7)
				End If
				
				If DoRunCommands Then
					If (VB.Left(Value, 1) = "/") Then
                        If (Len(Config.ServerCommandList) > 0) Then

                            CommandList = Config.ServerCommandList

                            ' please note: only commands with "%" as the first argument (or no "%" arguments)
                            ' (i.e. "/w %", "/ban %") are supported to be correctly no-processed
                            ' complexify here if that is needed
                            Args = Split(Value, Space(1), 3)
                            If UBound(Args) > 0 Then UserArg = Args(1)

                            CommandList = Replace(CommandList, "%", UserArg)

                            NoProcs = Split(CommandList, ",")
                        Else
                            ReDim NoProcs(0)
                        End If
						
						For i = LBound(NoProcs) To UBound(NoProcs)
                            If (Len(NoProcs(i)) > 0) And (StrComp(VB.Left(Value, Len(NoProcs(i)) + 2), "/" & NoProcs(i) & Space(1), CompareMethod.Text) = 0) Then
                                DoRunCommands = False
                                StartOutfilterPos = Len(NoProcs(i)) + 3
                                Exit For
                            End If
						Next i
						
						If DoRunCommands Then
							ProcessCommand(GetCurrentUsername, Value, True, False)
						Else
							DoRunCommands = False
						End If
					Else
						DoRunCommands = False
					End If
				End If
				
				If Not DoRunCommands Then
					' Don't do replacements for a command unless it involves text that will be seen by someone else
					'  and don't replace text in the command itself or the target username
					Value = VB.Left(Value, StartOutfilterPos) & OutFilterMsg(Mid(Value, StartOutfilterPos + 1))
					
					Call AddQ(Value, (modQueueObj.PRIORITY.CONSOLE_MESSAGE))
				End If
				
				'Ignore rest of code as the bot is closing
				If BotIsClosing Then
					Exit Sub
				End If
				
				cboSend.Items.Insert(0, Value)
				
				cboSend.Text = vbNullString
				
				If Me.WindowState <> System.Windows.Forms.FormWindowState.Minimized Then
					On Error Resume Next
					cboSend.Focus()
				End If
		End Select
		
		CurrentList.HideSelection = True
		
		If (KeyCode <> System.Windows.Forms.Keys.Tab) Then
			ACUserIndex = 0
		End If
		
		Exit Sub
		
ERROR_HANDLER: 
		AddChat(RTBColors.ErrorMessageText, "Error " & Err.Number & " (" & Err.Description & ") " & "in procedure cboSend_KeyDown")
		
		Exit Sub
	End Sub
	
	Private Sub cboSend_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cboSend.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim oldSelStart As Short
		Dim sClosest As String
		
		'AddChat vbBlue, KeyAscii
		
		Select Case KeyAscii
			Case 1, 19, 4, 2, 9, 21, 13, 10
				KeyAscii = 0
			Case 22
				If MultiLinePaste Then
					KeyAscii = 0
					MultiLinePaste = False
				End If
				
		End Select
		
		With cboSend
			If (KeyAscii > 0) Then
				If (.Items.Count > 15) Then
					.Items.RemoveAt(15)
				End If
			End If
		End With
		
		If Len(cboSend.Text) > 223 Then
			cboSend.ForeColor = System.Drawing.Color.Red
		Else
			cboSend.ForeColor = System.Drawing.Color.White
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
    Public Sub SControl_ErrorEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SControl.ErrorEvent
        Call modScripting.SC_Error()
    End Sub
	
	Private Sub sckBNet_CloseEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sckBNet.CloseEvent
		sckBNet.Close()
        If sckBNLS.CtlState <> 0 Then sckBNLS.Close()
		
		Call Event_BNetDisconnected()
		
		ds.ClientToken = 0
		g_Connected = False
	End Sub
	
	Private Sub sckBNet_ConnectEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sckBNet.ConnectEvent
		On Error Resume Next
		
		Call Event_BNetConnected()
		
		If MDebug("all") Then
			AddChat(COLOR_BLUE, "BNET CONNECT")
		End If
		
		Call modWarden.WardenCleanup(WardenInstance)
		WardenInstance = modWarden.WardenInitilize(sckBNet.SocketHandle)
		ds.Reset_Renamed()
		
		If (Not (BotVars.UseProxy)) Then
			InitBNetConnection()
		Else
			LogonToProxy(sckBNet, BotVars.Server, 6112, BotVars.ProxyIsSocks5)
		End If
		
	End Sub
	
	Sub InitBNetConnection()
		g_Connected = True
		
		'sckBNet.SendData ChrW(1)
        Call Send(sckBNet.SocketHandle, New Byte() {1}, 1, 0)
		
		If BotVars.BNLS Then
			modBNLS.SEND_BNLS_REQUESTVERSIONBYTE()
		Else
			Select Case modBNCS.GetLogonSystem()
				Case modBNCS.BNCS_NLS : Call modBNCS.SEND_SID_AUTH_INFO()
				Case modBNCS.BNCS_OLS
					modBNCS.SEND_SID_CLIENTID2()
					modBNCS.SEND_SID_LOCALEINFO()
					modBNCS.SEND_SID_STARTVERSIONING()
				Case modBNCS.BNCS_LLS
					modBNCS.SEND_SID_CLIENTID()
					modBNCS.SEND_SID_STARTVERSIONING()
				Case Else
					AddChat(RTBColors.ErrorMessageText, StringFormat("Unknown Logon System Type: {0}", modBNCS.GetLogonSystem()))
					AddChat(RTBColors.ErrorMessageText, "Please visit http://www.stealthbot.net/sb/issues/?unknownLogonType for information regarding this error.")
					DoDisconnect()
			End Select
		End If
	End Sub
	
	Private Sub sckBNet_Error(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_ErrorEvent) Handles sckBNet.Error
		Call Event_BNetError(eventArgs.Number, eventArgs.Description)
	End Sub
	
	Private Sub sckMCP_CloseEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sckMCP.CloseEvent
		AddChat(RTBColors.ErrorMessageText, "[REALM] Connection closed.")
		
		If Not ds.MCPHandler Is Nothing Then
			ds.MCPHandler.IsRealmError = True
			
			Call DoDisconnect()
		End If
	End Sub
	
	Private Sub sckMCP_ConnectEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sckMCP.ConnectEvent
		On Error Resume Next
		
		If MDebug("all") Then
			AddChat(COLOR_BLUE, "MCP CONNECT")
		End If
		
		AddChat(RTBColors.SuccessText, "[REALM] Connected!")
		
		'sckMCP.SendData ChrW(1)
        Call Send(sckMCP.SocketHandle, New Byte() {1}, 1, 0)
		
		If Not ds.MCPHandler Is Nothing Then
			ds.MCPHandler.SEND_MCP_STARTUP()
		End If
	End Sub
	
	Private Sub sckMCP_DataArrival(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_DataArrivalEvent) Handles sckMCP.DataArrival
		On Error GoTo ERROR_HANDLER

        Dim bData() As Byte

        ' Get the data and add it to the buffer
        sckMCP.GetData(bData, vbArray + vbByte, eventArgs.bytesTotal)
        MCPBuffer.AddData(bData)

        ' If we have full packets, parse them
		While MCPBuffer.FullPacket
            bData = MCPBuffer.GetPacket
			
			If Not ds.MCPHandler Is Nothing Then
                Call ds.MCPHandler.ParsePacket(bData)
			End If
		End While
		
		Exit Sub
		
ERROR_HANDLER: 
		AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in sckMCP_DataArrival().")
		
		Exit Sub
	End Sub
	
	Private Sub sckMCP_Error(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_ErrorEvent) Handles sckMCP.Error
		If Not g_Online Then
			' This message is ignored if we've entered chat
			AddChat(RTBColors.ErrorMessageText, "[REALM] Server error " & eventArgs.Number & ": " & eventArgs.Description)
			
			If Not ds.MCPHandler Is Nothing Then
				If ds.MCPHandler.FormActive Then
					frmRealm.UnloadRealmError()
				End If
			End If
		End If
	End Sub
	
	
	'// Written by Swent. Executes plugin timer subs
	Private Sub scTimer_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles scTimer.Tick
		On Error Resume Next
		
		RunInAll("scTimer_Timer")
	End Sub
	
	' centralized "idle" events:
	' IDLE MESSAGE (.5*IDLEMESSAGEDELAY minutes - user setting)
	' PROFILE AMP (1 minute - useful minimum song length)
	' BNCS.SID_NULL (2 minutes - keep alive)
	' BNCS.SID_CLANMOTD (10 minutes - may change)
	' BNCS.SID_FRIENDSLIST (5 minutes - for D1,W2,D2 [no update], SC,W3 [bug in SID_FRIENDSUPDATE])
	Private Sub tmrIdleTimer_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrIdleTimer.Tick
		On Error GoTo ERROR_HANDLER
		
		' long-counter
		Static lCounter As Integer
		
		lCounter = lCounter + 1
		
		Dim IdleWait As Integer
		If g_Online Then
			' bot idle (30 second interval (x config value), offset 0 seconds)
			If Config.IdleMessage Then
				
				IdleWait = Config.IdleMessageDelay * 30
				
				If (lCounter Mod IdleWait) = 0 Then
					'AddQ "IDLE"
					Call tmrIdleTimer_Timer_IdleMsg()
				End If
			End If
			
			' bot ProfileAmp (1 minute interval, offset 5 seconds)
			If g_Online And Config.ProfileAmp Then
				If (lCounter Mod 60) = 5 Then
					'AddQ "PROFILE AMP"
					Call UpdateProfile()
				End If
			End If
			
			' BNCS keepalive... (2 minute interval; offset -15 seconds)
			If (lCounter Mod 120) = 105 Then
				' if W3 & in clan, then (10 minute interval; offset -15 seconds from 2nd minute)
				If IsW3 And Clan.isUsed And (lCounter Mod 600) = 105 Then
					' request clan MOTD instead of NULL
					'AddQ "CLAN MOTD"
					RequestClanMOTD()
					' if friend list updates enabled, then (5 minute interval; offset -15 seconds from 4th minute)
				ElseIf Config.FriendsListTab And (lCounter Mod 300) = 225 Then 
					' request friendlist instead of FL
					'AddQ "FRIENDS"
					If (lvFriendList.Items.Count > 0) Then
						Call FriendListHandler.RequestFriendsList(PBuffer)
					Else
						PBuffer.SendPacket(SID_NULL)
					End If
					' else standard null
				Else
					'AddQ "NULL"
					PBuffer.SendPacket(SID_NULL)
				End If
			End If
		End If
		
		Exit Sub
		
ERROR_HANDLER: 
		
		AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in tmrIdleTimer_Timer().")
		
		Exit Sub
		
	End Sub
	
	Private Sub tmrIdleTimer_Timer_IdleMsg()
		On Error GoTo ERROR_HANDLER
		
		Dim U, IdleMsg As String
		Dim s() As String
		Dim IdleWaitS, IdleType As String
		Dim IdleWait As Short
		Dim UDP As Byte
		'UPGRADE_NOTE: IsError was upgraded to IsError_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim IsError_Renamed As Boolean
		
		BotVars.JoinWatch = 0
		
		If Not Config.IdleMessage Then Exit Sub
		
		IdleMsg = Config.IdleMessageText
		IdleWait = Config.IdleMessageDelay
		IdleType = Config.IdleMessageType
		
		If IdleWait < 2 Then Exit Sub
		
		'If iCounter >= IdleWait Then
		'iCounter = 0
		'on error resume next
		Dim WindowTitle As String
		If IdleType = "msg" Or IdleType = vbNullString Then
			
			If StrComp(IdleMsg, "null", CompareMethod.Text) = 0 Or IdleMsg = vbNullString Then
				Exit Sub
			End If
			
			IdleMsg = Replace(IdleMsg, "%cpuup", ConvertTime(GetUptimeMS))
			IdleMsg = Replace(IdleMsg, "%chan", g_Channel.Name)
			IdleMsg = Replace(IdleMsg, "%c", g_Channel.Name)
			IdleMsg = Replace(IdleMsg, "%me", GetCurrentUsername)
			IdleMsg = Replace(IdleMsg, "%v", CVERSION)
			IdleMsg = Replace(IdleMsg, "%ver", CVERSION)
			IdleMsg = Replace(IdleMsg, "%bc", CStr(BanCount))
			IdleMsg = Replace(IdleMsg, "%botup", ConvertTime(uTicks))
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.TrackName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			IdleMsg = Replace(IdleMsg, "%mp3", Replace(MediaPlayer.TrackName, "&", "+"))
			IdleMsg = Replace(IdleMsg, "%quote", g_Quotes.GetRandomQuote)
			IdleMsg = Replace(IdleMsg, "%rnd", GetRandomPerson)
			IdleMsg = Replace(IdleMsg, "%t", TimeString)
			
			If (IdleMsg = vbNullString) Then
				IsError_Renamed = True
			End If
			
		ElseIf IdleType = "uptime" Then 
			IdleMsg = "/me -: System Uptime: " & ConvertTime(GetUptimeMS()) & " :: Connection Uptime: " & ConvertTime(uTicks) & " :: " & CVERSION & " :-"
			
		ElseIf IdleType = "mp3" Then 
			
			'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.TrackName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			WindowTitle = MediaPlayer.TrackName
			
			If WindowTitle = vbNullString Then
				IsError_Renamed = True
				'UPGRADE_WARNING: Couldn't resolve default property of object MediaPlayer.IsPaused. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf (MediaPlayer.IsPaused) Then 
				IdleMsg = "/me -: Now Playing: " & WindowTitle & " (paused) :: " & CVERSION & " :-"
			Else
				IdleMsg = "/me -: Now Playing: " & WindowTitle & " :: " & CVERSION & " :-"
			End If
			
		ElseIf IdleType = "quote" Then 
			U = g_Quotes.GetRandomQuote
			
			IdleMsg = "/me : " & U
			
			If Len(U) > 217 Then
				IsError_Renamed = True
			End If
			
		End If
		
		If (IsError_Renamed) Then
			IdleMsg = "/me -- " & CVERSION
		End If
		
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If sckBNet.CtlState = 7 Then
			If InStr(1, IdleMsg, "& ", CompareMethod.Text) And IdleType = "msg" Then
				s = Split(IdleMsg, "& ")
				
				For IdleWait = LBound(s) To UBound(s)
					If Len(s(IdleWait)) > 215 Then
						s(IdleWait) = VB.Left(s(IdleWait), 215)
					End If
					AddQ(s(IdleWait))
				Next 
			Else
				If Len(IdleMsg) > 215 Then
					IdleMsg = VB.Left(IdleMsg, 215)
				End If
				
				Me.AddQ(IdleMsg)
			End If
		End If
		'End If
		
		Exit Sub
		
ERROR_HANDLER: 
		
		AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in tmrIdleTimer_Timer_IdleMsg().")
		
		Exit Sub
		
	End Sub
	
	Private Sub tmrSilentChannel_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrSilentChannel.Tick
		Dim Index As Short = tmrSilentChannel.GetIndex(eventSender)
		On Error GoTo ERROR_HANDLER
		
		Dim User As clsUserObj
		Dim Item As System.Windows.Forms.ListViewItem
		
		Dim i As Short
		Dim j As Short
		Dim found As Boolean
		Dim WasZero As Boolean
		
		If (g_Channel.IsSilent = False) Then
			Exit Sub
		End If
		
		If (Index = 0) Then
			If (Me.mnuDisableVoidView.Checked = False) Then
				'For i = 1 To g_Channel.Users.Count
				'    ' with our doevents, we can miss our cue indicating that we
				'    ' need to stop silent channel processing and cause an rte.
				'    If (i > g_Channel.Users.Count) Then
				'        Exit For
				'    End If
				'
				'     Set user = g_Channel.Users(i)
				'
				'    If (lvChannel.FindItem(user.DisplayName) Is Nothing) Then
				'        Dim Stats As String
				'        Dim Clan  As String
				'
				'        ParseStatstring user.Statstring, Stats, Clan
				'
				'        AddName user.DisplayName, user.Game, user.Flags, user.Ping, user.Clan
				'    End If
				'Next i
				
				Call LockWindowUpdate(&H0)
				
				lblCurrentChannel.Text = GetChannelString()
			End If
			
			tmrSilentChannel(0).Enabled = False
		ElseIf (Index = 1) Then 
			If (mnuDisableVoidView.Checked = False) Then
				If (g_Channel.IsSilent) Then
					Call g_Channel.ClearUsers()
					
					Me.lvChannel.Items.Clear()
				End If
				
				Call AddQ("/unsquelch " & GetCurrentUsername, (modQueueObj.PRIORITY.SPECIAL_MESSAGE))
			End If
		End If
		
		Exit Sub
		
ERROR_HANDLER: 
		AddChat(RTBColors.ErrorMessageText, "Error: " & Err.Description & " in tmrSilentChannel_Timer(" & Index & ").")
		
		Exit Sub
	End Sub
	
	Private Sub txtPre_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPre.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If (KeyAscii = 13) Then
			Call cboSend_KeyPress(cboSend, New System.Windows.Forms.KeyPressEventArgs(Chr(KeyAscii)))
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtPost_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPost.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If (KeyAscii = 13) Then
			Call cboSend_KeyPress(cboSend, New System.Windows.Forms.KeyPressEventArgs(Chr(KeyAscii)))
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Sub ConnectBNLS()
		' Don't try and connect if we don't have a server to connect to.
		If Len(BotVars.BNLSServer) = 0 Then
			AddChat(RTBColors.ErrorMessageText, "[BNLS] A working BNLS server could not be found.")
			AddChat(RTBColors.ErrorMessageText, "[BNLS]   Go to Settings -> Bot Settings -> Connection Settings -> Advanced and either set a server or use the automatic server finder.")
			Call DoDisconnect()
			Exit Sub
		End If
		
		Call Event_BNLSConnecting()
		
		With sckBNLS
			'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			If .CtlState <> 0 Then .Close()
			
			.RemoteHost = BotVars.BNLSServer
			.RemotePort = 9367
			.Connect()
		End With
	End Sub
	
	Sub Connect()
		Dim NotEnoughInfo As Boolean
		Dim MissingInfo As String
		
		'g_username = BotVars.Username
		
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If sckBNet.CtlState = 0 And sckBNLS.CtlState = 0 Then
			
			'Vars
			NotEnoughInfo = False
			MissingInfo = "Information required to connect: "
			If BotVars.Username = vbNullString Then
				MissingInfo = MissingInfo & "Username, "
				NotEnoughInfo = True
			End If
			If BotVars.Password = vbNullString Then
				MissingInfo = MissingInfo & "Password, "
				NotEnoughInfo = True
			End If
			If BotVars.Server = vbNullString Then
				MissingInfo = MissingInfo & "Server, "
				NotEnoughInfo = True
			End If
			If BotVars.Product = vbNullString Then
				MissingInfo = MissingInfo & "your choice of Client, "
				NotEnoughInfo = True
			Else
				Select Case GetProductInfo(BotVars.Product).KeyCount
					Case 0
					Case 1
						If BotVars.CDKey = vbNullString Then
							MissingInfo = MissingInfo & "CDKey, "
							NotEnoughInfo = True
						End If
					Case 2
						If BotVars.CDKey = vbNullString Then
							MissingInfo = MissingInfo & "CDKey, "
							NotEnoughInfo = True
						End If
						If BotVars.ExpKey = vbNullString Then
							MissingInfo = MissingInfo & "expansion CDKey, "
							NotEnoughInfo = True
						End If
				End Select
			End If
			
			If NotEnoughInfo Then
				MsgBox("You haven't provided enough information to connect! " & "Please edit your connection settings by choosing Bot Settings under the Settings menu." & vbNewLine & VB.Left(MissingInfo, Len(MissingInfo) - 2) & ".", MsgBoxStyle.Information)
				
				Call DoDisconnect(1)
				
				Exit Sub
			End If
			
			SetTitle("Connecting...")
			
			If ((StrComp(BotVars.Product, "PX2D", CompareMethod.Text) = 0) Or (StrComp(BotVars.Product, "VD2D", CompareMethod.Text) = 0)) Then
				
				Dii = True
			Else
				Dii = False
			End If
			
			'Changed 10-07-2009 - Hdx - Are we going to have private betas anymore?
#If (BETA = 2) Then
			'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression (BETA = 2) did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
			Call AddChat(RTBColors.InformationText, "Authorizing your private-release, please wait...")
			
			If (GetAuth(BotVars.Username)) Then
			Call AddChat(RTBColors.SuccessText, "Private usage authorized, connecting your bot...")
			
			' was auth function bypassed?
			If (AUTH_CHECKED = False) Then BotVars.Password = Chr$(0)
			Else
			Call AddChat(RTBColors.ErrorMessageText, "- - - - - YOU ARE NOT AUTHORIZED TO USE THIS PROGRAM - - - - -")
			
			Call DoDisconnect
			UserCancelledConnect = False
			Exit Sub
			End If
#Else
			AddChat(RTBColors.InformationText, "Connecting your bot...")
#End If
			
			
			If BotVars.BNLS Then
				If Len(BotVars.BNLSServer) = 0 Then
					If BotVars.UseAltBnls Then
						Call FindBNLSServer()
						
						Exit Sub
					End If
				End If
				
				Call ConnectBNLS()
			Else
				Call Event_BNetConnecting()
			End If
			
			With sckBNet
				'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
				If .CtlState <> 0 Then
					AddChat(RTBColors.ErrorMessageText, "Already connected.")
					Exit Sub
				End If
				
				
				If Not BotVars.UseProxy Then
					.RemoteHost = BotVars.Server
					.RemotePort = 6112
				Else
					'// PROXY
                    If BotVars.ProxyPort > 0 And Len(BotVars.ProxyIP) > 0 Then
                        .RemoteHost = BotVars.ProxyIP
                        .RemotePort = BotVars.ProxyPort
                    Else
                        MsgBox("You have selected to use proxies, but no proxy is configured. Please set one up in the Advanced " & " section of Bot Settings.", MsgBoxStyle.Information)

                        DoDisconnect()

                        Exit Sub
                    End If
				End If
				
				If Not BotVars.BNLS Then .Connect()
				
			End With
			
		End If
		
		Exit Sub
		
Error_Renamed: 
		MsgBox("Configuration file error. Please re-write your configuration file " & "using the Setup dialog.", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error")
		
		Call SetTitle("Disconnected")
		
		Exit Sub
	End Sub
	
	Public Sub Pause(ByVal fSeconds As Single, Optional ByVal AllowEvents As Boolean = True)
		Dim i As Short
		If AllowEvents Then
			For i = 0 To (1000 * fSeconds) \ 100
				Sleep(100)
				System.Windows.Forms.Application.DoEvents()
			Next i
		Else
			Sleep(fSeconds * 1000)
		End If
	End Sub
	
	'/* Fires every second */
	Private Sub UpTimer_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles UpTimer.Tick
		
		On Error GoTo ERROR_HANDLER
		
		Dim newColor As Integer
		Dim i As Short
		Dim pos As Short
		Dim doCheck As Boolean
		
		uTicks = (uTicks + 1000)
		
		If (floodCap > 2) Then
			floodCap = floodCap - 3
		End If
		
		Dim s As String
		If (VoteDuration > 0) Then
			VoteDuration = VoteDuration - 1
			
			If (VoteDuration = 0) Then
				
				s = Voting(BVT_VOTE_END)
				
				If (Len(s) > 1) Then
					AddQ(s)
				End If
			End If
		End If
		
		If (g_Queue.Count > 0) Then
			Ban(vbNullString, 0, 3)
		End If
		
		If (g_Channel.IsSilent = False) Then
			doCheck = True
			
			For i = 1 To g_Channel.Users.Count()
				With g_Channel.Users.Item(i)
					If (g_Channel.Self.IsOperator) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users(i).IsOperator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (.IsOperator = False) Then
							' channel password
							If ((BotVars.ChannelPasswordDelay > 0) And (Len(BotVars.ChannelPassword) > 0)) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users(i).PassedChannelAuth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If (.PassedChannelAuth = False) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().TimeInChannel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If (.TimeInChannel() > BotVars.ChannelPasswordDelay) Then
										'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If (GetSafelist(.DisplayName) = False) Then
											'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											Ban(.DisplayName & " Password time is up", AutoModSafelistValue - 1)
											
											doCheck = False
										End If
									End If
								End If
							End If
							
							' idle bans
							If ((doCheck) And ((BotVars.IB_On = BTRUE) And (BotVars.IB_Wait > 0))) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().TimeSinceTalk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If (.TimeSinceTalk() > BotVars.IB_Wait) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If (GetSafelist(.DisplayName) = False) Then
										'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										Ban(.DisplayName & " Idle for " & BotVars.IB_Wait & "+ seconds", AutoModSafelistValue - 1, IIf(BotVars.IB_Kick, 1, 0))
										
										doCheck = False
									End If
								End If
							End If
						End If
					End If
					
					If (BotVars.NoColoring = False) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pos = checkChannel(.DisplayName)
						
						If (pos > 0) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users(i).TimeSinceTalk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().Flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							newColor = GetNameColor(.Flags, .TimeSinceTalk, StrComp(.DisplayName, GetCurrentUsername, CompareMethod.Binary) = 0)
							
							'UPGRADE_WARNING: Lower bound of collection lvChannel.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
							If (System.Drawing.ColorTranslator.ToOle(lvChannel.Items.Item(pos).ForeColor) <> newColor) Then
								'UPGRADE_WARNING: Lower bound of collection lvChannel.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
								lvChannel.Items.Item(pos).ForeColor = System.Drawing.ColorTranslator.FromOle(newColor)
							End If
						End If
					End If
				End With
				
				doCheck = True
			Next i
		End If
		
		Exit Sub
		
ERROR_HANDLER: 
		AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in UpTimer_Timer().")
		
		Exit Sub
		
	End Sub
	
	'StealthLock (c) 2003 Stealth, Please do not remove this header
	Private Function GetAuth(ByVal Username As String) As Integer
		On Error GoTo ERROR_HANDLER
		
		Static lastAuth As Integer
		Static lastAuthName As String
		
		Dim clsCRC32 As New clsCRC32
		Dim hostFile As String
		Dim hostPath As String
		Dim f As Short
		Dim tmp As String
		Dim Result As Short ' string variable for storing beta authorization result
		' 0  == unauthorized
		' >0 == authorized
		
		If (lastAuth) Then
			If (StrComp(Username, lastAuthName, CompareMethod.Text) = 0) Then
				GetAuth = lastAuth
				
				Exit Function
			End If
		End If
		
		f = FreeFile
		
		If (g_OSVersion.IsWindowsNT) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object GetRegistryValue(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			hostPath = GetRegistryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\", "DatabasePath")
			
			If (Len(hostPath) = 0) Then
				hostPath = "%SystemRoot%\system32\drivers\etc\"
			End If
		Else
			hostPath = "%WinDir%"
		End If
		
		hostPath = ReplaceEnvironmentVars(hostPath & "\hosts")
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If (Len(Dir(hostPath)) > 0) Then
            FileOpen(f, hostPath, OpenMode.Input)
            Do While (EOF(f) = False)
                tmp = LineInput(f)

                hostFile = hostFile & tmp
            Loop
            FileClose(f)
        End If
		
		If (clsCRC32.CRC32(BETA_AUTH_URL) = BETA_AUTH_URL_CRC32) Then
			If (InStr(1, hostFile, Split(BETA_AUTH_URL, ".")(1), CompareMethod.Text) = 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object INet.OpenURL(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Result = CShort(Val(INet.OpenURL(BETA_AUTH_URL & Username, 0)))
			End If
		End If
		
		Do While INet.StillExecuting
			System.Windows.Forms.Application.DoEvents()
		Loop 
		
		If (Result = 1) Then
			lastAuth = Result
			lastAuthName = Username
			
			GetAuth = True
			
			AUTH_CHECKED = True
		End If
		
		'UPGRADE_NOTE: Object clsCRC32 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		clsCRC32 = Nothing
		
		Exit Function
		
ERROR_HANDLER: 
		
		AddChat(RTBColors.ErrorMessageText, "Beta Auth Error: #", System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), Err.Number, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), ": ", System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), Err.Description)
		'UPGRADE_NOTE: Object clsCRC32 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		clsCRC32 = Nothing
		GetAuth = False
		Exit Function
		
	End Function
	
	' http://www.go4expert.com/forums/showthread.php?t=208
	'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Function ReplaceEnvironmentVars(ByVal str_Renamed As String) As String
		
		Dim i As Short
		'UPGRADE_NOTE: Name was upgraded to Name_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Name_Renamed As String
		Dim Value As String
		Dim tmp As String
		
		tmp = str_Renamed
		
		i = 1
		
		While (Environ(i) <> "")
			Name_Renamed = Mid(Environ(i), 1, InStr(1, Environ(i), "=") - 1)
			
			Value = Mid(Environ(i), InStr(1, Environ(i), "=") + 1)
			
			tmp = Replace(tmp, "%" & Name_Renamed & "%", Value)
			
			i = i + 1
		End While
		
		ReplaceEnvironmentVars = tmp
		
	End Function
	
	'UPGRADE_NOTE: Tag was upgraded to Tag_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Function AddQ(ByVal Message As String, Optional ByRef msg_priority As Short = -1, Optional ByVal User As String = vbNullString, Optional ByVal Tag_Renamed As String = vbNullString, Optional ByRef OversizeDelimiter As String = " ") As Short
		
		On Error GoTo ERROR_HANDLER
		
		Dim Splt() As String
		Dim strTmp As String
		Dim i As Integer
		Dim currChar As Integer
		Dim Send As String
		'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Command_Renamed As String
		Dim GTC As Double
		Dim Q As clsQueueOBj
		Dim delay As Integer
		Dim Index As Integer
		Dim s As String ' temp string for settings
		Dim MaxLength As Short ' stores max length for split (with override)
		
		Static LastGTC As Double
		Static BanCount As Short
		
		strTmp = Message
		
		' cap priority at 100
		If (msg_priority > 100) Then
			msg_priority = 100
		End If
		
		If (g_Queue.Count = 0) Then
			BanCount = 0
		End If
		
		Dim cmdName As String
		Dim spaceIndex As Integer
		If (strTmp <> vbNullString) Then
			ReDim Splt(0)
			
			' check for tabs and replace with spaces (2005-09-23)
			If (InStr(1, strTmp, Chr(9), CompareMethod.Binary) <> 0) Then
				strTmp = Replace(strTmp, Chr(9), Space(4))
			End If
			
			' check for invalid characters in the message
			For i = 1 To Len(strTmp)
				currChar = Asc(Mid(strTmp, i, 1))
				
				If (currChar < 32) Then
					Exit Function
				End If
			Next i
			
			' is this an internal or battle.net command?
			If (StrComp(VB.Left(strTmp, 1), "/", CompareMethod.Binary) = 0) Then
				' if so, we have extra work to do
				For i = 2 To Len(strTmp)
					currChar = Asc(Mid(strTmp, i, 1))
					
					' find the first non-space after the /
					If (Not currChar = Asc(Space(1))) Then
						Exit For
					End If
				Next i
				
				' if we found a non-space, strip everything
				If (i >= 2) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strTmp = StringFormat("/{0}", Mid(strTmp, i))
				End If
				
				' Find the next instance of a space (the end of the command word)
				Index = InStr(1, strTmp, Space(1), CompareMethod.Binary)
				
				' is it a valid command word?
				If (Index > 2) Then
					' extract the command word
					Command_Renamed = LCase(Mid(strTmp, 2, Index - 2))
					
					' test it for being battle.net commands we need to process now
					If ((Command_Renamed = "w") Or (Command_Renamed = "whisper") Or (Command_Renamed = "m") Or (Command_Renamed = "msg") Or (Command_Renamed = "message") Or (Command_Renamed = "whois") Or (Command_Renamed = "where") Or (Command_Renamed = "whereis") Or (Command_Renamed = "squelch") Or (Command_Renamed = "unsquelch") Or (Command_Renamed = "ignore") Or (Command_Renamed = "unignore") Or (Command_Renamed = "ban") Or (Command_Renamed = "unban") Or (Command_Renamed = "kick") Or (Command_Renamed = "designate")) Then
						
						Splt = Split(strTmp, Space(1), 3)
						
						If (UBound(Splt) > 0) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Command_Renamed = StringFormat("{0} {1}", Splt(0), ReverseConvertUsernameGateway(Splt(1)))
							
							If ((g_Channel.IsSilent) And (Me.mnuDisableVoidView.Checked = False)) Then
								If ((LCase(Splt(0)) = "/unignore") Or (LCase(Splt(0)) = "/unsquelch")) Then
									If (StrComp(Splt(1), GetCurrentUsername, CompareMethod.Text) = 0) Then
										lvChannel.Items.Clear()
									End If
								End If
							End If
							
							If (UBound(Splt) > 1) Then
								ReDim Preserve Splt(UBound(Splt) - 1)
							End If
						End If
						
					ElseIf ((Command_Renamed = "f") Or (Command_Renamed = "friends")) Then 
						
						Splt = Split(strTmp, Space(1), 3)
						
						Command_Renamed = Splt(0)
						
						If (UBound(Splt) >= 1) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Command_Renamed = StringFormat("{0} {1}", Command_Renamed, Splt(1))
							
							If (UBound(Splt) >= 2) Then
								Select Case (LCase(Splt(1)))
									Case "m", "msg"
										ReDim Preserve Splt(UBound(Splt) - 1)
										
									Case Else
										Splt = Split(strTmp, Space(1), 4)
										
										If ((StrReverse(BotVars.Product) = PRODUCT_WAR3) Or (StrReverse(BotVars.Product) = PRODUCT_W3XP)) Then
											
											'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											Command_Renamed = StringFormat("{0} {1}", Command_Renamed, ReverseConvertUsernameGateway(Splt(2)))
										Else
											'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											Command_Renamed = StringFormat("{0} {1}", Command_Renamed, Splt(2))
										End If
										
										If (UBound(Splt) >= 3) Then
											'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											Command_Renamed = StringFormat("{0} {1}", Command_Renamed, Splt(3))
										End If
								End Select
							End If
						End If
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Command_Renamed = StringFormat("/{0}", Command_Renamed)
						strTmp = Mid(strTmp, Len(Command_Renamed) + 2)
					End If
					
					If (Len(Command_Renamed) >= BNET_MSG_LENGTH) Then
						Exit Function
					End If
					
					If (UBound(Splt) > 0) Then
						strTmp = Mid(strTmp, Len(Join(Splt, Space(1))) + (Len(Space(1))) + 1)
					End If
				End If
			End If
			
			If (msg_priority < 0) Then
				
				spaceIndex = InStr(1, Message, Space(1), CompareMethod.Binary)
				
				If (spaceIndex >= 2) Then
					cmdName = LCase(VB.Left(Mid(Message, 2), spaceIndex - 2))
				Else
					cmdName = LCase(Mid(Message, 2))
				End If
				
				Select Case (cmdName)
					Case "designate" : msg_priority = modQueueObj.PRIORITY.SPECIAL_MESSAGE
					Case "resign" : msg_priority = modQueueObj.PRIORITY.SPECIAL_MESSAGE
					Case "who" : msg_priority = modQueueObj.PRIORITY.SPECIAL_MESSAGE
					Case "unban" : msg_priority = modQueueObj.PRIORITY.SPECIAL_MESSAGE
					Case "clan", "c" : msg_priority = modQueueObj.PRIORITY.SPECIAL_MESSAGE
					Case "ban" : msg_priority = modQueueObj.PRIORITY.CHANNEL_MODERATION_MESSAGE
					Case "kick" : msg_priority = modQueueObj.PRIORITY.CHANNEL_MODERATION_MESSAGE
					Case Else : msg_priority = modQueueObj.PRIORITY.MESSAGE_DEFAULT
				End Select
			End If
			
			MaxLength = Config.MaxMessageLength
			
			Call SplitByLen(strTmp, MaxLength - Len(Command_Renamed), Splt, vbNullString,  , OversizeDelimiter)
			
			ReDim Preserve Splt(UBound(Splt))
			
			' add to the queue!
			For i = LBound(Splt) To UBound(Splt)
				' store current tick
				GTC = GetTickCount()
				
				' store working copy
				Send = Splt(i)
                If (Len(Command_Renamed) > 0) Then
                    If (Len(Send) > 0) Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Send = StringFormat("{0} {1}", Command_Renamed, Send)
                    Else
                        Send = Command_Renamed
                    End If
                ElseIf (VB.Left(Send, 1) = "/" And i > LBound(Splt)) Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Send = StringFormat(" {0}", Send)
                End If
				
				' create the queue object
				Q = New clsQueueOBj
				
				With Q
					.Message = Send
					.PRIORITY = msg_priority
					.Tag = Tag_Renamed
				End With
				
				' add it
				g_Queue.Push(Q)
				
				' should we subject this message to the typical delay,
				' or can we get it out of here a bit faster?  If we
				' want it out of here quick, we need an empty queue
				' and have had at least 10 seconds elapse since the
				' previous message.
				If (g_Queue.Count() = 1) Then
					If (GTC - LastGTC >= 10000) Then
						' set default message delay when queue is empty (in ms)
						delay = 10
						
						' are we issuing a ban or kick command?
						If (msg_priority = modQueueObj.PRIORITY.CHANNEL_MODERATION_MESSAGE) Then
							delay = g_BNCSQueue.BanDelay()
						End If
					End If
				End If
				
				' If queueTimerID is 0 the timer is idle right now, so reset it
				If QueueTimerID = 0 Then
					If delay = 0 Then
						delay = g_BNCSQueue.GetDelay(g_Queue.Peek.Message)
					End If
					
					' set the delay before our next queue cycle
                    QueueTimerID = SetTimer(0, 0, delay, AddressOf QueueTimerProc)
				End If
			Next i
			
			AddQ = UBound(Splt) + 1
			
			' store our tick for future reference
			LastGTC = GTC
		End If
		
		Exit Function
		
ERROR_HANDLER: 
		Call AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: " & Err.Description & " in AddQ().")
		
		Exit Function
	End Function
	
	
	Sub ClearChannel()
		' reset channel object
		'UPGRADE_NOTE: Object g_Channel may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		g_Channel = Nothing
		g_Channel = New clsChannelObj
		
		' clear channel UI elements
		lvChannel.Items.Clear()
		lblCurrentChannel.Text = vbNullString
		
		' reset this random boolean
		PassedClanMotdCheck = False
	End Sub
	
	
	Sub ReloadConfig(Optional ByRef Mode As Byte = 0)
		On Error GoTo ERROR_HANDLER
		
		Dim default_group_access As udtGetAccessResponse
		Dim s As String
		Dim i As Short
		Dim f As Short
		Dim Index As Short
		Dim bln As Boolean
		Dim doConvert As Boolean
		Dim command_output() As String
		
		Dim oCommandGenerator As New clsCommandGeneratorObj
		
		If Mode <> 0 Then
			Config.Load(GetConfigFilePath())
		End If
		
		BotVars.TSSetting = Config.TimestampMode
		
		' Client settings
        If Len(BotVars.Username) > 0 And StrComp(BotVars.Username, Config.Username, CompareMethod.Text) <> 0 Then
            AddChat(RTBColors.ConsoleText, "Username set to " & Config.Username & ".")
        End If
		BotVars.Username = Config.Username
		BotVars.Password = Config.Password
		BotVars.CDKey = Config.CDKey
		BotVars.ExpKey = Config.ExpKey
		BotVars.Product = StrReverse(GetProductInfo(Config.Game).Code)
		BotVars.Server = Config.Server
		BotVars.HomeChannel = Config.HomeChannel
		BotVars.BotOwner = Config.BotOwner
		
		BotVars.Trigger = Config.Trigger
		'If (BotVars.TriggerLong = vbNullString) Then
		'    BotVars.Trigger = "."
		'End If
		
		BotVars.BNLSServer = Config.BNLSServer
		
		' Load database and commands
		Call LoadDatabase()
		Call oCommandGenerator.GenerateCommands()
		
		' Set UI fonts
		Dim ResizeChatElements As Boolean
		Dim ResizeChannelElements As Boolean
		Dim lblHeight As Single
		If Mode <> 1 Then
			
			s = Config.ChatFont
			If s <> vbNullString And s <> rtbChat.Font.Name Then
				'rtbChat.Font.Name = s
				cboSend.Font = VB6.FontChangeName(cboSend.Font, s)
				'rtbWhispers.Font.Name = s
				txtPre.Font = VB6.FontChangeName(txtPre.Font, s)
				txtPost.Font = VB6.FontChangeName(txtPost.Font, s)
				
				ResizeChatElements = True
			End If
			
			s = Config.ChannelListFont
			If s <> vbNullString And s <> lvChannel.Font.Name Then
				lvChannel.Font = VB6.FontChangeName(lvChannel.Font, s)
				lvClanList.Font = VB6.FontChangeName(lvClanList.Font, s)
				lvFriendList.Font = VB6.FontChangeName(lvFriendList.Font, s)
				lblCurrentChannel.Font = VB6.FontChangeName(lblCurrentChannel.Font, s)
				ListviewTabs.Font = VB6.FontChangeName(ListviewTabs.Font, s)
				
				ResizeChatElements = True
			End If
			
			s = CStr(Config.ChatFontSize)
			If StrictIsNumeric(s) Then
				If CShort(s) <> rtbChat.Font.SizeInPoints Then
					'rtbChat.Font.Size = s
					cboSend.Font = VB6.FontChangeSize(cboSend.Font, CDec(s))
					'rtbWhispers.Font.Size = s
					txtPre.Font = VB6.FontChangeSize(txtPre.Font, CDec(s))
					txtPost.Font = VB6.FontChangeSize(txtPost.Font, CDec(s))
					
					ResizeChannelElements = True
				End If
			End If
			
			s = CStr(Config.ChannelListFontSize)
			If StrictIsNumeric(s) Then
				If CShort(s) <> lvChannel.Font.SizeInPoints Then
					lvChannel.Font = VB6.FontChangeSize(lvChannel.Font, CDec(s))
					lvClanList.Font = VB6.FontChangeSize(lvClanList.Font, CDec(s))
					lvFriendList.Font = VB6.FontChangeSize(lvFriendList.Font, CDec(s))
					lblCurrentChannel.Font = VB6.FontChangeSize(lblCurrentChannel.Font, CDec(s))
					ListviewTabs.Font = VB6.FontChangeSize(ListviewTabs.Font, CDec(s))
					
					ResizeChannelElements = True
				End If
			End If
			
			If ResizeChannelElements Then
				
				lblCurrentChannel.AutoSize = True
				lblHeight = VB6.PixelsToTwipsY(lblCurrentChannel.Height) + 40
				lblCurrentChannel.AutoSize = False
				lblCurrentChannel.Height = VB6.TwipsToPixelsY(lblHeight)
				
				ResizeChatElements = True
			End If
			
			If ResizeChatElements Then
				Call ChangeRTBFont(rtbChat, Config.ChatFont, Config.ChannelListFontSize)
				Call ChangeRTBFont(rtbWhispers, Config.ChatFont, Config.ChannelListFontSize)
				
				frmChat_Resize(Me, New System.EventArgs())
			End If
		End If
		
		Filters = Config.ChatFilters
		mnuToggleFilters.Checked = Filters
		If (Not Filters) Then
			BotVars.JoinWatch = 0
		End If
		
		BotVars.AutofilterMS = 0
		
		AutoModSafelistValue = Config.AutoSafelistLevel
		BotVars.ShowOfflineFriends = Config.ShowOfflineFriends
		
		If Config.HideClanDisplay Then
			With lvChannel
				'UPGRADE_WARNING: Lower bound of collection lvChannel.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(.Width) - VB6.PixelsToTwipsX(.Columns.Item(2).Width)))
				'UPGRADE_WARNING: Lower bound of collection lvChannel.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				.Columns.Item(2).Width = 0
			End With
		End If
		
		If Config.HidePingDisplay Then
			With lvChannel
				'UPGRADE_WARNING: Lower bound of collection lvChannel.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(.Width) - VB6.PixelsToTwipsX(.Columns.Item(3).Width)))
				'UPGRADE_WARNING: Lower bound of collection lvChannel.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				.Columns.Item(3).Width = 0
			End With
		End If
		
		BotVars.RetainOldBans = Config.RetainOldBans
		BotVars.StoreAllBans = Config.StoreAllBans
		
		BotVars.GatewayConventions = Config.NamespaceConvention
		BotVars.UseD2Naming = Config.UseD2Naming
		BotVars.D2NamingFormat = Config.D2NamingFormat
		
		BotVars.ShowStatsIcons = Config.ShowStatsIcons
		BotVars.ShowFlagsIcons = Config.ShowFlagIcons
		
		Dim found As System.Windows.Forms.ListViewItem
		Dim CurrentUser As Object
		Dim outbuf As String
		If (g_Online) Then
			
			SetTitle(GetCurrentUsername & ", online in channel " & g_Channel.Name)
			
			Me.UpdateTrayTooltip()
			
			lvChannel.Items.Clear()
			
			For i = 1 To g_Channel.Users.Count()
				CurrentUser = g_Channel.Users.Item(i)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrentUser.Clan. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrentUser.Stats. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrentUser.Ping. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrentUser.Flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrentUser.Game. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrentUser.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrentUser.DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AddName(CurrentUser.DisplayName, CurrentUser.Name, CurrentUser.Game, CurrentUser.Flags, CurrentUser.Ping, CurrentUser.Stats.IconCode, CurrentUser.Clan)
			Next i
			
			Me.lvFriendList.Items.Clear()
			
			For i = 1 To g_Friends.Count()
				CurrentUser = g_Friends.Item(i)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrentUser.Status. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrentUser.Game. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrentUser.DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AddFriend(CurrentUser.DisplayName, CurrentUser.Game, CurrentUser.Status)
			Next i
		End If
		
		JoinMessagesOff = Not Config.ShowJoinLeaves
		mnuToggle.Checked = JoinMessagesOff
		
		mail = Config.BotMail
		
		BotVars.BanEvasion = Config.BanEvasion
		BotVars.Logging = Config.LoggingMode
		
		mnuToggleWWUse.Checked = Config.WhisperWindows
		BotVars.WhisperCmds = Config.WhisperCommands
		PhraseBans = Config.PhraseBans
		BotVars.CaseSensitiveFlags = Config.CaseSensitiveDBFlags
		BotVars.AutoCompletePostfix = Config.AutoCompletePostfix
		BotVars.BNLS = Config.UseBNLS
		BotVars.LogDBActions = Config.LogDBActions
		BotVars.LogCommands = Config.LogCommands
		
		'/* time to idle: defaults to 600 seconds / 10 minutes idle */
		BotVars.SecondsToIdle = Config.SecondsToIdle
		
		BotVars.BanUnderLevel = Config.LevelBanW3
		BotVars.BanUnderLevelMsg = Config.LevelBanMessage
		BotVars.BanPeons = Config.PeonBan
		
		BotVars.KickOnYell = Config.KickOnYell
		
		' Capped at 32767, topic=29986 -Andy
		BotVars.IB_Wait = Config.IdleBanDelay
		
		BotVars.DefaultShitlistGroup = Config.ShitlistGroup
		If (BotVars.DefaultShitlistGroup <> vbNullString) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object default_group_access. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			default_group_access = GetAccess(BotVars.DefaultShitlistGroup, "GROUP")
			
			If (default_group_access.Username = vbNullString) Then
				Call ProcessCommand(GetCurrentUsername, "/add " & BotVars.DefaultShitlistGroup & " B --type group --banmsg Shitlisted", True, False, False)
			End If
		End If
		
		BotVars.DefaultTagbansGroup = Config.TagbanGroup
		If (BotVars.DefaultTagbansGroup <> vbNullString) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object default_group_access. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			default_group_access = GetAccess(BotVars.DefaultTagbansGroup, "GROUP")
			
			If (default_group_access.Username = vbNullString) Then
				Call ProcessCommand(CurrentUsername, "/add " & BotVars.DefaultTagbansGroup & " B --type group --banmsg Tagbanned", True, False, False)
			End If
		End If
		
		BotVars.DefaultSafelistGroup = Config.SafelistGroup
		If (BotVars.DefaultSafelistGroup <> vbNullString) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object default_group_access. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			default_group_access = GetAccess(BotVars.DefaultSafelistGroup, "GROUP")
			
			If (default_group_access.Username = vbNullString) Then
				Call ProcessCommand(GetCurrentUsername, "/add " & BotVars.DefaultSafelistGroup & " S --type group", True, False, False)
			End If
		End If
		
		BotVars.DisableMP3Commands = Not Config.Mp3Commands
		
		BotVars.MaxBacklogSize = Config.MaxBacklogSize
		BotVars.MaxLogFileSize = Config.MaxLogFileSize
		
		If Config.UrlDetection Then
			EnableURLDetect(rtbChat.Handle.ToInt32)
		Else
			DisableURLDetect(rtbChat.Handle.ToInt32)
		End If
		
		' reload quotes file
		g_Quotes = New clsQuotesObj
		
		BotVars.UseBackupChan = Config.UseBackupChannel
		BotVars.BackupChan = Config.BackupChannel
		
		mnuUTF8.Checked = Config.UseUTF8
		
		mnuToggleShowOutgoing.Checked = Config.ShowOutgoingWhispers
		mnuHideWhispersInrtbChat.Checked = Config.HideWhispersInMain
		mnuIgnoreInvites.Checked = Config.IgnoreClanInvites
		
		'LoadSafelist
		LoadArray(LOAD_PHRASES, Phrases)
		LoadArray(LOAD_FILTERS, gFilters)
		
		ProtectMsg = Config.ChannelProtectionMessage
		
		Call LoadOutFilters()
		
		BotVars.IB_On = Config.IdleBan
		BotVars.IB_Kick = Config.IdleBanKick
		BotVars.IB_Wait = Config.IdleBanDelay
		
		BotVars.Spoof = Config.PingSpoofing
		
		Protect = Config.ChannelProtection
		BotVars.UseUDP = Config.UseUDP
		BotVars.IPBans = Config.IPBans
		BotVars.UseAltBnls = Config.BNLSFinder
		BotVars.QuietTime = Config.QuietTime
		
		mnuFlash.Checked = Config.FlashOnEvents
		
		BotVars.UseProxy = Config.UseProxy
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If BotVars.UseProxy And sckBNet.CtlState = MSWinsockLib.StateConstants.sckConnected Then BotVars.ProxyStatus = modEnum.enuProxyStatus.psOnline
		BotVars.ProxyPort = Config.ProxyPort
		BotVars.ProxyIsSocks5 = CBool(Config.ProxyType)
		BotVars.NoTray = Not Config.MinimizeToTray
		BotVars.NoAutocompletion = Not Config.NameAutoComplete
		BotVars.NoColoring = Not Config.NameColoring
		
		mnuDisableVoidView.Checked = Not Config.VoidView
		
		BotVars.MediaPlayer = Config.MediaPlayer
		
		
		' Load some queue stuff, reluctantly
		BotVars.QueueMaxCredits = Config.QueueMaxCredits
		BotVars.QueueCostPerPacket = Config.QueueCostPerPacket
		BotVars.QueueCostPerByte = Config.QueueCostPerByte
		BotVars.QueueCostPerByteOverThreshhold = Config.QueueCostPerByteOver
		BotVars.QueueStartingCredits = Config.QueueStartingCredits
		BotVars.QueueThreshholdBytes = Config.QueueThresholdBytes
		BotVars.QueueCreditRate = Config.QueueCreditRate
		
		BotVars.UseRealm = Config.UseD2Realms
		
		txtPre.Text = ""
		txtPost.Text = ""
		
		txtPre.Visible = Not Config.DisablePrefixBox
		mnuPopAddLeft.Enabled = Not Config.DisablePrefixBox
		mnuPopFLAddLeft.Enabled = Not Config.DisablePrefixBox
		mnuPopClanAddLeft.Enabled = Not Config.DisablePrefixBox
		
		txtPost.Visible = Not Config.DisableSuffixBox
		
		'[Other] MathAllowUI - Will allow People to use MessageBox/InputBox or other UI related commands in the .eval/.math commands ~Hdx 09-25-07
		SCRestricted.AllowUI = Config.MathAllowUI
		BotVars.NoRTBAutomaticCopy = Config.DisableRTBAutoCopy
		
		BotVars.UseGreet = Config.GreetMessage
		BotVars.GreetMsg = Config.GreetMessageText
		BotVars.WhisperGreet = Config.WhisperGreet
		
		BotVars.ProxyIP = Config.ProxyIP
		
		BotVars.ChatDelay = Config.ChatDelay
		
		s = Config.GetFilePath("Logs")
		If (Not (s = vbNullString)) Then
			g_Logger.LogPath = s
		End If
		
		Call ChatQueue_Initialize()
		
		' I reluctantly add the queue variables here.
		
		If (g_Online) Then
			Call g_Channel.CheckUsers()
		Else
			Err.Clear()
			
			'//Removed 10/29/09 - Hdx - I'll add in this feature later properly, does not work as is.
			'If (ReadCfg(OV, "LocalIP") <> vbNullString) Then
			'    If (Err.Number = 0) Then: sckBNet.bind , ReadCfg(OV, "LocalIP")
			'    If (Err.Number = 0) Then: sckBNLS.bind , ReadCfg(OV, "LocalIP")
			'    If (Err.Number = 0) Then: sckMCP.bind , ReadCfg(OV, "LocalIP")
			'End If
		End If
		
		' disable the script system if override is set
		modScripting.SetScriptSystemDisabled(Config.DisableScripting)
		
		'UPGRADE_NOTE: Object oCommandGenerator may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oCommandGenerator = Nothing
		
		Exit Sub
		
ERROR_HANDLER: 
		If (Err.Number = 10049) Then
			AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in ReloadConfig().")
		End If
		
		Resume Next
	End Sub
	
	Private Sub ChangeRTBFont(ByRef rtb As System.Windows.Forms.RichTextBox, ByVal NewFont As String, ByVal NewSize As Short)
		Dim tmpBuffer As String
		
		With rtb
			.SelectionStart = 0
			.SelectionLength = Len(.Text)
			.SelectionFont = VB6.FontChangeSize(.SelectionFont, NewSize)
			.SelectionFont = VB6.FontChangeName(.SelectionFont, NewFont)
			tmpBuffer = .RTF
			.Text = vbNullString
			.Font = VB6.FontChangeName(.Font, NewFont)
			.Font = VB6.FontChangeSize(.Font, NewSize)
			'UPGRADE_WARNING: TextRTF was upgraded to Text and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			.Text = tmpBuffer
			.SelectionStart = Len(.Text)
		End With
	End Sub
	
	'returns OK to Proceed
	Function DisplayError(ByVal ErrorNumber As Short, ByRef bytType As Byte, ByVal source As modEnum.enuErrorSources) As Boolean
		
		Dim s As String
		
		s = GErrorHandler.GetErrorString(ErrorNumber, source)
		
        If (Len(s) > 0) Then
            Select Case (bytType)
                Case 0 : s = "[BNLS] " & s
                Case 1 : s = "[BNCS] " & s
                Case 2 : s = "[PROXY] " & s
            End Select

            AddChat(RTBColors.ErrorMessageText, s)
        End If
		
		DisplayError = GErrorHandler.OKToProceed()
	End Function
	
	Sub LoadOutFilters()
		Const o As String = "Outgoing"
		Const f As String = FILE_FILTERS
		
		Dim s As String
		Dim i As Short
		
		'UPGRADE_WARNING: Lower bound of array gOutFilters was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim gOutFilters(1)
		'UPGRADE_NOTE: Catch was upgraded to Catch_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		ReDim Catch_Renamed(0)
		
		Catch_Renamed(0) = vbNullString
		
		s = ReadINI(o, "Total", f)
		
		If (Not (StrictIsNumeric(s))) Then
			Exit Sub
		End If
		
		For i = 1 To Val(s)
			gOutFilters(i).ofFind = Replace(LCase(ReadINI(o, "Find" & i, f)), "¦", " ")
			gOutFilters(i).ofReplace = Replace(ReadINI(o, "Replace" & i, f), "¦", " ")
			
			If (i <> Val(s)) Then
				'UPGRADE_WARNING: Lower bound of array gOutFilters was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
				ReDim Preserve gOutFilters(i + 1)
			End If
		Next i
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_NOTE: Catch was upgraded to Catch_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		ReDim Preserve Catch_Renamed(UBound(Catch_Renamed) + 1)
		If (Dir(GetFilePath(FILE_CATCH_PHRASES)) <> vbNullString) Then
			i = FreeFile
			
			FileOpen(i, GetFilePath(FILE_CATCH_PHRASES), OpenMode.Input)
			
			If (LOF(i) < 2) Then
				FileClose(i)
				
				Exit Sub
			End If
			
			Do While Not EOF(i)
				s = LineInput(i)
				
				If ((s <> vbNullString) And (s <> " ")) Then
					Catch_Renamed(UBound(Catch_Renamed)) = LCase(s)
					
				End If
			Loop 
			
			'Note: Why did this happen?
			'If Catch(0) = vbNullString Then Catch(0) = "¯"
			
			FileClose(i)
		End If
	End Sub
	
	Function OutFilterMsg(ByVal strOut As String) As String
		Dim i As Short
		
		If (UBound(gOutFilters) > 0) Then
			For i = LBound(gOutFilters) To UBound(gOutFilters)
				strOut = Replace(strOut, gOutFilters(i).ofFind, gOutFilters(i).ofReplace)
			Next i
		End If
		
		OutFilterMsg = strOut
	End Function
	
	Private Sub sckBNet_DataArrival(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_DataArrivalEvent) Handles sckBNet.DataArrival
		'On Error GoTo ERROR_HANDLER
		
        Dim bData() As Byte
		Dim fTemp As String
		Dim BufferLimit As Integer
		Dim interations As Short
		
        sckBNet.GetData(bData, vbArray + vbByte, eventArgs.bytesTotal)
		
		If Not BotVars.UseProxy Or BotVars.ProxyStatus = modEnum.enuProxyStatus.psOnline Then
            BNCSBuffer.AddData(bData)

            While BNCSBuffer.FullPacket And BufferLimit < 20

                bData = BNCSBuffer.GetPacket

                Call BNCSParsePacket(bData)

                'interations = (interations + 1)

                'If (interations >= 2000) Then
                '    MsgBox "ahhhh!"
                '
                '    Exit Sub
                'End If

                ' Why do we need this?  Anyway, it's causing topic id #26093
                ' (The Void issue).
                'BufferLimit = (BufferLimit + 1) 'DebugOutput Left$(strBuffer, lngLen)
            End While
        Else
            'proxy is ON and NOT CONNECTED
            'parse incoming data
            ParseProxyPacket(bData)
        End If
		
		Exit Sub
		
ERROR_HANDLER: 
		AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in sckBNet_DataArrival().")
		
		Exit Sub
	End Sub
	
	Sub LoadArray(ByVal Mode As Byte, ByRef tArray() As String)
		Dim f As Short
		Dim Path As String
		Dim temp As String
		Dim i As Short
		Dim c As Short
		
		f = FreeFile
		
		Const FI As String = "TextFilters"
		
		Select Case Mode
			Case LOAD_FILTERS
				Path = GetFilePath(FILE_FILTERS)
			Case LOAD_PHRASES
				Path = GetFilePath(FILE_PHRASE_BANS)
			Case LOAD_DB
				Path = GetFilePath(FILE_USERDB)
				Exit Sub
		End Select
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Dir(Path) <> vbNullString Then
			FileOpen(f, Path, OpenMode.Input)
			If LOF(f) > 2 Then
				ReDim tArray(0)
				If Mode <> LOAD_FILTERS Then
					Do 
						temp = LineInput(f)
						If Len(temp) > 0 Then
							' removed for 2.5 - why am I PCing it ?
							'If Mode = LOAD_SAFELIST Then temp = PrepareCheck(temp)
							tArray(UBound(tArray)) = LCase(temp)
							ReDim Preserve tArray(UBound(tArray) + 1)
						End If
					Loop While Not EOF(f)
				Else
					temp = ReadINI(FI, "Total", FILE_FILTERS)
					If temp <> vbNullString And CShort(temp) > -1 Then
						c = Int(CDbl(temp))
						For i = 1 To c
							temp = ReadINI(FI, "Filter" & i, FILE_FILTERS)
							If temp <> vbNullString Then
								tArray(UBound(tArray)) = LCase(temp)
								If i <> c Then ReDim Preserve tArray(UBound(tArray) + 1)
							End If
						Next i
					End If
				End If
			End If
			FileClose(f)
		End If
	End Sub
	
	Private Sub sckBNLS_CloseEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sckBNLS.CloseEvent
		If MDebug("all") Then
			AddChat(COLOR_BLUE, "BNLS CLOSE")
		End If
		If (Not BNLSAuthorized) Then
			AddChat(RTBColors.ErrorMessageText, "You have been disconnected from the BNLS server. You may be IPbanned from the server, it may be having issues, or there is something blocking your connection.")
			AddChat(RTBColors.ErrorMessageText, "Try using another BNLS server to connect, and check your firewall settings.")
		End If
	End Sub
	
	Private Sub sckBNLS_ConnectEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sckBNLS.ConnectEvent
		If MDebug("all") Then
			AddChat(COLOR_BLUE, "BNLS CONNECT")
		End If
		
		Call Event_BNLSConnected()
		
		'With PBuffer
		'    .InsertNTString "stealth"
		'    .vLSendPacket &HE
		'End With
		modBNLS.SEND_BNLS_AUTHORIZE()
		
		SetNagelStatus(sckBNLS.SocketHandle, False)
		
		'frmChat.sckBNet.Connect 'BNLS is authorized, proceed to initiate BNet connection.
	End Sub
	
	Private Sub sckBNLS_DataArrival(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_DataArrivalEvent) Handles sckBNLS.DataArrival
		On Error GoTo ERROR_HANDLER

        Dim bData() As Byte
		
        sckBNLS.GetData(bData, vbArray + vbByte, eventArgs.bytesTotal)
		
		If BotVars.UseProxy And (BotVars.ProxyStatus = modEnum.enuProxyStatus.psConnecting Or BotVars.ProxyStatus = modEnum.enuProxyStatus.psLoggingIn) Then
			
			'        Debug.Print "prox input: " & DebugOutput(strTemp)
			
			Select Case BotVars.ProxyStatus
				Case modEnum.enuProxyStatus.psConnecting
					'chr(5) or chr(4) depending on version & method
					'do an instr search to find the method number.
					'For public proxys you are looking for chr(0)
					'<macyui>

                    If ArrayContains(bData, 4) Or ArrayContains(bData, 5) Then
                        If ArrayContains(bData, 0) Then
                            UpdateProxyStatus(modEnum.enuProxyStatus.psLoggingIn, PROXY_LOGGING_IN)

                            LogonToProxy(sckBNLS, BotVars.BNLSServer, 9367, False)
                        Else
                            UpdateProxyStatus(modEnum.enuProxyStatus.psNotConnected, PROXY_IS_NOT_PUBLIC)
                            sckBNLS.Close()
                        End If
                    Else
                        UpdateProxyStatus(modEnum.enuProxyStatus.psNotConnected, PROXY_IS_NOT_PUBLIC)
                        sckBNLS.Close()
                    End If

                Case modEnum.enuProxyStatus.psLoggingIn
                    'Then when it sends back chr(5) & chr(0) indicating
                    'connection success, login and proceed as usual.

                    If ArrayContains(bData, 5) And ArrayContains(bData, 0) Then
                        UpdateProxyStatus(modEnum.enuProxyStatus.psOnline, PROXY_LOGIN_SUCCESS)
                    Else
                        UpdateProxyStatus(modEnum.enuProxyStatus.psNotConnected, PROXY_LOGIN_FAILED)
                        sckBNLS.Close()
                    End If

            End Select
		Else
            BNLSBuffer.AddData(bData)
			
			While BNLSBuffer.FullPacket
				modBNLS.BNLSRecvPacket(BNLSBuffer.GetPacket)
			End While
		End If
		
		Exit Sub
		
ERROR_HANDLER: 
		AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in sckBNLS_DataArrival().")
		
		Exit Sub
	End Sub
	
	Private Sub sckBNLS_Error(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_ErrorEvent) Handles sckBNLS.Error
		Call Event_BNLSError(eventArgs.Number, eventArgs.Description)
	End Sub
	
	'This function checks if the user selected when right-clicked is the same one when they click on the menu option. - FrOzeN
	Private Function PopupMenuUserCheck() As Boolean
		If Not (lvChannel.FocusedItem Is Nothing) Then
			If mnuPopup.Tag <> lvChannel.FocusedItem.Text Then
				PopupMenuUserCheck = False
				Exit Function
			End If
		End If
		
		PopupMenuUserCheck = True
	End Function
	
	Private Function PopupMenuFLUserCheck() As Boolean
		If Not (lvFriendList.FocusedItem Is Nothing) Then
			If mnuPopFList.Tag <> lvFriendList.FocusedItem.Text Then
				PopupMenuFLUserCheck = False
				Exit Function
			End If
		End If
		
		PopupMenuFLUserCheck = True
	End Function
	
	Private Function PopupMenuCLUserCheck() As Boolean
		If Not (lvClanList.FocusedItem Is Nothing) Then
			If mnuPopClanList.Tag <> lvClanList.FocusedItem.Text Then
				PopupMenuCLUserCheck = False
				Exit Function
			End If
		End If
		
		PopupMenuCLUserCheck = True
	End Function
	
	Function GetSelectedUsers() As Collection
		Dim i As Short
		
		GetSelectedUsers = New Collection
		
		For i = 1 To lvChannel.Items.Count
			'UPGRADE_WARNING: Lower bound of collection lvChannel.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			If (lvChannel.Items.Item(i).Selected) Then
				'UPGRADE_WARNING: Lower bound of collection lvChannel.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				Call GetSelectedUsers.Add(lvChannel.Items.Item(i).Text)
			End If
		Next i
	End Function
	
	Function GetSelectedUser() As String
		If (lvChannel.FocusedItem Is Nothing) Then
			GetSelectedUser = vbNullString
			
			Exit Function
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object lvChannel.SelectedItem.Tag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetSelectedUser = lvChannel.FocusedItem.Tag
	End Function
	
	Function GetFriendsSelectedUser() As String
		If (lvFriendList.FocusedItem Is Nothing) Then
			GetFriendsSelectedUser = vbNullString
			
			Exit Function
		End If
		
		GetFriendsSelectedUser = CleanUsername(ReverseConvertUsernameGateway(lvFriendList.FocusedItem.Text))
	End Function
	
	Function GetRandomPerson() As String
		Dim i As Short
		
		If (g_Channel.Users.Count() > 0) Then
			Randomize()
			
			i = Int(g_Channel.Users.Count() * Rnd() + 1)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object g_Channel.Users().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetRandomPerson = g_Channel.Users.Item(i).DisplayName
		End If
	End Function
	
	Function MatchClosest(ByVal toMatch As String, Optional ByRef startIndex As Integer = 1) As String
		Dim lstView As System.Windows.Forms.ListView
		
		Dim i As Short
		Dim CurrentName As String
		Dim atChar As Short
		Dim Index As Short
		Dim Loops As Short
		
		i = InStr(1, toMatch, " ", CompareMethod.Binary)
		
		If (i > 0) Then
			toMatch = Mid(toMatch, i + 1)
		End If
		
		Select Case (ListviewTabs.SelectedIndex)
			Case 0
				lstView = lvChannel
			Case 1
				lstView = lvFriendList
			Case 2
				lstView = lvClanList
		End Select
		
		Dim c As Short
		With lstView.Items
			If (.Count > 0) Then
				
				If (startIndex > .Count) Then
					Index = 1
				Else
					Index = startIndex
				End If
				
				While (Loops < 2)
					For i = Index To .Count 'for each user
						'UPGRADE_WARNING: Lower bound of collection lstView.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
						CurrentName = .Item(i).Text
						
						If (Len(CurrentName) >= Len(toMatch)) Then
							For c = 1 To Len(toMatch) 'for each letter in their name
								If (StrComp(Mid(toMatch, c, 1), Mid(CurrentName, c, 1), CompareMethod.Text) <> 0) Then
									
									Exit For
								End If
							Next c
							
							If (c >= (Len(toMatch) + 1)) Then
								'UPGRADE_WARNING: Lower bound of collection lstView.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
								MatchClosest = .Item(i).Text
								
								MatchIndex = i
								
								Exit Function
							End If
						End If
					Next i
					
					Index = 1
					
					Loops = (Loops + 1)
				End While
				
				Loops = 0
			End If
		End With
		
		atChar = InStr(1, toMatch, "@", CompareMethod.Binary)
		
		Dim tmp As String
		Dim Gateways(5, 2) As String
		Dim OtherGateway As String
		Dim CurrentGateway As Short
		Dim j As Short
		If (atChar <> 0) Then
			
			' populate list
			Gateways(0, 0) = "USWest"
			Gateways(0, 1) = "Lordaeron"
			Gateways(1, 0) = "USEast"
			Gateways(1, 1) = "Azeroth"
			Gateways(2, 0) = "Asia"
			Gateways(2, 1) = "Kalimdor"
			Gateways(3, 0) = "Europe"
			Gateways(3, 1) = "Northrend"
			Gateways(4, 0) = "Beta"
			Gateways(4, 1) = "Westfall"
			
			CurrentGateway = -1
            If (Len(BotVars.Gateway) > 0) Then
                For i = 0 To UBound(Gateways, 1)
                    If (StrComp(BotVars.Gateway, Gateways(i, 0)) = 0) Then
                        CurrentGateway = i
                        OtherGateway = Gateways(i, 1)
                        Exit For
                    End If
                    If (StrComp(BotVars.Gateway, Gateways(i, 1)) = 0) Then
                        CurrentGateway = i
                        OtherGateway = Gateways(i, 0)
                        Exit For
                    End If
                Next i
                If (CurrentGateway = -1) Then ' BotVars.Gateway not known, @[tab]=@BotVars.Gateway
                    OtherGateway = BotVars.Gateway
                    CurrentGateway = 0
                End If
            Else ' BotVars.Gateway is nothing, @[tab]
                MatchClosest = vbNullString

                MatchIndex = 1

                Exit Function
            End If
			
			
			If (startIndex > UBound(Gateways, 2)) Then
				Index = 0
			Else
				Index = (startIndex - 1)
			End If
			
			If (Len(toMatch) >= (atChar + 1)) Then
				tmp = Mid(toMatch, atChar + 1)
				
				While (Loops < 2)
					If (Len(OtherGateway) >= Len(tmp)) Then
						If (StrComp(VB.Left(OtherGateway, Len(tmp)), tmp, CompareMethod.Text) = 0) Then
							
							
							MatchClosest = VB.Left(toMatch, atChar) & Gateways(CurrentGateway, i)
							
							MatchIndex = (i + 1)
							
							Exit Function
						End If
					End If
					
					Index = 0
					
					Loops = (Loops + 1)
				End While
			Else
				If (tmp = vbNullString) Then
					MatchClosest = VB.Left(toMatch, atChar) & OtherGateway
					
					MatchIndex = (Index + 1)
					
					Exit Function
				End If
			End If
		End If
		
		MatchClosest = vbNullString
		
		MatchIndex = 1
	End Function
	
	Function GetChannelString() As String
		If Not g_Online Then
			GetChannelString = vbNullString
		Else
			Select Case ListviewTabs.SelectedIndex
				Case 0
                    If Len(g_Channel.Name) = 0 Then
                        GetChannelString = BotVars.Gateway
                    Else
                        GetChannelString = g_Channel.Name & " (" & lvChannel.Items.Count & ")"
                    End If
				Case 1 : GetChannelString = "Friends (" & lvFriendList.Items.Count & ")"
				Case 2 : GetChannelString = "Clan " & g_Clan.Name & " (" & lvClanList.Items.Count & " members)"
			End Select
		End If
	End Function
	
	'this is a fucking mess. It reads:
	'"This copy of StealthBot has been tampered with. Please get a new copy of StealthBot at http://www.stealthbot.net.
	'Additionally, please report the website at which you downloaded StealthBot in an e-mail to abuse@stealthbot.net. Thanks!"
	
	Function GetHexProtectionMessage() As String
		GetHexProtectionMessage = Chr(Asc("T")) & Chr(Asc("h")) & Chr(Asc("i")) & Chr(Asc("s")) & Chr(Asc(" ")) & Chr(Asc("c")) & Chr(Asc("o")) & Chr(Asc("p")) & Chr(Asc("y")) & Chr(Asc(" ")) & Chr(Asc("o")) & Chr(Asc("f")) & Chr(Asc(" ")) & Chr(Asc("S")) & Chr(Asc("t")) & Chr(Asc("e")) & Chr(Asc("a")) & Chr(Asc("l")) & Chr(Asc("t")) & Chr(Asc("h")) & Chr(Asc("B")) & Chr(Asc("o")) & Chr(Asc("t")) & Chr(Asc(" ")) & Chr(Asc("h")) & Chr(Asc("a")) & Chr(Asc("s")) & Chr(Asc(" ")) & Chr(Asc("b")) & Chr(Asc("e")) & Chr(Asc("e")) & Chr(Asc("n")) & Chr(Asc(" ")) & Chr(Asc("t")) & Chr(Asc("a")) & Chr(Asc("m")) & Chr(Asc("p")) & Chr(Asc("e")) & Chr(Asc("r")) & Chr(Asc("e")) & Chr(Asc("d")) & Chr(Asc(" ")) & Chr(Asc("w")) & Chr(Asc("i")) & Chr(Asc("t")) & Chr(Asc("h")) & Chr(Asc(".")) & Chr(Asc(" ")) & Chr(Asc("P")) & Chr(Asc("l")) & Chr(Asc("e")) & Chr(Asc("a")) & Chr(Asc("s")) & Chr(Asc("e")) & Chr(Asc(" ")) & Chr(Asc("g")) & Chr(Asc("e")) & Chr(Asc("t")) & Chr(Asc(" ")) & Chr(Asc("a")) & Chr(Asc(" ")) & Chr(Asc("n")) & Chr(Asc("e")) & Chr(Asc("w")) & Chr(Asc(" ")) & Chr(Asc("c")) & Chr(Asc("o")) & Chr(Asc("p")) & Chr(Asc("y")) & Chr(Asc(" ")) & Chr(Asc("o")) & Chr(Asc("f")) & Chr(Asc(" ")) & Chr(Asc("S")) & Chr(Asc("t")) & Chr(Asc("e")) & Chr(Asc("a")) & Chr(Asc("l")) & Chr(Asc("t")) & Chr(Asc("h")) & Chr(Asc("B")) & Chr(Asc("o")) & Chr(Asc("t")) & Chr(Asc(" ")) & Chr(Asc("a")) & Chr(Asc("t")) & Chr(Asc(" ")) & Chr(Asc("h")) & Chr(Asc("t")) & Chr(Asc("t")) & Chr(Asc("p")) & Chr(Asc(":")) & Chr(Asc("/")) & Chr(Asc("/")) & Chr(Asc("w")) & Chr(Asc("w")) & Chr(Asc("w")) & Chr(Asc(".")) & Chr(Asc("s")) & Chr(Asc("t")) & Chr(Asc("e")) & Chr(Asc("a")) & Chr(Asc("l")) & Chr(Asc("t")) & Chr(Asc("h")) & Chr(Asc("b")) & Chr(Asc("o")) & Chr(Asc("t")) & Chr(Asc(".")) & Chr(Asc("n")) & Chr(Asc("e")) & Chr(Asc("t")) & Chr(Asc(".")) & Chr(Asc(" ")) & Chr(Asc("A")) & Chr(Asc("d")) & Chr(Asc("d")) & Chr(Asc("i")) & Chr(Asc("t")) & Chr(Asc("i")) & Chr(Asc("o")) & Chr(Asc("n")) & Chr(Asc("a")) & Chr(Asc("l")) & Chr(Asc("l")) & Chr(Asc("y")) & Chr(Asc(",")) & Chr(Asc(" ")) & Chr(Asc("p")) & Chr(Asc("l")) & Chr(Asc("e")) & Chr(Asc("a")) & Chr(Asc("s")) & Chr(Asc("e")) & Chr(Asc(" ")) & Chr(Asc("r")) & Chr(Asc("e")) & Chr(Asc("p")) & Chr(Asc("o")) & Chr(Asc("r")) & Chr(Asc("t")) & Chr(Asc(" ")) & Chr(Asc("t")) & Chr(Asc("h")) & Chr(Asc("e")) & Chr(Asc(" ")) & Chr(Asc("w")) & Chr(Asc("e")) & Chr(Asc("b")) & Chr(Asc("s")) & Chr(Asc("i")) & Chr(Asc("t")) & Chr(Asc("e")) & Chr(Asc(" ")) & Chr(Asc("a")) & Chr(Asc("t")) & Chr(Asc(" ")) & Chr(Asc("w")) & Chr(Asc("h")) & Chr(Asc("i")) & Chr(Asc("c")) & Chr(Asc("h")) & Chr(Asc(" ")) & Chr(Asc("y")) & Chr(Asc("o")) & Chr(Asc("u")) & Chr(Asc(" ")) & Chr(Asc("d")) & Chr(Asc("o")) & Chr(Asc("w")) & Chr(Asc("n")) & Chr(Asc("l")) & Chr(Asc("o")) & Chr(Asc("a")) & Chr(Asc("d")) & Chr(Asc("e")) & Chr(Asc("d")) & Chr(Asc(" ")) & Chr(Asc("S")) & Chr(Asc("t")) & Chr(Asc("e")) & Chr(Asc("a")) & Chr(Asc("l")) & Chr(Asc("t")) & Chr(Asc("h")) & Chr(Asc("B")) & Chr(Asc("o")) & Chr(Asc("t")) & Chr(Asc(" ")) & Chr(Asc("i")) & Chr(Asc("n")) & Chr(Asc(" ")) & Chr(Asc("a")) & Chr(Asc("n")) & Chr(Asc(" ")) & Chr(Asc("e")) & Chr(Asc("-")) & Chr(Asc("m")) & Chr(Asc("a")) & Chr(Asc("i")) & Chr(Asc("l")) & Chr(Asc(" ")) & Chr(Asc("t")) & Chr(Asc("o")) & Chr(Asc(" ")) & Chr(Asc("a")) & Chr(Asc("b")) & Chr(Asc("u")) & Chr(Asc("s")) & Chr(Asc("e")) & Chr(Asc("@")) & Chr(Asc("s")) & Chr(Asc("t")) & Chr(Asc("e")) & Chr(Asc("a")) & Chr(Asc("l")) & Chr(Asc("t")) & Chr(Asc("h")) & Chr(Asc("b")) & Chr(Asc("o")) & Chr(Asc("t")) & Chr(Asc(".")) & Chr(Asc("n")) & Chr(Asc("e")) & Chr(Asc("t")) & Chr(Asc(".")) & Chr(Asc(" ")) & Chr(Asc("T")) & Chr(Asc("h")) & Chr(Asc("a")) & Chr(Asc("n")) & Chr(Asc("k")) & Chr(Asc("s")) & Chr(Asc("!"))
	End Function
	
	Sub DeconstructSettings()
		If Not (SettingsForm Is Nothing) Then
			'UPGRADE_NOTE: Object SettingsForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			SettingsForm = Nothing
		End If
	End Sub
	
	'SHOW/HIDE STUFF
	Public Sub cmdShowHide_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShowHide.Click
		rtbWhispersVisible = (StrComp(cmdShowHide.Text, CAP_HIDE))
		rtbWhispers.Visible = rtbWhispersVisible
		Config.WhisperWindows = CBool(rtbWhispers.Visible)
		Call Config.Save()
		
		If Me.WindowState <> System.Windows.Forms.FormWindowState.Maximized And Me.WindowState <> System.Windows.Forms.FormWindowState.Minimized Then
			If rtbWhispersVisible Then
				Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) + VB6.PixelsToTwipsY(rtbWhispers.Height) - VB6.TwipsPerPixelY)
			Else
				Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(rtbWhispers.Height) + VB6.TwipsPerPixelY)
			End If
		End If
		
		Call frmChat_Resize(Me, New System.EventArgs())
	End Sub
	
	'// to be called on every successful login
	Sub InitListviewTabs()
		Dim toSet As Boolean
		
		If IsW3() Then
			If Clan.isUsed Then
				toSet = True
			Else
				toSet = False
			End If
		Else
			toSet = False
		End If
		
		ListviewTabs.TabPages.Item(LVW_BUTTON_CLAN).Enabled = toSet
		ListviewTabs.TabPages.Item(LVW_BUTTON_FRIENDS).Enabled = Config.FriendsListTab
	End Sub
	
	'// to be called at disconnect time
	Sub DisableListviewTabs()
		ListviewTabs.TabPages.Item(LVW_BUTTON_FRIENDS).Enabled = False
		ListviewTabs.TabPages.Item(LVW_BUTTON_CLAN).Enabled = False
	End Sub
	
	'UPGRADE_NOTE: Name was upgraded to Name_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Sub AddClanMember(ByVal Name_Renamed As String, ByRef Rank As Short, ByRef Online As Short)
		On Error GoTo ERROR_HANDLER
		Dim visible_rank As Short
		
		visible_rank = Rank
		
		If (visible_rank = 0) Then visible_rank = 1
		If (visible_rank > 4) Then visible_rank = 5 '// handle bad ranks
		
		'// add user
		
		Name_Renamed = KillNull(Name_Renamed)
		
		If (Not Online = 0) Then Online = 1
		
		With lvClanList
			'UPGRADE_WARNING: Lower bound of collection lvClanList.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_WARNING: Lower bound of collection lvClanList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
			.Items.Insert(.Items.Count + 1, Name_Renamed, visible_rank)
			If (BotVars.NoColoring = False) Then
				If (StrComp(StripAccountNumber(GetCurrentUsername), Name_Renamed) = 0) Then
					'UPGRADE_WARNING: Lower bound of collection lvClanList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					.Items.Item(.Items.Count).ForeColor = System.Drawing.ColorTranslator.FromOle(FormColors.ChannelListSelf)
				End If
			End If
			'UPGRADE_WARNING: Lower bound of collection lvClanList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
            .Items.Item(.Items.Count).SubItems.Add(Online + 6)
			'UPGRADE_WARNING: Lower bound of collection lvClanList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			.Items.Item(.Items.Count).SubItems.Add(CStr(visible_rank))
            .ListViewItemSorter = New modOtherCode.ListViewComparer(2)
			.Sorting = System.Windows.Forms.SortOrder.Descending
			.Sort()
		End With
		
		lblCurrentChannel.Text = GetChannelString()
		
        'Me.ListviewTabs_Click(0)
		
		RunInAll("Event_ClanInfo", Name_Renamed, Rank, Online)
		Exit Sub
ERROR_HANDLER: 
		AddChat(RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in frmChat.AddClanMember", Err.Number, Err.Description))
	End Sub
	
	Private Function GetClanSelectedUser() As String
		With lvClanList
			If Not (.FocusedItem Is Nothing) Then
				If .FocusedItem.Index < 1 Then
					GetClanSelectedUser = vbNullString : Exit Function
				Else
					GetClanSelectedUser = CleanUsername(ReverseConvertUsernameGateway(.FocusedItem.Text))
				End If
			End If
		End With
	End Function
	
	Private Sub lvClanList_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lvClanList.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim aInx As Short
		Dim bIsOn As Boolean
		Dim MyRank As Integer
		Dim TheirRank As Integer
		Dim CanMoveUp As Boolean
		Dim CanMoveDown As Boolean
		Dim CanRemove As Boolean
		Dim CanMakeChief As Boolean
		Dim CanDisband As Boolean
		Dim IsSelf As Boolean
		Dim CanLeave As Boolean
		
		If (lvClanList.FocusedItem Is Nothing) Then
			Exit Sub
		End If
		
		If Button = VB6.MouseButtonConstants.RightButton Then
			If Not (lvClanList.FocusedItem Is Nothing) Then
				aInx = lvClanList.FocusedItem.Index
				
				If aInx > 0 Then
					MyRank = g_Clan.Self.Rank
					'UPGRADE_WARNING: Couldn't resolve default property of object g_Clan.GetUser().Rank. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TheirRank = g_Clan.GetUser(GetClanSelectedUser).Rank
					IsSelf = StrComp(g_Clan.Self.Name, GetClanSelectedUser, CompareMethod.Binary) = 0
					
					CanMoveUp = Not IsSelf And (MyRank > (TheirRank + 1)) And (TheirRank > 0)
					CanMoveDown = Not IsSelf And (MyRank > TheirRank) And (TheirRank > 1)
					CanRemove = Not IsSelf And (CanMoveUp Or CanMoveDown)
					CanMakeChief = Not IsSelf And (MyRank = 4) And (TheirRank > 1)
					CanLeave = IsSelf And (MyRank > 0 And MyRank < 4)
					CanDisband = IsSelf And (MyRank = 4)
					
					'UPGRADE_WARNING: Couldn't resolve default property of object g_Clan.GetUser().IsOnline. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					bIsOn = g_Clan.GetUser(GetClanSelectedUser).IsOnline
					
					mnuPopClanWhisper.Enabled = bIsOn
					
					mnuPopClanPromote.Enabled = CanMoveUp
					mnuPopClanDemote.Enabled = CanMoveDown
					mnuPopClanRemove.Enabled = CanRemove
					mnuPopClanLeave.Enabled = CanLeave
					mnuPopClanDisband.Enabled = CanDisband
					mnuPopClanMakeChief.Enabled = CanMakeChief
					
					mnuPopClanPromote.Visible = Not IsSelf
					mnuPopClanDemote.Visible = Not IsSelf
					mnuPopClanRemove.Visible = Not IsSelf
					mnuPopClanLeave.Visible = IsSelf And (MyRank <> 4)
					mnuPopClanDisband.Visible = IsSelf And (MyRank = 4)
					mnuPopClanMakeChief.Visible = Not IsSelf And (MyRank = 4)
				End If
			End If
			
			mnuPopClanList.Tag = lvClanList.FocusedItem.Text 'Record which user is selected at time of right-clicking. - FrOzeN
			
            mnuPopClanList.ShowDropDown()
		End If
	End Sub
	
	Sub DoConnect()
		
		'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If ((sckBNLS.CtlState <> MSWinsockLib.StateConstants.sckClosed) Or (sckBNet.CtlState <> MSWinsockLib.StateConstants.sckClosed)) Then
			Call DoDisconnect()
		End If
		
		uTicks = 0
		
		UserCancelledConnect = False
		
		'Reset the BNLS auto-locator list
		BNLSFinderGotList = False
		
		'If Not IsValidIPAddress(BotVars.Server) And BotVars.UseProxy Then
		'AddChat RTBColors.ErrorMessageText, "[PROXY] Proxied connections must use a direct server IP address, such as those listed below your desired gateway in the Connection Settings menu, to connect."
		'AddChat RTBColors.ErrorMessageText, "[PROXY] Please change servers and try connecting again."
		'Else
		Call Connect()
		'End If
	End Sub
	
	Sub DoDisconnect(Optional ByVal DoNotShow As Byte = 0, Optional ByVal LeaveUCCAlone As Boolean = False)
		On Error GoTo ERROR_HANDLER
		
		Dim i As Short
		
		If (Not (UserCancelledConnect)) Then
			tmrAccountLock.Enabled = False
			
			SetTitle("Disconnected")
			
			Me.UpdateTrayTooltip()
			
			Call CloseAllConnections(DoNotShow = 0)
			
			'UPGRADE_NOTE: Object g_Channel may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_Channel = Nothing
			'UPGRADE_NOTE: Object g_Clan may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_Clan = Nothing
			'UPGRADE_NOTE: Object g_Friends may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			g_Friends = Nothing
			
			BotVars.Gateway = vbNullString
			
			CurrentUsername = vbNullString
			
			ListviewTabs.SelectedIndex = 0
			
			Call g_Queue.Clear()
			
			If Not LeaveUCCAlone Then
				UserCancelledConnect = True
			End If
			
			If (UserCancelledConnect) Then
				'AddChat vbRed, "DISC!"
				
				If ReconnectTimerID > 0 Then
					KillTimer(0, ReconnectTimerID)
					ReconnectTimerID = 0
				End If
				
				If ExReconnectTimerID > 0 Then
					KillTimer(0, ExReconnectTimerID)
					ExReconnectTimerID = 0
				End If
				
				If SCReloadTimerID > 0 Then
					KillTimer(0, SCReloadTimerID)
					SCReloadTimerID = 0
				End If
			Else
				'ReconnectTimerID = SetTimer(0, 0, BotVars.ReconnectDelay, _
				''    AddressOf Reconnect_TimerProc)
				'
				'ExReconnectTimerID = SetTimer(0, ExReconnectTimerID, _
				''    BotVars.ReconnectDelay, AddressOf ExtendedReconnect_TimerProc)
			End If
			
			DisableListviewTabs()
			
			BotVars.ProxyStatus = modEnum.enuProxyStatus.psNotConnected
			
			Clan.isUsed = False
			lvClanList.Items.Clear()
			
			BNLSBuffer.ClearBuffer()
			BNCSBuffer.ClearBuffer()
			MCPBuffer.ClearBuffer()
			
			g_Connected = False
			g_Online = False
			ds.EnteredChatFirstTime = False
			ds.ClientToken = 0
			
			Call ClearChannel()
			lvClanList.Items.Clear()
			lvFriendList.Items.Clear()
			
			'tmrSilentChannel(0).Enabled = False
			
			Call g_Queue.Clear()
			
			BNLSAuthorized = False
			uTicks = 0
			
			mnuSepZ.Visible = False
			mnuIgnoreInvites.Visible = False
			mnuRealmSwitch.Visible = False
			
			BotVars.LastChannel = vbNullString
			PrepareHomeChannelMenu()
			PrepareQuickChannelMenu()
			
			'UPGRADE_NOTE: Object BotVars.PublicChannels may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			BotVars.PublicChannels = Nothing
			PreparePublicChannelMenu()
			
			If ((Me.WindowState = System.Windows.Forms.FormWindowState.Normal) And (DoNotShow = 0)) Then
				
				'This SetFocus() call causes an error if any script have InputBoxes open.
				'This is the best fix I could come up with. :( -Pyro
				On Error Resume Next
				Call cboSend.Focus()
				On Error GoTo ERROR_HANDLER
			End If
			
			' clean up realms
			If Not ds.MCPHandler Is Nothing Then
				If ds.MCPHandler.FormActive Then
					frmRealm.UnloadAfterBNCSClose()
				End If
				
				'UPGRADE_NOTE: Object ds.MCPHandler may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				ds.MCPHandler = Nothing
			End If
			
			' clean up email reg
			frmEMailReg.Close()
			
			' close any pending INet
			INet.Tag = SB_INET_UNSET
			'UPGRADE_ISSUE: VBControlExtender property INet.Cancel was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			INet.Cancel()
			
			' reset BNLS finder
			BNLSFinderGotList = False
			BNLSFinderIndex = 0
			
			PassedClanMotdCheck = False
			
			On Error Resume Next
			RunInAll("Event_LoggedOff")
		End If
		
		Exit Sub
		
ERROR_HANDLER: 
		AddChat(RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in DoDisconnect().")
		
		Exit Sub
	End Sub
	
    Public Sub ParseFriendsPacket(ByVal PacketID As Integer, ByVal Data() As Byte)
        FriendListHandler.ParsePacket(PacketID, Data)
    End Sub
	
    Public Sub ParseClanPacket(ByVal PacketID As Integer, ByVal Data() As Byte)
        ClanHandler.ParseClanPacket(PacketID, Data)
    End Sub
	
	Public Sub RecordWindowPosition(Optional ByRef Maximized As Boolean = False)
		'Don't record other position information if maximized, otherwise when they unmaximize it will be fullscreen width and height. - FrOzeN
		If Not Maximized Then
			Config.PositionLeft = Int(VB6.PixelsToTwipsX(Me.Left) / VB6.TwipsPerPixelX)
			Config.PositionTop = Int(VB6.PixelsToTwipsY(Me.Top) / VB6.TwipsPerPixelY)
			Config.PositionHeight = Int(VB6.PixelsToTwipsY(Me.Height) / VB6.TwipsPerPixelY)
			Config.PositionWidth = Int(VB6.PixelsToTwipsX(Me.Width) / VB6.TwipsPerPixelX)
		End If
		
		Config.IsMaximized = Maximized
		Call Config.Save()
	End Sub
	
	Public Sub MakeLoggingDirectory()
		On Error Resume Next
		MkDir(GetFolderPath("Logs"))
	End Sub
	
	' Called from several points to keep accurate tabs on the user's prior selection
	'  in the send combo
	Public Sub RecordcboSendSelInfo()
		'Debug.Print "SelStart: " & cboSend.SelStart & ", SelLength: " & cboSend.SelLength
		cboSendSelLength = cboSend.SelectionLength
		cboSendSelStart = cboSend.SelectionStart
	End Sub
End Class