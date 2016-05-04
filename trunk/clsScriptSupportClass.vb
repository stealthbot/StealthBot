Option Strict Off
Option Explicit On
Friend Class clsScriptSupportClass
	'/*
	' * StealthBot Shared Scripting Class
	' *
	' * clsSharedScriptSupport.cls
	' *
	' *
	' * This class basically mirrors the signatures of several important StealthBot functions
	' * for the purpose of allowing your scripts to interact with the rest of the program.
	' *
	' * I'm very accomodating to the SB scripting community. If you have something you want
	' * to see mirrored or some change you want made, don't hesitate to e-mail me about it
	' * at stealth@stealthbot.net or bring it up on our scripting forums at stealthbot.net.
	' *
	' * SCRIPTING SYSTEM CHANGELOG
	' *     (version 2.7)
	' *     - Ping() function renamed PingByName() to avoid variable name conflicts (thanks raylu)
	' *     - Added GetQueueSize function (thanks 111787)
	' *     - Added Event_MessageSent() event (thanks Snap)
	' *     - Added Event_ClanInfo() event (thanks Jack)
	' *         Sub Event_ClanInfo(Name, Rank, Online)
	' *         Called once for each member of the clan - use it to fill a list of clan members
	' *     - Fixed the method called in MonitoredUserIsOnline() (thanks Snap)
	' *     - Added GetLastMonitorWhois() function (thanks Snap)
	' *     - Added GetMonitorUserData() function (thanks Snap)
	' *     - New argument in Event_UserJoins() event, thanks to Z1g0rro
	' *         Banned will contain a boolean (TRUE = banned by the bot, FALSE = normal user)
	' *         Add this argument at the end of the event signature
	' *     - The message "All connections closed." will raise a ServerError message (thanks Jack)
	' *     - Added ReloadScript() function (thanks various)
	' *     - Added AddChatFont() function (thanks Imhotep[Nu])
	' *     - Fixed a bug with GetInternalData() (thanks Jack)
	' *     - Added FlashBotWindow() function (thanks LuC1Fr)
	' *     - Fixed GetPositionByName() description (thanks J3m)
	' *     - Added SetSCTimeout() function (thanks WoD[ActionD])
	' *     - Added GetScriptControl() function (thanks HdxBmx)
	' *     - Added CommandEx() function documented below (thanks Imhotep[Nu])
	' *     - Removed the Sleep() function (thanks Draco)
	' *     - 3() should now properly report the bot's idle time (thanks Konohamaru)
	' *     - Fixed GetUserProfile() -- keys should now return properly to you one at a time in
	' *         Event_KeyReturn(KeyName, KeyValue) (thanks Jack, Sinaps)
	' *     - Added WhisperCmds to the BotVars object so you can toggle command-whisperedness
	' *         in scripting (thanks Jack)
	' *     - %me now works in calls to Command() (thanks Jack)
	' *     - Added Event_FirstRun() which will execute only the first time the bot starts up
	' *         and not on subsequent script control reloads (thanks Swent)
	' *     - Signature change to Event_UserInChannel(): The new signature is
	' *         Sub Event_UserInChannel(Username, Flags, Message, Ping, Product, StatUpdate)
	' *          StatUpdate is a boolean that tells you whether or not the person is
	' *          already in the channel and is merely having their information updated.
	' *     - Added the following clan-related events: (thanks raylu)
	' *         Event_ClanMemberList(Username, Rank, Online)
	' *         Event_ClanMemberUpdate(Username, Rank, Online)
	' *         Event_ClanMOTD(Message)
	' *         Event_ClanMemberLeaves(Username)
	' *         Event_BotRemovedFromClan()
	' *         Event_BotClanRankChanged(NewRank)
	' *         Event_BotJoinedClan(ClanTag)
	' *         Event_BotClanInfo(ClanTag, Rank)
	' *     - Added an Event_Shutdown() that executes only when the bot is actually closing
	' *         and not on script reloads (thanks Swent)
	' *     - Added a ssc.ClearScreen() command (thanks Imhotep[Nu])
	' *     - Added the C_Dec() function (thanks Imhotep[Nu])
	' *     - Added the DeleteURLCache() mirror function (thanks Jack)
	' *     - PadQueue() now inserts a blank queue message to add a delay before
	' *         the queue's next message goes out. The old PadQueue is still
	' *         present but has been more-accurately renamed to PadQueueCounter
	' *          (thanks Snap)
	' *     - SetBotProfile() no longer allows you to edit the Sex field (this is a
	' *         Blizzard change)
	' *     - Added GetApphInstance() which gives you the instance handle to StealthBot (thanks FrostWraith)
	' *     - Added DoStatstringParse() which allows you to parse a statstring from GetInternalData() out
	' *         just like the bot does (thanks ZergMasterI)
	' *     - Added a number of Windows API function mirrors (thanks FiftyToo)
	' *     - Added GetBotVersionNumber() function which returns only the numerical value of the bot version
	' *     - VetoThisMessage() Now works in PacketSent, And PacketReceived, Will prevent it from being sent,
	' *         or parsed, respectivly.
	' *     - Added GetCommands(Option FilePath) function, Will return a collection of CommandDocObjs for each <command> element in the passed file
	' *     - Added GetCurrentUsername() function, Will return the Bot's current username, as Battle.net sees it.
	' *     - Added ObserveScript() function, Will duplicate script specific events for the specified script, to yours
	' *     - Added GetObserved() function, Will return a collection of script names you are observing
	' *     - Added GetObservers() function, Will return a collection of script names that are obsering your script
	' *
	' *     (version 2.6R3, scripting system build 21)
	' *     - GetNameByPosition() boundary checks fixed (thanks Scio)
	' *
	' *     (version 2.6, scripting system build 20)
	' *     - Exposed the entire internal bot variable class to the scripting system
	' *         clsBotVars.txt shows you what you can access, BotVars.varName
	' *         Suggested by Imhotep[Nu]
	' *     - Fixed the MonitoredUserIsOnline() function (thanks Cnegurozka)
	' *     - Added BotPath function (thanks werehamster)
	' *     - Added IsOnline function (thanks Xelloss)
	' *     - Added Sleep function (thanks Imhotep[Nu])
	' *     - Added GetPositionByName function (thanks werehamster)
	' *     - Added GetNameByPosition function (thanks werehamster)
	' *     - Added GetBotVersion function (thanks Imhotep[Nu])
	' *     - Changed the ReloadSettings function (thanks Imhotep[Nu])
	' *     - New scripting event: Event_LoggedOff() (thanks Imhotep[Nu])
	' *     - Added Connect() and Disconnect() functions (thanks Imhotep[Nu])
	' *     - Added BotClose() function (thanks Imhotep[Nu])
	' *     - Changed GetInternalData() function and added GetInternalDataByUsername() function
	' *     - Added GetInternalUserCount() function
	' *     - Added Event_ChannelLeave() function (request of Imhotep[Nu])
	' *     - Added GetConfigEntry() and WriteConfigEntry() functions
	' *     - Added PrintURLToFile() function (thanks SoCxFiftyToo)
	' *     - Added VetoThisMessage() function -- use in Event_PressedEnter to prevent the
	' *         message in the event's arguments from being sent to Battle.net
	' *
	' *     (version 2.5, scripting system build 19)
	' *     - Command() now returns the command response string (requested)
	' *     - New SSC function GetInternalData(sUser, lDataType) - see the function in this file for details
	' *     - New SSC function IsShitlisted()
	' *     - New SSC function PadQueue() added
	' *
	' *     (version 2.4R2, scripting system build 18)
	' *     - AddChat now loops from 0 to ubound, which is correct. (Thanks Imhotep[Nu])
	' *     - Added the ReloadSettings function
	' *     - The Event_Close() sub is called when a user reloads the config
	' *     - Fixed the signature for Event_KeyPress in script.txt (should be Event_PressedEnter)
	' *     - Added Event_UserInChannel()
	' *     - Clarified how #include works in script.txt
	' *
	' *     (version 2.4, scripting system build 17)
	' *     - The Event_ChannelJoin scripted subroutine is now usable (thanks -)(nsane-)
	' *     - Exposed a MSINET control to the Scripting system, for use in script-to-website
	' *         communication
	' *     - Added the #include keyword for script files -- more information is in script.txt
	' *     - Added the MonitoredUserIsOnline() function
	' *     - The Level variable is now properly passed to the script control
	' *     - Added scripting events:
	' *         > ServerError messages
	' *         > PressedEnter
	' *     - The script control class now has access to GetTickCount and Beep API calls
	' *         and has been improved based on user requests
	' *     - Added the Match, DoReplace, DoBeep and GetGTC functions
	' *     - Added the myChannel, BotFlags and myUsername publicly accessible variables
	' *     - Added the _KeyReturn() event, for processing profile keys returned from the server
	' *     - Added the RequestUserProfile() method, for requesting any user's profile
	' *     - Added the SetBotProfile() method, for setting the bot's current profile
	' *     - Added the Event_Close() event, which executes on Form_Unload()
	' *     - Event_Load() is now called when you reload the script.txt file
	' *     - Added the OriginalStatstring variable to Event_Join(). It contains the unparsed statstring of the joining user
	' */
	
	
	Public MyChannel As String '// will contain the bot's current channel at runtime.
	Public BotFlags As Integer '// will contain the bot's current battle.net flags at runtime.
	Public myUsername As String '// will contain the bot's current username at runtime.
	'// NOTE: This may be different than the bot's config.ini username
	
	'// myTrigger has been replaced by BotVars.Trigger
	'Public myTrigger As String '// contains the bot's current trigger at runtime
	
	Public Enum BanTypes
		btBan = 0 '// used in calling the BanKickUnban() subroutine
		btKick = 1
		btUnban = 2
	End Enum
	
	'This is the name to use if you want to refer to the Internal 'Script' of the bot.
	'Right now only 'Event_Command' is run in the bot.
	'So ObserveScript(SSC.InternalScript) would cause Event_Command to be fired for you for all Built-In commands.
	Public ReadOnly Property InternalScript() As String
		Get
			InternalScript = vbNullString
		End Get
	End Property
	
	'/* ******************************************************************************************
	' *
	' *
	' *
	' *
	' * INTERNAL BOT "MIRROR" FUNCTIONS
	' *         Usage: ssc.function(arguments)
	' *         Example:    ssc.AddChat vbBlue, "Hello world!"
	' *
	' *
	' *
	' *
	' * ******************************************************************************************/
	
	'// ADDCHAT
	'// Grok's famous AddChat subroutine. Processes Starcraft/Diablo II color codes automatically.
	'// Format: AddChat(Color, Text)
	'// Extensible as far as you need:
	'// AddChat(Color, Text, Color, Text, Color, Text) -- will all display on one line.
	'// For example:
	'//     AddChat vbRed, "Hello, world!"
	'// will display that phrase in red.
	'UPGRADE_WARNING: ParamArray saElements was changed from ByRef to ByVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93C6A0DC-8C99-429A-8696-35FC4DCEFCCC"'
	Public Sub AddChat(ParamArray ByVal saElements() As Object)
		Dim arr() As Object
		
		arr = VB6.CopyArray(saElements)
		
		Call DisplayRichText((frmChat.rtbChat), arr)
	End Sub
	
	
	'// ADDQ (ADD QUEUE)
	'// Adds a string to the message send queue.
	'// Nonzero priority messages will be sent with precedence over 0-priority messages.
	Public Function AddQ(ByVal sText As String, Optional ByVal msg_priority As Short = -1, Optional ByVal Tag As String = vbNullString) As Short
		
		AddQ = frmChat.AddQ(sText, msg_priority, g_lastQueueUser, Tag)
	End Function
	
	
	'// COMMAND
	'// Calls StealthBot's command processor
	'// Messages passed to the processor will be evaluated as commands
	'// Public Function Commands(ByRef dbAccess As udtGetAccessResponse, Username As String, _
	''//     Message As String, InBot As Boolean, Optional CC As Byte = 0) As String
	
	'// RETURNS: The response string, or an empty string if there is no response
	
	'// Detailed description of each variable:
	'// dbAccess    is assembled below. It consists of the speaker username's access within the bot.
	'//             For scripting purposes, this module will assemble dbAccess by calling GetAccess() on
	'//             the username you specify.
	'// Username    is the speaker's username (the username of the person using the command.)
	'// Message     is the raw command message from Battle.net. If the user says ".say test", the raw
	'//             command message is ".say test". This is the method by which you should call the commands.
	'// InBot       defines whether or not the command has been issued from inside the bot. If it has,
	'//             the trigger is temporarily changed to "/". Basically, for scripting purposes you can
	'//             use it to control whether or not your command responses display publicly.
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Command_Renamed(ByVal Username As String, ByVal Message As String, Optional ByVal IsLocal As Boolean = False, Optional ByVal Whispered As Boolean = False)
		
		' execute command
		Call ProcessCommand(Username, Message, IsLocal, Whispered)
	End Sub
	
	
	
	'// PINGBYNAME
	'// Returns the cached ping of the specified user.
	'// If the user is not present in the channel, it returns -3.
	Public Function PingByName(ByVal Username As String) As Integer
		PingByName = GetPing(Username)
	End Function
	
	
	
	'// BANKICKUNBAN
	'// Returns a string corresponding to the success or failure of a ban attempt.
	'// Responses should be directly queued using AddQ().
	'// Example response strings:
	'//     That user is safelisted.
	'//     The bot does not have ops.
	'//     /ban thePerson Your mother!
	
	'// Variable descriptions:
	'// INPT    - Contains the username of the person followed by any extension to it, such as ban message
	'//         - Examples: "thePerson Your Mother has a very extremely unequivocally long ban message!"
	'//         -           "thePerson"
	'//         -           "thePerson Short ban message
	
	'// SPEAKERACCESS   contains the access of the person attempting to ban/kick. This is not applied in
	'//                 unban situations.
	'//                 In Kick and Ban situations, the target's access must be less than or equal to
	'//                 this value -- use it to control inherent safelisting (ie all users with > 20 access
	'//                 are not affected by it)
	
	'// MODE        contains the purpose of the subroutine call. The same routine is used to ban, kick and
	'//             unban users, so make that choice when calling it.
	'//             Ban = 0; Kick = 1; Unban = 2. Any other value will cause the function to die a horrible
	'//             death. (not really, it just won't do anything.)
	
	Public Function BanKickUnban(ByVal Inpt As String, ByVal SpeakerAccess As Short, Optional ByVal Mode As BanTypes = 0) As String
		
		BanKickUnban = Ban(Inpt, SpeakerAccess, CByte(Mode))
	End Function
	
	
	
	'// ISSAFELISTED
	'// Returns True if the user is safelisted, False if they're not. Pretty simple.
	Public Function isSafelisted(ByVal Username As String) As Boolean
		isSafelisted = GetSafelist(Username)
	End Function
	
	
	
	'// ISSHITLISTED
	'// Returns a null string if the user is not shitlisted, otherwise returns the shitlist message.
	Public Function isShitlisted(ByVal Username As String) As String
		isShitlisted = GetShitlist(Username)
	End Function
	
	
	
	'// GETDBENTRY
	'// Bit of a modification to my existing GetAccess() call to return the data to you effectively.
	'// The scripting control isn't the greatest.
	'// Pass it the username and it will pass you the person's access and flags.
	'// If the name is not in the database, it will return -1 / null flags.
	Public Function GetDBEntry(ByVal Username As String, Optional ByRef Access As Object = Nothing, Optional ByRef Flags As Object = Nothing, Optional ByRef EntryType As String = "USER") As Object
		
		Dim temp As udtGetAccessResponse
		Dim i As Short
		Dim Splt() As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		temp = GetCumulativeAccess(Username, EntryType)
		
		With temp
			'UPGRADE_WARNING: Couldn't resolve default property of object Access. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Access = .Rank
			'UPGRADE_WARNING: Couldn't resolve default property of object Flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Flags = .Flags
		End With
		
		GetDBEntry = New clsDBEntryObj
		
		With GetDBEntry
			'UPGRADE_WARNING: Couldn't resolve default property of object GetDBEntry.EntryType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.EntryType = temp.Type
			'UPGRADE_WARNING: Couldn't resolve default property of object GetDBEntry.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Name = temp.Username
			'UPGRADE_WARNING: Couldn't resolve default property of object GetDBEntry.Rank. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Rank = temp.Rank
			'UPGRADE_WARNING: Couldn't resolve default property of object GetDBEntry.Flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Flags = temp.Flags
			'UPGRADE_WARNING: Couldn't resolve default property of object GetDBEntry.CreatedOn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.CreatedOn = temp.AddedOn
			'UPGRADE_WARNING: Couldn't resolve default property of object GetDBEntry.CreatedBy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.CreatedBy = temp.AddedBy
			'UPGRADE_WARNING: Couldn't resolve default property of object GetDBEntry.ModifiedOn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.ModifiedOn = temp.ModifiedOn
			'UPGRADE_WARNING: Couldn't resolve default property of object GetDBEntry.ModifiedBy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.ModifiedBy = temp.ModifiedBy
			'UPGRADE_WARNING: Couldn't resolve default property of object GetDBEntry.BanMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.BanMessage = temp.BanMessage
		End With
		
		If ((temp.Groups <> vbNullString) And (temp.Groups <> "%")) Then
			If (InStr(1, temp.Groups, ",", CompareMethod.Binary) <> 0) Then
				Splt = Split(temp.Groups, ",")
			Else
				ReDim Preserve Splt(0)
				
				Splt(0) = temp.Groups
			End If
			
			For i = LBound(Splt) To UBound(Splt)
				'UPGRADE_WARNING: Couldn't resolve default property of object GetDBEntry.Groups. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetDBEntry.Groups.Add(GetDBEntry(Splt(i),  ,  , "GROUP"))
			Next i
		End If
	End Function
	
	'// GETSTDDBENTRY
	'// Bit of a modification to my existing GetCumulativeAccess() call to return the data to you
	'// effectively.
	'// The difference between GetDBEntry and GetCumulativeDBEntry is that this function will
	'// return the cumulative access of a particular user.  This includes both dynamic and static
	'// group memberships and all wildcard matches.  This function should almost always be used over
	'// GetDBEntry().
	'// The scripting control isn't the greatest.
	'// Pass it the username and it will pass you the person's access and flags.
	'// If the name is not in the database, it will return -1 / null flags.
	Public Sub GetStdDBEntry(ByVal Username As String, ByRef Access As Object, ByRef Flags As Object, Optional ByRef EntryType As String = "USER") '// yum, Variants >:\
		
		Dim temp As udtGetAccessResponse
		
		'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		temp = GetAccess(Username, EntryType)
		
		With temp
			'UPGRADE_WARNING: Couldn't resolve default property of object Access. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Access = .Rank
			'UPGRADE_WARNING: Couldn't resolve default property of object Flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Flags = .Flags
		End With
	End Sub
	
	
	'// PREPARELIKECHECK
	'// Prepares a string for comparison using the Visual Basic LIKE operator
	'// Originally written by Zorm, since expanded
	Public Function PrepareLikeCheck(ByVal sText As String) As String
		PrepareLikeCheck = PrepareCheck(sText)
	End Function
	
	
	'// GETGTC
	'// Returns the current system uptime in milliseconds as reported by the GetTickCount() API call
	Public Function GetGTC() As Integer
		GetGTC = GetTickCount()
	End Function
	
	
	
	'// DOBEEP
	'// Executes a call to the Beep() API function
	Public Function DoBeep(ByVal lFreq As Integer, ByVal lDuration As Integer) As Integer
		DoBeep = Beep_Renamed(lFreq, lDuration)
	End Function
	
	
	
	'// MATCH
	'// Allows VBScripters to use the Like comparison operator in VB
	'// Specify TRUE to the third argument (DoPreparation) to automatically prepare both inbound strings
	'//     for compatibility with Like
	Public Function match(ByVal sString As String, ByVal sPattern As String, ByVal DoPreparation As Boolean) As Boolean
		
		If DoPreparation Then
			sString = PrepareCheck(sString)
			sPattern = PrepareCheck(sPattern)
		End If
		
		match = (sString Like sPattern)
	End Function
	
	
	'// SETBOTPROFILE
	'// Sets the bot's current profile to the specified value(s).
	'// If passed as null, the Values() will not be reset, so profile data you are not changing will not be overwritten.
	'// As of Starcraft version 1.15, Blizzard removed the Sex field from user profiles
	'//     so this data is no longer writeable.
	'// To maintain backwards-compatibility this method's signature will not change,
	'//     but be aware that the sNewSex value will not affect anything.
	Public Sub SetBotProfile(ByVal sNewSex As String, ByVal sNewLocation As String, ByVal sNewDescription As String)
		Call SetProfileEx(sNewLocation, sNewDescription)
	End Sub
	
	
	'// GETUSERPROFILE
	'// Gets the profile of a specified user. The profile is returned in three pieces via the _KeyReturn() event.
	'// If Username is null, the bot's current username will be used instead.
	Public Sub GetUserProfile(Optional ByVal Username As String = "")
		SuppressProfileOutput = True
		
        If Len(Username) > 0 Then
            Call RequestProfile(Username)
        Else
            Call RequestProfile(GetCurrentUsername)
        End If
	End Sub
	
	
	'// RELOADSETTINGS
	'// Reloads the bot's configuration settings, userlist, safelist, tagban list, and script.txt files - equivalent to
	'//     clicking "Reload Config" under the Settings menu inside the bot.
	'// @param DoNotLoadFontSettings - when passed a value of 1 the bot will not
	'//     attempt to alter the main richtextbox font settings, which causes its contents to be erased
	Public Sub ReloadSettings(ByVal DoNotLoadFontSettings As Byte)
		Call frmChat.ReloadConfig(DoNotLoadFontSettings)
	End Sub
	
	
	'// BOTPATH
	'// Returns the bot's current path. Future compatibility with multiple user profiles is already in place.
	'//     Return value includes the trailing "\".
	Public Function BotPath() As String
		BotPath = GetProfilePath()
	End Function
	
	
	'// GETINTERNALUSERCOUNT
	'// Returns the highest index for use when calling GetInternalDataByIndex
	'//     this allows you to call that function with (1 to GetInternalUserCount())
	Public Function GetInternalUserCount() As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Channel.Users. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetInternalUserCount = Channel.Users.Count
	End Function
	
	
	'// GETINTERNALDATABYUSERNAME
	'// Retrieves the specified stored internal data for a given user in the channel
	'//  If the specified user isn't present, the return value is -5
	'//  See lDataType constants in GetInternalData() below
	Public Function GetInternalDataByUsername(ByVal sUser As String, ByVal lDataType As Integer) As Object
		Dim i As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Channel.GetUserIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		i = Channel.GetUserIndex(sUser)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object GetInternalData(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetInternalDataByUsername = GetInternalData(i, lDataType)
	End Function
	
	
	'// GETINTERNALDATA
	'// Retrieves the specified stored internal data for a given user in the channel
	'//  If the specified user is not present, return value is '-5'
	Public Function GetInternalData(ByVal iIndex As Short, ByVal lDataType As Integer) As Object
		' -- '
		'       these constants will be useful in making calls to this function
		'                           |   <purpose>
		Const GID_CLAN As Short = 0 '-> retrieves 4-character clan name
		Const GID_FLAGS As Short = 1 '-> retrieves Battle.net flags
		Const GID_PING As Short = 2 '-> retrieves ping on login
		Const GID_PRODUCT As Short = 3 '-> retrieves 4-digit product code
		Const GID_ISSAFELISTED As Short = 4 '-> retrieves Boolean value denoting safelistedness
		Const GID_STATSTRING As Short = 5 '-> retrieves unparsed statstring
		Const GID_TIMEINCHANNEL As Short = 6 '-> retrieves time in channel in seconds
		Const GID_TIMESINCETALK As Short = 7 '-> retrieves time since the user's last message in seconds
		' -- '
		
		If iIndex > 0 Then
			Select Case lDataType
				Case GID_CLAN
					'UPGRADE_WARNING: Couldn't resolve default property of object Channel.Users. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object GetInternalData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetInternalData = Channel.Users(iIndex).Clan
					
				Case GID_FLAGS
					'UPGRADE_WARNING: Couldn't resolve default property of object Channel.Users. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object GetInternalData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetInternalData = Channel.Users(iIndex).Flags
					
				Case GID_PING
					'UPGRADE_WARNING: Couldn't resolve default property of object Channel.Users. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object GetInternalData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetInternalData = Channel.Users(iIndex).Ping
					
				Case GID_PRODUCT
					'UPGRADE_WARNING: Couldn't resolve default property of object Channel.Users. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object GetInternalData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetInternalData = Channel.Users(iIndex).Game
					
				Case GID_ISSAFELISTED
					'UPGRADE_WARNING: Couldn't resolve default property of object Channel.Users. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetInternalData = GetSafelist(Channel.Users(iIndex).DisplayName)
					
				Case GID_STATSTRING
					'UPGRADE_WARNING: Couldn't resolve default property of object Channel.Users. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object GetInternalData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetInternalData = Channel.Users(iIndex).Statstring
					
				Case GID_TIMEINCHANNEL
					'UPGRADE_WARNING: Couldn't resolve default property of object Channel.Users. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object GetInternalData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetInternalData = Channel.Users(iIndex).TimeInChannel
					
				Case GID_TIMESINCETALK
					'UPGRADE_WARNING: Couldn't resolve default property of object Channel.Users. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object GetInternalData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetInternalData = Channel.Users(iIndex).TimeSinceTalk
					
				Case Else
					'UPGRADE_WARNING: Couldn't resolve default property of object GetInternalData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetInternalData = 0
					
			End Select
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object GetInternalData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetInternalData = -5
		End If
	End Function
	
	
	'// ISONLINE
	'// Returns a boolean denoting teh bot's status ONLINE=TRUE, OFFLINE=FALSE.
	Public Function IsOnline() As Boolean
		IsOnline = g_Online
	End Function
	
	
	'// GETPOSITIONBYNAME
	'// Returns the channel list position of a user by their username
	'// Returns 0 if the user is not present
	Public Function GetPositionByName(ByVal sUser As String) As Short
		GetPositionByName = checkChannel(sUser)
	End Function
	
	
	'// GETNAMEBYPOSITION
	'// Returns the name of the person at position X in the channel list.
	'// Positions are 1-based. Returns an empty string if the user isn't present
	Public Function GetNameByPosition(ByVal iPosition As Short) As String
		With frmChat.lvChannel.Items
			If iPosition > 0 And iPosition <= .Count Then
				'UPGRADE_WARNING: Lower bound of collection frmChat.lvChannel.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				GetNameByPosition = .Item(iPosition).Text
			Else
				GetNameByPosition = vbNullString
			End If
		End With
	End Function
	
	
	'// GETBOTVERSION
	'// Returns the current StealthBot app version as a string.
	Public Function GetBotVersion() As String
		GetBotVersion = CVERSION
	End Function
	
	
	'// GETBOTVERSIONNUMBER
	'// Returns numerical value of the current StealthBot version.
	Public Function GetBotVersionNumber() As Double
		Dim strVersion() As String
		
		strVersion = Split(GetBotVersion(), Space(1))
		
		If (UBound(strVersion)) Then
			If (strVersion(1) = "Beta") Then
				If (UBound(strVersion) > 1) Then
					GetBotVersionNumber = Val(Mid(strVersion(2), 2))
				End If
			Else
				GetBotVersionNumber = Val(Mid(strVersion(1), 2))
			End If
		End If
	End Function
	
	
	'// CONNECT
	'// Connects the bot. Will disconnect an already-existent connection.
	Public Sub Connect()
		Call frmChat.DoConnect()
	End Sub
	
	
	
	'// DISCONNECT
	'// Closes any current connections within the bot.
	Public Sub Disconnect()
		Call frmChat.DoDisconnect()
	End Sub
	
	
	'// BOTCLOSE
	'// Shuts down StealthBot
	Public Sub BotClose()
		'Call frmChat.Form_Unload(0)
		frmChat.Close()
	End Sub
	
	
	'// GETCONFIGENTRY
	'// Reads a value from config.ini and returns it as a string
	'// If no value is present an empty string will be returned
	'// PARAMETERS
	'//     sSection - Section heading from the INI file - examples: "Main", "Other"
	'//     sEntryName - Entry you want to read - examples: "Server", "Username"
	'//     sFileName - File you're reading from - examples: "config.ini", "definitions.ini"
	'// This function will adapt to any filepath hacks the user has in place
	'// You can also use it to read out of your own config file, by specifying a full path
	'//     in the sFileName argument
	Public Function GetConfigEntry(ByVal sSection As String, ByVal sEntryName As String, ByVal sFileName As String) As String
        If Len(sFileName) = 0 Then
            sFileName = GetConfigFilePath()
        End If
		
		sFileName = GetFilePath(sFileName)
		
		If (StrComp(sFileName, StringFormat("{0}\Access.ini", CurDir()), CompareMethod.Text) = 0) Then
			If (StrComp(sSection, "Flags", CompareMethod.Text) = 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object OpenCommand().RequiredFlags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetConfigEntry = OpenCommand(sEntryName, modScripting.GetScriptName).RequiredFlags
			ElseIf (StrComp(sSection, "Numeric", CompareMethod.Text) = 0) Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object OpenCommand().RequiredRank. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetConfigEntry = OpenCommand(sEntryName, modScripting.GetScriptName).RequiredRank
			End If
			
			Exit Function
		End If
		
		GetConfigEntry = ReadINI(sSection, sEntryName, sFileName)
	End Function
	
	
	
	'// WRITECONFIGENTRY
	'// Writes a value to config.ini
	'// PARAMETERS
	'//     sSection - Section heading from the INI file - examples: "Main", "Other"
	'//     sEntryName - Entry you want to read - examples: "Server", "Username"
	'//     sValue - Value to be written to the file
	'//     sFileName - File you're reading from - examples: "config.ini", "definitions.ini"
	'// This function will adapt to any filepath hacks the user has in place
	'// You can also use it to read out of your own config file, by specifying a full path
	'//     in the sFileName argument
	Public Sub WriteConfigEntry(ByVal sSection As String, ByVal sEntryName As String, ByVal sValue As String, ByVal sFileName As String)
        If Len(sFileName) = 0 Then
            sFileName = GetConfigFilePath()
        End If
		
		sFileName = GetFilePath(sFileName)
		
		Dim cmdObj As Object
		If (StrComp(sFileName, StringFormat("{0}\Access.ini", CurDir()), CompareMethod.Text) = 0) Then
			
			cmdObj = OpenCommand(sEntryName, modScripting.GetScriptName)
			
			If (StrComp(sSection, "Flags", CompareMethod.Text) = 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object cmdObj.RequiredFlags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				cmdObj.RequiredFlags = sValue
			ElseIf (StrComp(sSection, "Numeric", CompareMethod.Text) = 0) Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object cmdObj.RequiredRank. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				cmdObj.RequiredRank = sValue
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object cmdObj.Save. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cmdObj.Save()
			
			Exit Sub
		End If
		
		WriteINI(sSection, sEntryName, sValue, sFileName)
	End Sub
	
	
	
	'// VETOTHISMESSAGE
	'// Used with PressedEnter event to prevent a message from being sent to Battle.net
	'// For use processing scripts entirely internally
	Public Sub VetoThisMessage()
		SetVeto(True)
	End Sub
	
	'// PRINTURLTOFILE
	'// Mirror function for the Windows API URLDownloadToFile() function
	'// Currently you are restricted to placing files in the StealthBot install directory only
	Public Sub PrintURLToFile(ByVal sFileName As String, ByVal sURL As String)
		sFileName = StringFormat("{0}\{1}", CurDir(), sFileName)
		
		URLDownloadToFile(0, sURL, sFileName, 0, 0)
	End Sub
	
	
	
	'// DELETEURLCACHE
	'// Mirror function for the Windows API DeleteUrlCacheEntry() function
	'// Call before using PrintURLToFile() to clear any residual IE cache entries for
	'//     the URL you're retrieving
	Public Sub DeleteURLCache(ByVal sURL As String)
		DeleteURLCacheEntry(sURL)
	End Sub
	
	
	
	'// PADQUEUECOUNTER
	'// Pads the queue so further messages will be sent more slowly
	Public Sub PadQueueCounter()
		'QueueLoad = QueueLoad + 1
		'Does nothing now
	End Sub
	
	
	
	'// PADQUEUE
	'// Inserts a blank message into the queue
	Public Sub PadQueue()
		InsertDummyQueueEntry()
	End Sub
	
	
	
	'// GETQUEUESIZE
	'// Returns the number of items currently in the outgoing message queue
	Public Function GetQueueSize() As Short
		GetQueueSize = g_Queue.Count
	End Function
	
	
	
	'// FLASHBOTWINDOW
	'// Flashes the bot's entry in the taskbar to get attention.
	Public Sub FlashBotWindow()
		Call FlashWindow()
	End Sub
	
	
	
	'// RELOADSCRIPT
	'// Reloads the base script.txt file with any includes, equivalent to choosing
	'//  that menu option on the bot's Settings menu
	'// The command must wait before reloading the script so all operations are cleared.
	Public Sub ReloadScript()
        SCReloadTimerID = SetTimer(frmChat.Handle.ToInt32, 0, 400, AddressOf ScriptReload_TimerProc)
	End Sub
	
	
	
	'// SETSCTIMEOUT
	'//   Recommended to modify the timeout setting in your settings.ini file rather than using this sub directly.
	Public Sub SetSCTimeout(ByVal newValue As Integer)
		If (newValue > 1 And newValue < 60001) Then
            SCReloadTimerID = SetTimer(frmChat.Handle.ToInt32, newValue, 400, AddressOf ScriptReload_TimerProc)
		End If
	End Sub
	
	
	'// COLOR
	'//   Gets a reference to the clsColor object. This object has properties that represent the CSS color constants
	Public Function Color() As clsColor
		Color = New clsColor
	End Function
	
	'// GetScriptControl
	'// Returns the Script Control as an object
	Public Function GetScriptControl() As Object
		GetScriptControl = frmChat.SControl
	End Function
	
	
	
	'// GETCOMMANDLINE
	'// Returns the command line arguments specified at the bot's runtime,
	'//  or later during the bot's operation using the /setcl console command.
	Public Function GetCommandLine() As String
		GetCommandLine = CommandLine
	End Function
	
	
	
	'// GETCONFIGPATH
	'// Returns the current full path to the bot's config.ini, accounting for
	'//  any -cpath overrides
	Public Function GetConfigPath() As String
		GetConfigPath = GetConfigFilePath()
	End Function
	
	
	
	'// CLEARSCREEN
	'// Empties the bot's current chat window
	'//  CLEAROPTION:
	'//   1 = clear chat window
	'//   2 = clear whisper window
	'//   3 (default) = clear both chat and whisper window
	Public Function ClearScreen(Optional ByVal ClearOption As Short = 0) As String
		Call frmChat.ClearChatScreen(ClearOption)
	End Function
	
	
	
	'// CDEC
	'// Typecasts a VBS variant to the vbDecimal datatype
	'//  By request from Imhotep[Nu]
	Public Sub C_Dec(ByRef vToCast As Object)
		'UPGRADE_WARNING: Couldn't resolve default property of object vToCast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		vToCast = CDec(vToCast)
	End Sub
	
	
	
	'// REQUESTPROFILEKEY
	'//  Requests a specific user's profile key. Use with care as Blizzard will
	'//  ip-ban you for requesting some keys.
	'//  The result will come back to you in an Event_KeyReturn.
	Public Sub RequestProfileKey(ByVal sUsername As String, ByVal sKey As String)
		SuppressProfileOutput = True
		
		SpecificProfileKey = sKey
		
		RequestSpecificKey(sUsername, sKey)
	End Sub
	
	
	'// GETAPPHINSTANCE
	'//     Returns the App.hInstance value
	Public Function GetApphInstance() As Integer
		GetApphInstance = VB6.GetHInstance.ToInt32
	End Function
	
	'// GETCHATHWND
	'//     Returns the hWnd for the main chat window
	Public Function GetChathWnd() As Integer
		GetChathWnd = frmChat.Handle.ToInt32
	End Function
	
	
	'// DOSTATSTRINGPARSE
	'//     Parses a statstring given to you by GetInternalData (or elsewhere)
	'//     Returns the parsed user-message string you see in join/leave messages
	Public Function DoStatstringParse(ByVal sStatstring As String) As String
		'Dim sBuffer As String
		Dim UserStats As clsUserStats
		
		'Call ParseStatstring(sStatstring, sBuffer, sClanTag)
		
		UserStats = New clsUserStats
		
		UserStats.Statstring = sStatstring
		
		DoStatstringParse = UserStats.ToString_Renamed
		
		'UPGRADE_NOTE: Object UserStats may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		UserStats = Nothing
	End Function
	
	
	'// GETUSERSTATS
	'//     Parses a statstring given to you by GetInerrnalData
	'//     Returns the parsed stats object.
	Public Function GetUserStats(ByVal sStatstring As String) As Object
		Dim UserStats As clsUserStats
		
		' create userstats object
		UserStats = New clsUserStats
		
		' set statstring in object, parses
		UserStats.Statstring = sStatstring
		
		' set as return
		GetUserStats = UserStats
		
		' clean up
		'UPGRADE_NOTE: Object UserStats may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		UserStats = Nothing
	End Function
	
	
	'// GETWINCURSORPOS
	'//     Mirror function for the Windows API function GetCursorPos()
	Public Sub GetWinCursorPos(ByRef lCursorX As Integer, ByRef lCursorY As Integer)
		Dim PAPI As POINTAPI
		
		GetCursorPos(PAPI)
		
		lCursorX = PAPI.x
		lCursorY = PAPI.y
	End Sub
	
	
	'// SETWINCURSORPOS
	'//     Mirror function for the Windows API function SetCursorPos()
	Public Sub SetWinCursorPos(ByVal lNewX As Integer, ByVal lNewY As Integer)
		SetCursorPos(lNewX, lNewY)
	End Sub
	
	
	'// WINFINDWINDOW
	'//     Mirror function for the Windows API function FindWindow()
	Public Function WinFindWindow(ByVal lpClassName As Integer, ByVal lpWindowName As String) As Integer
		WinFindWindow = FindWindow(CStr(lpClassName), lpWindowName)
	End Function
	
	
	'// WINFINDWINDOWEX
	'//     Mirror function for the Windows API function FindWindowEx()
	Public Function WinFindWindowEx(ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Object
		WinFindWindowEx = FindWindowEx(hWnd1, hWnd2, lpsz1, lpsz2)
	End Function
	
	
	'// WINSENDMESSAGE
	'//     Mirror function for the Windows API function
	Public Function WinSendMessage(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
		WinSendMessage = SendMessage(hWnd, wMsg, wParam, lParam)
	End Function
	
	'// WINGETWINDOWLONG
	'//     Mirror function for the Windows API function
	Public Function WinGetWindowLong(ByVal hWnd As Integer, ByVal nIndex As Short) As Integer
		WinGetWindowLong = GetWindowLong(hWnd, nIndex)
	End Function
	
	'// WINSETWINDOWLONG
	'//     Mirror function for the Windows API function
	Public Function WinSetWindowLong(ByVal hWnd As Integer, ByVal nIndex As Short, ByVal dwNewLong As Integer) As Integer
		WinSetWindowLong = SetWindowLong(hWnd, nIndex, dwNewLong)
	End Function
	
	'// DSP
	'//     Displays messages via the specified output type and allows for unlimited message Lengths.
	'//     Syntax: Dsp Output, Message, [Username], [Color]
	'//     Output must be one of the following integer Values():
	'//         1 = AddQ
	'//             Example: Dsp 1, "Example message"
	'//         2 = Emote
	'//             Example: Dsp 2, "Example message"
	'//         3 = Whisper
	'//             Example: Dsp 3, "Example message", "JoeUser"
	'//         4 = AddChat
	'//             Example: Dsp 4, "Example message", , vbCyan
	Public Sub Dsp(ByVal intID As Short, ByVal strMessage As String, Optional ByVal strUsername As String = "", Optional ByVal lngColor As Integer = 16777215)
		
		Select Case intID
			Case 1
				Me.AddQ(strMessage)
			Case 2
				Me.AddQ("/me " & strMessage)
			Case 3
				If Len(strUsername) = 0 Then
					AddChat(RTBColors.ErrorMessageText, "Dsp whisper error: You did not supply a username.")
					Exit Sub
				End If
				Me.AddQ("/w " & strUsername & " " & strMessage)
			Case 4
				AddChat(lngColor, strMessage)
		End Select
	End Sub
	
	
	'// STRINGFORMAT
	'// Replaces {#} parameters in source with the elements passed after it.
	'// EX: StringFormat("Hello {1}, your ping is {2}ms", Username, Ping)
	'UPGRADE_WARNING: ParamArray params was changed from ByRef to ByVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93C6A0DC-8C99-429A-8696-35FC4DCEFCCC"'
	Public Function StringFormat(ByRef source As String, ParamArray ByVal params() As Object) As String
		Dim arr() As Object
		arr = VB6.CopyArray(params)
		StringFormat = modOtherCode.StringFormatA(source, arr)
	End Function
	
	
	'// UTCNOW
	'// Returns the current UTC time and date value
	Public Function UtcNow() As Date
		UtcNow = modDateTime.UtcNow
	End Function
	
	'// LOGGER
	'// Returns the logger object
	Public Function Logger() As Object
		Logger = g_Logger
	End Function
	
	'// ISDEBUG
	'// Returns whether the bot was started with the debug command line
	Public Function isDebug() As Boolean
		isDebug = modGlobals.isDebug()
	End Function
	
	'// FILETIMETODATE
	'// Returns a date/time from two long values, as either "value value" or value, value
	Public Function FileTimeToDate(ByVal strTime As Object, Optional ByVal HighValue As Integer = 0) As Date
		Dim FTime As FILETIME
		
		If (HighValue <> 0) Then
			With FTime
				'UPGRADE_WARNING: Couldn't resolve default property of object strTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.dwLowDateTime = CInt(Val(strTime))
				.dwHighDateTime = CInt(Val(CStr(HighValue)))
			End With
		Else
			With FTime
				'UPGRADE_WARNING: Couldn't resolve default property of object strTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.dwLowDateTime = UnsignedToLong(CDbl(Mid(strTime, InStr(1, strTime, " ", CompareMethod.Binary) + 1)))
				
				'UPGRADE_WARNING: Couldn't resolve default property of object strTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.dwHighDateTime = UnsignedToLong(CDbl(Left(strTime, InStr(1, strTime, " ", CompareMethod.Binary))))
			End With
		End If
		
		FileTimeToDate = modDateTime.FileTimeToDate(FTime)
	End Function
	
	'// QUOTES
	'// Returns the quotes object
	Public Function Quotes() As Object
		Quotes = g_Quotes
	End Function
	
	'// MEDIAPLAYER
	'// Returns the WinAmp or iTunes media player object
	Public Function MediaPlayer() As Object
		MediaPlayer = modMediaPlayer.MediaPlayer()
	End Function
	
	Public Function Config() As Object
		Config = modGlobals.Config
	End Function
	
	'// CHANNEL
	'// Returns the channel object
	Public Function Channel() As Object
		Channel = g_Channel.Clone()
	End Function
	
	'// CLAN
	'// Returns the clan object
	Public Function Clan() As Object
		Clan = g_Clan.Clone()
	End Function
	
	'// FRIENDS
	'// Returns a collection of friends objects
	Public Function Friends() As Object
		Dim i As Short
		
		Friends = New Collection
		
		For i = 1 To g_Friends.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends().Clone. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Friends.Add. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Friends.Add(g_Friends.Item(i).Clone())
		Next i
	End Function
	
	'// QUEUE
	'// Returns the queue object
	Public Function Queue() As Object
		Queue = g_Queue
	End Function
	
	'// OSVERSION
	'// Returns an object with properties returning operating system information
	Public Function OSVersion() As Object
		OSVersion = New clsOSVersion
	End Function
	
	'// GETSYSTEMUPTIME
	'// Returns the amount of time your system has been up
	Public Function GetSystemUptime() As String
		GetSystemUptime = ConvertTime(GetUptimeMS)
	End Function
	
	'// GETCONNECTIONUPTIME
	'// Returns the amount of time your bot has been online
	Public Function GetConnectionUptime() As String
		GetConnectionUptime = ConvertTime(uTicks)
	End Function
	
	'// CRC32
	'// Returns the result of a standard CRC32 hash
	'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function CRC32(ByRef str_Renamed As String) As Integer
		Dim clsCRC32 As New clsCRC32
		
		CRC32 = clsCRC32.CRC32(str_Renamed)
		
		'UPGRADE_NOTE: Object clsCRC32 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		clsCRC32 = Nothing
	End Function
	
	'// SHA1
	'// Returns the result of a standard Sha-1 hash
	Public Function Sha1(ByRef Data As String, Optional ByVal inHex As Boolean = False) As String
		Dim a As Integer
		Dim B As Integer
		Dim c As Integer
		Dim d As Integer
		Dim e As Integer
		
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		Call modSHA1.DefaultSHA1( System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(Data, vbFromUnicode)), a, B, c, d, e)
		
		If inHex Then
			Sha1 = LCase(Hex(a) & Hex(B) & Hex(c) & Hex(d) & Hex(e))
		Else
			Sha1 = LongToStr(a) & LongToStr(B) & LongToStr(c) & LongToStr(d) & LongToStr(e)
		End If
	End Function
	
	'// XSHA1
	'// Returns the result of a non-standard Battle.net "broken" Sha-1 hash
	'// Changed 10/18/2009 - Hdx
	'//    Added Optional Argument "Spacer", This will go between every hex digit, people usually use Space$(1) for nice debugging print
	'//    inHex will now properly pad characters that are < 0x10 with a leading 0, Chr$(6) becomes "06" instead of "6"
	'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function XSHA1(ByRef str_Renamed As String, Optional ByVal inHex As Boolean = False, Optional ByRef SPACER As String = vbNullString) As String
		Dim i As Short
		Dim s As String
		
		XSHA1 = modBNCSutil.hashPassword(str_Renamed)
		
		If inHex Then
			For i = 1 To Len(XSHA1)
				s = StringFormat("{0}{1}{2}", s, IIf(i > 1, SPACER, vbNullString), ZeroOffset(Asc(Mid(XSHA1, i, 1)), 2))
			Next 
			XSHA1 = s
		End If
	End Function
	
	'// CREATENLS
	'// Creates an NLS object to log on with a WarCraft III account
	Public Function CreateNls(ByVal Username As String, ByVal Password As String) As clsNLS
		Dim NLS As clsNLS
		
		NLS = New clsNLS
		
		If NLS.Initialize(Username, Password) Then
			CreateNls = NLS
		End If
		
		'UPGRADE_NOTE: Object NLS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		NLS = Nothing
	End Function
	
	'// VERIFYSERVERSIGNATURE
	'// Verifies a WarCraft III server signature to be for the passed IP address
	Public Function VerifyServerSignature(ByVal IPAddress As String, ByVal ServerSignature As String) As Boolean
		Dim NLS As clsNLS
		
		NLS = New clsNLS
		
		VerifyServerSignature = NLS.VerifyServerSignature(IPAddress, ServerSignature)
		
		'UPGRADE_NOTE: Object NLS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		NLS = Nothing
	End Function
	
	'// CREATEKEYDECODER
	'// Creates a KeyDecoder object for decoding and hashing CD keys.
	Public Function CreateKeyDecoder(ByVal CDKey As String) As clsKeyDecoder
		Dim KD As New clsKeyDecoder
		
		If KD.Initialize(CDKey) Then
			CreateKeyDecoder = KD
		End If
		
		'UPGRADE_NOTE: Object KD may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		KD = Nothing
	End Function
	
	'// GETALLFONTS
	'// Gets an array of all currently installed fonts on this system.
	Public Function GetAllFonts() As String()
		Dim Fonts() As String
		Dim i As Short
		
		'UPGRADE_ISSUE: Screen property Screen.FontCount was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		ReDim Fonts(Screen.FontCount - 1)
		'UPGRADE_ISSUE: Screen property Screen.FontCount was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		For i = 0 To Screen.FontCount - 1
			'UPGRADE_ISSUE: Screen property Screen.Fonts was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Fonts(i) = Screen.Fonts(i)
		Next i
		
		GetAllFonts = VB6.CopyArray(Fonts)
	End Function
	
	'// CLIPBOARD
	'// Access to the system clipboard.
	Public Function Clipboard() As Microsoft.VisualBasic.PC.Clipboard
		My.Computer.Clipboard = My.Computer.Clipboard
	End Function
	
	'// TWIPSPERPIXELX
	'// Get the horizontal ratio of twips/pixel.
	Public Function TwipsPerPixelX() As Single
		TwipsPerPixelX = VB6.TwipsPerPixelX
	End Function
	
	'// TWIPSPERPIXELY
	'// Get the vertical ratio of twips/pixel.
	Public Function TwipsPerPixelY() As Single
		TwipsPerPixelY = VB6.TwipsPerPixelY
	End Function
	
	'// SCREENWIDTH
	'// Get the width of the current display in twips.
	Public Function ScreenWidth() As Single
		ScreenWidth = VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width)
	End Function
	
	'// SCREENHEIGHT
	'// Get the height of the current display in twips.
	Public Function ScreenHeight() As Single
		ScreenHeight = VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height)
	End Function
	
	'// GETSCRIPTMODULE
	'// Returns the currently executing script module object
	'// PARAMETERS
	'//     ScriptName - this function will return that script's module
	Public Function GetScriptModule(ByVal scriptName As String) As Object
		GetScriptModule = modScripting.GetScriptModule(scriptName)
	End Function
	
	'// GETMODULEID
	'// Returns a script module's ID
	'// This ID is a string from 2 to the amount of scripts loaded + 1
	'// PARAMETERS
	'//     ScriptName - this function will return that script's module ID
	Public Function GetModuleID(ByVal scriptName As String) As String
		GetModuleID = modScripting.GetModuleID(scriptName)
	End Function
	
	'// GETSCRIPTNAME
	'// Returns a script's stored name
	'// PARAMETERS
	'//     ModuleID - this function will return the script name of the provided script module
	Public Function GetScriptName(ByVal ModuleID As String) As String
		GetScriptName = modScripting.GetScriptName(ModuleID)
	End Function
	
	'// GETWORKINGDIRECTORY
	'// Returns the working directory for your script
	'// Use this as a place to store script configurations, databases, or other data for your script
	'// It is recommended to use this over BotPath() for script-specific information
	'// PARAMETERS
	'//     ScriptName - this function will return that script's working directory
	Public Function GetWorkingDirectory(ByVal scriptName As String) As String
		' if none provided, get current script name
        If Len(scriptName) = 0 Then scriptName = modScripting.GetScriptName
		
		' if we are in /exec or something, return vbnullstring
        If Len(scriptName) = 0 Then Exit Function
		
		' return working directory
		GetWorkingDirectory = BotPath() & "Scripts\" & scriptName & "\"
		
		On Error Resume Next
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Len(Dir(GetWorkingDirectory, FileAttribute.Directory)) = 0 Then
            MkDir(GetWorkingDirectory)
        End If
	End Function
	
	'// CREATEOBJ
	'// Creates a script-specific object for use with your script
	'// Returns the object as a result of the function, and makes ObjName directly accessible by the ObjName.Method syntax
	'// PARAMETERS
	'//     ObjType -    this will be one of the following:
	'//                  "Timer" - a script timer to do stuff after an interval (in milliseconds)
	'//                  "LongTimer" - a script timer to do stuff after an interval (in seconds - like the old PluginSystem)
	'//                  "Winsock" - a Windows Socket, allowing you to connect to remote servers
	'//                  "Inet" - a script Inet control allowing easy access to the HTML source of a website
	'//                  "Form" - a window UI fully accessible to the script
	'//                  "Menu" - a script menu UI that appears under the Scripting > (Script Name) menu, but can be moved elsewhere (such as into a form)
	'//     ObjName -    the name of the object
	'//                  this must be valid as a variable name in script
	'//     ScriptName - this function will create the object for the provided script
	Public Function CreateObj(ByVal ObjType As String, ByVal ObjName As String, ByVal scriptName As String) As Object
		Dim ModuleID As String
		
		ModuleID = modScripting.GetModuleID(scriptName)
		
		CreateObj = modScripting.CreateObj(frmChat.SControl.Modules(ModuleID), ObjType, ObjName)
	End Function
	
	'// DESTROYOBJ
	'// Destroys an object created with CreateObj
	'// PARAMETERS
	'//     ObjName -    the name of the object
	'//     ScriptName - this function will destroy the object for the provided script
	Public Sub DestroyObj(ByVal ObjName As String, ByVal scriptName As String)
		Dim ModuleID As String
		
		ModuleID = modScripting.GetModuleID(scriptName)
		
		modScripting.DestroyObj(frmChat.SControl.Modules(ModuleID), ObjName)
	End Sub
	
	'// GETOBJBYNAME
	'// Returns an object created with CreateObj
	'// PARAMETERS
	'//     ObjName -    the name of the object
	'//     ScriptName - this function will get the object for the provided script
	Public Function GetObjByName(ByVal ObjName As String, ByVal scriptName As String) As Object
		Dim ModuleID As String
		
		ModuleID = modScripting.GetModuleID(scriptName)
		
		GetObjByName = modScripting.GetObjByName(frmChat.SControl.Modules(ModuleID), ObjName)
	End Function
	
	'// GETCOMMANDS
	'// Gets a collection of stored commands
	'// PARAMETERS
	'//     ScriptName -  this function will get the commands for the specified script
	'//                   vbNullString to get internal commands
	'//                   Chr(0) to get commands from any script that is enabled
	'UPGRADE_NOTE: clsCommandDocObj was upgraded to clsCommandDocObj_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function GetCommands(ByVal scriptName As String) As Collection
		Dim clsCommandDocObj_Renamed As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object clsCommandDocObj.GetCommands. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetCommands = clsCommandDocObj.GetCommands(scriptName)
	End Function
	
	'// CREATECOMMAND
	'// Creates a command for your script
	'// Returns the CommandDocs object after creation for modifications and saving
	'// PARAMETERS
	'//     commandName - the name of the new command
	'//     ScriptName -  this function will create the command for the specified script
	'//                   vbNullString to get internal commands
	'//                   Chr(0) to get commands from any script
	Public Function CreateCommand(ByVal commandName As String, ByVal scriptName As String) As Object
		'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Command_Renamed As clsCommandDocObj
		
        If ((Len(scriptName) = 0) Or scriptName = Chr(0)) Then
            frmChat.AddChat(RTBColors.ErrorMessageText, "Error: You can not create a command without an owner.")
            Exit Function
        End If
		
		Command_Renamed = New clsCommandDocObj
		
		Call Command_Renamed.CreateCommand(commandName, scriptName)
		Call Command_Renamed.OpenCommand(commandName, scriptName)
		
		CreateCommand = Command_Renamed
		
		'UPGRADE_NOTE: Object Command_Renamed may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Command_Renamed = Nothing
	End Function
	
	'// OPENCOMMAND
	'// Returns the CommandDocs object of a script command
	'// PARAMETERS
	'//     commandName - the name of the command
	'//     ScriptName -  this function will get the command for the specified script
	'//                   vbNullString to get internal commands
	'//                   Chr(0) to get commands from any script that is enabled
	Public Function OpenCommand(ByVal commandName As String, ByVal scriptName As String) As Object
		'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Command_Renamed As clsCommandDocObj
		
		Command_Renamed = New clsCommandDocObj
		
		If Command_Renamed.OpenCommand(commandName, scriptName) Then
			OpenCommand = Command_Renamed
		End If
		
		'UPGRADE_NOTE: Object Command_Renamed may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Command_Renamed = Nothing
	End Function
	
	'// DELETECOMMAND
	'// Deletes a command for a script
	'// PARAMETERS
	'//     commandName - the name of the command
	'//     ScriptName -  this function will delete the command for the specified script
	'//                   vbNullString to get internal commands
	'//                   Chr(0) to get commands from any script
	Public Function DeleteCommand(ByVal commandName As String, ByVal scriptName As String) As Object
		'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Command_Renamed As clsCommandDocObj
		
		Command_Renamed = New clsCommandDocObj
		
		Call Command_Renamed.OpenCommand(commandName, scriptName)
		Call Command_Renamed.Delete()
		
		DeleteCommand = Command_Renamed
		
		'UPGRADE_NOTE: Object Command_Renamed may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Command_Renamed = Nothing
	End Function
	
	'// OBSERVESCRIPT
	'// Add the currently runnign script as an Observer of the Specified script
	'// The current script will have all events of the observed script duplicated and executed
	Public Sub ObserveScript(ByVal Script As String)
		modScripting.AddScriptObserver(modScripting.GetScriptName, Script)
	End Sub
	
	'// GETOBSERVED
	'// Returns a collection of script names that this script is observing
	Public Function GetObserved() As Collection
		GetObserved = modScripting.GetScriptObservers(modScripting.GetScriptName, True)
	End Function
	
	'// GETOBSERVERS
	'// Returns a collection of script names that are observing this script
	Public Function GetObservers() As Collection
		GetObservers = modScripting.GetScriptObservers(modScripting.GetScriptName, False)
	End Function
	
	'// OBSERVEFUNCTION
	'// Add the currently running script as on Observer for this function, in all other scripts.
	'// The current script will have this function executed every time it is executed in another script.
	Public Sub ObserveFunction(ByVal sFunction As String)
		modScripting.AddFunctionObserver(modScripting.GetScriptName, sFunction)
	End Sub
	
	'// GETSETTINGSENTRY
	'// Retrieves a value from settings.ini for your script
	'// PARAMETERS
	'//     sEntryName - the name of the entry to look up
	'//     ScriptName - this function will return the specified script's setting
	Public Function GetSettingsEntry(ByVal sEntryName As String, ByVal scriptName As String) As String
		On Error Resume Next
		
		Dim Path As String
		
		If scriptName = vbNullString Then scriptName = modScripting.GetScriptName
		
		Path = GetFilePath(FILE_SCRIPT_INI, GetFolderPath("Scripts"))
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If (Len(Dir(Path)) = 0) Then
            FileOpen(1, Path, OpenMode.Output)
            FileClose(1)
        End If
		
		GetSettingsEntry = ReadINI(scriptName, sEntryName, Path)
		
		If (Not InStr(1, GetSettingsEntry, " ;") = 0) Then
			GetSettingsEntry = Mid(GetSettingsEntry, 1, InStr(1, GetSettingsEntry, " ;") - 1)
		End If
		
		If (Not InStr(1, GetSettingsEntry, " #") = 0) Then
			GetSettingsEntry = Mid(GetSettingsEntry, 1, InStr(1, GetSettingsEntry, " #") - 1)
		End If
	End Function
	
	'// WRITESETTINGSENTRY
	'// Stores a value to settings.ini for your script
	'// PARAMETERS
	'//     sEntryName - the name of the entry to store in
	'//     sValue -     the value to store in the entry
	'//     ScriptName - this function will write the setting for the specified script
	Public Sub WriteSettingsEntry(ByVal sEntryName As String, ByVal sValue As String, Optional ByVal sDescription As String = "", Optional ByVal scriptName As String = "")
		On Error GoTo ERROR_HANDLER
		
		Dim Path As String
		
		If scriptName = vbNullString Then scriptName = modScripting.GetScriptName
		
		Path = GetFilePath(FILE_SCRIPT_INI, GetFolderPath("Scripts"))
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If (Len(Dir(Path)) = 0) Then
            FileOpen(1, Path, OpenMode.Output)
            FileClose(1)
        End If
		
		WriteINI(scriptName, sEntryName, sValue & IIf(Len(sDescription), " ; " & sDescription, ""), Path)
		
		Exit Sub
		
ERROR_HANDLER: 
		
		Dim f As Short
		
		f = FreeFile
		
		On Error Resume Next
		
		MkDir(GetFolderPath("Scripts"))
		
		On Error GoTo ERROR_HANDLER
		
		FileOpen(f, Path, OpenMode.Output)
		
		FileClose(f)
		
		Resume Next
	End Sub
	
	'// SCRIPTS
	'// Gets a collection of all script CodeObjects, allowing direct access via the object.method syntax
	Public Function Scripts() As Object
		On Error Resume Next
		
		Scripts = modScripting.Scripts()
	End Function
	
	'// GETSCRIPTBYNAME
	'// Gets a script by its name
	Public Function GetScriptByName(ByVal scriptName As String) As Object
		On Error Resume Next
		
		GetScriptByName = Scripts(scriptName)
	End Function
	
	'// STRCONVEX
	'// Mirrors the VB6 StrConv function
	'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function StrConvEx(ByVal str_Renamed As String, ByVal Conv As VbStrConv, Optional ByVal locale As Integer = 0) As Object
		StrConvEx = StrConv(str_Renamed, Conv, locale)
	End Function
	
	'// DATABUFFEREX
	'// Gets an instance of a databuffer for the currently executing script
	'// This will return a new one everytime!
	Public Function DataBufferEx() As Object
		DataBufferEx = New clsDataBuffer
		
		'UPGRADE_WARNING: Couldn't resolve default property of object DataBufferEx.setCripple. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DataBufferEx.setCripple()
	End Function
	
	'// RESOLVEHOSTNAME
	'// Returns the IP address resolved from the given host
	Public Function ResolveHostName(ByVal strHostName As String, Optional ByRef errCode As Integer = 0) As String
		Dim Result As String
		Result = ResolveHost(strHostName)
		If Result = vbNullString Then
			errCode = WSAGetLastError()
		End If
		ResolveHostName = Result
	End Function
	
	'// FORCEBNCSPACKETPARSE
	'// Pass a complete Battle.net packet into this and the bot will parse it
	Public Sub ForceBNCSPacketParse(ByVal PacketData As String)
		Call BNCSParsePacket(PacketData)
	End Sub
	
	'// GETUSERDATABASE
	'// Gets a collection of users in the database
	'// 10/10/2009 B#003 52 - Changed code to destroy the temp object after each use
	Public Function GetUserDatabase() As Collection
		Dim x As Short
		Dim temp As clsDBEntryObj
		GetUserDatabase = New Collection
		For x = LBound(DB) To UBound(DB)
			If (Len(DB(x).Username) > 0) Then
				temp = New clsDBEntryObj
				With temp
					.Name = DB(x).Username
					.Rank = DB(x).Rank
					.CreatedOn = DB(x).AddedOn
					.CreatedBy = DB(x).AddedBy
					.BanMessage = DB(x).BanMessage
					.Flags = DB(x).Flags
					.AddGroup(DB(x).Groups)
					.ModifiedBy = DB(x).ModifiedBy
					.ModifiedOn = DB(x).ModifiedOn
					.EntryType = DB(x).Type
				End With
				GetUserDatabase.Add(temp, DB(x).Username)
				'UPGRADE_NOTE: Object temp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				temp = Nothing
			End If
		Next x
	End Function
	
	'// GETCURRENTUSERNAME
	'// Will return the current username for the bot, as Battle.net sees it
	Public Function GetCurrentUsername() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object GetCurrentUsername. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetCurrentUsername = modGlobals.CurrentUsername
	End Function
	
	'// GETCURRENTSERVERIP
	'// Returns the IP of the server the bot is currently connected to.
	Public Function GetCurrentServerIP() As Object
		If modGlobals.g_Online Then
			'UPGRADE_WARNING: Couldn't resolve default property of object GetCurrentServerIP. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetCurrentServerIP = frmChat.sckBNet.RemoteHostIP
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object GetCurrentServerIP. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetCurrentServerIP = vbNullString
		End If
	End Function
	
	'// CHOOSE
	'// Returns the chosen value based on the given index.
	'//  see: https://msdn.microsoft.com/en-us/library/aa262690(v=vs.60).aspx
	'UPGRADE_NOTE: Choose was upgraded to Choose_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_WARNING: ParamArray pChoice was changed from ByRef to ByVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93C6A0DC-8C99-429A-8696-35FC4DCEFCCC"'
	Public Function Choose_Renamed(ByVal aIndex As Single, ParamArray ByVal pChoice() As Object) As Object
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Choose_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Choose_Renamed = System.DBNull.Value
		If aIndex > 0 And aIndex <= UBound(pChoice) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object pChoice(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Choose_Renamed = pChoice(aIndex - 1)
		End If
	End Function
	
	'// IIF
	'// Returns a different value based on if a statement is true or false.
	'//  see: https://msdn.microsoft.com/en-us/library/aa445024(v=vs.60).aspx
	'UPGRADE_NOTE: IIf was upgraded to IIf_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function IIf_Renamed(ByVal bExpression As Boolean, ByVal vTruePart As Object, ByVal vFalsePart As Object) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object Interaction.IIf(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object IIf_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		IIf_Renamed = IIf(bExpression, vTruePart, vFalsePart)
	End Function
	
	'// PARTITION
	'// Returns the range that the specified Number appears in in the given interval defined by Start, Stop, and Interval.
	'//  see: https://msdn.microsoft.com/en-us/library/aa445092(v=vs.60).aspx
	'UPGRADE_NOTE: Partition was upgraded to Partition_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function Partition_Renamed(ByVal aNumber As Integer, ByVal aStart As Integer, ByVal aStop As Integer, ByVal aInterval As Integer) As String
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		Partition_Renamed = System.DBNull.Value
		If aInterval > 0 And aStop > aStart Then
			Partition_Renamed = Partition(aNumber, aStart, aStop, aInterval)
		End If
	End Function
	
	'// SWITCH
	'// Returns the first even item in the parameter list after the odd expression that is true.
	'// There must be an even number of arguments.
	'//  see: https://msdn.microsoft.com/en-us/library/aa263378(v=vs.60).aspx
	'UPGRADE_NOTE: Switch was upgraded to Switch_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_WARNING: ParamArray pVarExpr was changed from ByRef to ByVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93C6A0DC-8C99-429A-8696-35FC4DCEFCCC"'
	Public Function Switch_Renamed(ParamArray ByVal pVarExpr() As Object) As Object
		Dim i As Integer
		
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Switch_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Switch_Renamed = System.DBNull.Value
		For i = 0 To UBound(pVarExpr) - 1 Step 2
			'UPGRADE_WARNING: Couldn't resolve default property of object pVarExpr(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If pVarExpr(i) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object pVarExpr(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Switch_Renamed = pVarExpr(i + 1)
				Exit Function
			End If
		Next i
	End Function
	
	'Class_Terminate doesn't run when a UI form is created via a script. I made this sub so the bot's Unload sub can explicitly _
	'call the code within Class_Terminate to kill all the objects. - FrOzeN
	Public Sub Dispose()
		Call Class_Terminate_Renamed()
	End Sub
	
	'Re-coded this sub so it properly kills all the UI forms created that are still open. - FrOzeN
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		frmChat.SControl.Reset()
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class