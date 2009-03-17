VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScriptSupportClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' *     - Rewrote and integrated several Plugin System functions
' *         Dsp(), SharedScriptSupport.GetSetting(), SetSetting(), TimerInterval(), TimerEnabled(),
' *         GetTimerEnabled(), GetTimeLeft(), and GetTimerWaiting()
' *         Note: All integrated functions are backwards compatiable. No current calls need to be modified.
' *     - Added two new timer functions
' *         GetTimerInterval, TimeLeft()
' *     - Added two new functions to work with plugins
' *         PluginEnabled(), GetPluginEnabled()
' *     - Added six functions to work with the new Plugins menu
' *         AddMenuItem(), SetMenuItemCheck(), GetMenuItemCheck(), ToggleMenuItemCheck(),
' *         SetMenuItemEnabled(), and ReloadPluginMenus()
' *         For a more detailed explanation, see the function comments below, or visit:
' *            http://www.stealthbot.net/board/index.php?showtopic=29141
' *     - Added new functions for creating UI forms
' *         CreateForm(), DestroyForm(), AddFormObject(), GetUIObject()
' *         For a detailed explination of each event and examples visit:
' *            http://www.stealthbot.net/board/index.php?showtopic=27969
' *     - Added GetBotVersionNumber() function which returns only the numerical value of the bot version
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

Option Explicit

Public MyChannel As String  '// will contain the bot's current channel at runtime.
Public BotFlags As Long     '// will contain the bot's current battle.net flags at runtime.
Public myUsername As String '// will contain the bot's current username at runtime.
                            '// NOTE: This may be different than the bot's config.ini username

'// myTrigger has been replaced by BotVars.Trigger
'Public myTrigger As String '// contains the bot's current trigger at runtime
         
Public Enum BanTypes
    btBan = 0     '// used in calling the BanKickUnban() subroutine
    btKick = 1
    btUnban = 2
End Enum

Public Enum ControlTypes
    btCommandBox = 0
    btLabel = 1
    btTextBox = 2
    btTimer = 3
    btPictureBox = 4
    btCheckBox = 5
    btOptionBox = 6
    btComboBox = 7
    btListBox = 8
    btShape = 9
    btLine = 10
    btListView = 11
    btImageList = 12
    btINet = 13
    btRichTextBox = 14
End Enum

Private obScriptForms As Object
Private DictionaryKeys As String

Private Type Timer
    id       As Integer
    Script   As Module
    code     As String
    Interval As Integer
End Type

'Private Timers     As Collection
'Private TimerCount As Integer

' ...
'Public Function setInterval(ScriptModule As Module, strCode As String, intInterval As Integer) As Integer
'
'    On Error Resume Next
'
'    Dim tmr As Timer ' ...
'
'    ' ...
'    With tmr
'        .ID = TimerCount + 1
'        Set .Script = ScriptModule
'        .Code = strCode
'        .Interval = intInterval
'    End With
'
'    ' ...
'    Timers.Add tmr, TimerCount
'
'    ' ...
'    setInterval = TimerCount
'
'    ' ...
'    TimerCount = TimerCount + 1
'
'End Function

' ...
'Public Sub clearInterval(ScriptModule As Module, intIntervalID As Integer)
'
'    On Error Resume Next'
'
'    ' ...
'    Timers.Remove intIntervalID
'
'End Sub


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
Public Sub AddChat(ParamArray saElements() As Variant)
    Dim arr() As Variant ' ...
    
    ' ...
    arr() = saElements
    
    ' ...
    Call DisplayRichText(frmChat.rtbChat, arr)
End Sub


'// ADDQ (ADD QUEUE)
'// Adds a string to the message send queue.
'// Nonzero priority messages will be sent with precedence over 0-priority messages.
Public Sub AddQ(ByVal sText As String, Optional ByVal msg_priority As Integer = -1, Optional ByVal _
    Tag As String = vbNullString)
    
    '// Disabled while in Clan SBs.
    If ((StrComp(g_Channel.Name, "Clan SBs", vbTextCompare) = 0) And _
        (IsStealthBotTech() = False)) Then Exit Sub
    
    '// priority zero reserved for internal usage only
    'If (msg_priority = 0) Then
    '    msg_priority = 1
    'End If

    '// ...
    Call frmChat.AddQ(sText, msg_priority, g_lastQueueUser, Tag)
    
    '// ...
    'If (g_lastQueueUser <> vbNullString) Then
    '    g_lastQueueUser = vbNullString
    'End If
End Sub


'// COMMAND
'// Calls StealthBot's command processor
'// Messages passed to the processor will be evaluated as commands
'// Public Function Commands(ByRef dbAccess As udtGetAccessResponse, Username As String, _
'//     Message As String, InBot As Boolean, Optional CC As Byte = 0) As String

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
'// CC          CC controls whether or not the command string is passed to the Custom Command processor.
'//             For scripting purposes I will not allow the user to set CC, it is always 0.
Public Sub Command(ByVal Username As String, ByVal Message As String, Optional ByVal IsLocal As Boolean = _
    False, Optional ByVal Whispered As Boolean = False)
    
    ' execute command
    Call ProcessCommand(Username, Message, IsLocal, Whispered)
End Sub


'// COMMANDEX
'// Calls StealthBot's command processor, allowing YOU to specify the user's access
'// Messages passed to the processor will be evaluated as commands
'// Public Function Commands(ByRef dbAccess As udtGetAccessResponse, Username As String, _
'//     Message As String, InBot As Boolean, Optional CC As Byte = 0) As String

'// RETURNS: The response string, or an empty string if there is no response

'// Description of each variable:
'//     All variables not mentioned are the same as above
'//     uAccess: access to grant this user
'//     uFlags:  flags to grant this user
'//  *** I will NOT be checking your flags for sanity.
'//  *** Please ensure that your flags are in a string format, "ABCDEF",
'//         ALL uppercase and with no repeating letters, or the bot will become
'//         somewhat confused.

Public Function CommandEx(ByVal Username As String, ByVal Message As String, Optional ByVal IsLocal As Boolean = _
    False) As String
    
    Dim dbAccess           As udtGetAccessResponse
    Dim command_response() As String
    Dim bln                As Boolean

    ' ...
    If (IsLocal) Then
        ' ...
        Username = "(console)"
    
        ' ...
        With dbAccess
            .Access = 200
            .Flags = "A"
        End With
    Else
        ' ...
        dbAccess = GetCumulativeAccess(Username)
    End If

    ' execute command
    bln = executeCommand(Username, dbAccess, Message, IsLocal, command_response)
    
    ' ...
    CommandEx = Join(command_response, Chr$(0))
End Function



'// PING
'// Returns the cached ping of the specified user.
'// If the user is not present in the channel, it returns -3.
Public Function PingByName(ByVal Username As String) As Long
    
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
         
Public Function BanKickUnban(ByVal Inpt As String, ByVal SpeakerAccess As Integer, _
    Optional ByVal Mode As BanTypes = 0) As String

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
Public Function GetDBEntry(ByVal Username As String, Optional ByRef Access, Optional ByRef Flags, _
    Optional ByRef entryType As String = "USER") As Object

    Dim temp   As udtGetAccessResponse
    Dim I      As Integer
    Dim Splt() As String
    
    temp = GetCumulativeAccess(Username, entryType)
    
    With temp
        Access = .Access
        Flags = .Flags
    End With
    
    ' ...
    Set GetDBEntry = New clsDBEntryObj
    
    ' ...
    With GetDBEntry
        .entryType = temp.Type
        .Name = temp.Username
        .rank = temp.Access
        .Flags = temp.Flags
        .CreatedOn = temp.AddedOn
        .CreatedBy = temp.AddedBy
        .ModifiedOn = temp.ModifiedOn
        .ModifiedBy = temp.ModifiedBy
    End With

    If ((temp.Groups <> vbNullString) And (temp.Groups <> "%")) Then
        If (InStr(1, temp.Groups, ",", vbBinaryCompare) <> 0) Then
            Splt() = Split(temp.Groups, ",")
        Else
            ReDim Preserve Splt(0)
        
            Splt(0) = temp.Groups
        End If
    
        For I = LBound(Splt) To UBound(Splt)
            GetDBEntry.Groups.Add GetDBEntry(Splt(I), , , "GROUP")
        Next I
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
Public Sub GetStdDBEntry(ByVal Username As String, ByRef Access, ByRef Flags, _
    Optional ByRef entryType As String = "USER") '// yum, Variants >:\

    Dim temp As udtGetAccessResponse
    
    temp = GetAccess(Username, entryType)
    
    With temp
        Access = .Access
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
Public Function GetGTC() As Long
    
    GetGTC = GetTickCount()
    
End Function



'// DOBEEP
'// Executes a call to the Beep() API function
Public Function DoBeep(ByVal lFreq As Long, ByVal lDuration As Long) As Long

    DoBeep = Beep(lFreq, lDuration)
    
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
Public Sub GetUserProfile(Optional ByVal Username As String)
    
    SuppressProfileOutput = True
    
    If LenB(Username) > 0 Then
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
Public Function GetInternalUserCount() As Integer
    
    GetInternalUserCount = Channel.Users.Count
    
End Function


'// GETINTERNALDATABYUSERNAME
'// Retrieves the specified stored internal data for a given user in the channel
'//  If the specified user isn't present, the return value is -5
'//  See lDataType constants in GetInternalData() below
Public Function GetInternalDataByUsername(ByVal sUser As String, ByVal lDataType As Long) As Variant

    Dim I As Integer
    
    I = Channel.GetUserIndex(sUser)
    
    GetInternalDataByUsername = GetInternalData(I, lDataType)

End Function


'// GETINTERNALDATA
'// Retrieves the specified stored internal data for a given user in the channel
'//  If the specified user is not present, return value is '-5'
Public Function GetInternalData(ByVal iIndex As Integer, ByVal lDataType As Long) As Variant
    ' -- '
    '       these constants will be useful in making calls to this function
    '                           |   <purpose>
    Const GID_CLAN = 0          '-> retrieves 4-character clan name
    Const GID_FLAGS = 1         '-> retrieves Battle.net flags
    Const GID_PING = 2          '-> retrieves ping on login
    Const GID_PRODUCT = 3       '-> retrieves 4-digit product code
    Const GID_ISSAFELISTED = 4  '-> retrieves Boolean value denoting safelistedness
    Const GID_STATSTRING = 5    '-> retrieves unparsed statstring
    Const GID_TIMEINCHANNEL = 6 '-> retrieves time in channel in seconds
    Const GID_TIMESINCETALK = 7 '-> retrieves time since the user's last message in seconds
    ' -- '
    
    If iIndex > 0 Then
        Select Case lDataType
            Case GID_CLAN
                GetInternalData = Channel.Users(iIndex).Clan
                
            Case GID_FLAGS
                GetInternalData = Channel.Users(iIndex).Flags
            
            Case GID_PING
                GetInternalData = Channel.Users(iIndex).Ping
            
            Case GID_PRODUCT
                GetInternalData = Channel.Users(iIndex).game
            
            Case GID_ISSAFELISTED
                GetInternalData = GetSafelist(Channel.Users(iIndex).DisplayName)
            
            Case GID_STATSTRING
                GetInternalData = Channel.Users(iIndex).Statstring
            
            Case GID_TIMEINCHANNEL
                GetInternalData = Channel.Users(iIndex).TimeInChannel
            
            Case GID_TIMESINCETALK
                GetInternalData = Channel.Users(iIndex).TimeSinceTalk
            
            Case Else
                GetInternalData = 0
                
        End Select
    Else
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
Public Function GetPositionByName(ByVal sUser As String) As Integer
    
    GetPositionByName = checkChannel(sUser)
    
End Function


'// GETNAMEBYPOSITION
'// Returns the name of the person at position X in the channel list.
'// Positions are 1-based. Returns an empty string if the user isn't present
Public Function GetNameByPosition(ByVal X As Integer) As String
    
    With frmChat.lvChannel.ListItems
        If X > 0 And X <= .Count Then
            GetNameByPosition = .Item(X).text
        Else
            GetNameByPosition = ""
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
Function GetBotVersionNumber()
   Dim strVersion() As String
   
   strVersion = Split(GetBotVersion(), Space(1))

   If (UBound(strVersion)) Then
      If (strVersion(1) = "Beta") Then
         If (UBound(strVersion) > 1) Then
            GetBotVersionNumber = Mid$(strVersion(2), 2)
         End If
      Else
         GetBotVersionNumber = Mid$(strVersion(1), 2)
      End If
   End If
End Function


'// CONNECT
'// Connects the bot. Will disconnect an already-existent connection.
Public Sub Connect()
    
    Call frmChat.DoConnect
    
End Sub



'// DISCONNECT
'// Closes any current connections within the bot.
Public Sub Disconnect()
    
    Call frmChat.DoDisconnect
    
End Sub


'// BOTCLOSE
'// Shuts down StealthBot
Public Sub BotClose()
    
    'Call frmChat.Form_Unload(0)
    Unload frmChat
    
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
    
    If LenB(sFileName) = 0 Then
        sFileName = GetConfigFilePath()
    End If
    
    sFileName = GetFilePath(sFileName)

    If (StrComp(sFileName, App.Path & "\access.ini", vbTextCompare) = 0) Then
        If (StrComp(sSection, "Flags", vbTextCompare) = 0) Then
            GetConfigEntry = OpenCommand(sEntryName).RequiredFlags
        ElseIf (StrComp(sSection, "Numeric", vbTextCompare) = 0) Then
            GetConfigEntry = OpenCommand(sEntryName).RequiredRank
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
    
    If LenB(sFileName) = 0 Then
        sFileName = GetConfigFilePath()
    End If
    
    sFileName = GetFilePath(sFileName)
    
    If (StrComp(sFileName, App.Path & "\access.ini", vbTextCompare) = 0) Then
        Dim cmdObj As Object
        
        Set cmdObj = OpenCommand(sEntryName)
    
        If (StrComp(sSection, "Flags", vbTextCompare) = 0) Then
            cmdObj.RequiredFlags = sValue
        ElseIf (StrComp(sSection, "Numeric", vbTextCompare) = 0) Then
            cmdObj.RequiredRank = sValue
        End If
        
        cmdObj.Save
    
        Exit Sub
    End If

    WriteINI sSection, sEntryName, sValue, sFileName
    
End Sub



'// VETOTHISMESSAGE
'// Used with PressedEnter event to prevent a message from being sent to Battle.net
'// For use processing scripts entirely internally
Public Sub VetoThisMessage()

    SetVeto True
    
End Sub



'// PRINTURLTOFILE
'// Mirror function for the Windows API URLDownloadToFile() function
'// Currently you are restricted to placing files in the StealthBot install directory only
Public Sub PrintURLToFile(ByVal sFileName As String, ByVal sURL As String)
    
    sFileName = App.Path & "\" & sFileName
    
    URLDownloadToFile 0, sURL, sFileName, 0, 0
    
End Sub



'// DELETEURLCACHE
'// Mirror function for the Windows API DeleteUrlCacheEntry() function
'// Call before using PrintURLToFile() to clear any residual IE cache entries for
'//     the URL you're retrieving
Public Sub DeleteURLCache(ByVal sURL As String)
    
    DeleteURLCacheEntry sURL
    
End Sub



'// PADQUEUECOUNTER
'// Pads the queue so further messages will be sent more slowly
Public Sub PadQueueCounter()
    
    QueueLoad = QueueLoad + 1

End Sub



'// PADQUEUE
'// Inserts a blank message into the queue
Public Sub PadQueue()
    
    InsertDummyQueueEntry
    
End Sub



'// GETQUEUESIZE
'// Returns the number of items currently in the outgoing message queue
Public Function GetQueueSize() As Integer
    
    GetQueueSize = g_Queue.Count

End Function



'// FLASHBOTWINDOW
'// Flashes the bot's entry in the taskbar to get attention.
Public Sub FlashBotWindow()

    Call FlashWindow
    
End Sub



'// RELOADSCRIPT
'// Reloads the base script.txt file with any includes, equivalent to choosing
'//  that menu option on the bot's Settings menu
'// The command must wait before reloading the script so all operations are cleared.
Public Sub ReloadScript()

    SCReloadTimerID = SetTimer(frmChat.hWnd, 0, 400, AddressOf ScriptReload_TimerProc)
    
End Sub



'// SETSCTIMEOUT
'//   Recommended to modify the timeout setting in your settings.ini file rather than using this sub directly.
Public Sub SetSCTimeout(ByVal newValue As Long)

    If (newValue > 1 And newValue < 60001) Then
        SCReloadTimerID = SetTimer(frmChat.hWnd, newValue, 400, AddressOf ScriptReload_TimerProc)
    End If
End Sub



'// GETSCRIPTCONTROL
'// Returns the Script Control as an object
Public Function GetScriptControl() As Object

    Set GetScriptControl = frmChat.SControl
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
'//  By default, also empties the whisper window; pass an argument of TRUE to the
'//   "DoNotClearWhispers" parameter and it will skip that behavior
Public Function ClearScreen(Optional ByVal DoNotClearWhispers As Boolean) As String
    Call frmChat.mnuClear_Click
    
    If Not DoNotClearWhispers Then
        Call frmChat.mnuClearWW_Click
    End If
End Function



'// CDEC
'// Typecasts a VBS variant to the vbDecimal datatype
'//  By request from Imhotep[Nu]
Public Sub C_Dec(ByRef vToCast As Variant)
    vToCast = CDec(vToCast)
End Sub



'// REQUESTPROFILEKEY
'//  Requests a specific user's profile key. Use with care as Blizzard will
'//  ip-ban you for requesting some keys.
'//  The result will come back to you in an Event_KeyReturn.
Public Sub RequestProfileKey(ByVal sUsername As String, ByVal sKey As String)
    SuppressProfileOutput = True

    RequestSpecificKey sUsername, sKey
End Sub


'// GETAPPHINSTANCE
'//     Returns the App.hInstance value
Public Function GetApphInstance() As Long
    GetApphInstance = App.hInstance
End Function



'// DOSTATSTRINGPARSE
'//     Parses a statstring given to you by GetInternalData (or elsewhere)
'//     Returns the parsed user-message string you see in join/leave messages
'//     You will need to provide the clan tag separately by another GetInternalData() request
'//         if you wish to use it in parsing
Public Function DoStatstringParse(ByVal sStatstring As String, ByVal sClanTag As String) As String
    'Dim sBuffer As String
    Dim UserStats As clsUserStats
    
    'Call ParseStatstring(sStatstring, sBuffer, sClanTag)
    
    ' ...
    Set UserStats = New clsUserStats
    
    ' ...
    UserStats.Statstring = sStatstring
    
    ' ...
    DoStatstringParse = UserStats.ToString
    
    ' ...
    Set UserStats = Nothing
End Function


'// GETWINCURSORPOS
'//     Mirror function for the Windows API function GetCursorPos()
Public Sub GetWinCursorPos(ByRef lCursorX As Long, ByRef lCursorY As Long)
    Dim PAPI As POINTAPI
    
    GetCursorPos PAPI
    
    lCursorX = PAPI.X
    lCursorY = PAPI.Y
End Sub


'// SETWINCURSORPOS
'//     Mirror function for the Windows API function SetCursorPos()
Public Sub SetWinCursorPos(ByVal lNewX As Long, ByVal lNewY As Long)
    SetCursorPos lNewX, lNewY
End Sub


'// WINFINDWINDOW
'//     Mirror function for the Windows API function FindWindow()
Public Function WinFindWindow(ByVal lpClassName As Long, ByVal lpWindowName As String) As Long
    WinFindWindow = FindWindow(lpClassName, lpWindowName)
End Function


'// WINFINDWINDOWEX
'//     Mirror function for the Windows API function FindWindowEx()
Public Function WinFindWindowEx(ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String)
    WinFindWindowEx = FindWindowEx(hWnd1, hWnd2, lpsz1, lpsz2)
End Function


'// WINSENDMESSAGE
'//     Mirror function for the Windows API function
Public Function WinSendMessage(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    WinSendMessage = SendMessage(hWnd, wMsg, wParam, lParam)
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
Public Sub Dsp(ByVal intID As Integer, ByVal strMessage As String, Optional ByVal strUsername As String, Optional ByVal lngColor As Long = 16777215)

    Select Case intID
        Case 1
            Me.AddQ strMessage
        Case 2
            Me.AddQ "/me " & strMessage
        Case 3
            If Len(strUsername) = 0 Then
                AddChat vbRed, "Dsp whisper error: You did not supply a username."
                Exit Sub
            End If
            Me.AddQ "/w " & strUsername & " " & strMessage
        Case 4
            AddChat lngColor, strMessage
    End Select
End Sub


'// RELOADPLUGINMENUS
'//   Reloads the all plugin menus. All user-added items will be cleared.
Public Sub ReloadPluginMenus()
    Call DeletePluginMenus
    Call RegisterPluginMenus
End Sub


'// REMOVEPLUGINMENU
'//   Removes an entire plugin menu.
Public Sub RemoveMenuItem(ByVal strPrefix As String, Optional ByVal strName As String)

    If frmChat.SControl.Modules(1).CodeObject.GetSetting("ps", "menusDisabled") Then Exit Sub
    
    DeleteMenuItemByID GetPluginMenu(strPrefix)
End Sub


'// ADDMENUITEM
'//   Written by Swent. Allows users to add custom items to a plugin's menu.
'//   Syntax: AddMenuItem "prefix", "item name", [boolSeparator], [boolDisabled], [Checkmark]
'//     If boolSeparator is True, this item is a separator and all other arguments are ignored.
'//     If boolDisabled is True, this item is greyed out and can't be clicked.
'//     If boolCheckmark is True, this item has a checkmark
'//   Add the following sub to your plugin which will be called when any user-added item is clicked:
'//     Sub <prefix>_Menu_Callback(ItemName)
Public Sub AddMenuItem(ByVal strPrefix As String, ByVal strName As String, Optional ByVal boolSeparator As Boolean, _
                Optional ByVal boolDisabled As Boolean, Optional ByVal boolCheckmark As Boolean)

    If frmChat.SControl.Modules(1).CodeObject.GetSetting("ps", "menusDisabled") Then Exit Sub

    Dim strCallback As String
    Dim lngMenu As Long, lngItem As Long, intItemCount As Integer
    
    lngMenu = GetPluginMenu(strPrefix)
    intItemCount = GetMenuItemCount(lngMenu)

    '// Insert a separator and warning if this is the first user-added item
    If intItemCount < 7 Then
        AddScriptMenuItem lngMenu, 0, 0, 1
        'AddScriptMenuItem lngMenu, "(use at your own risk)", 0, , 1
    End If

    '// Create the initial callback sub
    strCallback = strPrefix & "_item" & intItemCount + 1 & "_callback"
    frmChat.SControl.Modules(1).AddCode "Sub " & strCallback & ":" & strPrefix & "_Menu_Callback """ & strName & """:End " & "Sub"

    '// Add the menu item
    lngItem = AddScriptMenuItem(lngMenu, strName, strCallback, boolSeparator, boolDisabled, boolCheckmark)
    
    RegisterPluginItem strPrefix, strName, lngItem
End Sub


'// SETMENUITEMCHECK
'//   Changes checked status of an item in a plugin menu.
Public Sub SetMenuItemCheck(ByVal strPrefix As String, ByVal strName As String, ByVal boolChecked As Boolean)

    If frmChat.SControl.Modules(1).CodeObject.GetSetting("ps", "menusDisabled") Then Exit Sub
    
    SetMenuCheck GetPluginMenu(strPrefix), GetPluginItem(strPrefix, strName), boolChecked
End Sub


'// GETMENUITEMCHECK
'//   Get's the checked status of an item in a plugin menu.
'//   Returns True (checked), False (unchecked), or -1 if it's not a checkmark item
Public Function GetMenuItemCheck(ByVal strPrefix As String, ByVal strName As String)
    If frmChat.SControl.Modules(1).CodeObject.GetSetting("ps", "menusDisabled") Then Exit Function
    
    Dim intChange As Integer
    intChange = SetMenuCheck(GetPluginMenu(strPrefix), GetPluginItem(strPrefix, strName), True)
    
    If intChange = 0 Then
        SetMenuCheck GetPluginMenu(strPrefix), GetPluginItem(strPrefix, strName), False
        GetMenuItemCheck = False
    ElseIf intChange = 1 Then
        GetMenuItemCheck = True
    Else
        GetMenuItemCheck = -1
    End If
End Function


'// TOGGLEMENUITEMCHECK
'//   Reverses the checked status of a plugin menu item.
Public Sub ToggleMenuItemCheck(ByVal strPrefix As String, ByVal strName As String)

    If frmChat.SControl.Modules(1).CodeObject.GetSetting("ps", "menusDisabled") Then Exit Sub
    
    SetMenuItemCheck strPrefix, strName, Not GetMenuItemCheck(strPrefix, strName)
End Sub


'// SETMENUITEMENABLED
'//   Enables or disables a plugin menu item. Disabled items are greyed out and not clickable.
Public Sub SetMenuItemEnabled(ByVal strPrefix As String, ByVal strName As String, ByVal boolEnabled As Boolean)

    If frmChat.SControl.Modules(1).CodeObject.GetSetting("ps", "menusDisabled") Then Exit Sub
    
    SetMenuEnabled GetPluginMenu(strPrefix), GetPluginItem(strPrefix, strName), boolEnabled
End Sub


'// CREATEFORM
'//    Written by Hdx. Creates a Form object that the user can use.
'//    Returns true if the name/prefix combo is free, and the form is created
'Public Function CreateForm(ByVal strPrefix As String, ByVal strName As String) As Boolean
'    If (obScriptForms.Exists(strPrefix & "|" & strName)) Then
'        CreateForm = False
'        Exit Function
'    Else
'        Dim frmUI As New frmScriptUI
'        frmUI.FillPrefixName strPrefix, strName
'        Load frmUI
'        Call obScriptForms.Add(strPrefix & "|" & strName, frmUI)
'        DictionaryKeys = DictionaryKeys & strPrefix & "|" & strName & vbCrLf
'        CreateForm = True
'    End If
'End Function

'// DESTROYFORM
'//   Writtin by Hdx - Will destroy a ScruptUI form, freeing up its name, and memory from the Dictionary
'//   The Form calls this function upon successful "Unload()", Formes must be re-created after this is called.
'Public Sub DestroyForm(ByVal strPrefix As String, ByVal strForm As String)
'    If (obScriptForms.Exists(strPrefix & "|" & strForm)) Then
'        Dim frmUI As frmScriptUI
'        Set frmUI = obScriptForms.Item(strPrefix & "|" & strForm)
'        frmUI.DestroyObjects
'        Unload frmUI
'        Set frmUI = Nothing
'    End If
'End Sub

'// ADDFORMOBJECT
'//    Writtin by Hdx - Adds an object of the selected type to a Script's form.
'//    Returns False if the form does not exits, or the name is already taken
'Public Function AddFormObject(ByVal strPrefix As String, ByVal strFormName As String, ByVal strName As String, ByVal btType As ControlTypes) As Boolean
'    On Error Resume Next
'    If (obScriptForms.Exists(strPrefix & "|" & strFormName)) Then
'        Dim frmUI As frmScriptUI
'        Set frmUI = obScriptForms.Item(strPrefix & "|" & strFormName)
'        Select Case btType
'            Case ControlTypes.btCommandBox:  AddFormObject = frmUI.AddCommandButton(strName)
'            Case ControlTypes.btLabel:       AddFormObject = frmUI.AddLabel(strName)
'            Case ControlTypes.btTextBox:     AddFormObject = frmUI.AddTextBox(strName)
'            Case ControlTypes.btTimer:       AddFormObject = frmUI.AddTimer(strName)
'            Case ControlTypes.btPictureBox:  AddFormObject = frmUI.AddPictureBox(strName)
'            Case ControlTypes.btCheckBox:    AddFormObject = frmUI.AddCheckBox(strName)
'            Case ControlTypes.btOptionBox:   AddFormObject = frmUI.AddOptionBox(strName)
'            Case ControlTypes.btComboBox:    AddFormObject = frmUI.AddComboBox(strName)
'            Case ControlTypes.btListBox:     AddFormObject = frmUI.AddListBox(strName)
'            Case ControlTypes.btShape:       AddFormObject = frmUI.AddShape(strName)
'            Case ControlTypes.btLine:        AddFormObject = frmUI.AddLine(strName)
'            Case ControlTypes.btListView:    AddFormObject = frmUI.AddListView(strName)
'            Case ControlTypes.btImageList:   AddFormObject = frmUI.AddImageList(strName)
'            Case ControlTypes.btINet:        AddFormObject = frmUI.AddINet(strName)
'            Case ControlTypes.btRichTextBox: AddFormObject = frmUI.AddRichTextBox(strName)
'            Case Else: AddFormObject = False
'        End Select
'    Else
'        AddFormObject = False
'    End If
'End Function

'// GETUIOBJECT
'//    Written by Hdx - Retrns the control/form the the specified name. Returns Null if not found.
'Public Function GetUIObject(ByVal strPrefix As String, ByVal strForm As String, Optional ByVal strObject As String = vbNullString) As Object
'    If (obScriptForms.Exists(strPrefix & "|" & strForm)) Then
'        Dim frmUI As frmScriptUI
'        Set frmUI = obScriptForms.Item(strPrefix & "|" & strForm)
'        If (strObject = vbNullString) Then
'            Set GetUIObject = frmUI
'        Else
'            Set GetUIObject = frmUI.GetControl(strObject)
'        End If
'        Set frmUI = Nothing
'    Else
'        Set GetUIObject = Nothing
'    End If
'End Function

' // ...
Public Function IsCommand(Optional ByVal str As String = vbNullString, Optional IsLocal As Boolean = False) _
    As Object
    
    Set IsCommand = modOtherCode.IsCommand(str, IsLocal)
End Function

Public Function OpenCommand(ByVal strCommand As String) As Object
    '// ...
    Set OpenCommand = modCommandDocsObj.OpenCommand(strCommand)
End Function

Public Function CreateCommand(ByVal strCommand As String) As Object
    '// ...
    Set CreateCommand = modCommandDocsObj.CreateCommand(strCommand)
End Function

Public Function IsCorrectSyntax(ByVal CommandName As String, ByVal CommandArgs As String) As Boolean
    '// ...
    IsCorrectSyntax = modCommandCode.IsCorrectSyntax(CommandName, CommandArgs)
End Function

Public Function HasAccess(ByVal Username As String, ByVal CommandName As String, Optional ByVal CommandArgs As _
    String = vbNullString, Optional ByRef outbuf As String) As Boolean
    
    HasAccess = modCommandCode.HasAccess(Username, CommandName, CommandArgs, outbuf)
End Function

Public Function UtcNow() As Date
    UtcNow = modDateTime.UtcNow
End Function

Public Function Logger() As Object

    Set Logger = g_Logger

End Function

Public Function isDebug() As Boolean

    isDebug = modGlobals.isDebug()
    
End Function

Public Function FileTimeToDate(ByRef strTime As String) As Date

    Dim FTime As FILETIME ' ...
    
    ' ...
    With FTime
        .dwHighDateTime = _
            CLng(Left$(strTime, InStr(1, strTime, " ", vbBinaryCompare)))

        .dwLowDateTime = _
            CLng(Mid$(strTime, InStr(1, strTime, " ", vbBinaryCompare) + 1))
    End With
    
    ' ...
    FileTimeToDate = modDateTime.FileTimeToDate(FTime)
    
End Function

Public Function Quotes() As Object

    Dim I As Integer ' ...
    
    ' ...
    Set Quotes = New Collection
    
    ' ...
    For I = 1 To g_Quotes.Count
        Quotes.Add g_Quotes(I)
    Next I
    
End Function

' // ...
Public Function MediaPlayer() As Object

    Set MediaPlayer = modMediaPlayer.MediaPlayer()
    
End Function

Public Function Channel() As Object

    Set Channel = g_Channel.Clone()
    
End Function

Public Function Clan() As Object

    Set Clan = g_Clan.Clone()
    
End Function

Public Function Friends() As Object
    
    Dim I As Integer ' ...

    ' ...
    Set Friends = New Collection
    
    ' ...
    For I = 1 To g_Friends.Count
        Friends.Add g_Friends(I).Clone()
    Next I
    
End Function

Public Function Queue() As Object

    Set Queue = g_Queue

End Function

Public Function OSVersion() As Object

    Set OSVersion = New clsOSVersion

End Function

Public Function GetSystemUptime() As String
    GetSystemUptime = ConvertTime(GetUptimeMS)
End Function

Public Function GetConnectionUptime() As String
    GetConnectionUptime = ConvertTime(uTicks)
End Function

Public Function CRC32(ByVal str As String) As String

    Dim clsCRC32 As New clsCRC32
    
    CRC32 = clsCRC32.CRC32(str)
    
    Set clsCRC32 = Nothing

End Function

Public Function Sha1(ByVal data As String) As String
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim e As Long
    
    Call modSHA1.DefaultSHA1(StrConv(data, vbFromUnicode), a, b, c, d, e)
    Sha1 = LongToStr(a) & LongToStr(b) & LongToStr(c) & LongToStr(d) & LongToStr(e)
End Function

Public Function CreateObjEx(ByVal ModuleName As String, ByVal ObjType As String, ByVal ObjName As String) As Object

    Set CreateObjEx = _
            modScripting.CreateObjEx(frmChat.SControl.Modules(ModuleName), _
                    ObjType, ObjName)

End Function

Public Function DestroyObjEx(ByVal ModuleName As String, ByVal ObjName As String) As Object

    modScripting.DestroyObjEx frmChat.SControl.Modules(ModuleName), ObjName

End Function

Public Function GetObjByNameEx(ByVal ModuleName As String, ByVal ObjName As String) As Object

    Set GetObjByNameEx = _
            modScripting.GetObjByNameEx(frmChat.SControl.Modules(ModuleName), _
                    ObjName)

End Function

Public Function Scripts() As Object

    On Error Resume Next

    Dim I   As Integer ' ...
    Dim str As String  ' ...

    Set Scripts = New Collection

    For I = 1 To frmChat.SControl.Modules.Count
        str = _
            frmChat.SControl.Modules(I).CodeObject.GetSettingsEntry("Public")
                
        If (StrComp(str, "False", vbTextCompare) <> 0) Then
            Scripts.Add frmChat.SControl.Modules(I).CodeObject, _
                frmChat.SControl.Modules(I).CodeObject.Script("Name")
        End If
    Next I

End Function

Public Function StrConvEx(ByVal str As String, ByVal Conv As VbStrConv, Optional ByVal locale As Long) As Variant

    StrConvEx = StrConv(str, Conv, locale)

End Function

Public Function DataBufferEx() As Object

    Set DataBufferEx = New clsPacketBuffer
    
    DataBufferEx.setCripple

End Function

Public Function GetScriptObjByName(ByVal scriptName As String) As Object

    On Error Resume Next

    Dim I   As Integer ' ...
    Dim str As String ' ...
    
    For I = 1 To frmChat.SControl.Modules.Count
        str = frmChat.SControl.Modules(I).Eval("Name()")
        
        If (StrComp(scriptName, str, vbTextCompare) = 0) Then
            str = _
                frmChat.SControl.Modules(I).CodeObject.GetSettingsEntry("Public")
            
            If (StrComp(str, "False", vbTextCompare) = 0) Then
                Exit Function
            End If
        
            Set GetScriptObjByName = _
                    frmChat.SControl.Modules(I)
             
            Exit Function
        End If
    Next I
    
    For I = 1 To frmChat.SControl.Modules.Count
        str = frmChat.SControl.Modules(I).Name
    
        If (StrComp(scriptName, str, vbTextCompare) = 0) Then
            str = _
                frmChat.SControl.Modules(I).CodeObject.GetSettingsEntry("Public")
            
            If (StrComp(str, "False", vbTextCompare) = 0) Then
                Exit Function
            End If
        
             Set GetScriptObjByName = _
                    frmChat.SControl.Modules(I)
             
             Exit Function
        End If
    Next I

End Function

Public Function GetSettingsEntryEx(ByVal scriptName As String, ByVal sEntryName As String)
    
    Dim Path As String ' ...
    
    ' ...
    Path = App.Path & "\scripts\scripts.ini"

    ' ...
    If (LenB(Dir(Path)) = 0) Then
        Open Path For Output As #1
        Close #1
    End If
    
    ' ...
    GetSettingsEntryEx = ReadINI(scriptName, sEntryName, Path)
    
    ' ...
    If (InStr(1, GetSettingsEntryEx, " ;") <> 0) Then
        GetSettingsEntryEx = _
            Mid$(GetSettingsEntryEx, 1, (InStr(1, GetSettingsEntryEx, " ;") - 1))
    End If
    
    ' ...
    If (InStr(1, GetSettingsEntryEx, " #") <> 0) Then
        GetSettingsEntryEx = _
            Mid$(GetSettingsEntryEx, 1, (InStr(1, GetSettingsEntryEx, " #") - 1))
    End If

End Function

Public Function WriteSettingsEntryEx(ByVal scriptName As String, ByVal sEntryName As String, ByVal sValue As String, Optional sDescription As String)

    Dim Path As String ' ...
    
    ' ...
    Path = App.Path & "\scripts\scripts.ini"

    ' ...
    If (LenB(Dir(Path)) = 0) Then
        Open Path For Output As #1
        Close #1
    End If
    
    ' ...
    WriteINI scriptName, sEntryName, sValue & _
        IIf(Len(sDescription), " ; " & sDescription, ""), Path

End Function


'Class_Terminate doesn't run when a UI form is created via a script. I made this sub so the bot's Unload sub can explicitly _
    call the code within Class_Terminate to kill all the objects. - FrOzeN
Public Sub Dispose()
    Call Class_Terminate
End Sub

Private Sub Class_Initialize()
    'Set obScriptForms = CreateObject("Scripting.Dictionary")
End Sub

'Re-coded this sub so it properly kills all the UI forms created that are still open. - FrOzeN
Private Sub Class_Terminate()

    'Static Disposed As Boolean
    '
    'If Not Disposed Then
    '    Dim I As Integer, strParts() As String, frmTemp As frmScriptUI
    '    strParts() = Split(DictionaryKeys, vbCrLf)
    '
    '    For I = 0 To UBound(strParts) - 1
    '        If obScriptForms.Exists(strParts(I)) Then
    '            Set frmTemp = obScriptForms.Item(strParts(I))
    '            frmTemp.DestroyObjects
    '            Unload frmTemp
    '            Set frmTemp = Nothing
    '        End If
    '    Next I
    '
    '    obScriptForms.RemoveAll
    '    Set obScriptForms = Nothing
    '
    '    Disposed = True
    'End If

    frmChat.SControl.Reset
    'frmChat.SCRestricted.Reset
    
'    'Check objects exist before attempting to unload them - FrOzeN
'    If obScriptForms.Count <> 0 Then
'
'        'Go through and destroy all the Forms the user has created
'        Dim X As Integer
'
'        For X = 0 To obScriptForms.Count
'            Unload obScriptForms.Item(X)
'            Call obScriptForms.Remove(X)
'        Next X
'
'        Set obScriptForms = Nothing
'    End If
End Sub

'// SETSETTING
'//   Written by Swent. Creates/modifies plugin settings.
'//   Usage - new setting: SetSetting "prefix", "settingName", value, "setting description", False
'//   Usage - existing setting: SetSetting "prefix", "settingName", value
'//   Setting Values() of any Length are now set and retrieved properly (no longer a 255 char max)
'//   You can now have spaces in your setting names.
Public Sub SetSetting(ByVal strPrefix As String, ByVal strName As String, ByVal varValue As Variant, _
                      Optional ByVal strDescription As String, Optional ByVal boolOverwrite As Boolean = True)
                      
    Dim strPath As String, arrSetting() As String, I As Integer
    strPath = BotPath() & "plugins\settings.ini"

    If Len(strDescription) > 0 Then
        SetSetting strPrefix, strName & " (info)", strDescription, "", True
    End If
    
    If boolOverwrite Or Len(GetConfigEntry(strPrefix, strName, strPath)) = 0 Then
        If Len(strName & varValue) > 254 Then
            WriteConfigEntry UCase(strPrefix), strName, Left(varValue, 254 - Len(strName)), strPath
            arrSetting = Split(strName, "*")
            If UBound(arrSetting) > 0 Then I = arrSetting(1) + 1 Else I = 1
            SetSetting strPrefix, arrSetting(0) & "*" & I, Mid(varValue, 255 - Len(strName)), strDescription, boolOverwrite
        Else
            WriteConfigEntry UCase(strPrefix), strName, varValue, strPath
        End If
    End If
End Sub
