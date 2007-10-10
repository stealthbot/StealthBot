VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScriptSupportClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*
' * StealthBot Shared Scripting Class
' *
' * ScriptSupportClass.cls
' *
' *     Last modified 8:45 PM 10/7/2007 by Swent
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
' *     - Added RegisterScriptMenu() and AddScriptMenuItem() [see end of changelog] which allow
' *         scripts to have menus inside StealthBot. See those functions' comments for more
' *         information along with their support functions (thanks FiftyToo)
' *     - Added GetApphInstance() which gives you the instance handle to StealthBot (thanks FrostWraith)
' *     - Added DoStatstringParse() which allows you to parse a statstring from GetInternalData() out
' *         just like the bot does (thanks ZergMasterI)
' *     - Added a number of Windows API function mirrors (thanks FiftyToo)
' *     - Added the Dsp() function (see notes below) - Swent
' *
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

Public myChannel As String  '// will contain the bot's current channel at runtime.
Public BotFlags As Long     '// will contain the bot's current battle.net flags at runtime.
Public myUsername As String '// will contain the bot's current username at runtime.
                            '// NOTE: This may be different than the bot's config.ini username
                            
Public strSettingNotes      '// Used to display notes on your plugin settings

'// myTrigger has been replaced by BotVars.Trigger
'Public myTrigger As String '// contains the bot's current trigger at runtime
         
Public Enum BanTypes
    btBan = 0     '// used in calling the BanKickUnban() subroutine
    btKick = 1
    btUnban = 2
End Enum

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
    On Error Resume Next
    Dim s As String
    Dim l As Long, lngVerticalPos As Long
    Dim i As Integer, intRange As Integer, f As Integer
    Dim blUnlock As Boolean, LogThis As Boolean
    
    If Not BotVars.LockChat Then
    
        If IsWin2000Plus() Then
            intRange = 40
            
            GetScrollRange frmChat.rtbChat.hWnd, SB_VERT, 0, intRange
            lngVerticalPos = SendMessage(frmChat.rtbChat.hWnd, EM_GETTHUMB, 0&, 0&)
            
            'Debug.Print "ScrollRange: " & intRange & " ; VerticalPos: " & lngVerticalPos & " ; frmchat.rtbchatHeight " & frmchat.rtbchat.Height & " ; pix " & frmchat.rtbchat.Height / Screen.TwipsPerPixelY
            
            If (lngVerticalPos + (frmChat.rtbChat.Height / Screen.TwipsPerPixelY)) <= intRange Then
                'LockWindowUpdate frmchat.rtbchat.hWnd
                frmChat.rtbChat.Visible = False
                blUnlock = True
            End If
        End If
        
        LogThis = (BotVars.Logging < 2)
        
        If rtbChatLength > 10000 Then
            With frmChat.rtbChat
                .Visible = False
                .SelStart = 0
                .SelLength = InStr(1, .text, vbLf, vbBinaryCompare)
                
                rtbChatLength = rtbChatLength - .SelLength
                
'                If BotVars.Logging < 2 And LOF(i) < BotVars.MaxLogfileSize Then
'                    i = FreeFile
'                    Open (GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & ".txt") For Append As #i
'                        Print #i, Left$(.SelText, Len(.SelText) - 2)
'                    Close #i
'                End If
                
                .SelText = ""
                .Visible = True
            End With
        End If
        
        s = GetTimestamp
        
        With frmChat.rtbChat
            .SelStart = Len(.text)
            .SelLength = 0
            .SelColor = RTBColors.TimeStamps
            If .SelBold = True Then .SelBold = False
            If .SelItalic = True Then .SelItalic = False
            If .SelUnderline = True Then .SelUnderline = False
            .SelText = s
            .SelStart = Len(.text)
        End With
        
        If LogThis Then
            f = FreeFile
            
            Open (GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & ".txt") For Append As #f
            
            If (LOF(f) >= BotVars.MaxLogfileSize) Then
                LogThis = False
                Close #f
            End If
        End If
        
        For i = LBound(saElements) To UBound(saElements) Step 2
        
            If InStr(1, saElements(i + 1), Chr(0), vbBinaryCompare) > 0 Then _
                KillNull saElements(i + 1)
            
            If Len(saElements(i + 1)) > 0 Then
                l = InStr(1, saElements(i + 1), "{\rtf", vbTextCompare)
                
                While (l > 0)
                    Mid$(saElements(i + 1), l + 1, 1) = "/"
                    
                    l = InStr(1, saElements(i + 1), "{\rtf", vbTextCompare)
                Wend
            
                With frmChat.rtbChat
                    .SelStart = Len(.text)
                    l = .SelStart
                    .SelLength = 0
                    .SelColor = saElements(i)
                    .SelText = saElements(i + 1) & Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)))
                    
                    rtbChatLength = rtbChatLength + Len(s) + Len(saElements(i + 1)) + Len(Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements))))
                    
                    .SelStart = Len(.text)
                End With
                
                ' Fixed 11/21/06 to properly log timestamps
                If LogThis Then
                    Print #f, s & saElements(i + 1) & Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)));
                End If
            End If
            
        Next i
        
        Call ColorModify(frmChat.rtbChat, l)
    
        If LogThis Then
            Close #f
            LogThis = False
        End If

        If blUnlock Then
            SendMessage frmChat.rtbChat.hWnd, WM_VSCROLL, SB_THUMBPOSITION + &H10000 * lngVerticalPos, 0&
            'LockWindowUpdate 0&
            frmChat.rtbChat.Visible = True
        End If

    End If
End Sub


'// ADDCHATFONT
'// Same as the function above, but allows for you to specify the FONT used in
'//     each message.
'// Format: AddChatFont(Font, Color, Text)
'// Extensible as far as you need:
'// AddChat(Font, Color, Text, Font, Color, Text, Font, Color, Text)
'//     -- will all display on one line.
'// For example:
'//     AddChatFont "Courier New", vbRed, "Hello, world!"
'// will display that phrase in the Courier New font and red color.
Public Sub AddChatFont(ParamArray saElements() As Variant)
    On Error Resume Next
    Dim s As String
    Dim l As Long, lngVerticalPos As Long
    Dim i As Integer, intRange As Integer, f As Integer
    Dim blUnlock As Boolean, LogThis As Boolean
    
    If Not BotVars.LockChat Then
    
        If IsWin2000Plus() Then
            GetScrollRange frmChat.rtbChat.hWnd, SB_VERT, 0, intRange
            lngVerticalPos = SendMessage(frmChat.rtbChat.hWnd, EM_GETTHUMB, 0&, 0&)
            
            'Debug.Print "ScrollRange: " & intRange & " ; VerticalPos: " & lngVerticalPos & " ; frmchat.rtbchatHeight " & frmchat.rtbchat.Height & " ; pix " & frmchat.rtbchat.Height / Screen.TwipsPerPixelY
            
            If (lngVerticalPos + (frmChat.rtbChat.Height / Screen.TwipsPerPixelY)) <= intRange Then
                'LockWindowUpdate frmchat.rtbchat.hWnd
                frmChat.rtbChat.Visible = False
                blUnlock = True
            End If
        End If
        
        LogThis = (BotVars.Logging < 2)
        
        If rtbChatLength > 10000 Then
            With frmChat.rtbChat
                .Visible = False
                .SelStart = 0
                .SelLength = InStr(1, .text, vbLf, vbBinaryCompare)
                
                rtbChatLength = rtbChatLength - .SelLength
                
'                If BotVars.Logging < 2 And LOF(i) < BotVars.MaxLogfileSize Then
'                    i = FreeFile
'                    Open (GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & ".txt") For Append As #i
'                        Print #i, Left$(.SelText, Len(.SelText) - 2)
'                    Close #i
'                End If
                
                .SelText = ""
                .Visible = True
            End With
        End If
        
        s = GetTimestamp
        
        With frmChat.rtbChat
            .SelStart = Len(.text)
            .SelLength = 0
            .SelColor = RTBColors.TimeStamps
            If .SelBold = True Then .SelBold = False
            If .SelItalic = True Then .SelItalic = False
            If .SelUnderline = True Then .SelUnderline = False
            .SelText = s
            .SelStart = Len(.text)
        End With
        
        If LogThis Then
            f = FreeFile
            
            Open (GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & ".txt") For Append As #f
            
            If (LOF(f) >= BotVars.MaxLogfileSize) Then
                LogThis = False
                Close #f
            End If
        End If
        
        For i = LBound(saElements) To UBound(saElements) Step 3
        
            If InStr(1, saElements(i + 1), Chr(0), vbBinaryCompare) > 0 Then _
                KillNull saElements(i + 1)
            
            If InStr(1, saElements(i + 2), Chr(0), vbBinaryCompare) > 0 Then _
                KillNull saElements(i + 2)
            
            If Len(saElements(i + 2)) > 0 Then
                l = InStr(1, saElements(i + 1), "{\rtf", vbTextCompare)
                
                While (l > 0)
                    Mid$(saElements(i + 2), l + 2, 1) = "/"
                    
                    l = InStr(1, saElements(i + 2), "{\rtf", vbTextCompare)
                Wend
            
                With frmChat.rtbChat
                    .SelStart = Len(.text)
                    l = .SelStart
                    .SelLength = 0
                    .SelFontName = saElements(i)
                    .SelColor = saElements(i + 1)
                    .SelText = saElements(i + 2) & Left$(vbCrLf, -2 * CLng((i + 2) = UBound(saElements)))
                    
                    rtbChatLength = rtbChatLength + Len(s) + Len(saElements(i + 2)) + Len(Left$(vbCrLf, -2 * CLng((i + 2) = UBound(saElements))))
                    
                    .SelStart = Len(.text)
                End With
                
                ' Fixed 11/21/06 to properly log timestamps
                If LogThis Then
                    Print #f, s & saElements(i + 2) & Left$(vbCrLf, -2 * CLng((i + 2) = UBound(saElements)));
                End If
            End If
            
        Next i
        
        Call ColorModify(frmChat.rtbChat, l)
    
        If LogThis Then
            Close #f
            LogThis = False
        End If

        If blUnlock Then
            SendMessage frmChat.rtbChat.hWnd, WM_VSCROLL, SB_THUMBPOSITION + &H10000 * lngVerticalPos, 0&
            'LockWindowUpdate 0&
            frmChat.rtbChat.Visible = True
        End If

    End If
End Sub


'// ADDQ (ADD QUEUE)
'// Adds a string to the message send queue.
'// Nonzero priority messages will be sent with precedence over 0-priority messages.
Public Sub AddQ(ByVal sText As String, Optional ByVal Priority As Byte = 0)

    frmChat.AddQ sText, Priority
    
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
Public Function Command(ByVal Username As String, ByVal Message As String, ByVal InBot As Boolean) As Boolean
    Dim dbAccess As udtGetAccessResponse
    '// GetAccess() will return the access type the commands sub needs
    
    ' execute command
    Command = ProcessCommand(Username, Message, InBot, False)
    
    ' backwards compatible (may be removed in future!)
    If (Command = False) Then
        Dim tmpMsg As String ' ...
        
        ' store local copy of message
        tmpMsg = Message
    
        ' manipulate message for deprecated
        ' calling format
        If (InBot) Then
            tmpMsg = "/" & tmpMsg
        Else
            tmpMsg = BotVars.Trigger & tmpMsg
        End If
        
        ' execute command
        Command = ProcessCommand(Username, Message, InBot, False)
    End If
End Function


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

Public Function CommandEx(ByVal Username As String, ByVal Message As String, ByVal InBot As Boolean, ByVal uAccess As Integer, ByVal uFlags As String, ByRef cmdRet() As String) As Boolean
    Dim dbAccess As udtGetAccessResponse
    '// GetAccess() will return the access type the commands sub needs
    
    dbAccess.Access = uAccess
    dbAccess.Flags = uFlags

    ' execute command
    CommandEx = ExecuteCommand(Username, dbAccess, Message, InBot, cmdRet())
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
Public Sub GetDBEntry(ByVal Username As String, ByRef Access, ByRef Flags) '// yum, Variants >:\

    Dim Temp As udtGetAccessResponse
    
    Temp = GetAccess(Username)
    
    Access = Temp.Access
    Flags = Temp.Flags

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
Public Function Match(ByVal sString As String, ByVal sPattern As String, ByVal DoPreparation As Boolean) As Boolean

    If DoPreparation Then
        sString = PrepareCheck(sString)
        sPattern = PrepareCheck(sPattern)
    End If
    
    Match = (sString Like sPattern)
    
End Function


'// SETBOTPROFILE
'// Sets the bot's current profile to the specified value(s).
'// If passed as null, the values will not be reset, so profile data you are not changing will not be overwritten.
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
        Call RequestProfile(CurrentUsername)
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
    
    GetInternalUserCount = colUsersInChannel.Count
    
End Function


'// GETINTERNALDATABYUSERNAME
'// Retrieves the specified stored internal data for a given user in the channel
'//  If the specified user isn't present, the return value is -5
'//  See lDataType constants in GetInternalData() below
Public Function GetInternalDataByUsername(ByVal sUser As String, ByVal lDataType As Long) As Variant

    Dim i As Integer
    
    i = UsernameToIndex(sUser)
    
    GetInternalDataByUsername = GetInternalData(i, lDataType)

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
                GetInternalData = colUsersInChannel.Item(iIndex).Clan
                
            Case GID_FLAGS
                GetInternalData = colUsersInChannel.Item(iIndex).Flags
            
            Case GID_PING
                GetInternalData = colUsersInChannel.Item(iIndex).Ping
            
            Case GID_PRODUCT
                GetInternalData = colUsersInChannel.Item(iIndex).Product
            
            Case GID_ISSAFELISTED
                GetInternalData = colUsersInChannel.Item(iIndex).Safelisted
            
            Case GID_STATSTRING
                GetInternalData = colUsersInChannel.Item(iIndex).Statstring
            
            Case GID_TIMEINCHANNEL
                GetInternalData = colUsersInChannel.Item(iIndex).TimeInChannel()
            
            Case GID_TIMESINCETALK
                GetInternalData = colUsersInChannel.Item(iIndex).TimeSinceTalk()
            
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
    
    GetPositionByName = CheckChannel(sUser)
    
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
    
    GetBotVersion = "StealthBot v" & App.Major & "." & App.Minor
    
    '// Compiler flag - BETA = 0 in public releases
    #If BETA = 1 Then
        GetBotVersion = GetBotVersion & App.REVISION & " Development Edition"
    #End If
    
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
    
    Call frmChat.Form_Unload(0)
    
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
    
    GetQueueSize = colQueue.Count

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
'// Sets the number of milliseconds before the script control times your script's execution out
'//  Values range from 1 to 60000
Public Sub SetSCTimeout(ByVal newValue As Long)

    If (newValue > 1 And newValue < 60001) Then
        SCReloadTimerID = SetTimer(frmChat.hWnd, newValue, 400, AddressOf ScriptReload_TimerProc)
    End If
    
End Sub



'// GETSCRIPTCONTROL
'// Returns the Script Control as an object
Public Function GetScriptControl() As Variant
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
    RequestSpecificKey sUsername, sKey
End Sub



'// REGISTERSCRIPTMENU
'//     Creates a first-level menu item under the Scripts menu item
'//     Call, then call AddScriptMenuItem to add menu items using the handle
'//       this function returns to you
Public Function RegisterScriptMenu(ByVal sMenuCaption As String) As Long
    
    Dim lMenu As Long
    Dim ThisScript_MenuID As Long
    
    lMenu = GetMenu(frmChat.hWnd)
    
    If ScriptMenu_ParentID = 0 Then
        ScriptMenu_ParentID = AddParentMenu(lMenu, "Plugins", , 5)
        AddMenuItem ScriptMenu_ParentID, "(use at your own risk)", , , True
    End If
    
    ThisScript_MenuID = AddParentMenu(ScriptMenu_ParentID, sMenuCaption)
    
    RegisterScriptMenu = ThisScript_MenuID
        DrawMenuBar frmChat.hWnd
    colDynamicMenus.Add ThisScript_MenuID
    
End Function



'// ADDSCRIPTMENUITEM
'//     Adds a second-level menu item to an already-created menu
'//     ARGUMENTS:
'//     lMenuHandle should be a menu handle returned by RegisterScriptMenu
'//     sItemCaption is the caption of the menu item you are adding
'//     sCallbackFunction is the name of the script.txt method you want called
'//         when someone clicks on this menu item
'//     PARAMETER VALUES
'//         MSeparator      if true, this menu item will be a separator and all other values are ignored (for caption & callback)
'//         MDisabled       if true, this menu item will be disabled
'//         MChecked        if true, this menu item will be checked
'//     RETURNS
'//         Returns the menu item's individual Command ID for your tracking pleasure
Public Function AddScriptMenuItem(ByVal lMenuHandle As Long, ByVal sItemCaption As String, ByVal sCallbackFunction As String, _
        Optional ByVal MSeparator As Boolean, Optional ByVal MDisabled As Boolean, Optional ByVal MChecked As Boolean) As Long
    
    Dim lCallbackID As Long
    lCallbackID = AddMenuItem(lMenuHandle, sItemCaption, MSeparator, MChecked, MDisabled)
    
    dctCallbacks.Add CStr(lCallbackID), sCallbackFunction
    
    DrawMenuBar frmChat.hWnd
    
    AddScriptMenuItem = lCallbackID
    
End Function



'// SETMENUCHECK
'//     Toggles the checkmark on an item.
'//     Returns 1 if the item was previously checked, 0 if it was unchecked, and -1 if the menu item doesn't exist.
Public Function SetMenuCheck(ByVal lMenuHandle As Long, ByVal lMenuCommandID As Long, ByVal bNewCheckState As Boolean) As Long
    
    Dim l As Long
    
    l = CheckMenuItem(lMenuHandle, lMenuCommandID, IIf(bNewCheckState, MF_CHECKED, MF_UNCHECKED))
    
    DrawMenuBar frmChat.hWnd
    
    If (l And MF_CHECKED) = MF_CHECKED Then
        l = 1
    End If
    
    SetMenuCheck = l
    
End Function



'// SETMENUENABLED
'//     Toggles whether or not a menu item is grayed out.
Public Sub SetMenuEnabled(ByVal lMenuHandle As Long, ByVal lMenuCommandID As Long, ByVal bNewEnabledState As Boolean)
    
    Dim l As Long
    Dim s As String
    
    s = GetMenuCaptionByCommand(frmChat.hWnd, lMenuCommandID)
    l = ModifyMenu(lMenuHandle, lMenuCommandID, IIf(bNewEnabledState, MF_STRING, MF_GRAYED), lMenuCommandID, s)
    
    DrawMenuBar frmChat.hWnd
    
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
    Dim sBuffer As String
    
    Call ParseStatstring(sStatstring, sBuffer, sClanTag)
    
    DoStatstringParse = sBuffer
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


'// SETSETTING
'//     Creates and modifies plugin settings.
'//     Overwrite is a boolean (True or False). Set to True if you want it to overwrite an existing value.
'//     Use "" for the strDescription parameter and the existing description will be used.
'//     Setting values of any length can now be set and retrieved properly (no 255 char max)
Sub SetSetting(ByVal strPrefix As String, strName As String, varValue As Variant, strDescription As String, boolOverwrite As Boolean)
    Dim strSettingsPath As String, arrSetting() As String, i As Integer
    strSettingsPath = BotPath() & "plugins\settings.ini"

    If Len(strDescription) > 0 Then
        If InStr(strSettingNotes, LCase(strPrefix) & "|" & strName & "|") = 0 Then
            strSettingNotes = strSettingNotes & LCase(strPrefix) & "|" & strName & "|" & strDescription & "||"
        End If
    End If
    If boolOverwrite Or Len(GetConfigEntry(strPrefix, strName, strSettingsPath)) = 0 Then
        If Len(strName & varValue) > 254 Then
            WriteConfigEntry UCase(strPrefix), strName, Left(varValue, 254 - Len(strName)), strSettingsPath
            arrSetting = Split(strName, "*")
            If UBound(arrSetting) > 0 Then i = arrSetting(1) + 1 Else i = 1
            SetSetting strPrefix, arrSetting(0) & "*" & i, Mid(varValue, 255 - Len(strName)), strDescription, boolOverwrite
        Else
            WriteConfigEntry UCase(strPrefix), strName, varValue, strSettingsPath
        End If
    End If
End Sub


'// GETSETTING
'//   Retrieves plugin settings
Function GetSetting(strPrefix, strName)
    Dim strSettingsPath As String, i As Integer, strAppend As String, tmpValue As String, varValue As Variant
    strSettingsPath = BotPath() & "plugins\settings.ini"
    
    Do
        If i > 0 Then strAppend = "*" & i
        tmpValue = GetConfigEntry(strPrefix, Left(strName, Len(strName)) & strAppend, strSettingsPath)
        varValue = varValue & tmpValue
        If Len(tmpValue) = 0 Then
            GetSetting = varValue
            Exit Do
        End If
        i = i + 1
    Loop
End Function


'// DSP
'// Displays messages via the specified output type and allows for unlimited message lengths.
'//     intID (required) must be one of the following integer values:
'//         1 = AddQ
'//             Usage: Dsp 1, Message, [Delimiter]
'//         2 = Emote
'//             Usage: Dsp 2, Message, [Delimiter]
'//         3 = Whisper
'//             Usage: Dsp 3, Message, Username, [Delimiter]
'//         4 = AddChat
'//             Usage: Dsp 4, Message, Color
'//     Color must be a vbColor or any RGB value
Public Sub Dsp(ByVal intID As Integer, ByVal strMessage As String, Optional ByVal strOption As String, Optional ByVal strDelimiter As String)

    If intID < 4 Then
        Dim arrQueue() As String, arrMessage() As String, tmpMsgPart As String, i As Integer
        Dim strUsername As String, strDelim As String, strProduct As String, strD2 As String
        ReDim arrQueue(0)
        
        '// Get username and delimiter if applicable
        If intID < 3 Then
            If Len(strOption) > 0 Then
              strDelim = strOption
            Else
              strDelim = " "
            End If
        Else
            strUsername = strOption
            strProduct = GetInternalDataByUsername(strOption, 3)
            If strProduct = "D2DV" Or strProduct = "D2PX" Then strUsername = "*" & strUsername
            If Len(strDelimiter) > 0 Then
                strDelim = strDelimiter
            Else
                strDelim = " "
            End If
        End If
        
        '// If message is too long to display in a single AddQ, split into multiple messages
        If Len(strMessage) > 220 And InStr(strMessage, strDelim) Then
            arrMessage = Split(strMessage)
            
            For i = 0 To UBound(arrMessage)
                If Len(tmpMsgPart & arrMessage(i) & strDelim) < 205 Then
                    tmpMsgPart = tmpMsgPart & arrMessage(i) & strDelim
                Else
                    arrQueue(UBound(arrQueue)) = tmpMsgPart & "[more]"
                    ReDim Preserve arrQueue(UBound(arrQueue) + 1)
                    tmpMsgPart = arrMessage(i) & strDelim
                End If
            Next
            arrQueue(UBound(arrQueue)) = Left(tmpMsgPart, Len(tmpMsgPart) - 1)
        Else
            arrQueue(0) = strMessage
        End If
        If UBound(arrQueue) > 0 Then PadQueueCounter

        '// Display the message(s)
        For i = 0 To UBound(arrQueue)
            If Len(arrQueue(i)) > 0 Then
                Select Case intID
                    Case 1: AddQ arrQueue(i)
                    Case 2: AddQ "/me " & arrQueue(i)
                    Case 3: AddQ "/w " & strUsername & " " & arrQueue(i)
                End Select
            End If
        Next
    Else
        If Len(strDelimiter) > 0 Then strOption = strDelimiter
        AddChat strOption, strMessage
    End If
End Sub

