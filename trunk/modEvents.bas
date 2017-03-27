Attribute VB_Name = "modEvents"
'StealthBot Project - modEvents.bas
' Andy T (andy@stealthbot.net) March 2005

Option Explicit
Private Const OBJECT_NAME As String = "modEvents"

Private Const MSG_FILTER_MAX_EVENTS As Long = 100 ' maximum number of storable events
Private Const MSG_FILTER_DELAY_INT  As Long = 500 ' interval for event count measuring
Private Const MSG_FILTER_MSG_COUNT  As Long = 3   ' message count maximums

Private Type MSGFILTER
    UserObj   As Object
    EventObj  As Object
    EventTime As Date
End Type

Private m_arrMsgEvents()  As MSGFILTER
Private m_eventCount      As Integer

Public Sub Event_FlagsUpdate(ByVal Username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long, Optional QueuedEventID As Integer = 0)
    
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If

    Dim UserObj         As clsUserObj
    Dim PreviousUserObj As clsUserObj
    Dim UserEvent       As clsUserEventObj
    
    Dim UserIndex       As Integer
    Dim i               As Integer
    Dim PreviousFlags   As Long
    Dim Clan            As String
    Dim parsed          As String
    Dim pos             As Integer
    Dim doUpdate        As Boolean
    Dim Displayed       As Boolean  ' stores whether this event has been displayed by another event in the RTB

    ' if our username is for some reason null, we don't
    ' want to continue, possibly causing further errors
    If (LenB(Username) < 1) Then
        Exit Sub
    End If
 
    
    UserIndex = g_Channel.GetUserIndexEx(CleanUsername(Username))
    
    If (UserIndex > 0) Then
        Set UserObj = g_Channel.Users(UserIndex)
        
        If (QueuedEventID = 0) Then
            If (UserObj.Queue.Count > 0) Then
                Set UserEvent = New clsUserEventObj
                
                With UserEvent
                    .EventID = ID_USERFLAGS
                    .Flags = Flags
                    .Ping = Ping
                    .GameID = UserObj.Game
                End With
                
                UserObj.Queue.Add UserEvent
            Else
                PreviousFlags = UserObj.Flags
            End If
        Else
            PreviousFlags = _
                UserObj.Queue(QueuedEventID - 1).Flags
        End If
        
        Clan = UserObj.Clan
    Else
        If (g_Channel.IsSilent = False) Then
            frmChat.AddChat RTBColors.ErrorMessageText, "Warning! There was a flags update received for a user that we do " & _
                    "not have a record for.  This may be indicative of a server split or other technical difficulty."
                    
            Exit Sub
        Else
            If (g_Channel.Users.Count >= 200) Then
                Exit Sub
            End If
        
            Set UserObj = New clsUserObj

            With UserObj
                .Name = Username
                .Statstring = Message
            End With
        End If
    End If
    
    With UserObj
        .Flags = Flags
        .Ping = Ping
    End With
    
    If (g_Channel.IsSilent) Then
        g_Channel.Users.Add UserObj
    End If

    ' convert username to appropriate
    ' display format
    Username = UserObj.DisplayName
    
    ' are we receiving a flag update for ourselves?
    If (StrComp(Username, GetCurrentUsername, vbBinaryCompare) = 0) Then
        ' assign my current flags to the
        ' relevant internal variable
        MyFlags = Flags
        
        ' assign my current flags to the
        ' relevant scripting variable
        SharedScriptSupport.BotFlags = MyFlags
    End If
    
    ' we aren't in a silent channel, are we?
    If (g_Channel.IsSilent) Then
        AddName Username, UserObj.Name, UserObj.Game, Flags, Ping, UserObj.Stats.IconCode, _
            UserObj.Clan
    Else
        If ((UserObj.Queue.Count = 0) Or (QueuedEventID > 0)) Then
            If (Flags <> PreviousFlags) Then
                If (g_Channel.Self.IsOperator) Then
                    If ((Username = GetCurrentUsername) And _
                            ((PreviousFlags And USER_CHANNELOP) <> USER_CHANNELOP)) Then
                            
                        g_Channel.CheckUsers
                    Else
                        g_Channel.CheckUser Username
                    End If
                End If
                
                pos = checkChannel(Username)
                
                If (pos) Then
                    Dim NewFlags As Long
                    Dim LostFlags As Long
                    
                    frmChat.lvChannel.ListItems.Remove pos
                    
                    ' voodoo magic: only show flags that are new
                    NewFlags = Not (Flags Imp PreviousFlags)
                    LostFlags = Not (PreviousFlags Imp Flags)
                
                    If (NewFlags And USER_CHANNELOP) = USER_CHANNELOP Or _
                        (NewFlags And USER_BLIZZREP) = USER_BLIZZREP Or _
                        (NewFlags And USER_SYSOP) = USER_SYSOP Then
                        pos = 1
                    End If
                    
                    AddName Username, UserObj.Name, UserObj.Game, Flags, Ping, UserObj.Stats.IconCode, _
                        UserObj.Clan, pos
                    
                    ' default to display this event
                    Displayed = False
                    
                    ' check whether it has been
                    If QueuedEventID > 0 And UserObj.Queue.Count >= QueuedEventID Then
                        Set UserEvent = UserObj.Queue(QueuedEventID)
                        Displayed = UserEvent.Displayed
                    End If
                    
                    ' display if it has not
                    If Not Displayed Then
                        Dim FDescN As String
                        Dim FDescO As String
                        FDescN = FlagDescription(NewFlags, False)
                        FDescO = FlagDescription(LostFlags, False)
                        
                        If LenB(FDescN) > 0 Then
                            frmChat.AddChat RTBColors.JoinUsername, "-- ", RTBColors.JoinedChannelName, _
                                Username, RTBColors.JoinText, " is now a " & FDescN & "."
                        End If
                        
                        If LenB(FDescO) > 0 Then
                            frmChat.AddChat RTBColors.JoinUsername, "-- ", RTBColors.JoinedChannelName, _
                                Username, RTBColors.JoinText, " is no longer a " & FDescO & "."
                        End If
                    End If
                End If
            End If
        End If
    End If

    If ((UserObj.Queue.Count = 0) Or (QueuedEventID > 0)) Then
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' call event script function
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        On Error Resume Next
        
        RunInAll "Event_FlagUpdate", Username, Flags, Ping
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_FlagsUpdate()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Public Sub Event_JoinedChannel(ByVal ChannelName As String, ByVal Flags As Long)
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If

    Dim mailCount   As Integer
    Dim ToANSI      As String
    Dim LastChannel As String
    Dim sChannel    As String
    
    ' if our channel is for some reason null, we don't
    ' want to continue, possibly causing further errors
    If (LenB(ChannelName) = 0) Then
        Exit Sub
    End If
    
    LastChannel = g_Channel.Name
    
    Call frmChat.ClearChannel
    
    ' we want to reset our filter
    ' Values() when we join a new channel
    'BotVars.JoinWatch = 0
    
    'frmChat.tmrSilentChannel(0).Enabled = False
    
    'If (StrComp(g_Channel.Name, "Clan " & Clan.Name, vbTextCompare) = 0) Then
    '    PassedClanMotdCheck = False
    'End If

    ' show home channel in menu
    BotVars.LastChannel = LastChannel

    ' if we've just left another channel, call event script
    ' function indicating that we've done so.
    If (LenB(LastChannel) > 0) Then
        On Error Resume Next
        
        RunInAll "Event_ChannelLeave"
        
        On Error GoTo ERROR_HANDLER
    End If
    
    With g_Channel
        .Name = ChannelName
        .Flags = Flags
        .JoinTime = UtcNow
    End With
    
    PrepareHomeChannelMenu
    PrepareQuickChannelMenu
    
    SharedScriptSupport.MyChannel = ChannelName

    frmChat.AddChat RTBColors.JoinedChannelText, "-- Joined channel: ", _
        RTBColors.JoinedChannelName, ChannelName, RTBColors.JoinedChannelText, " --"
    
    SetTitle GetCurrentUsername & ", online in channel " & g_Channel.Name
    
    frmChat.UpdateTrayTooltip
    
    frmChat.ListviewTabs.Tab = LVW_BUTTON_CHANNEL
    Call frmChat.ListviewTabs_Click(LVW_BUTTON_CHANNEL)
    
    ' have we just joined the void?
    If (g_Channel.IsSilent) Then
        ' if we've joined the void, lets try to grab the list of
        ' users within the channel by attempting to force a user
        ' update message using Battle.net's unignore command.
        If (frmChat.mnuDisableVoidView.Checked = False) Then
            ' lets inform user of potential lag issues while in this channel
            frmChat.AddChat RTBColors.InformationText, "If you experience a lot of lag while within " & _
                    "this channel, try selecting 'Disable Silent Channel View' from the Window menu."
            
            frmChat.tmrSilentChannel(1).Enabled = True
        
            frmChat.AddQ "/unignore " & GetCurrentUsername
        End If
    Else
        frmChat.tmrSilentChannel(1).Enabled = False
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' check for mail
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    mailCount = GetMailCount(GetCurrentUsername)
        
    If (mailCount) Then
        frmChat.AddChat RTBColors.ConsoleText, "You have " & _
            mailCount & " new message" & IIf(mailCount = 1, "", "s") & _
                ". Type /inbox to retrieve."
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' call event script function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    
    RunInAll "Event_ChannelJoin", ChannelName, Flags
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_JoinedChannel()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Public Sub Event_UserDataReceived(ByRef oRequest As udtServerRequest, ByVal sUsername As String, ByRef Keys() As String, ByRef Values() As String)
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If

    Dim sKeyShort As String
    Dim sValue As String
    Dim aOutput() As String
    Dim i As Integer
    Dim j As Integer
    Dim s As String
    Dim d As Double

    Dim oFT As FILETIME
    Dim oST As SYSTEMTIME

    Const LONG_MAX_VALUE As Double = 2147483647

    RunInAll "Event_UserDataReceived", oRequest.Tag, Keys, Values

    For i = LBound(Keys) To UBound(Keys)
        RunInAll "Event_KeyReturn", Keys(i), Values(i)

        sKeyShort = Mid$(Keys(i), InStr(1, Keys(i), "\", vbTextCompare) + 1)
        sValue = Values(i)

        Select Case oRequest.HandlerType
            Case reqUserInterface
                frmProfile.SetKey Keys(i), sValue
            Case reqUserCommand, reqInternal
                If StrComp(Left$(Keys(i), 7), "System\", vbTextCompare) = 0 Then
                    j = InStr(1, sValue, Space$(1), vbBinaryCompare)

                    If j > 0 Then    ' Probably a FILETIME
                        With oFT
                            .dwLowDateTime = UnsignedToLong(CDbl(Mid$(KillNull(sValue), j + 1)))
                            .dwHighDateTime = UnsignedToLong(CDbl(Left$(sValue, j)))
                        End With

                        FileTimeToSystemTime oFT, oST

                        s = StringFormat("{0}: {1} (Battle.net time)", sKeyShort, SystemTimeToString(oST))
                    ElseIf StrictIsNumeric(sValue) Then
                        s = StringFormat("{0}: {1}", sKeyShort, ConvertTimeInterval(sValue, True))
                    End If

                    If oRequest.HandlerType = reqUserCommand Then
                        oRequest.Command.Respond s
                    Else
                        frmChat.AddChat RTBColors.ServerInfoText, s
                    End If
                Else
                    aOutput = Split(sValue, Chr(13))
                    For j = 0 To UBound(aOutput)
                        s = StringFormat("[{0}] {1}", sKeyShort, aOutput(j))

                        If oRequest.HandlerType = reqUserCommand Then
                            oRequest.Command.Respond s
                        Else
                            frmChat.AddChat RTBColors.ServerInfoText, s
                        End If
                    Next
                End If
        End Select
    Next
    
    ' If this request was triggered by a command, send the response.
    If oRequest.HandlerType = reqUserCommand Then
        If oRequest.Command.GetResponse().Count = 0 Then
            oRequest.Command.Respond StringFormat("{0} has not configured a profile.", sUsername)
        End If
        oRequest.Command.SendResponse
    End If
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_UserDataReceived()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Public Sub Event_LeftChatEnvironment()
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    BotVars.LastChannel = g_Channel.Name
    PrepareHomeChannelMenu
    PrepareQuickChannelMenu
    
    frmChat.ClearChannel
    
    SetTitle GetCurrentUsername & ", online on " & BotVars.Gateway
    
    frmChat.ListviewTabs_Click 0
    
    frmChat.AddChat RTBColors.JoinedChannelText, "-- Left channel --"
    
    On Error Resume Next
    
    RunInAll "Event_ChannelLeave"
    
    On Error GoTo ERROR_HANDLER
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_LeftChatEnvironment()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Public Sub Event_LoggedOnAs(Username As String, Statstring As String, AccountName As String)
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    Dim sChannel As String
    Dim ShowW3   As Boolean
    Dim ShowD2   As Boolean
    Dim Stats    As clsUserStats

    LastWhisper = vbNullString

    'If InStr(1, Username, "*", vbBinaryCompare) <> 0 Then
    '    Username = Right(Username, Len(Username) - InStr(1, Username, "*", vbBinaryCompare))
    'End If
    
    Call g_Queue.Clear
    
    Set g_Channel = New clsChannelObj
    Set g_Friends = New Collection
    ' reset Clan if we didn't just receive a SID_CLANINFO
    If Not g_Clan.InClan Then
        Set g_Clan = New clsClanObj
    End If
    
    g_Online = True
    
    ConnectionTickCount = GetTickCountMS()
    
    ' in case this wasn't set before
    ds.EnteredChatFirstTime = True
    
    Set Stats = New clsUserStats
    Stats.Statstring = Statstring
    
    CurrentUsername = KillNull(Username)
    
    'RequestSystemKeys
    
    Call SetNagelStatus(frmChat.sckBNet.SocketHandle, True)
    
    Call EnableSO_KEEPALIVE(frmChat.sckBNet.SocketHandle)
    
    If (StrComp(Left$(CurrentUsername, 2), "w#", vbTextCompare) = 0) Then
        CurrentUsername = Mid(CurrentUsername, 3)
    End If

    ' if D2 and on a char, we need to tell the whole world this so that Self is known later on
    If (StrComp(Stats.Game, PRODUCT_D2DV, vbBinaryCompare) = 0) Or (StrComp(Stats.Game, PRODUCT_D2XP, vbBinaryCompare) = 0) Then
        If (LenB(Stats.CharacterName) > 0) Then
            CurrentUsername = Stats.CharacterName & "*" & CurrentUsername
        End If
    End If

    ' show home channel in menu
    PrepareHomeChannelMenu
    PrepareQuickChannelMenu

    ' setup Bot menu game-specific features
    ShowW3 = (StrComp(Stats.Game, PRODUCT_WAR3, vbBinaryCompare) = 0) Or (StrComp(Stats.Game, PRODUCT_W3XP, vbBinaryCompare) = 0)
    ShowD2 = (StrComp(Stats.Game, PRODUCT_D2DV, vbBinaryCompare) = 0) Or (StrComp(Stats.Game, PRODUCT_D2XP, vbBinaryCompare) = 0)
    frmChat.mnuProfile.Enabled = True
    'frmChat.mnuClanCreate.Visible = ShowW3
    frmChat.mnuRealmSwitch.Visible = ShowD2

    SharedScriptSupport.myUsername = GetCurrentUsername
    
    With frmChat
        .InitListviewTabs
    
        .AddChat RTBColors.InformationText, "[BNCS] Logged on as ", RTBColors.SuccessText, Username, _
            RTBColors.InformationText, StringFormat(" using {0}.", Stats.ToString)
            
        .tmrAccountLock.Enabled = False
        .UpTimer.Enabled = True
    End With
    
    If (frmChat.sckBNLS.State <> sckClosed) Then
        frmChat.sckBNLS.Close
    End If
    
    If (ExReconnectTimerID > 0) Then
        Call KillTimer(0, ExReconnectTimerID)
        
        ExReconnectTimerID = 0
    End If
    
    If Config.FriendsListTab Then
        Call frmChat.FriendListHandler.RequestFriendsList
    End If
    
    Set Stats = Nothing
    
    RequestSystemKeys reqInternal
    If (LenB(BotVars.Gateway) > 0) Then
        ' PvPGN: we already have our gateway, we're logged on
        SetTitle GetCurrentUsername & ", online in channel " & g_Channel.Name
        
        Call InsertDummyQueueEntry
        
        On Error Resume Next
        
        RunInAll "Event_LoggedOn", CurrentUsername, BotVars.Product
    End If
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_LoggedOnAs()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

' updated 8-10-05 for new logging system
Public Sub Event_LogonEvent(ByVal Action As String, ByVal Result As Long, Optional ByVal ExtraInfo As String)
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    Dim lColor       As Long
    Dim sMessage     As String

    lColor = RTBColors.ErrorMessageText
    
    ' get starting text
    Select Case UCase$(Action)
        Case ACCOUNT_MODE_LOGON
            sMessage = "Logon error - "
        Case ACCOUNT_MODE_CREAT
            sMessage = "Account creation error - "
        Case ACCOUNT_MODE_CHPWD
            sMessage = "Password change error - "
    End Select
    
    ' choose result code
    Select Case Result
        Case &H0
            lColor = RTBColors.SuccessText
            ' replace with specific success message
            Select Case UCase$(Action)
                Case ACCOUNT_MODE_LOGON
                    sMessage = "Logon successful."
                Case ACCOUNT_MODE_CREAT
                    sMessage = "Account created successfully."
                Case ACCOUNT_MODE_CHPWD
                    sMessage = "Account password changed successfully."
                Case ACCOUNT_MODE_RSPWD
                    sMessage = "Sent the request to reset password. You will receive an email to continue this process."
                Case ACCOUNT_MODE_CHREG
                    sMessage = "Sent the request to change email associated with the account."
            End Select
        Case &H1
            sMessage = sMessage & "account does not exist."
        Case &H2
            Select Case UCase$(Action)
                Case ACCOUNT_MODE_CHPWD
                    sMessage = sMessage & "invalid old password."
                Case Else
                    sMessage = sMessage & "invalid password."
            End Select
        Case &H4
            sMessage = sMessage & "account already exists."
        Case &H5
            sMessage = sMessage & "account requires upgrade."
        Case &H6
            sMessage = sMessage & "account closed - " & ExtraInfo & "."
        Case &H7
            sMessage = sMessage & "name too short."
        Case &H8
            sMessage = sMessage & "name contains invalid characters."
        Case &H9
            sMessage = sMessage & "name contains banned word."
        Case &HA
            sMessage = sMessage & "name contains too few alphanumeric charaters."
        Case &HB
            sMessage = sMessage & "name contains adjacent punctuation."
        Case &HC
            sMessage = sMessage & "name contains too many punctuation characters."
        Case &HE
            sMessage = sMessage & "account email registration."
        Case &HF
            sMessage = sMessage & ExtraInfo & "."
        Case &H3101 ' actually status 0x01 from SID_CHANGEPASSWORD
            sMessage = sMessage & "account does not exist or invalid old password."
        Case &H3D05 ' actually status 0x05 from SID_CREATEACCOUNT2
            sMessage = sMessage & "account is still being created."
        Case -3& ' parameter empty
            Select Case UCase$(Action)
                Case ACCOUNT_MODE_LOGON, ACCOUNT_MODE_CREAT
                    sMessage = sMessage & "username or password not provided."
                Case ACCOUNT_MODE_CHPWD
                    sMessage = sMessage & "new password not provided."
                Case ACCOUNT_MODE_RSPWD
                    sMessage = sMessage & "email address not provided."
                Case ACCOUNT_MODE_CHREG
                    sMessage = sMessage & "new email address not provided."
            End Select
        Case -2& ' time out
            sMessage = sMessage & "timed out."
        Case -1& ' attempt
            lColor = RTBColors.InformationText
            ' replace with specific in-progress message
            Select Case UCase$(Action)
                Case ACCOUNT_MODE_LOGON
                    sMessage = "Sending logon information..."
                Case ACCOUNT_MODE_CREAT
                    sMessage = "Attempting to create account..."
                Case ACCOUNT_MODE_CHPWD
                    sMessage = "Attempting to change password..."
            End Select
        Case Else
            sMessage = sMessage & "unknown response code (0x" & Hex(Result) & ": " & ExtraInfo & ")."
    End Select
    
    frmChat.AddChat lColor, "[BNCS] " & sMessage

    Exit Sub

ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_LogonEvent()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Public Sub Event_ServerError(ByVal Message As String)
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    frmChat.AddChat RTBColors.ErrorMessageText, Message
    
    RunInAll "Event_ServerError", Message
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_ServerError()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Public Sub Event_ChannelJoinError(ByVal EventID As Integer, ByVal ChannelName As String)
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    Dim ChannelJoinError As String
    Dim ChannelJoinButtons As VbMsgBoxStyle
    Dim ChannelJoinResult As VbMsgBoxResult
    Dim Message As String
    Dim ChannelCreateOption As String

    'frmChat.AddChat RTBColors.ErrorMessageText, Message
    
    If (LenB(BotVars.Gateway) = 0) Then
        ' continue gateway discovery
        SEND_SID_CHATCOMMAND "/whoami"
    Else
        ChannelCreateOption = Config.AutoCreateChannels
    
        Select Case ChannelCreateOption
            Case "ALERT"
                Select Case EventID
                    Case ID_CHANNELDOESNOTEXIST
                        ChannelJoinError = "Channel does not exist." & vbNewLine & "Do you want to create it?"
                        ChannelJoinButtons = vbYesNo Or vbQuestion Or vbDefaultButton1
                    Case ID_CHANNELFULL
                        ChannelJoinError = "Channel is full."
                        ChannelJoinButtons = vbOKOnly Or vbExclamation Or vbDefaultButton1
                    Case ID_CHANNELRESTRICTED
                        ChannelJoinError = "Channel is restricted."
                        ChannelJoinButtons = vbOKOnly Or vbExclamation Or vbDefaultButton1
                End Select
                
                ChannelJoinResult = MsgBox("Failed to join " & ChannelName & ":" & vbNewLine & _
                    ChannelJoinError, ChannelJoinButtons, "StealthBot")
                
                If ChannelJoinResult = vbYes Then
                    Call FullJoin(ChannelName, 2)
                End If
                
            Case Else
            ' "ALWAYS" - handle it as error to bot
            ' "NEVER" - failed to join or create
                Select Case EventID
                    Case ID_CHANNELDOESNOTEXIST
                        Message = "[BNCS] Channel does not exist."
                    Case ID_CHANNELFULL
                        Message = "[BNCS] Channel is full."
                    Case ID_CHANNELRESTRICTED
                        Message = "[BNCS] Channel is restricted."
                End Select
                
                frmChat.AddChat RTBColors.ErrorMessageText, Message
                
        End Select
        
        'should we expose?
        'RunInAll "Event_ChannelJoinError", EventID, ChannelName
    End If
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_ChannelJoinError()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Public Sub Event_ServerInfo(ByVal Username As String, ByVal Message As String)
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If

    Const MSG_BANNED      As String = " was banned by "
    Const MSG_UNBANNED    As String = " was unbanned by "
    Const MSG_SQUELCHED   As String = " has been squelched."
    Const MSG_UNSQUELCHED As String = " has been unsquelched."
    Const MSG_KICKEDOUT   As String = " kicked you out of the channel!"
    Const MSG_FRIENDS     As String = "Your friends are:"
    
    Dim i      As Integer
    Dim temp   As String
    Dim bHide  As Boolean
    Dim ToANSI As String
    
    If (Message = vbNullString) Then
        Exit Sub
    End If
    
    Username = ConvertUsername(Username)

    If g_Clan.InClan Then
        If (StrComp(g_Channel.Name, "Clan " & g_Clan.Name, vbTextCompare) = 0) Then
            If (g_Clan.PendingClanMOTD) Then
                Call frmChat.AddChat(RTBColors.ServerInfoText, Message)
                g_Clan.PendingClanMOTD = False
                Exit Sub
            End If
        End If
    End If
    
    If (g_request_receipt) Then ' for .cs and .cb commands
        Caching = True
    
        
        ' Changed 08-18-09 - Hdx - Uses the new Channel cache function, Eventually to beremoved to script
        'Call CacheChannelList(Message, 1)
        Call CacheChannelList(enAdd, Message)
        
        'With frmChat.cacheTimer
        '    .Enabled = False
        '    .Enabled = True
        'End With
    End If
    
    ' what is our current gateway name?
    If (BotVars.Gateway = vbNullString) Then
        If (InStr(1, Message, "You are ", vbTextCompare) > 0) And (InStr(1, Message, ", using ", _
                vbTextCompare) > 0) Then
                
            If ((InStr(1, Message, " in channel ", vbTextCompare) = 0) And _
                    (InStr(1, Message, " in game ", vbTextCompare) = 0) And _
                    (InStr(1, Message, " a private ", vbTextCompare) = 0)) Then
                    
                i = InStrRev(Message, Space$(1))
                
                BotVars.Gateway = Mid$(Message, i + 1)
                
                SetTitle GetCurrentUsername & ", online on " & BotVars.Gateway
                
                Call DoChannelJoinHome
                
                Call InsertDummyQueueEntry
                
                RunInAll "Event_LoggedOn", CurrentUsername, BotVars.Product

                Exit Sub
            End If
        End If
    End If

    If (InStr(1, Message, Space$(1), vbBinaryCompare) <> 0) Then
        If (InStr(1, Message, "are still marked", vbTextCompare) <> 0) Then
            Exit Sub
        End If
        
        If ((InStr(1, Message, " from your friends list.", vbBinaryCompare) > 0) Or _
            (InStr(1, Message, " to your friends list.", vbBinaryCompare) > 0) Or _
            (InStr(1, Message, " in your friends list.", vbBinaryCompare) > 0) Or _
            (InStr(1, Message, " of your friends list.", vbBinaryCompare) > 0)) Then
            
            If Config.FriendsListTab Then
                If Not frmChat.FriendListHandler.SupportsFriendPackets(Config.Game) Then
                    Call frmChat.FriendListHandler.RequestFriendsList
                End If
            End If
        End If
        
        'Ban Evasion and banned-user tracking
        temp = Split(Message, " ")(1)
        
        ' added 1/21/06 thanks to
        ' http://www.stealthbot.net/forum/index.php?showtopic=24582
        
        If (Len(temp) > 0) Then
            Dim Banning    As Boolean
            Dim Unbanning  As Boolean
            Dim User       As String
            Dim cOperator  As String
            Dim msgPos     As Integer
            Dim pos        As Integer
            Dim tmp        As String
            Dim banpos     As Integer
            Dim j          As Integer
            Dim Reason     As String
            
            If (InStr(1, Message, MSG_BANNED, vbTextCompare) > 0) Then
                User = Left$(Message, _
                    (InStr(1, Message, MSG_BANNED, vbBinaryCompare) - 1))
                
                Reason = Mid$(Message, InStr(1, Message, MSG_BANNED, vbBinaryCompare) + Len(MSG_BANNED) + 1) ' trim out username and banned message
                If (InStr(1, Reason, " (", vbBinaryCompare)) Then 'Did they give a message?
                  Reason = Mid$(Reason, InStr(1, Reason, " (") + 2) 'trim out the banning name (Note, when banned by a rep using Len(Username) won't work as its banned "By a Blizzard Representative")
                  Reason = Left$(Reason, Len(Reason) - 2) 'Trim off the trailing ")."
                Else
                  Reason = vbNullString
                End If
                
                If (Len(User) > 0) Then
                    pos = g_Channel.GetUserIndex(Username)
                    
                    If (pos > 0) Then
                        Dim BanlistObj As clsBanlistUserObj
                                                
                        banpos = g_Channel.IsOnBanList(User, Username)
                        
                        If (banpos > 0) Then
                            g_Channel.Banlist.Remove banpos
                        Else
                            g_Channel.BanCount = (g_Channel.BanCount + 1)
                        End If
                        
                        If ((BotVars.StoreAllBans) Or _
                                (StrComp(Username, GetCurrentUsername, vbBinaryCompare) = 0)) Then
                            
                            Set BanlistObj = New clsBanlistUserObj
                            
                            With BanlistObj
                                .Name = User
                                .Operator = Username
                                .DateOfBan = UtcNow
                                .IsDuplicateBan = (g_Channel.IsOnBanList(User) > 0)
                                .Reason = Reason
                            End With
                        
                            If (BanlistObj.IsDuplicateBan) Then
                                With g_Channel.Banlist(g_Channel.IsOnBanList(User))
                                    .IsDuplicateBan = False
                                End With
                            End If
                            
                            g_Channel.Banlist.Add BanlistObj
                        End If
                    End If
                    
                    Call RemoveBanFromQueue(User)
                End If
                
                If (frmChat.mnuHideBans.Checked) Then
                    bHide = True
                End If
            ElseIf (InStr(1, Message, MSG_UNBANNED, vbTextCompare) > 0) Then
                User = Left$(Message, _
                    (InStr(1, Message, MSG_UNBANNED, vbBinaryCompare) - 1))
                                
                If (Len(User) > 0) Then
                    g_Channel.BanCount = (g_Channel.BanCount - 1)
                    
                    Do
                        banpos = g_Channel.IsOnBanList(User)
                    
                        If (banpos > 0) Then
                            g_Channel.Banlist.Remove banpos
                        End If
                    Loop While (banpos <> 0)
                End If
            End If
    
            '// backup channel
            If (InStr(1, Message, "kicked you out", vbTextCompare) > 0) Then
                If (BotVars.UseBackupChan) Then
                    If (Len(BotVars.BackupChan) > 0) Then
                        frmChat.AddQ "/join " & BotVars.BackupChan
                    End If
                Else
                    frmChat.AddQ "/join " & g_Channel.Name
                End If
            End If
            
            If (InStr(1, Message, " has been unsquelched", vbTextCompare) > 0) Then
                If ((g_Channel.IsSilent) And (frmChat.mnuDisableVoidView.Checked = False)) Then
                    frmChat.lvChannel.ListItems.Clear
                End If
            End If
        End If
        
        If (InStr(1, Message, "designated heir", vbTextCompare) <> 0) Then
            g_Channel.OperatorHeir = Left$(Message, Len(Message) - 29)
        End If
        
        
        temp = "Your friends are:"
        
        If (StrComp(Left$(Message, Len(temp)), temp) = 0) Then
            If (Not (BotVars.ShowOfflineFriends)) Then
                Message = Message & _
                    "  ÿci(StealthBot is hiding your offline friends)"
            End If
        End If
    
    End If ' message contains a space
    
    If (StrComp(Right$(Message, 9), ", offline", vbTextCompare) = 0) Then
        If (BotVars.ShowOfflineFriends) Then
            frmChat.AddChat RTBColors.ServerInfoText, Message
        End If
    Else
        If (Not (bHide)) Then
            frmChat.AddChat RTBColors.ServerInfoText, Message
        End If
    End If

    RunInAll "Event_ServerInfo", Message
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_ServerInfo()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Public Sub Event_UserEmote(ByVal Username As String, ByVal Flags As Long, ByVal Message As String, _
    Optional QueuedEventID As Integer = 0)
    
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
        
    Dim UserEvent   As clsUserEventObj
    Dim UserObj     As clsUserObj
    
    Dim i           As Integer
    Dim ToANSI      As String
    Dim pos         As Integer
    Dim PassedQueue As Boolean
    
    pos = g_Channel.GetUserIndexEx(CleanUsername(Username))
    
    If (pos > 0) Then
        Set UserObj = g_Channel.Users(pos)
        
        If (QueuedEventID = 0) Then
            UserObj.LastTalkTime = UtcNow
            
            If (UserObj.Queue.Count > 0) Then
                Set UserEvent = New clsUserEventObj
                
                With UserEvent
                    .EventID = ID_EMOTE
                    .Flags = Flags
                    .Message = Message
                End With
                
                UserObj.Queue.Add UserEvent
            End If
        End If
    Else
        ' create new user object for invisible representatives...
        Set UserObj = New clsUserObj
        
        ' store user name
        UserObj.Name = Username
    End If
    
    ' convert user name
    Username = UserObj.DisplayName
    
    If (QueuedEventID = 0) Then
        If (g_Channel.Self.IsOperator) Then
            If (GetSafelist(Username) = False) Then
                CheckMessage Username, Message
            End If
        End If
    End If
    
    If ((UserObj.Queue.Count = 0) Or (QueuedEventID > 0)) Then
        If (AllowedToTalk(Username, Message)) Then
            'If (GetVeto = False) Then
                frmChat.AddChat RTBColors.EmoteText, "<", RTBColors.EmoteUsernames, Username & _
                    Space$(1), RTBColors.EmoteText, Message & ">"
            'End If
            
            If (Catch(0) <> vbNullString) Then
                CheckPhrase Username, Message, CPEMOTE
            End If
            
            If (frmChat.mnuFlash.Checked) Then
                FlashWindow
            End If
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' call event script function
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        On Error Resume Next
        
        If ((BotVars.NoSupportMultiCharTrigger) And (Len(BotVars.TriggerLong) > 1)) Then
            If (StrComp(Left$(Message, Len(BotVars.TriggerLong)), BotVars.TriggerLong, _
                vbBinaryCompare) = 0) Then
                
                Message = BotVars.Trigger & Mid$(Message, Len(BotVars.TriggerLong) + 1)
            End If
        End If
        
        RunInAll "Event_UserEmote", Username, Flags, Message
    End If

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_UserEmote()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Public Sub Event_UserInChannel(ByVal Username As String, ByVal Flags As Long, ByVal Statstring As String, ByVal Ping As Long, Optional QueuedEventID As Integer = 0)
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If

    Dim UserEvent    As clsUserEventObj
    Dim UserObj      As clsUserObj
    Dim found        As ListItem
    
    Dim UserIndex    As Integer
    Dim i            As Integer
    Dim strCompare   As String
    Dim Level        As Byte
    Dim StatUpdate   As Boolean
    Dim Index        As Long
    Dim Stats        As String
    Dim Clan         As String
    Dim pos          As Integer
    Dim showUpdate   As Boolean
    Dim Displayed    As Boolean ' whether this event has been displayed in the RTB (if combined with another)
    Dim AcqOps       As Boolean
    Dim NewIcon      As Long    ' temp store new icon

    If (LenB(Username) < 1) Then
        Exit Sub
    End If

    UserIndex = g_Channel.GetUserIndexEx(CleanUsername(Username))

    If (UserIndex > 0) Then

        Set UserObj = g_Channel.Users(UserIndex)
        
        If (QueuedEventID = 0) Then
            If (UserObj.Queue.Count > 0) Then
                If (UserObj.Stats.Statstring = vbNullString) Then
                    showUpdate = True
                End If
                
                Set UserEvent = New clsUserEventObj
                
                With UserEvent
                    .EventID = ID_USER
                    .Flags = Flags
                    .Ping = Ping
                    .GameID = UserObj.Game
                    .Clan = UserObj.Clan
                    .Statstring = Statstring
                End With
                
                UserObj.Queue.Add UserEvent
            End If
        End If
        
        StatUpdate = True
    Else
        Set UserObj = New clsUserObj
    End If
    
    With UserObj
        .Name = Username
        .Flags = Flags
        .Ping = Ping
        .JoinTime = g_Channel.JoinTime
        .Statstring = Statstring
    End With
    
    If (UserIndex = 0) Then
        g_Channel.Users.Add UserObj
    End If
    
    Username = UserObj.DisplayName
    
    'ParseStatstring OriginalStatstring, Stats, Clan
    If (StatUpdate = False) Then
        'frmChat.AddChat vbRed, UserObj.Stats.IconCode
    
        AddName Username, UserObj.Name, UserObj.Game, Flags, Ping, UserObj.Stats.IconCode, UserObj.Clan
        
        frmChat.ListviewTabs_Click 0
        
        DoLastSeen Username
    Else
        If ((UserObj.Queue.Count = 0) Or (QueuedEventID > 0)) Then
            If (JoinMessagesOff = False) Then
                ' default to display this event
                Displayed = False
                
                ' check whether it has been
                If QueuedEventID > 0 And UserObj.Queue.Count >= QueuedEventID Then
                    Set UserEvent = UserObj.Queue(QueuedEventID)
                    Displayed = UserEvent.Displayed
                End If
                
                ' display if it has not already been
                If Not Displayed Then
                    Dim UserColor As Long
                    Dim FDesc As String
                    FDesc = FlagDescription(Flags, False)
                    
                    If LenB(FDesc) > 0 Then
                        FDesc = " as a " & FDesc
                    End If
                    
                    ' display message
                    If (Flags And USER_BLIZZREP) Then
                        UserColor = RGB(97, 105, 255)
                    ElseIf (Flags And USER_SYSOP) Then
                        UserColor = RGB(97, 105, 255)
                    ElseIf (Flags And USER_CHANNELOP) Then
                        UserColor = RTBColors.TalkUsernameOp
                    Else
                        UserColor = RTBColors.JoinUsername
                    End If
                    
                    frmChat.AddChat RTBColors.JoinText, "-- Stats updated: ", _
                        UserColor, Username, _
                        RTBColors.JoinUsername, " [" & Ping & "ms]", _
                        RTBColors.JoinText, " is using " & UserObj.Stats.ToString, _
                        RTBColors.JoinUsername, FDesc, _
                        RTBColors.JoinText, "."
                End If
            End If
            
            pos = checkChannel(Username)

            If (pos > 0) Then
            
                Set found = frmChat.lvChannel.ListItems(pos)
                
                ' if the update occured to a D2 user ...
                If ((StrComp(UserObj.Game, PRODUCT_D2DV) = 0) Or (StrComp(UserObj.Game, PRODUCT_D2XP) = 0)) Then
                    ' the username could have changed!
                    If (StrComp(UserObj.DisplayName, found.Text, vbBinaryCompare) <> 0) Then
                        ' it did, so update user name text in channel list
                        found.Text = UserObj.DisplayName
                        
                        ' now check if this is Self
                        If (StrComp(UserObj.Name, CleanUsername(CurrentUsername), vbBinaryCompare) = 0) Then
                            ' it is! we have to do some magic to tell SB we have a new name
                            CurrentUsername = UserObj.Stats.CharacterName & "*" & CleanUsername(CurrentUsername)
                            
                            ' tell scripting
                            SharedScriptSupport.myUsername = GetCurrentUsername
                            
                            ' set form title
                            SetTitle GetCurrentUsername & ", online in channel " & _
                                    g_Channel.Name
                            
                            ' tell tray icon
                            Call frmChat.UpdateTrayTooltip
                        End If
                    End If
                End If
                
                ' if we are showing stats icons ...
                If (BotVars.ShowStatsIcons) Then 'and the icon code is valid
                    If (UserObj.Stats.IconCode <> -1) Then
                        ' if the icon in the list is not the icon found by stats, update
                        NewIcon = GetSmallIcon(UserObj.Game, UserObj.Flags, UserObj.Stats.IconCode)
                        If (found.SmallIcon <> NewIcon) Then
                            found.SmallIcon = NewIcon
                        End If
                    End If
                End If
                
                If (found.ListSubItems.Count > 0) Then
                    found.ListSubItems(1).Text = UserObj.Clan
                End If
                
                Set found = Nothing
            End If
        End If
    End If
    
    If ((UserObj.Queue.Count = 0) Or (QueuedEventID > 0)) Then
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' call event script function
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        On Error Resume Next
        
        RunInAll "Event_UserInChannel", Username, Flags, UserObj.Stats.ToString, Ping, _
            UserObj.Game, StatUpdate
    End If
    
    If (MDebug("statstrings")) Then
        frmChat.AddChat RTBColors.InformationText, "Username: " & Username & ", Statstring: " & _
            Statstring
    End If

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_UserInChannel()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Public Sub Event_UserJoins(ByVal Username As String, ByVal Flags As Long, ByVal Statstring As String, ByVal Ping As Long, Optional QueuedEventID As Integer = 0)
                
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    Dim UserObj     As clsUserObj
    Dim UserEvent   As clsUserEventObj
    
    Dim toCheck     As String
    Dim strCompare  As String
    Dim i           As Long
    Dim temp        As Byte
    Dim Level       As Byte
    Dim L           As Long
    Dim Banned      As Boolean
    Dim f           As Integer
    Dim UserIndex   As Integer
    Dim BanningUser As Boolean
    Dim pStats      As String
    Dim IsBanned    As Boolean
    Dim AcqFlags    As Long
    Dim ToDisplay   As Boolean
    
    If (Len(Username) < 1) Then
        Exit Sub
    End If

    UserIndex = g_Channel.GetUserIndexEx(CleanUsername(Username))
    
    If (QueuedEventID > 0) Then
        If (UserIndex = 0) Then
            frmChat.AddChat RTBColors.ErrorMessageText, "Error: We have received a queued join event for a user that we " & _
                "couldn't find in the channel."
        
            Exit Sub
        End If
    
        Set UserObj = g_Channel.Users(UserIndex)
    Else
        If (UserIndex = 0) Then
            Set UserObj = New clsUserObj
            
            With UserObj
                .Name = Username
                .Flags = Flags
                .Ping = Ping
                .JoinTime = UtcNow
                .Statstring = Statstring
            End With

            If (BotVars.ChatDelay > 0) Then
                Set UserEvent = New clsUserEventObj
                
                With UserEvent
                    .EventID = ID_JOIN
                    .Flags = Flags
                    .Ping = Ping
                    .GameID = UserObj.Game
                    .Statstring = Statstring
                    .Clan = UserObj.Clan
                    .IconCode = UserObj.Stats.Icon
                End With
                
                UserObj.Queue.Add UserEvent
            End If

            g_Channel.Users.Add UserObj
        Else
            frmChat.AddChat RTBColors.ErrorMessageText, "Warning! We have received a join event for a user that we had thought was " & _
                    "already present within the channel.  This may be indicative of a server split or other technical difficulty."
            
            Exit Sub
        End If
    End If
    
    Username = UserObj.DisplayName
    
    If ((UserObj.Queue.Count = 0) Or (QueuedEventID = 0)) Then
        If (g_Channel.Self.IsOperator) Then
            g_Channel.CheckUser Username, UserObj
        End If
    End If
    
    If ((UserObj.Queue.Count = 0) Or (QueuedEventID > 0)) Then
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' GUI
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        ' if we have join/leaves on
        If (JoinMessagesOff = False) Then

            ' does this event have events delayed after it?
            If QueuedEventID > 0 And UserObj.Queue.Count > 0 Then
                
                ' loop through the events occuring after this one
                For i = QueuedEventID To UserObj.Queue.Count
                
                    ' get the event
                    Set UserEvent = UserObj.Queue(i)
                    
                    ' default to not combine with userjoins
                    ToDisplay = False
                    
                    Select Case UserEvent.EventID
                    
                        ' user flags update
                        Case ID_USERFLAGS
                            ' will combine with userjoins
                            ToDisplay = True
                            
                            AcqFlags = UserEvent.Flags
                            
                        ' user stats update / user in channel
                        Case ID_USER
                            ' will combine with userjoins
                            ToDisplay = True
                            
                            ' is stats different / provided?
                            If LenB(UserEvent.Statstring) > 0 Then
                                If StrComp(UserEvent.Statstring, UserObj.Statstring) Then
                                    
                                    ' store stats update stats in object used in userjoins message generation
                                    UserObj.Statstring = UserEvent.Statstring
                                End If
                            End If
                        
                    End Select
                    
                    ' if we're going to combine this event with userjoins ...
                    If ToDisplay Then 'then set .displayed on the queue'd event so it is not displayed separately
                        UserEvent.Displayed = True
                        
                        ' also update in collection
                        UserObj.Queue.Remove i
                        UserObj.Queue.Add UserEvent, , , i - 1
                    End If
                    
                Next i
                
            End If
            
            If (Not Filters) Or (Not CheckBlock(Username)) Then
                Dim UserColor As Long
                Dim FDesc As String
                FDesc = FlagDescription(AcqFlags Or Flags, False)
                
                If LenB(FDesc) > 0 Then
                    FDesc = " as a " & FDesc
                End If
                
                ' display message
                If (AcqFlags And USER_BLIZZREP) Or (Flags And USER_BLIZZREP) Then
                    UserColor = RGB(97, 105, 255)
                ElseIf (AcqFlags And USER_SYSOP) Or (Flags And USER_SYSOP) Then
                    UserColor = RGB(97, 105, 255)
                ElseIf (AcqFlags And USER_CHANNELOP) Or (Flags And USER_CHANNELOP) Then
                    UserColor = RTBColors.TalkUsernameOp
                Else
                    UserColor = RTBColors.JoinUsername
                End If
                
                frmChat.AddChat RTBColors.JoinText, "-- ", _
                    UserColor, Username, _
                    RTBColors.JoinUsername, " [" & Ping & "ms]", _
                    RTBColors.JoinText, " has joined the channel using " & UserObj.Stats.ToString, _
                    RTBColors.JoinUsername, FDesc, _
                    RTBColors.JoinText, "."
            End If
        End If
        
        ' add to user list
        AddName Username, UserObj.Name, UserObj.Game, Flags, Ping, UserObj.Stats.IconCode, UserObj.Clan
        
        ' focus on channel tab
        frmChat.ListviewTabs_Click 0
        
        ' flash window
        If (frmChat.mnuFlash.Checked) Then
            FlashWindow
        End If
        
        ' update last seen info
        Call DoLastSeen(Username)
        
        ' check is banned
        IsBanned = (UserObj.PendingBan)
        
        'frmChat.AddChat vbRed, IsBanned
        
        ' if not banned...
        If (IsBanned = False) Then
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Greet message
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
            If (BotVars.UseGreet) Then
                If (LenB(BotVars.GreetMsg)) Then
                    If (BotVars.WhisperGreet) Then
                        frmChat.AddQ "/w " & Username & _
                            Space$(1) & DoReplacements(BotVars.GreetMsg, Username, Ping)
                    Else
                        frmChat.AddQ DoReplacements(BotVars.GreetMsg, Username, Ping)
                    End If
                End If
            End If
                
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Botmail
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If (mail) Then
                L = GetMailCount(Username)
                
                If (L > 0) Then
                    frmChat.AddQ "/w " & Username & " You have " & L & _
                        " new message" & IIf(L = 1, "", "s") & ". Type !inbox to retrieve."
                End If
            End If
        End If
            
        ' print their statstring, if desired
        If (MDebug("statstrings")) Then
            frmChat.AddChat RTBColors.ErrorMessageText, Statstring
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' call event script function
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        On Error Resume Next
        
        'frmChat.AddChat vbRed, frmChat.SControl.Error.Number
        
        RunInAll "Event_UserJoins", Username, Flags, UserObj.Stats.ToString, Ping, _
            UserObj.Game, UserObj.Stats.Level, Statstring, IsBanned
    End If
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_UserJoins()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Public Sub Event_UserLeaves(ByVal Username As String, ByVal Flags As Long)
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If

    Dim UserObj   As clsUserObj
    
    Dim UserIndex As Integer
    Dim i         As Integer
    Dim ii        As Integer
    Dim Holder()  As Variant
    Dim pos       As Integer
    Dim bln       As Boolean

    UserIndex = g_Channel.GetUserIndexEx(CleanUsername(Username))
    
    If (UserIndex > 0) Then
        If (g_Channel.Users(UserIndex).IsOperator) Then
            g_Channel.RemoveBansFromOperator Username
        End If
        
        If (g_Channel.Users(UserIndex).Queue.Count = 0) Then
            If ((Not JoinMessagesOff) And ((Not Filters) Or (Not CheckBlock(Username)))) Then
                'If (GetVeto = False) Then
                Dim UserColor As Long
                
                ' display message
                If (Flags And USER_BLIZZREP) Then
                    UserColor = RGB(97, 105, 255)
                ElseIf (Flags And USER_SYSOP) Then
                    UserColor = RGB(97, 105, 255)
                ElseIf (Flags And USER_CHANNELOP) Then
                    UserColor = RTBColors.TalkUsernameOp
                Else
                    UserColor = RTBColors.JoinUsername
                End If
                
                frmChat.AddChat RTBColors.JoinText, "-- ", _
                    UserColor, g_Channel.Users(UserIndex).DisplayName, _
                    RTBColors.JoinText, " has left the channel."
                'End If
            End If
        End If
        
        g_Channel.Users.Remove UserIndex
    Else
        frmChat.AddChat RTBColors.ErrorMessageText, "Warning! We have received a leave event for a user that we didn't know " & _
                "was in the channel.  This may be indicative of a server split or other technical difficulty."
    
        Exit Sub
    End If
    
    If (StrComp(Username, g_Channel.OperatorHeir, vbTextCompare) = 0) Then
        g_Channel.OperatorHeir = vbNullString
        
        Call g_Channel.CheckUsers
    End If
    
    Username = ConvertUsername(Username)
    
    RemoveBanFromQueue Username
    
    pos = checkChannel(Username)
    
    If (pos > 0) Then
        If (frmChat.mnuFlash.Checked) Then
            FlashWindow
        End If
    
        With frmChat.lvChannel
            .ListItems.Remove pos

            .Refresh
        End With
        
        frmChat.ListviewTabs_Click 0
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' call event script function
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        On Error Resume Next
        
        RunInAll "Event_UserLeaves", CleanUsername(Username), Flags
    End If

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_UserLeaves()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Public Sub Event_UserTalk(ByVal Username As String, ByVal Flags As Long, ByVal Message As String, _
        ByVal Ping As Long, Optional QueuedEventID As Integer = 0)
    
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    Dim UserObj       As clsUserObj
    Dim UserEvent     As clsUserEventObj
    
    Dim strSend       As String
    Dim s             As String
    Dim U             As String
    Dim strCompare    As String
    Dim i             As Integer
    Dim ColIndex      As Integer
    Dim b             As Boolean
    Dim ToANSI        As String
    Dim BanningUser   As Boolean
    Dim UsernameColor As Long
    Dim TextColor     As Long
    Dim CaratColor    As Long
    Dim pos           As Integer
    Dim blnCheck      As Boolean
    
    pos = g_Channel.GetUserIndexEx(CleanUsername(Username))
    
    If (pos > 0) Then
        Set UserObj = g_Channel.Users(pos)
        
        UserObj.LastTalkTime = UtcNow
        
        If (QueuedEventID = 0) Then
            If (UserObj.Queue.Count > 0) Then
                Set UserEvent = New clsUserEventObj
                
                With UserEvent
                    .EventID = ID_TALK
                    .Flags = Flags
                    .Ping = Ping
                    .Message = Message
                End With
                
                UserObj.Queue.Add UserEvent
            End If
        End If
    Else
        ' create new user object for invisible representatives...
        Set UserObj = New clsUserObj
        
        ' store user name
        UserObj.Name = Username
    End If
    
    ' convert user name
    Username = UserObj.DisplayName
    
    If (QueuedEventID = 0) Then
        If (g_Channel.Self.IsOperator) Then
            If (GetSafelist(Username) = False) Then
                CheckMessage Username, Message
            End If
        End If
    End If
    
    If ((UserObj.Queue.Count = 0) Or (QueuedEventID > 0)) Then
        If (Message <> vbNullString) Then
            If (AllowedToTalk(Username, Message)) Then
                ' are we watching the user?
                'If (StrComp(WatchUser, Username, vbTextCompare) = 0) Then
                If (PrepareCheck(Username) Like PrepareCheck(WatchUser)) Then
                    UsernameColor = RTBColors.ErrorMessageText
                    
                ' is user an operator?
                ElseIf ((Flags And USER_CHANNELOP&) = USER_CHANNELOP&) Then
                    UsernameColor = RTBColors.TalkUsernameOp
                Else
                    UsernameColor = RTBColors.TalkUsernameNormal
                End If
                
                If (((Flags And USER_BLIZZREP&) = USER_BLIZZREP&) Or ((Flags And USER_SYSOP&) = _
                        USER_SYSOP&)) Then
                        
                    TextColor = RGB(97, 105, 255)
                    
                    CaratColor = RGB(97, 105, 255)
                Else
                    TextColor = RTBColors.TalkNormalText
                    
                    CaratColor = RTBColors.Carats
                End If
                
                'If (GetVeto = False) Then
                    frmChat.AddChat CaratColor, "<", UsernameColor, Username, CaratColor, "> ", _
                        TextColor, Message
                'End If
                
                If (Catch(0) <> vbNullString) Then
                    CheckPhrase Username, Message, CPTALK
                End If
                    
                If (frmChat.mnuFlash.Checked) Then
                    FlashWindow
                End If
            End If
        End If
        
        If (VoteDuration > 0) Then
            If (InStr(1, LCase(Message), "yes", vbTextCompare) > 0) Then
                Call Voting(BVT_VOTE_ADD, BVT_VOTE_ADDYES, Username)
            ElseIf (InStr(1, LCase(Message), "no", vbTextCompare) > 0) Then
                Call Voting(BVT_VOTE_ADD, BVT_VOTE_ADDNO, Username)
            End If
        End If
        
        Call ProcessCommand(Username, Message, False, False)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' call event script function
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        On Error Resume Next
        
        If ((BotVars.NoSupportMultiCharTrigger) And (Len(BotVars.TriggerLong) > 1)) Then
            If (StrComp(Left$(Message, Len(BotVars.TriggerLong)), BotVars.TriggerLong, _
                    vbBinaryCompare) = 0) Then
                
                Message = BotVars.Trigger & Mid$(Message, Len(BotVars.TriggerLong) + 1)
            End If
        End If
    
        RunInAll "Event_UserTalk", Username, Flags, Message, Ping
    End If
    

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_UserTalk()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Private Function CheckMessage(Username As String, Message As String) As Boolean
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    Dim BanningUser As Boolean
    Dim i           As Integer
    
    If (PhraseBans) Then
        For i = LBound(Phrases) To UBound(Phrases)
            If ((Phrases(i) <> vbNullString) And (Phrases(i) <> Space$(1))) Then
                If ((InStr(1, Message, Phrases(i), vbTextCompare)) <> 0) Then
                    Ban Username & " Banned phrase: " & Phrases(i), _
                            (AutoModSafelistValue - 1), Abs(CLng(Config.PhraseKick))
                    
                    BanningUser = True
                    
                    Exit For
                End If
            End If
        Next i
    End If
    
    If (BanningUser = False) Then
        If (BotVars.QuietTime) Then
            Ban Username & " Quiet-time", (AutoModSafelistValue - 1), Abs(CLng(Config.QuietTimeKick))
        Else
            If (BotVars.KickOnYell = 1) Then
                If (Len(Message) > 5) Then
                    If (PercentActualUppercase(Message) > 90) Then
                        Ban Username & " Yelling", (AutoModSafelistValue - 1), 1
                    End If
                End If
            End If
        End If
        
        If ((BotVars.QuietTime) Or (BotVars.KickOnYell = 1)) Then
            BanningUser = True
        End If
    End If
    
    CheckMessage = BanningUser
    
    Exit Function
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.CheckMessage()", Err.Number, Err.Description, OBJECT_NAME))
End Function

Public Sub Event_VersionCheck(Message As Long, ExtraInfo As String)
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    Select Case (Message)
        Case 0:
            frmChat.AddChat RTBColors.SuccessText, "[BNCS] Client version accepted!"
            
            ' if using server finder
            If ((BotVars.BNLS) And (BotVars.UseAltBnls)) Then
                ' save BNLS server so future instances of the bot won't need to get the list, connection succeeded
                If Config.BNLSServer <> BotVars.BNLSServer Then
                    Config.BNLSServer = BotVars.BNLSServer
                    Call Config.Save
                End If
            End If
        
        Case 1:
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Version check failed! " & _
                "The version byte for this attempt was 0x" & Hex(GetVerByte(BotVars.Product)) & "." & _
                IIf(LenB(ExtraInfo) = 0, vbNullString, " Extra Information: " & ExtraInfo)

            If (BotVars.BNLS) Then
                If (frmChat.HandleBnlsError("[BNCS] BNLS has not been updated yet, " & _
                        "or you experienced an error. Try connecting again.")) Then
                    ' if we are using the finder, then don't close all connections
                    Message = 0
                End If
            Else
                frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Please ensure you " & _
                    "have updated your hash files using more current ones from the directory " & _
                        "of the game you're connecting with."
                
                frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] In addition, you can try " & _
                    "choosing ""Update version bytes from StealthBot.net"" from the Bot menu."
            End If
        
        Case 2:
            frmChat.AddChat RTBColors.SuccessText, "[BNCS] Version check passed!"
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Your CD-key is invalid!"
        
        Case 3:
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Version check failed! " & _
                "BNLS has not been updated yet.. Try reconnecting in an hour or two."
        
        Case 4:
            frmChat.AddChat RTBColors.SuccessText, "[BNCS] Version check passed!"
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Your CD-key is for another game."
        
        Case 5:
            frmChat.AddChat RTBColors.SuccessText, "[BNCS] Version check passed!"
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Your CD-key is banned. " & _
                "For more information, visit http://us.blizzard.com/support/article.xml?locale=en_US&articleId=20637 ."
        
        Case 6:
            frmChat.AddChat RTBColors.SuccessText, "[BNCS] Version check passed!"
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Your CD-key is currently in " & _
                "use under the owner name: " & ExtraInfo & "."
        
        Case 7:
            frmChat.AddChat RTBColors.SuccessText, "[BNCS] Version check passed!"
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Your expansion CD-key is invalid."
        
        Case 8:
            frmChat.AddChat RTBColors.SuccessText, "[BNCS] Version check passed!"
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Your expansion CD-key is currently " & _
                "in use under the owner name: " & ExtraInfo & "."
        
        Case 9:
            frmChat.AddChat RTBColors.SuccessText, "[BNCS] Version check passed!"
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Your expansion CD-key is banned. " & _
                "For more information, visit http://us.blizzard.com/support/article.xml?locale=en_US&articleId=20637 ."
        
        Case 10:
            frmChat.AddChat RTBColors.SuccessText, "[BNCS] Version check passed!"
            
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Your expansion CD-key is for another game."
        
        Case Else
            frmChat.AddChat RTBColors.ErrorMessageText, "Unhandled 0x51 response! Value: " & Message
    End Select
    
    If (Message > 0) Then
        Call frmChat.DoDisconnect
    End If
    
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_VersionCheck()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

Public Sub Event_WhisperFromUser(ByVal Username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    'Dim s       As String
    Dim lCarats As Long
    Dim WWIndex As Integer
    
    Username = ConvertUsername(Username)
    
    If (Catch(0) <> vbNullString) Then
        Call CheckPhrase(Username, Message, CPWHISPER)
    End If
    
    If (frmChat.mnuFlash.Checked) Then
        FlashWindow
    End If
    
    If (StrComp(Message, BotVars.ChannelPassword, vbTextCompare) = 0) Then
        lCarats = g_Channel.GetUserIndex(Username)
        
        If (lCarats > 0) Then
            With g_Channel.Users(lCarats)
                .PassedChannelAuth = True
            End With
            
            frmChat.AddQ "/w " & Username & " Password accepted."
        End If
    End If
    
    If (VoteDuration > 0) Then
        If (InStr(1, Message, "yes", vbTextCompare) > 0) Then
            Call Voting(BVT_VOTE_ADD, BVT_VOTE_ADDYES, Username)
        ElseIf (InStr(1, Message, "no", vbTextCompare) > 0) Then
            Call Voting(BVT_VOTE_ADD, BVT_VOTE_ADDNO, Username)
        End If
    End If
            
    lCarats = RTBColors.WhisperCarats
    
    If (Flags And &H1) Then
        lCarats = COLOR_BLUE
    End If
    
    '####### Mail check
    If (mail) Then
        If (StrComp(Left$(Message, 6), "!inbox", vbTextCompare) = 0) Then
            Dim Msg As udtMail
            
            If (GetMailCount(Username) > 0) Then
                Call GetMailMessage(Username, Msg)
                
                If (Len(RTrim(Msg.To)) > 0) Then
                    frmChat.AddQ "/w " & Username & " Message from " & _
                        RTrim$(Msg.From) & ": " & RTrim$(Msg.Message)
                End If
            End If
        End If
    End If
    '#######
    
    If ((Not Filters) Or ((Not (CheckMsg(Message, Username, -5))) And (Not (CheckBlock(Username))))) Then
    
        If (Not (frmChat.mnuHideWhispersInrtbChat.Checked)) Then
            frmChat.AddChat lCarats, "<From ", RTBColors.WhisperUsernames, _
                Username, lCarats, "> ", RTBColors.WhisperText, Message
        End If
        
        frmChat.AddWhisper lCarats, "<From ", RTBColors.WhisperUsernames, _
            Username, lCarats, "> ", RTBColors.WhisperText, Message
            
        frmChat.rtbWhispers.Visible = rtbWhispersVisible
                       
        If (frmChat.mnuToggleWWUse.Checked) Then
        'If ((frmChat.mnuToggleWWUse.Checked) And _
            '(frmChat.WindowState <> vbMinimized)) Then
            
            If (Not (IrrelevantWhisper(Message, Username))) Then
                WWIndex = AddWhisperWindow(Username)
                
                With colWhisperWindows.Item(WWIndex)
                    If (.Shown = False) Then
                        'window was previously hidden
                        
                        ShowWW WWIndex
                    End If
                    
                    .Caption = "Whisper Window: " & Username
                    
                    .AddWhisper RTBColors.WhisperUsernames, "> " & Username, lCarats, _
                        ": ", RTBColors.WhisperText, Message
                End With
            End If
        End If
    
        Call ProcessCommand(Username, Message, False, True)
    End If
    
    If (Not (CheckBlock(Username))) Then
        LastWhisper = Username
        LastWhisperFromTime = Now
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' call event script function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If BotIsClosing Then Exit Sub
    
    On Error Resume Next
    
    g_lastQueueUser = Username
    
    RunInAll "Event_WhisperFromUser", Username, Flags, Message, Ping
    'End If

    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_WhisperFromUser()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

' Flags and ping are deliberately not used at this time
Public Sub Event_WhisperToUser(ByVal Username As String, ByVal Flags As Long, ByVal Message As String, ByVal Ping As Long)
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    Dim WWIndex As Integer
    
    'frmChat.AddChat vbRed, Username
    
    If (StrComp(Username, "your friends", vbTextCompare) <> 0) Then
        Username = ConvertUsername(Username)
        
        LastWhisperTo = Username
    Else
        LastWhisperTo = "%f%"
    End If

    If (Not (frmChat.mnuHideWhispersInrtbChat.Checked)) Then
        frmChat.AddChat RTBColors.WhisperCarats, "<To ", RTBColors.WhisperUsernames, _
            Username, RTBColors.WhisperCarats, "> ", RTBColors.WhisperText, Message
    End If
    
    If ((frmChat.mnuHideWhispersInrtbChat.Checked) Or _
        (frmChat.mnuToggleShowOutgoing.Checked)) Then
        
        frmChat.AddWhisper RTBColors.WhisperCarats, "<To ", RTBColors.WhisperUsernames, _
            Username, RTBColors.WhisperCarats, "> ", RTBColors.WhisperText, Message
    End If

    If (frmChat.mnuToggleWWUse.Checked) Then
        If ((InStr(1, Message, "ß~ß") = 0) And _
            (StrComp(Username, "your friends") <> 0)) Then
            
            WWIndex = AddWhisperWindow(Username)
            
            If (frmChat.WindowState <> vbMinimized) Then
                Call ShowWW(WWIndex)
            End If
            
            colWhisperWindows.Item(WWIndex).Caption = "Whisper Window: " & Username
            colWhisperWindows.Item(WWIndex).AddWhisper RTBColors.TalkBotUsername, "> " & _
                GetCurrentUsername, RTBColors.WhisperCarats, ": ", RTBColors.WhisperText, Message
        End If
    End If
    
    If (Not (rtbWhispersVisible)) Then
        If (frmChat.rtbWhispers.Visible = True) Then
            frmChat.rtbWhispers.Visible = False
        End If
    End If
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_WhisperToUser()", Err.Number, Err.Description, OBJECT_NAME))
End Sub


'11/22/07 - Hdx - Pass the channel listing (0x0B) directly off to scriptors for there needs. (What other use is there?)
Public Sub Event_ChannelList(sChannels() As String)
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    Dim x As Integer
    Dim sChannel As String
        
    If (MDebug("all")) Then
        frmChat.AddChat RTBColors.InformationText, "Received Channel List: "
    End If
    
    ' save public channels
    Set BotVars.PublicChannels = New Collection
    
    For x = 0 To UBound(sChannels)
        sChannel = sChannels(x)

        If LenB(sChannel) > 0 Then
            BotVars.PublicChannels.Add sChannel
        End If
    Next x
    
    PreparePublicChannelMenu
    
    RunInAll "Event_ChannelList", ConvertStringArray(sChannels)
    Exit Sub
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_ChannelList()", Err.Number, Err.Description, OBJECT_NAME))
End Sub

'10/01/09 - Hdx - This is for SID_MESSAGEBOX, for now it'll raise it's own event, and Event_ServerError
Public Function Event_MessageBox(lStyle As Long, sText As String, sCaption As String)
On Error GoTo ERROR_HANDLER:
    Call Event_ServerError(sText)
    
    RunInAll "Event_MessageBox", lStyle, sText, sCaption

    Exit Function
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.Event_MessageBox()", Err.Number, Err.Description, OBJECT_NAME))
End Function

Public Function CleanUsername(ByVal Username As String, Optional ByVal PrependNamingStar As Boolean = False) As String
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    Dim tmp As String
    Dim pos As Integer
    
    tmp = Username
    
    If (tmp <> vbNullString) Then
        pos = InStr(1, tmp, "*", vbBinaryCompare)
    
        If (pos > 0) Then
            If (Right$(tmp, 1) = ")") Then
                ' fixed so that usernames actually ending in
                ' ")" don't get trimmed (ultimately messing up
                ' such bots with ops). ~Ribose/2009-11-15
                If pos > 3 Then
                    ' blah (*blah)
                    '     ^^^
                    If Mid$(tmp, pos - 2, 3) = " (*" Then
                        tmp = Left$(tmp, Len(tmp) - 1)
                    End If
                End If
            End If
            
            tmp = Mid$(tmp, pos + 1)
        End If
    End If
    
    If (Dii And PrependNamingStar And BotVars.UseD2Naming = False) Then
        tmp = "*" & tmp
    End If

    CleanUsername = tmp
    
    Exit Function
ERROR_HANDLER:
    Call frmChat.AddChat(RTBColors.ErrorMessageText, _
        StringFormat("Error: #{0}: {1} in {2}.CleanUsername()", Err.Number, Err.Description, OBJECT_NAME))
End Function

'Private Function GetDiablo2CharacterName(ByVal Username As String) As String
'
'    Dim tmp As String
'    Dim Pos As Integer
'
'    Pos = InStr(1, Username, "*", vbBinaryCompare)
'
'    If (Pos > 0) Then
'        tmp = Mid$(Username, 1, Pos - 1)
'    End If
'
'    GetDiablo2CharacterName = tmp
'
'End Function
