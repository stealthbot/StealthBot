Attribute VB_Name = "modCommandsChat"
Option Explicit
'This module will contain all the command code for commands that affect the chat environment
'Things like Say, DND, Squelch, etc...
Private m_AwayMessage As String

Public Sub OnAway(Command As clsCommandObj)
    If (LenB(m_AwayMessage) > 0) Then
        Call frmChat.AddQ("/away", enuPriority.COMMAND_RESPONSE_MESSAGE)
        If (Not Command.IsLocal) Then
            If (m_AwayMessage = " - ") Then
                Call frmChat.AddQ("/me is back.", enuPriority.COMMAND_RESPONSE_MESSAGE)
            Else
                Call frmChat.AddQ("/me is back from (" & m_AwayMessage & ")", enuPriority.COMMAND_RESPONSE_MESSAGE)
            End If
        End If
        m_AwayMessage = vbNullString
    Else
        If (LenB(Command.Argument("Message")) > 0) Then
            m_AwayMessage = Command.Argument("Message")
            Call frmChat.AddQ("/away " & m_AwayMessage, enuPriority.COMMAND_RESPONSE_MESSAGE)
            
            If (Command.IsLocal) Then Call frmChat.AddQ("/me is away (" & m_AwayMessage & ")")
        Else
            m_AwayMessage = " - "
            Call frmChat.AddQ("/away", enuPriority.COMMAND_RESPONSE_MESSAGE)
            If (Not Command.IsLocal) Then Call frmChat.AddQ("/me is away.")
        End If
    End If
End Sub

Public Sub OnBack(Command As clsCommandObj)
    If (LenB(m_AwayMessage) > 0) Then
        Call frmChat.AddQ("/away", enuPriority.COMMAND_RESPONSE_MESSAGE, Command.Username)
        
        If (Not Command.IsLocal) Then
            If (m_AwayMessage = " - ") Then
                Call frmChat.AddQ("/me is back.", enuPriority.COMMAND_RESPONSE_MESSAGE)
            Else
                Call frmChat.AddQ("/me is back from (" & m_AwayMessage & ")", enuPriority.COMMAND_RESPONSE_MESSAGE)
            End If
            m_AwayMessage = vbNullString
        End If
    End If
End Sub

Public Sub OnBlock(Command As clsCommandObj)
    Dim FiltersPath As String
    Dim Total       As Integer
    Dim TotalString As String
    
    If (Command.IsValid) Then
        FiltersPath = GetFilePath(FILE_FILTERS)
        If (CheckBlock(Command.Argument("Username"))) Then 'Should prevent adding filters that are the same as other filters
            Command.Respond "That username is already in the block list, or is under a wildcard block."
        Else
            TotalString = ReadINI("BlockList", "Total", FiltersPath)
            If (LenB(TotalString) = 0) Then TotalString = "0"
            
            If (StrictIsNumeric(TotalString)) Then
                Total = Int(TotalString)
                WriteINI "BlockList", "Filter" & (Total + 1), Command.Argument("Username"), FiltersPath
                WriteINI "BlockList", "Total", Total + 1, FiltersPath
                Command.Respond StringFormat("Added {0}{1}{0} to the username block list.", Chr$(34), Command.Argument("Username"))
                
                Call frmChat.LoadArray(LOAD_BLOCKLIST, g_Blocklist())
            Else
                Command.Respond "Your filters file has been edited manually and is no longer valid. Please delete it."
            End If
        End If
    End If
End Sub

Public Sub OnConnect(Command As clsCommandObj)
    frmChat.DoConnect
End Sub

Public Sub OnCQ(Command As clsCommandObj)
    g_Queue.Clear
    Command.Respond "Queue cleared."
End Sub

Public Sub OnDisconnect(Command As clsCommandObj)
    Call frmChat.DoDisconnect
End Sub

Public Sub OnExpand(Command As clsCommandObj)
    Dim sMessage As String
    Dim tmpSend As String
    Dim i As Integer
    If (Command.IsValid) Then
        sMessage = Command.Argument("Message")
        
        If (Len(tmpSend) > 223) Then tmpSend = Left$(tmpSend, 223)
        For i = 1 To Len(sMessage)
            tmpSend = StringFormat("{0}{1}{2}", tmpSend, Mid$(sMessage, i, 1), IIf(i = Len(sMessage), vbNullString, Space$(1)))
        Next i
        
        If (Not Command.Restriction("RAW_USAGE")) Then tmpSend = StringFormat("{0} Says: {1}", Command.Username, tmpSend)
        
        If (Len(tmpSend) > 223) Then
            tmpSend = Left$(tmpSend, 223)
        End If
        
        Command.PublicOutput = True
        Command.Respond tmpSend
    End If
End Sub

Public Sub OnFAdd(Command As clsCommandObj)
    If (Command.IsValid) Then
        Call frmChat.AddQ("/f a " & Command.Argument("Username"), enuPriority.COMMAND_RESPONSE_MESSAGE, Command.Username)
        Command.Respond StringFormat("Added user {0}{1}{0} to this account's friends list.", Chr$(34), Command.Argument("Username"))
    End If
End Sub

Public Sub OnFilter(Command As clsCommandObj)
    Dim FiltersPath As String
    Dim Total       As Integer
    Dim TotalString As String
    
    If (Command.IsValid) Then
        FiltersPath = GetFilePath(FILE_FILTERS)
        If (CheckMsg(Command.Argument("Filter"))) Then 'Should prevent adding filters that are the same as other filters
            Command.Respond "That filter is already in the list, or is under a wildcard."
        Else
            TotalString = ReadINI("TextFilters", "Total", FiltersPath)
            If (LenB(TotalString) = 0) Then TotalString = "0"
            
            If (StrictIsNumeric(TotalString)) Then
                Total = Int(TotalString)
                WriteINI "TextFilters", "Filter" & (Total + 1), Command.Argument("Filter"), FiltersPath
                WriteINI "TextFilters", "Total", Total + 1, FiltersPath
                Command.Respond StringFormat("Added {0}{1}{0} to the message filter list.", Chr$(34), Command.Argument("Filter"))
                Call frmChat.LoadArray(LOAD_FILTERS, gFilters())
            Else
                Command.Respond "Your filters file has been edited manually and is no longer valid. Please delete it."
            End If
        End If
    End If
End Sub

Public Sub OnForceJoin(Command As clsCommandObj)
    If (Command.IsValid) Then
        Call FullJoin(Command.Argument("Channel"))
    End If
End Sub

Public Sub OnFRem(Command As clsCommandObj)
    If (Command.IsValid) Then
        Call frmChat.AddQ("/f r " & Command.Argument("Username"), enuPriority.COMMAND_RESPONSE_MESSAGE, Command.Username)
        Command.Respond StringFormat("Removed user {0}{1}{0} from this account's friends list.", Chr$(34), Command.Argument("Username"))
    End If
End Sub

Public Sub OnHome(Command As clsCommandObj)
    If (LenB(Config.HomeChannel) = 0) Then
        ' do product home join instead
        Call DoChannelJoinProductHome
    Else
        ' go home
        Call FullJoin(Config.HomeChannel, 2)
    End If
End Sub

Public Sub OnReturn(Command As clsCommandObj)
    If (LenB(BotVars.LastChannel) > 0) Then
        ' go to last channel
        Call FullJoin(BotVars.LastChannel, 2)
    End If
End Sub

Public Sub OnIgnore(Command As clsCommandObj)
    If (Command.IsValid) Then
        If (Command.IsLocal) Then
            Command.Respond StringFormat("Ignoring messages from {0}{1}{0}.", Chr$(34), Command.Argument("Username"))
            frmChat.AddQ "/ignore " & Command.Argument("Username"), enuPriority.COMMAND_RESPONSE_MESSAGE
        Else
            Dim dbTarget As udtUserAccess
            Dim dbCaller As udtUserAccess
            dbTarget = Database.GetUserAccess(Command.Argument("Username"))
            dbCaller = Database.GetUserAccess(Command.Username)
            
            If (dbTarget.Rank > dbCaller.Rank) Then
                Command.Respond "That user has a higher rank then you."
            ElseIf (dbTarget.Rank = dbCaller.Rank) Then
                Command.Respond "That user has the same rank as you."
            Else
                Command.Respond StringFormat("Ignoring messages from {0}{1}{0}.", Chr$(34), Command.Argument("Username"))
                frmChat.AddQ "/ignore " & Command.Argument("Username"), enuPriority.COMMAND_RESPONSE_MESSAGE
            End If
        End If
    End If
End Sub

Public Sub OnIgPriv(Command As clsCommandObj)
    Call frmChat.AddQ("/o igpriv", enuPriority.COMMAND_RESPONSE_MESSAGE, Command.Username)
    Command.Respond "Ignoring messages from non-friends in private channels."
End Sub

Public Sub OnJoin(Command As clsCommandObj)
    If (Command.IsValid) Then
        Call frmChat.AddQ("/join " & Command.Argument("Channel"), enuPriority.COMMAND_RESPONSE_MESSAGE, Command.Username)
    End If
End Sub

Public Sub OnUnBlock(Command As clsCommandObj)
    
    Dim i           As Integer
    Dim Total       As Integer
    Dim TotalString As String
    Dim FiltersPath As String
    
    If (Command.IsValid) Then
        FiltersPath = GetFilePath(FILE_FILTERS)
        
        TotalString = ReadINI("BlockList", "Total", FiltersPath)
        If (LenB(TotalString) = 0) Then TotalString = "0"
        
        If (StrictIsNumeric(TotalString)) Then
            Total = Int(TotalString)
            For i = 0 To Total
                If (StrComp(Command.Argument("Username"), ReadINI("BlockList", "Filter" & (i + 1), FiltersPath), vbTextCompare) = 0) Then
                    Exit For
                ElseIf (i = Total) Then
                    Command.Respond StringFormat("{0}{1}{0} is not blocked.", Chr(34), Command.Argument("Username"))
                    Exit Sub
                End If
            Next i
            
            If (i < Total) Then
                For i = i To (Total - 1)
                    WriteINI "BlockList", "Filter" & (i + 1), ReadINI("BlockList", "Filter" & (i + 2), FiltersPath), FiltersPath
                Next i
                WriteINI "BlockList", "Total", (Total - 1), FiltersPath
                Command.Respond StringFormat("Removed {0}{1}{0} from the blocked users list.", Chr$(34), Command.Argument("Username"))
                Call frmChat.LoadArray(LOAD_BLOCKLIST, g_Blocklist())
            End If
        Else
            Command.Respond "Your filters file has been edited manually and is no longer valid. Please delete it."
        End If
    End If
End Sub

Public Sub OnUnFilter(Command As clsCommandObj)
    Dim i           As Integer
    Dim Total       As Integer
    Dim TotalString As String
    Dim FiltersPath As String
    
    If (Command.IsValid) Then
        FiltersPath = GetFilePath(FILE_FILTERS)
        
        TotalString = ReadINI("TextFilters", "Total", FiltersPath)
        If (LenB(TotalString) = 0) Then TotalString = "0"
        
        If (StrictIsNumeric(TotalString)) Then
            Total = Int(TotalString)
            For i = 0 To Total
                If (StrComp(Command.Argument("Filter"), ReadINI("TextFilters", "Filter" & (i + 1), FiltersPath), vbTextCompare) = 0) Then
                    Exit For
                ElseIf (i = Total) Then
                    Command.Respond StringFormat("{0}{1}{0} is not filtered.", Chr(34), Command.Argument("Filter"))
                    Exit Sub
                End If
            Next i
    
            If (i < Total) Then
                For i = i To (Total - 1)
                    WriteINI "TextFilters", "Filter" & (i + 1), ReadINI("TextFilters", "Filter" & (i + 2), FiltersPath), FiltersPath
                Next i
                WriteINI "TextFilters", "Total", (Total - 1), FiltersPath
                Command.Respond StringFormat("Removed {0}{1}{0} from the message filter list.", Chr$(34), Command.Argument("Filter"))
                Call frmChat.LoadArray(LOAD_FILTERS, gFilters())
            End If
        Else
            Command.Respond "Your filters file has been edited manually and is no longer valid. Please delete it."
        End If
    End If
End Sub

Public Sub OnUnIgnore(Command As clsCommandObj)
    If (Command.IsValid) Then
        Command.Respond StringFormat("Receiving messages from {0}{1}{0}.", Chr$(34), Command.Argument("Username"))
        frmChat.AddQ "/unignore " & Command.Argument("Username"), enuPriority.COMMAND_RESPONSE_MESSAGE
    End If
End Sub

Public Sub OnUnIgPriv(Command As clsCommandObj)
    Call frmChat.AddQ("/o unigpriv", enuPriority.COMMAND_RESPONSE_MESSAGE, Command.Username)
    Command.Respond "Allowing messages from non-friends in private channels."
End Sub

Public Sub OnQuickRejoin(Command As clsCommandObj)
    Call RejoinChannel(g_Channel.Name)
End Sub

Public Sub OnReconnect(Command As clsCommandObj)
    If (g_Online) Then
        Dim LastChannel As String
        
        If LenB(g_Channel.Name) = 0 And LenB(BotVars.LastChannel) > 0 Then
            ' already outside chat environment
            LastChannel = BotVars.LastChannel
        Else
            ' in chat room
            LastChannel = g_Channel.Name
        End If
        
        Call frmChat.DoDisconnect
        
        'frmChat.AddChat g_Color.ErrorMessageText, "[BNCS] Reconnecting by command, please wait..."
        
        Pause 1
        
        'frmChat.AddChat g_Color.SuccessText, "Connection initialized."
        
        Call frmChat.DoConnect
        
        ' reinstate last channel
        BotVars.LastChannel = LastChannel
        PrepareHomeChannelMenu
        PrepareQuickChannelMenu
    Else
        frmChat.AddChat g_Color.SuccessText, "Connection initialized."
        
        Call frmChat.DoConnect
    End If
End Sub

Public Sub OnReJoin(Command As clsCommandObj)
    Call frmChat.AddQ(StringFormat("/join {0} Rejoin", GetCurrentUsername), enuPriority.COMMAND_RESPONSE_MESSAGE, Command.Username)
    Call frmChat.AddQ("/join " & g_Channel.Name, enuPriority.COMMAND_RESPONSE_MESSAGE, Command.Username)
End Sub

Public Sub OnSay(Command As clsCommandObj)
    Dim tmpSend As String
    
    If (Command.IsValid) Then
        If (Not Command.Restriction("RAW_USAGE")) Then
            tmpSend = StringFormat("{0} Says: {1}", Command.Username, Command.Argument("Message"))
        Else
            tmpSend = Command.Argument("Message")
        End If
        
        If (Len(tmpSend) > 223) Then tmpSend = Left$(tmpSend, 223)
        
        Command.PublicOutput = True
        Command.Respond tmpSend
    End If
End Sub

Public Sub OnSCQ(Command As clsCommandObj)
    g_Queue.Clear
End Sub

Public Sub OnShout(Command As clsCommandObj)
    Dim tmpSend As String
    
    If (Command.IsValid) Then
    
        If (Not Command.Restriction("RAW_USAGE")) Then
            tmpSend = StringFormat("{0} Shouts: {1}", Command.Username, UCase$(Command.Argument("Message")))
        Else
            tmpSend = UCase$(Command.Argument("Message"))
        End If
        
        If (Len(tmpSend) > 223) Then tmpSend = Left$(tmpSend, 223)
    
        Command.PublicOutput = True
        Command.Respond tmpSend
    End If
End Sub

Public Sub OnWatch(Command As clsCommandObj)
    If (Command.IsValid) Then
        WatchUser = Command.Argument("Username")
        Command.Respond StringFormat("Now watching {0}{1}{0}.", Chr$(34), Command.Argument("Username"))
    Else
        If (LenB(WatchUser) > 0) Then
            Command.Respond StringFormat("Stopped watching {0}{1}{0}.", Chr$(34), WatchUser)
            WatchUser = vbNullString
        End If
    End If
End Sub

Public Sub OnWatchOff(Command As clsCommandObj)
    If (LenB(WatchUser) > 0) Then
        Command.Respond StringFormat("Stopped watching {0}{1}{0}.", Chr$(34), WatchUser)
        WatchUser = vbNullString
    Else
        Command.Respond "Watch is disabled."
    End If
End Sub


