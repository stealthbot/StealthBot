Attribute VB_Name = "modCommandsChat"
Option Explicit
'This module will contain all the command code for commands that affect the chat environment
'Things like Say, DND, Squelch, etc...
Private m_AwayMessage As String

Public Function OnAway(Command As clsCommandObj) As Boolean
    If (LenB(m_AwayMessage) > 0) Then
        Call frmChat.AddQ("/away", PRIORITY.COMMAND_RESPONSE_MESSAGE)
        If (Not Command.IsLocal) Then
            If (m_AwayMessage = " - ") Then
                Call frmChat.AddQ("/me is back.", PRIORITY.COMMAND_RESPONSE_MESSAGE)
            Else
                Call frmChat.AddQ("/me is back from (" & m_AwayMessage & ")", PRIORITY.COMMAND_RESPONSE_MESSAGE)
            End If
        End If
        m_AwayMessage = vbNullString
    Else
        If (LenB(Command.Argument("Message")) > 0) Then
            m_AwayMessage = Command.Argument("Message")
            Call frmChat.AddQ("/away " & m_AwayMessage, PRIORITY.COMMAND_RESPONSE_MESSAGE)
            
            If (Command.IsLocal) Then Call frmChat.AddQ("/me is away (" & m_AwayMessage & ")")
        Else
            m_AwayMessage = " - "
            Call frmChat.AddQ("/away", PRIORITY.COMMAND_RESPONSE_MESSAGE)
            If (Not Command.IsLocal) Then Call frmChat.AddQ("/me is away.")
        End If
    End If
End Function

Public Function OnBack(Command As clsCommandObj) As Boolean
    If (LenB(m_AwayMessage) > 0) Then
        Call frmChat.AddQ("/away", PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
        
        If (Not Command.IsLocal) Then
            If (m_AwayMessage = " - ") Then
                Call frmChat.AddQ("/me is back.", PRIORITY.COMMAND_RESPONSE_MESSAGE)
            Else
                Call frmChat.AddQ("/me is back from (" & m_AwayMessage & ")", PRIORITY.COMMAND_RESPONSE_MESSAGE)
            End If
            m_AwayMessage = vbNullString
        End If
    End If
End Function

Public Function OnBlock(Command As clsCommandObj) As Boolean
    Dim FiltersPath As String
    Dim Total       As Integer
    
    If (Command.IsValid) Then
        FiltersPath = GetFilePath("Filters.ini")
        If (CheckBlock(Command.Argument("Username"))) Then 'Should prevent adding filters that are the same as other filters
            Command.Respond "That username is already in the block list, or is under a wildcard block."
        Else
            If (StrictIsNumeric(ReadINI("BlockList", "Total", FiltersPath))) Then
                Total = Int(ReadINI("BlockList", "Total", FiltersPath))
                WriteINI "BlockList", "Filter" & (Total + 1), Command.Argument("Username"), FiltersPath
                WriteINI "BlockList", "Total", Total + 1, FiltersPath
                Command.Respond StringFormatA("Added {0}{1}{0} to the username block list.", Chr$(34), Command.Argument("Username"))
            Else
                Command.Respond "Your filters file has been edited manually and is no longer valid. Please delete it."
            End If
        End If
    End If
End Function

Public Function OnConnect(Command As clsCommandObj) As Boolean
    frmChat.DoConnect
End Function

Public Function OnCQ(Command As clsCommandObj) As Boolean
    g_Queue.Clear
    Command.Respond "Queue cleared."
End Function

Public Function OnDisconnect(Command As clsCommandObj) As Boolean
    frmChat.DoDisconnect
End Function

Public Function OnExpand(Command As clsCommandObj) As Boolean
    Dim sMessage As String
    Dim tmpSend As String
    Dim I As Integer
    If (Command.IsValid) Then
        sMessage = Command.Argument("Message")
        For I = 1 To Len(sMessage)
            tmpSend = StringFormatA("{0}{1}{2}", tmpSend, Mid$(sMessage, I, 1), IIf(I = Len(sMessage), vbNullString, Space$(1)))
        Next I
        
        If (Len(tmpSend) > 223) Then
            tmpSend = Left$(tmpSend, 223)
        End If
        If (StrComp(ReadCfg("Override", "ForcePublicSay"), "Y", vbTextCompare) = 0) Then
            Call frmChat.AddQ(tmpSend, PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
        Else
            Command.Respond tmpSend
        End If
    End If
End Function

Public Function OnFAdd(Command As clsCommandObj) As Boolean
    If (Command.IsValid) Then
        Call frmChat.AddQ("/f a " & Command.Argument("Username"), PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
        Command.Respond StringFormatA("Added user {0}{1}{0} to this account's friends list.", Chr$(34), Command.Argument("Username"))
    End If
End Function

Public Function OnFilter(Command As clsCommandObj) As Boolean
    Dim FiltersPath As String
    Dim Total       As Integer
    
    If (Command.IsValid) Then
        FiltersPath = GetFilePath("Filters.ini")
        If (CheckMsg(Command.Argument("Filter"))) Then 'Should prevent adding filters that are the same as other filters
            Command.Respond "That filter is already in the list, or is under a wildcard."
        Else
            If (StrictIsNumeric(ReadINI("TextFilters", "Total", FiltersPath))) Then
                Total = Int(ReadINI("TextFilters", "Total", FiltersPath))
                WriteINI "TextFilters", "Filter" & (Total + 1), Command.Argument("Filter"), FiltersPath
                WriteINI "TextFilters", "Total", Total + 1, FiltersPath
                Command.Respond StringFormatA("Added {0}{1}{0} to the message filter list.", Chr$(34), Command.Argument("Filter"))
            Else
                Command.Respond "Your filters file has been edited manually and is no longer valid. Please delete it."
            End If
        End If
    End If
End Function

Public Function OnForceJoin(Command As clsCommandObj) As Boolean
    If (Command.IsValid) Then
        Call FullJoin(Command.Argument("Channel"))
    End If
End Function

Public Function OnFRem(Command As clsCommandObj) As Boolean
    If (Command.IsValid) Then
        Call frmChat.AddQ("/f r " & Command.Argument("Username"), PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
        Command.Respond StringFormatA("Removed user {0}{1}{0} from this account's friends list.", Chr$(34), Command.Argument("Username"))
    End If
End Function

Public Function OnHome(Command As clsCommandObj) As Boolean
    Call frmChat.AddQ("/join " & BotVars.HomeChannel, PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
End Function

Public Function OnIgnore(Command As clsCommandObj) As Boolean
    If (Command.IsValid) Then
        If (Command.IsLocal) Then
            Command.Respond StringFormatA("Ignoring messages from {0}{1}{0}.", Chr$(34), Command.Argument("Username"))
            frmChat.AddQ "/ignore " & Command.Argument("Username"), PRIORITY.COMMAND_RESPONSE_MESSAGE
        Else
            Dim dbTarget As udtGetAccessResponse
            Dim dbCaller As udtGetAccessResponse
            dbTarget = GetCumulativeAccess(Command.Argument("Username"))
            dbCaller = GetCumulativeAccess(Command.Username)
            
            If (dbTarget.Rank > dbCaller.Rank) Then
                Command.Respond "That user has a higher rank then you."
            ElseIf (dbTarget.Rank = dbCaller.Rank) Then
                Command.Respond "That user has the same rank as you."
            Else
                Command.Respond StringFormatA("Ignoring messages from {0}{1}{0}.", Chr$(34), Command.Argument("Username"))
                frmChat.AddQ "/ignore " & Command.Argument("Username"), PRIORITY.COMMAND_RESPONSE_MESSAGE
            End If
        End If
    End If
End Function

Public Function OnIgPriv(Command As clsCommandObj) As Boolean
    Call frmChat.AddQ("/o igpriv", PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
    Command.Respond "Ignoring messages from non-friends in private channels."
End Function

Public Function OnJoin(Command As clsCommandObj) As Boolean
    If (Command.IsValid) Then
        Call frmChat.AddQ("/join " & Command.Argument("Channel"), PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
    End If
End Function

Public Function OnUnBlock(Command As clsCommandObj) As Boolean
    
    Dim I           As Integer
    Dim Total       As Integer
    Dim FiltersPath As String
    
    If (Command.IsValid) Then
        FiltersPath = GetFilePath("Filters.ini")
        If (StrictIsNumeric(ReadINI("BlockList", "Total", FiltersPath))) Then
            Total = Int(ReadINI("BlockList", "Total", FiltersPath))
            For I = 0 To Total
                If (StrComp(Command.Argument("Username"), ReadINI("BlockList", "Filter" & I, FiltersPath), vbTextCompare) = 0) Then
                    Exit For
                ElseIf (I = Total) Then
                    Command.Respond StringFormatA("{0}{1}{0} is not blocked.", Chr(34), Command.Argument("Username"))
                    Exit Function
                End If
            Next I
            
            If (I < Total) Then
                For I = I To (Total - 1)
                    WriteINI "BlockList", "Filter" & I, ReadINI("BlockList", "Filter" & (I + 1), FiltersPath), FiltersPath
                Next I
                WriteINI "BlockList", "Total", (Total - 1), FiltersPath
                Command.Respond StringFormatA("Removed {0}{1}{0} from the blocked users list.", Chr$(34), Command.Argument("Username"))
            End If
        Else
            Command.Respond "Your filters file has been edited manually and is no longer valid. Please delete it."
        End If
    End If
End Function

Public Function OnUnFilter(Command As clsCommandObj) As Boolean
    Dim z As String
    Dim I As Integer
    Dim Total As Integer
    Dim FiltersPath As String
    
    If (Command.IsValid) Then
        If (StrictIsNumeric(ReadINI("TextFilters", "Total", FiltersPath))) Then
            Total = Int(ReadINI("TextFilters", "Total", FiltersPath))
            For I = 0 To Total
                If (StrComp(Command.Argument("Filter"), ReadINI("TextFilters", "Filter" & I, FiltersPath), vbTextCompare) = 0) Then
                    Exit For
                ElseIf (I = Total) Then
                    Command.Respond StringFormatA("{0}{1}{0} is not filtered.", Chr(34), Command.Argument("Filter"))
                    Exit Function
                End If
            Next I
    
            If (I < Total) Then
                For I = I To (Total - 1)
                    WriteINI "TextFilters", "Filter" & I, ReadINI("TextFilters", "Filter" & (I + 1), FiltersPath), FiltersPath
                Next I
                WriteINI "TextFilters", "Total", (Total - 1), FiltersPath
                Command.Respond StringFormatA("Removed {0}{1}{0} from the message filter list.", Chr$(34), Command.Argument("Filter"))
                Call frmChat.LoadArray(LOAD_FILTERS, gFilters())
            End If
        Else
            Command.Respond "Your filters file has been edited manually and is no longer valid. Please delete it."
        End If
    End If
End Function

Public Function OnUnIgnore(Command As clsCommandObj) As Boolean
    If (Command.IsValid) Then
        Command.Respond StringFormatA("Receiving messages from {0}{1}{0}.", Chr$(34), Command.Argument("Username"))
        frmChat.AddQ "/unignore " & Command.Argument("Username"), PRIORITY.COMMAND_RESPONSE_MESSAGE
    End If
End Function

Public Function OnUnIgPriv(Command As clsCommandObj) As Boolean
    Call frmChat.AddQ("/o unigpriv", PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
    Command.Respond "Allowing messages from non-friends in private channels."
End Function

Public Function OnQuickRejoin(Command As clsCommandObj) As Boolean
    Call RejoinChannel(g_Channel.Name)
End Function

Public Function OnReconnect(Command As clsCommandObj) As Boolean
    If (g_Online) Then
        Dim temp As String
        temp = BotVars.HomeChannel
    
        BotVars.HomeChannel = g_Channel.Name
        
        Call frmChat.DoDisconnect
        
        frmChat.AddChat RTBColors.ErrorMessageText, "[BNET] Reconnecting by command, please wait..."
        
        Pause 1
        
        frmChat.AddChat RTBColors.SuccessText, "Connection initialized."
        
        Call frmChat.DoConnect
        
        Pause 3, True
        
        BotVars.HomeChannel = temp
    Else
        frmChat.AddChat RTBColors.SuccessText, "Connection initialized."
        
        Call frmChat.DoConnect
    End If
End Function

Public Function OnReJoin(Command As clsCommandObj) As Boolean
    Call frmChat.AddQ(StringFormatA("/join {0} Rejoin", GetCurrentUsername), PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
    Call frmChat.AddQ("/join " & g_Channel.Name, PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
End Function

Public Function OnSay(Command As clsCommandObj) As Boolean
    Dim tmpSend As String
    
    If (Command.IsValid) Then
        If (Len(Command.Argument("Message")) > 223) Then
            tmpSend = Left$(Command.Argument("Message"), 223)
        Else
            tmpSend = Command.Argument("Message")
        End If
    
        If (StrComp(ReadCfg("Override", "ForcePublicSay"), "Y", vbTextCompare) = 0) Then
            Call frmChat.AddQ(tmpSend, PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
        Else
            Command.Respond tmpSend
        End If
    End If
End Function

Public Function OnSCQ(Command As clsCommandObj) As Boolean
    g_Queue.Clear
End Function

Public Function OnShout(Command As clsCommandObj) As Boolean
    Dim tmpSend As String
    
    If (Command.IsValid) Then
        If (Len(Command.Argument("Message")) > 223) Then
            tmpSend = UCase$(Left$(Command.Argument("Message"), 223))
        Else
            tmpSend = UCase$(Command.Argument("Message"))
        End If
    
        If (StrComp(ReadCfg("Override", "ForcePublicSay"), "Y", vbTextCompare) = 0) Then
            Call frmChat.AddQ(tmpSend, PRIORITY.COMMAND_RESPONSE_MESSAGE, Command.Username)
        Else
            Command.Respond tmpSend
        End If
    End If
End Function

Public Function OnWatch(Command As clsCommandObj) As Boolean
    If (Command.IsValid) Then
        WatchUser = Command.Argument("Username")
        Command.Respond "Now watching " & Command.Argument("Username")
    Else
        If (LenB(WatchUser) > 0) Then
            Command.Respond "Stoped watching " & WatchUser
            WatchUser = vbNullString
        End If
    End If
End Function

Public Function OnWatchOff(Command As clsCommandObj) As Boolean
    If (LenB(WatchUser) > 0) Then
        Command.Respond "Stoped watching " & WatchUser
        WatchUser = vbNullString
    Else
        Command.Respond "Watch is disabled."
    End If
End Function


