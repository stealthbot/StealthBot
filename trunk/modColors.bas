Attribute VB_Name = "modColors"
Option Explicit

Public Type udtColorList
    ChannelListBack As Long
    ChannelListText As Long
    ChannelListSelf As Long
    ChannelListIdle As Long
    ChannelListSquelched As Long
    ChannelListOps As Long
    SendBoxesBack As Long
    SendBoxesText As Long
    ChannelLabelBack As Long
    ChannelLabelText As Long
    RTBBack As Long
End Type

Public Type udtColorListRTB
    TalkBotUsername As Long
    TalkUsernameNormal As Long
    TalkUsernameOp As Long
    TalkNormalText As Long
    Carats As Long
    InformationText As Long
    SuccessText As Long
    ErrorMessageText As Long
    TimeStamps As Long
    ServerInfoText As Long
    ConsoleText As Long
    WhisperCarats As Long
    WhisperUsernames As Long
    WhisperText As Long
    EmoteUsernames As Long
    EmoteText As Long
    JoinUsername As Long
    JoinText As Long
    JoinedChannelName As Long
    JoinedChannelText As Long
End Type

Public RTBColors As udtColorListRTB
Public FormColors As udtColorList

Public Sub SetFormColors()
    With FormColors
        frmChat.lvChannel.BackColor = .ChannelListBack
        frmChat.lvChannel.ForeColor = .ChannelListText
        frmChat.lvFriendList.BackColor = .ChannelListBack
        frmChat.lvFriendList.ForeColor = .ChannelListText
        frmChat.lvClanList.BackColor = .ChannelListBack
        frmChat.lvClanList.ForeColor = .ChannelListText
        frmChat.lblCurrentChannel.ForeColor = .ChannelLabelText
        frmChat.lblCurrentChannel.BackColor = .ChannelLabelBack
        frmChat.txtPost.ForeColor = .SendBoxesText
        frmChat.txtPre.ForeColor = .SendBoxesText '.SendBoxesText
        frmChat.cboSend.ForeColor = .SendBoxesText
        frmChat.txtPost.BackColor = .SendBoxesBack
        frmChat.txtPre.BackColor = .SendBoxesBack
        frmChat.rtbChat.BackColor = .RTBBack
        frmChat.cboSend.BackColor = .SendBoxesBack
    End With
End Sub

Public Sub GetColorLists(Optional sPath As String)
    Dim f As Integer
    f = FreeFile
    
    If sPath = vbNullString Then sPath = GetFilePath(FILE_COLORS)
    
    ' Initialize
    With FormColors
        .ChannelLabelBack = -1
        .ChannelLabelText = -1
        .ChannelListBack = -1
        .ChannelListText = -1
        .ChannelListSelf = -1
        .ChannelListIdle = -1
        .ChannelListSquelched = -1
        .ChannelListOps = -1
        .RTBBack = -1
        .SendBoxesBack = -1
        .SendBoxesText = -1
    End With
    
    With RTBColors
        .Carats = -1
        .ConsoleText = -1
        .EmoteText = -1
        .EmoteUsernames = -1
        .ErrorMessageText = -1
        .InformationText = -1
        .JoinedChannelName = -1
        .JoinedChannelText = -1
        .JoinText = -1
        .JoinUsername = -1
        .ServerInfoText = -1
        .SuccessText = -1
        .TalkBotUsername = -1
        .TalkNormalText = -1
        .TalkUsernameNormal = -1
        .TalkUsernameOp = -1
        .TimeStamps = -1
        .WhisperCarats = -1
        .WhisperText = -1
        .WhisperUsernames = -1
    End With
    
    
    'Attempt to read
    If LenB(Dir$(sPath)) > 0 Then
        Open sPath For Random As #f Len = 4
        
        Dim SCLFVer As Integer
        Dim i As Integer
        
        ' old size = 112
        ' new size = 128
        ' added four listview color settings ~Ribose/2010-03-01
        If LOF(f) = 112 Then
            SCLFVer = 1
        Else
            SCLFVer = 2
        End If
        
        i = 1
        
        If LOF(f) > 1 Then
            With FormColors
                .ChannelLabelBack = DoGet(f, i)
                .ChannelLabelText = DoGet(f, i)
                .ChannelListBack = DoGet(f, i)
                .ChannelListText = DoGet(f, i)
                If SCLFVer = 2 Then
                    .ChannelListSelf = DoGet(f, i)
                    .ChannelListIdle = DoGet(f, i)
                    .ChannelListSquelched = DoGet(f, i)
                    .ChannelListOps = DoGet(f, i)
                End If
                .RTBBack = DoGet(f, i)
                .SendBoxesBack = DoGet(f, i)
                .SendBoxesText = DoGet(f, i)
            End With
            
            With RTBColors
                .TalkBotUsername = DoGet(f, i)
                .TalkUsernameNormal = DoGet(f, i)
                .TalkUsernameOp = DoGet(f, i)
                .TalkNormalText = DoGet(f, i)
                .Carats = DoGet(f, i)
                .EmoteText = DoGet(f, i)
                .EmoteUsernames = DoGet(f, i)
                .InformationText = DoGet(f, i)
                .SuccessText = DoGet(f, i)
                .ErrorMessageText = DoGet(f, i)
                .TimeStamps = DoGet(f, i)
                .ServerInfoText = DoGet(f, i)
                .ConsoleText = DoGet(f, i)
                .JoinText = DoGet(f, i)
                .JoinUsername = DoGet(f, i)
                .JoinedChannelName = DoGet(f, i)
                .JoinedChannelText = DoGet(f, i)
                .WhisperCarats = DoGet(f, i)
                .WhisperText = DoGet(f, i)
                .WhisperUsernames = DoGet(f, i)
            End With
        End If
        Close #f
    End If
    
    With FormColors
        If .ChannelLabelBack = -1 Then .ChannelLabelBack = &HCC3300
        If .ChannelLabelText = -1 Then .ChannelLabelText = vbWhite
        If .ChannelListBack = -1 Then .ChannelListBack = vbBlack
        If .ChannelListText = -1 Then .ChannelListText = COLOR_TEAL
        If .ChannelListSelf = -1 Then .ChannelListSelf = vbWhite
        If .ChannelListIdle = -1 Then .ChannelListIdle = &HBBBBBB
        If .ChannelListSquelched = -1 Then .ChannelListSquelched = &H99
        If .ChannelListOps = -1 Then .ChannelListOps = &HDDDDDD
        If .RTBBack = -1 Then .RTBBack = vbBlack
        If .SendBoxesBack = -1 Then .SendBoxesBack = vbBlack
        If .SendBoxesText = -1 Then .SendBoxesText = vbWhite
    End With
        
    With RTBColors
        If .Carats = -1 Then .Carats = COLOR_BLUE
        If .ConsoleText = -1 Then .ConsoleText = COLOR_TEAL
        If .EmoteText = -1 Then .EmoteText = vbYellow
        If .EmoteUsernames = -1 Then .EmoteUsernames = vbWhite
        If .InformationText = -1 Then .InformationText = vbYellow
        If .ErrorMessageText = -1 Then .ErrorMessageText = vbRed
        If .JoinedChannelName = -1 Then .JoinedChannelName = vbYellow
        If .JoinedChannelText = -1 Then .JoinedChannelText = COLOR_TEAL
        If .JoinText = -1 Then .JoinText = vbGreen
        If .JoinUsername = -1 Then .JoinUsername = vbYellow
        If .TalkUsernameNormal = -1 Then .TalkUsernameNormal = vbYellow
        If .TalkUsernameOp = -1 Then .TalkUsernameOp = vbWhite
        If .SuccessText = -1 Then .SuccessText = vbGreen
        If .TalkBotUsername = -1 Then .TalkBotUsername = vbCyan
        If .TalkNormalText = -1 Then .TalkNormalText = vbWhite
        If .ServerInfoText = -1 Then .ServerInfoText = COLOR_BLUE
        If .TimeStamps = -1 Then .TimeStamps = vbWhite
        If .WhisperCarats = -1 Then .WhisperCarats = &H80FF&
        If .WhisperText = -1 Then .WhisperText = &H999999
        If .WhisperUsernames = -1 Then .WhisperUsernames = vbYellow
    End With
    
    Call SetFormColors
End Sub

' calls Get from the file at the specified position, and increases the position
Private Function DoGet(ByVal f As Integer, ByRef i As Integer) As Long
    Dim l As Long
    
    Get #f, i, l
    
    DoGet = l
    
    i = i + 1
End Function
