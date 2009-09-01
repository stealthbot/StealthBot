Attribute VB_Name = "modColors"
Option Explicit

Public Type udtColorList
    ChannelListBack As Long
    ChannelListText As Long
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
    
    If sPath = vbNullString Then sPath = GetFilePath("Colors.sclf")
    
    ' Initialize
    With FormColors
        .ChannelLabelBack = -1
        .ChannelLabelText = -1
        .ChannelListBack = -1
        .ChannelListText = -1
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
    If LenB(Dir(sPath)) > 0 Then
        Open sPath For Random As #f Len = 4
        
        If LOF(f) > 1 Then
            With FormColors
                Get #f, 1, .ChannelLabelBack
                Get #f, 2, .ChannelLabelText
                Get #f, 3, .ChannelListBack
                Get #f, 4, .ChannelListText
                Get #f, 5, .RTBBack
                Get #f, 6, .SendBoxesBack
                Get #f, 7, .SendBoxesText
            End With
            
            With RTBColors
                Get #f, 8, .TalkBotUsername
                Get #f, 9, .TalkUsernameNormal
                Get #f, 10, .TalkUsernameOp
                Get #f, 11, .TalkNormalText
                Get #f, 12, .Carats
                Get #f, 13, .EmoteText
                Get #f, 14, .EmoteUsernames
                Get #f, 15, .InformationText
                Get #f, 16, .SuccessText
                Get #f, 17, .ErrorMessageText
                Get #f, 18, .TimeStamps
                Get #f, 19, .ServerInfoText
                Get #f, 20, .ConsoleText
                Get #f, 21, .JoinText
                Get #f, 22, .JoinUsername
                Get #f, 23, .JoinedChannelName
                Get #f, 24, .JoinedChannelText
                Get #f, 25, .WhisperCarats
                Get #f, 26, .WhisperText
                Get #f, 27, .WhisperUsernames
            End With
        End If
        Close #f
    End If
    
    With FormColors
        If .ChannelLabelBack = -1 Then .ChannelLabelBack = &HCC3300
        If .ChannelLabelText = -1 Then .ChannelLabelText = vbWhite
        If .ChannelListBack = -1 Then .ChannelListBack = vbBlack
        If .ChannelListText = -1 Then .ChannelListText = COLOR_TEAL
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
