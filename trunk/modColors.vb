Option Strict Off
Option Explicit On
Module modColors
	
	Public Structure udtColorList
		Dim ChannelListBack As Integer
		Dim ChannelListText As Integer
		Dim ChannelListSelf As Integer
		Dim ChannelListIdle As Integer
		Dim ChannelListSquelched As Integer
		Dim ChannelListOps As Integer
		Dim SendBoxesBack As Integer
		Dim SendBoxesText As Integer
		Dim ChannelLabelBack As Integer
		Dim ChannelLabelText As Integer
		Dim RTBBack As Integer
	End Structure
	
	Public Structure udtColorListRTB
		Dim TalkBotUsername As Integer
		Dim TalkUsernameNormal As Integer
		Dim TalkUsernameOp As Integer
		Dim TalkNormalText As Integer
		Dim Carats As Integer
		Dim InformationText As Integer
		Dim SuccessText As Integer
		Dim ErrorMessageText As Integer
		Dim TimeStamps As Integer
		Dim ServerInfoText As Integer
		Dim ConsoleText As Integer
		Dim WhisperCarats As Integer
		Dim WhisperUsernames As Integer
		Dim WhisperText As Integer
		Dim EmoteUsernames As Integer
		Dim EmoteText As Integer
		Dim JoinUsername As Integer
		Dim JoinText As Integer
		Dim JoinedChannelName As Integer
		Dim JoinedChannelText As Integer
	End Structure
	
	Public RTBColors As udtColorListRTB
	Public FormColors As udtColorList
	
	Public Sub SetFormColors()
		With FormColors
			frmChat.lvChannel.BackColor = System.Drawing.ColorTranslator.FromOle(.ChannelListBack)
			frmChat.lvChannel.ForeColor = System.Drawing.ColorTranslator.FromOle(.ChannelListText)
			frmChat.lvFriendList.BackColor = System.Drawing.ColorTranslator.FromOle(.ChannelListBack)
			frmChat.lvFriendList.ForeColor = System.Drawing.ColorTranslator.FromOle(.ChannelListText)
			frmChat.lvClanList.BackColor = System.Drawing.ColorTranslator.FromOle(.ChannelListBack)
			frmChat.lvClanList.ForeColor = System.Drawing.ColorTranslator.FromOle(.ChannelListText)
			frmChat.lblCurrentChannel.ForeColor = System.Drawing.ColorTranslator.FromOle(.ChannelLabelText)
			frmChat.lblCurrentChannel.BackColor = System.Drawing.ColorTranslator.FromOle(.ChannelLabelBack)
			frmChat.txtPost.ForeColor = System.Drawing.ColorTranslator.FromOle(.SendBoxesText)
			frmChat.txtPre.ForeColor = System.Drawing.ColorTranslator.FromOle(.SendBoxesText) '.SendBoxesText
			frmChat.cboSend.ForeColor = System.Drawing.ColorTranslator.FromOle(.SendBoxesText)
			frmChat.txtPost.BackColor = System.Drawing.ColorTranslator.FromOle(.SendBoxesBack)
			frmChat.txtPre.BackColor = System.Drawing.ColorTranslator.FromOle(.SendBoxesBack)
			frmChat.rtbChat.BackColor = System.Drawing.ColorTranslator.FromOle(.RTBBack)
			frmChat.cboSend.BackColor = System.Drawing.ColorTranslator.FromOle(.SendBoxesBack)
		End With
	End Sub
	
	Public Sub GetColorLists(Optional ByRef sPath As String = "")
		Dim f As Short
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
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        Dim SCLFVer As Short
		Dim i As Short
        If Len(Dir(sPath)) > 0 Then
            FileOpen(f, sPath, OpenMode.Random, , , 4)


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
            FileClose(f)
        End If
		
		With FormColors
			If .ChannelLabelBack = -1 Then .ChannelLabelBack = &HCC3300
			If .ChannelLabelText = -1 Then .ChannelLabelText = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
			If .ChannelListBack = -1 Then .ChannelListBack = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
			If .ChannelListText = -1 Then .ChannelListText = COLOR_TEAL
			If .ChannelListSelf = -1 Then .ChannelListSelf = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
			If .ChannelListIdle = -1 Then .ChannelListIdle = &HBBBBBB
			If .ChannelListSquelched = -1 Then .ChannelListSquelched = &H99
			If .ChannelListOps = -1 Then .ChannelListOps = &HDDDDDD
			If .RTBBack = -1 Then .RTBBack = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
			If .SendBoxesBack = -1 Then .SendBoxesBack = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
			If .SendBoxesText = -1 Then .SendBoxesText = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
		End With
		
		With RTBColors
			If .Carats = -1 Then .Carats = COLOR_BLUE
			If .ConsoleText = -1 Then .ConsoleText = COLOR_TEAL
			If .EmoteText = -1 Then .EmoteText = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
			If .EmoteUsernames = -1 Then .EmoteUsernames = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
			If .InformationText = -1 Then .InformationText = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
			If .ErrorMessageText = -1 Then .ErrorMessageText = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
			If .JoinedChannelName = -1 Then .JoinedChannelName = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
			If .JoinedChannelText = -1 Then .JoinedChannelText = COLOR_TEAL
			If .JoinText = -1 Then .JoinText = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Lime)
			If .JoinUsername = -1 Then .JoinUsername = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
			If .TalkUsernameNormal = -1 Then .TalkUsernameNormal = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
			If .TalkUsernameOp = -1 Then .TalkUsernameOp = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
			If .SuccessText = -1 Then .SuccessText = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Lime)
			If .TalkBotUsername = -1 Then .TalkBotUsername = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Cyan)
			If .TalkNormalText = -1 Then .TalkNormalText = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
			If .ServerInfoText = -1 Then .ServerInfoText = COLOR_BLUE
			If .TimeStamps = -1 Then .TimeStamps = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
			If .WhisperCarats = -1 Then .WhisperCarats = &H80FF
			If .WhisperText = -1 Then .WhisperText = &H999999
			If .WhisperUsernames = -1 Then .WhisperUsernames = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
		End With
		
		Call SetFormColors()
	End Sub
	
	' calls Get from the file at the specified position, and increases the position
	Private Function DoGet(ByVal f As Short, ByRef i As Short) As Integer
		Dim l As Integer
		
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(f, l, i)
		
		DoGet = l
		
		i = i + 1
	End Function
End Module