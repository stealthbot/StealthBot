Option Strict Off
Option Explicit On
Module modQuickChannels
	'modQuickChannels - Project StealthBot
	'   by Andy
	'   June 2006
	
	Public Sub LoadQuickChannels()
		Dim UseDefaultQC As Boolean
		Dim i As Short
		Dim QCCollection As Collection
		Dim LineText As String
		
		UseDefaultQC = True
		
		' "upgrade" a QC.ini to a QC.txt, deletes QC.ini
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If (Len(Dir(GetFilePath("QuickChannels.ini"))) > 0) Then

            UseDefaultQC = False

            Call UpgradeQuickChannelsToList()

        End If
		
		' read QC.txt
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Len(Dir(GetFilePath(FILE_QUICK_CHANNELS))) > 0 Then

            UseDefaultQC = False

            QCCollection = ListFileLoad(GetFilePath(FILE_QUICK_CHANNELS), 9)

            For i = 1 To QCCollection.Count()

                'UPGRADE_WARNING: Couldn't resolve default property of object QCCollection.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                QC(i) = Trim(QCCollection.Item(i))

            Next i

            'UPGRADE_NOTE: Object QCCollection may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            QCCollection = Nothing

        End If
		
		' no QC files, use defaults and save it
		If UseDefaultQC Then
			
			i = 0
			Do 
				
				LineText = GetDefaultQC(i)
				
                If Len(LineText) > 0 Then

                    QC(i + 1) = LineText

                    i = i + 1

                End If
				
            Loop Until Len(LineText) = 0
			
			'Call SaveQuickChannels
			
		End If
		
		PrepareQuickChannelMenu()
	End Sub
	
	Public Function GetDefaultQC(ByVal Index As Short) As String
		
		' no default QC...
		GetDefaultQC = vbNullString
		'Select Case Index
		'Case 0: GetDefaultQC = "Clan BoT"
		'Case 1: GetDefaultQC = "Public Chat 1"
		'Case 2: GetDefaultQC = "Clan TDA"
		'Case 3: GetDefaultQC = "Clan BNU"
		'Case 4: GetDefaultQC = "Op W@R"
		'End Select
		
	End Function
	
	Public Sub UpgradeQuickChannelsToList()
		Dim sFileINI As String
		Dim i As Short
		
		sFileINI = GetFilePath("QuickChannels.ini")
		
		' old QC.ini bounds
		For i = 0 To 8
			QC(i + 1) = Trim(ReadINI("QuickChannels", CStr(i), sFileINI))
		Next i
		
		Kill(sFileINI)
		
		Call SaveQuickChannels()
	End Sub
	
	Public Sub SaveQuickChannels()
		Dim i As Short
		Dim QCCollection As New Collection
		
		For i = LBound(QC) To UBound(QC)
            If Len(QC(i)) = 0 Then
                QC(i) = " "
            End If
			QCCollection.Add(QC(i))
		Next i
		
		Call ListFileSave(GetFilePath(FILE_QUICK_CHANNELS), QCCollection)
		
		'UPGRADE_NOTE: Object QCCollection may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		QCCollection = Nothing
	End Sub
	
	Public Sub PrepareQuickChannelMenu()
		Dim i As Short
		Dim Caption As String
		Dim ShownAddQC As Boolean
		Dim FoundThisChannel As Boolean
		
		ShownAddQC = False
		FoundThisChannel = False
		
		' bounds of mnuCustomChannels
		For i = 0 To 8
			Caption = Trim(QC(i + 1))
			
			With frmChat.mnuCustomChannels(i)
                If Len(Caption) > 0 Then
                    If StrComp(Caption, g_Channel.Name, CompareMethod.Text) = 0 Then
                        FoundThisChannel = True
                    End If

                    .Visible = True
                    .Text = MakeChannelMenuItemSafe(Caption)
                Else
                    If Not ShownAddQC And Len(g_Channel.Name) > 0 Then
                        frmChat.mnuCustomChannelAdd.Visible = True
                        'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        frmChat.mnuCustomChannelAdd.Text = StringFormat("&Add {0}{1}{0} as F{2}", Chr(34), MakeChannelMenuItemSafe(g_Channel.Name, True), CStr(i + 1))

                        ShownAddQC = True
                    End If

                    .Visible = False
                    .Text = vbNullString
                End If
			End With
		Next i
		
		If FoundThisChannel Or Not ShownAddQC Then
			frmChat.mnuCustomChannelAdd.Visible = False
		End If
	End Sub
	
	Public Sub PrepareHomeChannelMenu()
		Dim ShowHome As Boolean
		Dim ShowLast As Boolean
		
        ShowHome = (Len(Config.HomeChannel) > 0 And StrComp(Config.HomeChannel, g_Channel.Name, CompareMethod.Text) <> 0)
		With frmChat.mnuHomeChannel
			.Text = MakeChannelMenuItemSafe(Config.HomeChannel, True) & " (&Home Channel)"
			.Visible = ShowHome
		End With
		
        ShowLast = (Len(BotVars.LastChannel) > 0 And StrComp(BotVars.LastChannel, g_Channel.Name, CompareMethod.Text) <> 0)
		With frmChat.mnuLastChannel
			.Text = MakeChannelMenuItemSafe(BotVars.LastChannel, True) & " (&Previous Channel)"
			.Visible = ShowLast
		End With
		
		frmChat.mnuQCDash.Visible = (ShowHome Or ShowLast)
	End Sub
	
	Public Sub PreparePublicChannelMenu()
		Dim i As Short
		Dim AnyVisible As Boolean
		
		' unload all menu items
		For i = 1 To frmChat.mnuPublicChannels.Count - 1
			Call frmChat.mnuPublicChannels.Unload(i)
		Next i
		
		' hide 0th menu item
		With frmChat.mnuPublicChannels(0)
			.Text = vbNullString
			.Visible = False
		End With
		
		AnyVisible = False
		
		If Not BotVars.PublicChannels Is Nothing Then
			If BotVars.PublicChannels.Count() > 0 Then
				For i = 1 To BotVars.PublicChannels.Count()
					AnyVisible = True
					
					If (i > 1) Then
						' load a new one
						Call frmChat.mnuPublicChannels.Load(i - 1)
					End If
					
					With frmChat.mnuPublicChannels(i - 1)
						'UPGRADE_WARNING: Couldn't resolve default property of object BotVars.PublicChannels.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.Text = MakeChannelMenuItemSafe(BotVars.PublicChannels.Item(i))
						.Visible = True
					End With
				Next i
			End If
		End If
		
		' dash and header visibility
		frmChat.mnuPCDash.Visible = AnyVisible
		frmChat.mnuPCHeader.Visible = AnyVisible
	End Sub
	
	Private Function MakeChannelMenuItemSafe(ByVal sChannel As String, Optional ByVal CanBeDash As Boolean = False) As String
		sChannel = Replace(sChannel, "&", "&&",  ,  , CompareMethod.Binary)
		
		If Not CanBeDash And StrComp(sChannel, "-", CompareMethod.Binary) = 0 Then
			sChannel = "&-"
		End If
		
		MakeChannelMenuItemSafe = sChannel
	End Function
End Module