Option Strict Off
Option Explicit On
Friend Class clsFriendlistHandler
	
	Private Const SID_FRIENDSLIST As Integer = &H65
	Private Const SID_FRIENDSUPDATE As Integer = &H66
	Private Const SID_FRIENDSADD As Integer = &H67
	Private Const SID_FRIENDSREMOVE As Integer = &H68
	Private Const SID_FRIENDSPOSITION As Integer = &H69
	
	Public Event FriendUpdate(ByVal Username As String, ByVal FLIndex As Byte)
	Public Event FriendAdded(ByVal Username As String, ByVal Product As String, ByVal Location As Byte, ByVal Status As Byte, ByVal Channel As String)
	Public Event FriendRemoved(ByVal Username As String)
	Public Event FriendListReceived(ByVal FriendCount As Byte)
	Public Event FriendListEntry(ByVal Username As String, ByVal Product As String, ByVal Channel As String, ByVal Status As Byte, ByVal Location As Byte)
	Public Event FriendMoved()
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		g_Friends = New Collection
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Public Sub RequestFriendsList(ByRef pBuff As clsDataBuffer)
		'ResetList
		pBuff.SendPacket(SID_FRIENDSLIST)
	End Sub
	
    Public Sub ParsePacket(ByVal PacketID As Integer, ByRef Data() As Byte)
        On Error GoTo ERROR_HANDLER

        Dim pBuff As New clsDataBuffer
        Dim flTemp As clsFriendObj
        Dim n As Short

        pBuff.Data = Data

        Select Case PacketID
            Case SID_FRIENDSLIST
                '0x65 packet format
                '(BYTE)       Number of Entries
                'For each entry:
                '(STRING)     Account
                '(BYTE)       Status
                '(BYTE)       Location
                '(DWORD)      ProductID
                '(STRING)     Location name

                Call ResetList()

                n = pBuff.GetByte() ' Number of entries
                RaiseEvent FriendListReceived(n)

                If (n > 0) Then

                    'For each entry
                    For n = 0 To n - 1
                        flTemp = New clsFriendObj

                        With flTemp
                            .Name = pBuff.GetString() ' Account
                            .Status = pBuff.GetByte() ' Status
                            .LocationID = pBuff.GetByte() ' Location

                            ' Product ID
                            .game = StrReverse(pBuff.GetRaw(4))
                            If Conv(.game) = 0 Then
                                .game = "OFFL"
                            End If

                            ' Location name
                            .Location = pBuff.GetString()
                        End With

                        ' Add to the internal list
                        g_Friends.Add(flTemp)

                        RaiseEvent FriendListEntry(flTemp.DisplayName, flTemp.game, flTemp.Location, flTemp.Status, flTemp.LocationID)

                        'UPGRADE_NOTE: Object flTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                        flTemp = Nothing
                    Next n
                End If

            Case SID_FRIENDSUPDATE
                '0x66 packet format
                '(BYTE)       Entry number
                '(BYTE)       Status
                '(BYTE)       Location
                '(DWORD)      ProductID
                '(STRING)     Location name

                n = pBuff.GetByte() + 1

                With g_Friends.Item(n)
                    'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends(n).Status. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Status = pBuff.GetByte() ' Status
                    'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends(n).LocationID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .LocationID = pBuff.GetByte() ' Location

                    ' NOTE: There is a server bug here where, when this packet is sent automaticlaly
                    '   (not requested), the ProductID field contains your own product instead.
                    '   Because of this, we ignore that field completely and wait for the periodic updates
                    '   to update the value.
                    '   (see: https://bnetdocs.org/packet/384/sid-friendsupdate)

                    pBuff.GetDWORD()
                    ' Product ID
                    '.Game = StrReverse(pBuff.GetRaw(4))
                    'If Conv(.Game) = 0 Then
                    '    .Game = "OFFL"
                    'End If

                    ' Location name
                    'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends(n).Location. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Location = pBuff.GetString()

                    'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends(n).DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    RaiseEvent FriendUpdate(.DisplayName, n)
                End With

            Case SID_FRIENDSADD
                '0x67 packet format
                '(STRING)       Account
                '(BYTE)         Status
                '(BYTE)         Location
                '(DWORD)        ProductID
                '(STRING)       Location name

                flTemp = New clsFriendObj

                With flTemp
                    .Name = pBuff.GetString() ' Account
                    .Status = pBuff.GetByte() ' Status
                    .LocationID = pBuff.GetByte() ' Location

                    ' Product ID
                    .game = StrReverse(pBuff.GetRaw(4))
                    If Conv(.game) = 0 Then
                        .game = "OFFL"
                    End If

                    ' Location name
                    .Location = pBuff.GetString()

                    RaiseEvent FriendAdded(.DisplayName, .game, .LocationID, .Status, .Location)
                End With

                ' Add to the internal list
                g_Friends.Add(flTemp)

                'UPGRADE_NOTE: Object flTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                flTemp = Nothing

            Case SID_FRIENDSREMOVE
                '0x68 packet format
                '(BYTE)       Entry Number

                n = pBuff.GetByte() + 1

                If n > 0 And n <= g_Friends.Count() Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends.Item(n).DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    RaiseEvent FriendRemoved(g_Friends.Item(n).DisplayName)

                    g_Friends.Remove(n)
                End If

            Case SID_FRIENDSPOSITION
                '0x69 packet format
                '(BYTE)     Old Position
                '(BYTE)     New Position

                'UPGRADE_NOTE: Object flTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                flTemp = Nothing
                RaiseEvent FriendMoved()

        End Select

        'UPGRADE_NOTE: Object flTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        flTemp = Nothing

        Exit Sub

ERROR_HANDLER:
        frmChat.AddChat(RTBColors.ErrorMessageText, "Error: " & Err.Description & " in ParsePacket().")

        Exit Sub

        'debug.print "Error " & Err.Number & " (" & Err.Description & ") in procedure ParsePacket of Class Module clsFriendListHandler"

    End Sub
	
	Public Sub ResetList()
		'frmChat.lvFriendList.ListItems.Clear
		
		'UPGRADE_NOTE: Object g_Friends may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		g_Friends = Nothing
		g_Friends = New Collection
	End Sub
	
	Public Function UsernameToFLIndex(ByVal sUsername As String) As Short
		Dim i As Short
		
		If g_Friends.Count() > 0 Then
			For i = 1 To g_Friends.Count()
				'UPGRADE_WARNING: Couldn't resolve default property of object g_Friends.Item().DisplayName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If StrComp(sUsername, g_Friends.Item(i).DisplayName, CompareMethod.Text) = 0 Then
					UsernameToFLIndex = i
					Exit Function
				End If
			Next i
		End If
	End Function
	
	' Returns true if the specified product automatically receives friend update packets.
	'   (SID_FRIENDSUPDATE, SID_FRIENDSADD, SID_FRIENDSREMOVE, SID_FRIENDSPOSITION)
	Public Function SupportsFriendPackets(ByVal sProduct As String) As Boolean
		Select Case GetProductInfo(sProduct).Code
			Case PRODUCT_STAR, PRODUCT_SEXP, PRODUCT_WAR3, PRODUCT_W3XP
				SupportsFriendPackets = True
			Case Else
				SupportsFriendPackets = False
		End Select
	End Function
	
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object g_Friends may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		g_Friends = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	'Public Sub WriteLog(ByVal s As String, Optional ByVal NoDebug As Boolean = False)
	'    If InStr(1, Command(), "-logFriends") Then
	'
	'        If Dir$(App.Path & "\friendlog.txt") = "" Then
	'            Open App.Path & "\friendlog.txt" For Output As #1
	'            Close #1
	'        End If
	'
	'        Open App.Path & "\friendlog.txt" For Append As #1
	'            If NoDebug Then
	'                Print #1, s
	'            Else
	'                Print #1, DebugOutput(s) & vbCrLf
	'            End If
	'        Close #1
	'
	'    End If
	'End Sub
End Class