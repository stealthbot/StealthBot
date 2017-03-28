VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFriendlistHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event FriendsListReply(ByVal Friends As Collection)
Public Event FriendsUpdate(ByVal EntryNumber As Byte, ByVal FriendObj As clsFriendObj)
Public Event FriendsAdd(ByVal FriendObj As clsFriendObj)
Public Event FriendsRemove(ByVal EntryNumber As Byte)
Public Event FriendsPosition(ByVal EntryNumber As Byte, ByVal NewPosition As Byte)

Public Sub ParsePacket(ByVal PacketID As Long, ByVal inBuf As clsDataBuffer)
    On Error GoTo ERROR_HANDLER

    Dim FriendObj As clsFriendObj
    Dim Friends   As Collection
    Dim i         As Integer
    Dim Count     As Byte
    Dim EntryNum  As Byte
    Dim NewPos    As Byte

    Select Case PacketID
        Case SID_FRIENDSLIST                        ' no-cookie server request
            Count = inBuf.GetByte                   ' (BYTE) Friend count
            Set Friends = New Collection            ' (MEMBER[]) Members
            For i = 1 To Count
                Set FriendObj = New clsFriendObj
                With FriendObj
                    .Name = inBuf.GetString         ' (STRING) Username
                    .Status = inBuf.GetByte         ' (BYTE) Status
                    .LocationID = inBuf.GetByte     ' (BYTE) Location
                    .Game = DWordToString(inBuf.GetDWord) ' (DWORD) Product
                    .Location = inBuf.GetString     ' (STRING) Location name
                End With
                Friends.Add FriendObj
                Set FriendObj = Nothing
            Next i

            RaiseEvent FriendsListReply(Friends)

        Case SID_FRIENDSUPDATE                      ' information event
            EntryNum = inBuf.GetByte                ' (BYTE) Entry number
            Set FriendObj = New clsFriendObj
            With FriendObj
                .Status = inBuf.GetByte             ' (BYTE) Status
                .LocationID = inBuf.GetByte         ' (BYTE) Location
                .Game = DWordToString(inBuf.GetDWord) ' (DWORD) Product
                .Location = inBuf.GetString         ' (STRING) Location name
            End With

            RaiseEvent FriendsUpdate(EntryNum, FriendObj)

        Case SID_FRIENDSADD                         ' information event
            Set FriendObj = New clsFriendObj
            With FriendObj
                .Name = inBuf.GetString             ' (STRING) Username
                .Status = inBuf.GetByte             ' (BYTE) Status
                .LocationID = inBuf.GetByte         ' (BYTE) Location
                .Game = DWordToString(inBuf.GetDWord) ' (DWORD) Product
                .Location = inBuf.GetString         ' (STRING) Location name
            End With

            RaiseEvent FriendsAdd(FriendObj)
        
        Case SID_FRIENDSREMOVE
            EntryNum = inBuf.GetByte                ' (BYTE) Entry number

            RaiseEvent FriendsRemove(EntryNum)

        Case SID_FRIENDSPOSITION
            EntryNum = inBuf.GetByte                ' (BYTE) Entry number
            NewPos = inBuf.GetByte             ' (BYTE) New position

            RaiseEvent FriendsPosition(EntryNum, NewPos)

    End Select

    Exit Sub

ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, "Error: " & Err.Description & " in clsFriendlistHandler.ParsePacket()."

    Exit Sub

End Sub

Public Sub RequestFriendsList()
    Dim pBuf As clsDataBuffer
    Set pBuf = New clsDataBuffer
    pBuf.SendPacket SID_FRIENDSLIST
    Set pBuf = Nothing
End Sub

Public Function UsernameToFLIndex(ByVal sUsername As String) As Integer
    Dim i As Integer
    
    If g_Friends.Count > 0 Then
        For i = 1 To g_Friends.Count
            If StrComp(sUsername, g_Friends.Item(i).DisplayName, vbTextCompare) = 0 Then
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


Private Sub Class_Terminate()
    Set g_Friends = Nothing
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
