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

Private Const SID_FRIENDSLIST& = &H65
Private Const SID_FRIENDSUPDATE& = &H66
Private Const SID_FRIENDSADD& = &H67
Private Const SID_FRIENDSREMOVE& = &H68
Private Const SID_FRIENDSPOSITION& = &H69

Public Event FriendUpdate(ByVal Username As String, ByVal FLIndex As Byte)
Public Event FriendAdded(ByVal Username As String, ByVal Product As String, ByVal Location As Byte, _
                            ByVal Status As Byte, ByVal Channel As String)
Public Event FriendRemoved(ByVal Username As String)
Public Event FriendListReceived(ByVal FriendCount As Byte)
Public Event FriendListEntry(ByVal Username As String, ByVal Product As String, ByVal Channel As String, _
                             ByVal Status As Byte, ByVal Location As Byte)
Public Event FriendMoved()

Private Sub Class_Initialize()
    Set g_Friends = New Collection
End Sub

Public Sub RequestFriendsList(ByRef pBuff As clsDataBuffer)
    'ResetList
    pBuff.SendPacket SID_FRIENDSLIST
End Sub

Public Sub ParsePacket(ByVal PacketID As Long, ByRef Data As String)
    On Error GoTo ERROR_HANDLER
    
    Dim pBuff As New clsDataBuffer
    Dim flTemp As clsFriendObj
    Dim n As Integer
    
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

            Call ResetList
            
            n = pBuff.GetByte()     ' Number of entries
            RaiseEvent FriendListReceived(n)
            
            If (n > 0) Then
            
                'For each entry
                For n = 0 To n - 1
                    Set flTemp = New clsFriendObj

                    With flTemp
                        .Name = pBuff.GetString()       ' Account
                        .Status = pBuff.GetByte()       ' Status
                        .LocationID = pBuff.GetByte()   ' Location
                        
                        ' Product ID
                        .Game = StrReverse(pBuff.GetRaw(4))
                        If Conv(.Game) = 0 Then
                            .Game = "OFFL"
                        End If
                        
                        ' Location name
                        .Location = pBuff.GetString()
                    End With
                    
                    ' Add to the internal list
                    g_Friends.Add flTemp

                    RaiseEvent FriendListEntry(flTemp.DisplayName, flTemp.Game, flTemp.Location, flTemp.Status, _
                         flTemp.LocationID)
                         
                    Set flTemp = Nothing
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
            
            With g_Friends(n)
                .Status = pBuff.GetByte()       ' Status
                .LocationID = pBuff.GetByte()   ' Location
                
                ' NOTE: There is a server bug here where, when this packet is sent automaticlaly
                '   (not requested), the ProductID field contains your own product instead.
                '   Because of this, we ignore that field completely and wait for the periodic updates
                '   to update the value.
                '   (see: https://bnetdocs.org/packet/384/sid-friendsupdate)
                
                pBuff.GetDWORD
                ' Product ID
                '.Game = StrReverse(pBuff.GetRaw(4))
                'If Conv(.Game) = 0 Then
                '    .Game = "OFFL"
                'End If
                
                ' Location name
                .Location = pBuff.GetString()
                
                RaiseEvent FriendUpdate(.DisplayName, n)
            End With
        
        Case SID_FRIENDSADD
            '0x67 packet format
            '(STRING)       Account
            '(BYTE)         Status
            '(BYTE)         Location
            '(DWORD)        ProductID
            '(STRING)       Location name
            
            Set flTemp = New clsFriendObj
        
            With flTemp
                .Name = pBuff.GetString()       ' Account
                .Status = pBuff.GetByte()       ' Status
                .LocationID = pBuff.GetByte()   ' Location
                
                ' Product ID
                .Game = StrReverse(pBuff.GetRaw(4))
                If Conv(.Game) = 0 Then
                    .Game = "OFFL"
                End If
                
                ' Location name
                .Location = pBuff.GetString()

                RaiseEvent FriendAdded(.DisplayName, .Game, .LocationID, .Status, .Location)
            End With
            
            ' Add to the internal list
            g_Friends.Add flTemp
            
            Set flTemp = Nothing
        
        Case SID_FRIENDSREMOVE
            '0x68 packet format
            '(BYTE)       Entry Number
            
            n = pBuff.GetByte() + 1
            
            If n > 0 And n <= g_Friends.Count Then
                RaiseEvent FriendRemoved(g_Friends.Item(n).DisplayName)
                
                g_Friends.Remove n
            End If
        
        Case SID_FRIENDSPOSITION
            '0x69 packet format
            '(BYTE)     Old Position
            '(BYTE)     New Position
            
            Set flTemp = Nothing
            RaiseEvent FriendMoved
    
    End Select
    
    Set flTemp = Nothing

    Exit Sub

ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, "Error: " & Err.description & " in ParsePacket()."

    Exit Sub

    'debug.print "Error " & Err.Number & " (" & Err.Description & ") in procedure ParsePacket of Class Module clsFriendListHandler"
    
End Sub

Public Sub ResetList()
    'frmChat.lvFriendList.ListItems.Clear
    
    Set g_Friends = Nothing
    Set g_Friends = New Collection
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
