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

'clsClanPacketHandler - project StealthBot - authored by Stealth (stealth@stealthbot.net)

'Special thanks:
'-  Ethereal packetlogger was used in my own research
'-  thanks to Arta[vL] and BNetDocs (http://bnetdocs.valhallalegends.com) for additional assistance

Private Const SID_FRIENDLIST& = &H65
Private Const SID_FRIENDUPDATE& = &H66
Private Const SID_FRIENDADDED& = &H67
Private Const SID_FRIENDREMOVED& = &H68
Private Const SID_FRIENDMOVED& = &H69
    
Public colFriends As Collection

Public Event FriendUpdate(ByVal Username As String, ByVal FLIndex As Byte)
Public Event FriendAdded(ByVal Username As String, ByVal Product As String, ByVal Location As Byte, _
                            ByVal Status As Byte, ByVal Channel As String)
Public Event FriendRemoved(ByVal Username As String)
Public Event FriendListEntry(ByVal Username As String, ByVal Product As String, ByVal Channel As String, _
                             ByVal Status As Byte, ByVal Location As Byte)
Public Event FriendMoved()


Private Sub Class_Initialize()
    Set colFriends = New Collection
End Sub

Public Sub RequestFriendsList(ByRef pBuff As PacketBuffer)
    ResetList
    pBuff.SendPacket &H65
End Sub

Public Sub ParsePacket(ByVal PacketID As Long, ByRef Data As String)
    
    Dim n As Long, bytTemp As Byte
    Dim s As String
    Dim Position As Long
    Dim flTemp As clsFriend
    
    On Error GoTo ParsePacket_Error

    Position = 1
    
    Select Case PacketID
        Case SID_FRIENDLIST
            '0x65 packet format
            '(BYTE)       Number of Entries
            n = Asc(Mid$(Data, 1, 1))
            
            Call ResetList
            
            Position = 2
            Set flTemp = Nothing
            
            For n = 0 To n - 1
                Set flTemp = New clsFriend
                
                ''debug.print DebugOutput(Mid$(Data, Position)) & vbCrLf
                s = ""
                
                With flTemp
                    'For each entry:
                    '(STRING)         Account
                    s = KillNull(Mid$(Data, Position))
                    .Username = s
                    'debug.print "Username: " & s
                    Position = Position + Len(.Username) + 1 'account for chr(0) at end
                    'debug.print "Position incremented " & Len(.Username) + 1 & " to " & Position
                    
                    '(BYTE)       Status
                    bytTemp = Asc(Mid$(Data, Position, 1))
                    .Status = bytTemp
                    Position = Position + 1
                    
                    'debug.print "Status: " & .Status
                    'debug.print "Position incremented 1 to " & Position
                    
                    '(BYTE)       Location
                    bytTemp = Asc(Mid$(Data, Position, 1))
                    .Location = bytTemp
                    Position = Position + 1
                    
                    'debug.print "Location: " & .Location
                    'debug.print "Position incremented 1 to " & Position
                    
                    '(DWORD)      ProductID
                    s = Mid$(Data, Position, 4)
                    If Conv(s) = 0 Then
                        .Product = "OFFL"
                    Else
                        .Product = StrReverse(s)
                    End If
                    Position = Position + 4
                    
                    'debug.print "ProductID: " & .Product
                    'debug.print "Position incremented 4 to " & Position
                    
                    '(STRING)         Channel
                    s = KillNull(Mid$(Data, Position))
                    Position = Position + Len(s) + 1
                    .Channel = s
                    
                End With
                
                ''debug.print "Successfully added: " & flTemp.Username
                
                colFriends.Add flTemp, flTemp.Username
                
                RaiseEvent FriendListEntry(flTemp.Username, flTemp.Product, flTemp.Channel, flTemp.Status, _
                    flTemp.Location)
                    
                Set flTemp = Nothing
                
            Next n
            
        Case SID_FRIENDUPDATE
            '(BYTE)       Entry number
            n = Asc(Mid$(Data, Position, 1)) + 1
            
            Position = Position + 1
            
            With colFriends.Item(n)
                '(BYTE) Flags
                bytTemp = Asc(Mid$(Data, Position, 1))
                Position = Position + 1
                .Status = bytTemp
                
                '(BYTE) Location
                bytTemp = Asc(Mid$(Data, Position, 1))
                Position = Position + 1
                .Location = bytTemp
                
                '(DWORD) Product ID
                s = Mid$(Data, Position, 4)
                Position = Position + 4
                .Product = StrReverse(s)
            
                '(STRING) Location str (blank = private)
                s = Mid$(Data, Position)
                s = Left$(s, Len(s) - 1)
                .Channel = IIf(LenB(s) > 0, s, "(private)")
                
                RaiseEvent FriendUpdate(.Username, n)
            End With
        
        Case SID_FRIENDADDED
            Set flTemp = New clsFriend
        
            With flTemp
                '(STRING)         Account
                s = KillNull(Mid$(Data, Position))
                Position = Position + Len(s) + 1
                .Username = s
                            
                '(BYTE)       Friend Status
                bytTemp = Asc(Mid$(Data, Position, 1))
                Position = Position + 1
                .Status = bytTemp
                
                '(BYTE)       Friend Location
                bytTemp = Asc(Mid$(Data, Position, 1))
                Position = Position + 1
                .Location = bytTemp
                
                '(DWORD)      ProductID
                s = Mid$(Data, Position, 4)
                Position = Position + 4
                .Product = IIf(Conv(s) > 0, StrReverse(s), "OFFL")
                
                '(STRING)         Channel
                s = KillNull(Mid$(Data, Position))
                .Channel = s
                
                RaiseEvent FriendAdded(.Username, .Product, .Location, .Status, .Channel)
            End With
            
            colFriends.Add flTemp, flTemp.Username
            
            Set flTemp = Nothing
        
        Case SID_FRIENDREMOVED
            '(BYTE)       Entry Number
            bytTemp = Asc(Mid$(Data, 1, 1)) + 1
            
            If bytTemp > 0 And bytTemp <= colFriends.Count Then
                RaiseEvent FriendRemoved(colFriends.Item(bytTemp).Username)
                colFriends.Remove bytTemp
            End If
            
        
        Case SID_FRIENDMOVED
            '(BYTE) Old Position
            '(BYTE) New Position
            Set flTemp = Nothing
            RaiseEvent FriendMoved
            
    
    End Select
    
    Set flTemp = Nothing

    On Error GoTo 0
    Exit Sub

ParsePacket_Error:

    'debug.print "Error " & Err.Number & " (" & Err.Description & ") in procedure ParsePacket of Class Module clsFriendlistHandler"
    
End Sub

Public Sub ResetList()
    Set colFriends = Nothing
    Set colFriends = New Collection
End Sub

Public Function UsernameToFLIndex(ByVal sUsername As String) As Integer
    Dim i As Integer
    
    If colFriends.Count > 0 Then
        For i = 1 To colFriends.Count
            If StrComp(sUsername, colFriends.Item(i).Username, vbTextCompare) = 0 Then
                UsernameToFLIndex = i
                Exit Function
            End If
        Next i
    End If
End Function


Private Sub Class_Terminate()
    Set colFriends = Nothing
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
