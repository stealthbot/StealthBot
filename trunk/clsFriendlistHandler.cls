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

Public Event FriendUpdate(ByVal Username As String, ByVal FLIndex As Byte)
Public Event FriendAdded(ByVal Username As String, ByVal Product As String, ByVal Location As Byte, _
                            ByVal Status As Byte, ByVal Channel As String)
Public Event FriendRemoved(ByVal Username As String)
Public Event FriendListEntry(ByVal Username As String, ByVal Product As String, ByVal Channel As String, _
                             ByVal Status As Byte, ByVal Location As Byte)
Public Event FriendMoved()


Private Sub Class_Initialize()
    Set g_Friends = New Collection
End Sub

Public Sub RequestFriendsList(ByRef pBuff As clsDataBuffer)
    'ResetList
    pBuff.SendPacket &H65
End Sub

Public Sub ParsePacket(ByVal PacketID As Long, ByRef Data As String)
    On Error GoTo ERROR_HANDLER
    
    Dim n As Integer, bytTemp As Byte
    Dim s As String
    Dim Position As Long
    Dim flTemp As clsFriendObj

    Position = 1
    
    Select Case PacketID
        Case SID_FRIENDLIST
            Set g_Friends = New Collection
        
            '0x65 packet format
            '(BYTE)       Number of Entries
            n = Asc(Mid$(Data, 1, 1))
            
            Call ResetList
            
            Position = 2
            'Set flTemp = Nothing
            
            If (n > 0) Then
                For n = 0 To n - 1
                    'If (g_Friends.Count > n) Then
                    '    Set flTemp = g_Friends(n + 1)
                    'Else
                        Set flTemp = New clsFriendObj
                    'End If
                    ''debug.print DebugOutput(Mid$(Data, Position)) & vbCrLf
                    s = ""
                    
                    With flTemp
                        'For each entry:
                        '(STRING)         Account
                        s = KillNull(Mid$(Data, Position))
                        .Name = s
                        'debug.print "Username: " & s
                        Position = Position + Len(.Name) + 1 'account for chr(0) at end
                        'debug.print "Position incremented " & Len(.Name) + 1 & " to " & Position
                        
                        '(BYTE)       Status
                        bytTemp = Asc(Mid$(Data, Position, 1))
                        .Status = bytTemp
                        Position = Position + 1
                        
                        'debug.print "Status: " & .Status
                        'debug.print "Position incremented 1 to " & Position
                        
                        '(BYTE)       Location
                        bytTemp = Asc(Mid$(Data, Position, 1))
                        .LocationID = bytTemp
                        Position = Position + 1
                        
                        'debug.print "Location: " & .Location
                        'debug.print "Position incremented 1 to " & Position
                        
                        '(DWORD)      ProductID
                        s = Mid$(Data, Position, 4)
                        If Conv(s) = 0 Then
                            .game = "OFFL"
                        Else
                            .game = StrReverse(s)
                        End If
                        Position = Position + 4
                        
                        'debug.print "ProductID: " & .Game
                        'debug.print "Position incremented 4 to " & Position
                        
                        '(STRING)         Channel
                        s = KillNull(Mid$(Data, Position))
                        Position = Position + Len(s) + 1
                        .Location = s
                    End With
                    
                    ''debug.print "Successfully added: " & flTemp.Name
                    
                    'If (g_Friends.Count > n) Then
                    '    RaiseEvent FriendUpdate(flTemp.DisplayName, n + 1)
                    'Else
                        g_Friends.Add flTemp
                        
                        'frmChat.AddChat vbRed, flTemp.Name & ":" & flTemp.DisplayName
                        
                        RaiseEvent FriendListEntry(flTemp.DisplayName, flTemp.game, flTemp.Location, flTemp.Status, _
                             flTemp.LocationID)
                    'End If
                        
                    Set flTemp = Nothing
                    
                Next n
            End If
            
        Case SID_FRIENDUPDATE
            Exit Sub
        
            '(BYTE)       Entry number
            n = Asc(Mid$(Data, Position, 1)) + 1
            
            Position = 2 ' Position + 1
            
            With g_Friends(n)
                '(BYTE) Flags
                bytTemp = Asc(Mid$(Data, Position, 1))
                Position = Position + 1
                .Status = bytTemp
                
                '(BYTE) Location
                bytTemp = Asc(Mid$(Data, Position, 1))
                Position = Position + 1
                .LocationID = bytTemp
                
                '(DWORD) Product ID
                s = Mid$(Data, Position, 4)
                Position = Position + 4
                .game = StrReverse(s)

                '(STRING) Location str (blank = private)
                s = Mid$(Data, Position)
                
                ' ...
                If (Len(s) > 1) Then
                    s = Left$(s, Len(s) - 1)
                End If
                
                .Location = IIf(LenB(s) > 0, s, "(private)")
                
                RaiseEvent FriendUpdate(.DisplayName, n)
            End With
        
        Case SID_FRIENDADDED
            Exit Sub
        
            Set flTemp = New clsFriendObj
        
            With flTemp
                '(STRING)         Account
                s = KillNull(Mid$(Data, Position))
                Position = Position + Len(s) + 1
                .Name = s
                            
                '(BYTE)       Friend Status
                bytTemp = Asc(Mid$(Data, Position, 1))
                Position = Position + 1
                .Status = bytTemp
                
                '(BYTE)       Friend Location
                bytTemp = Asc(Mid$(Data, Position, 1))
                Position = Position + 1
                .LocationID = bytTemp
                
                '(DWORD)      ProductID
                s = Mid$(Data, Position, 4)
                Position = Position + 4
                .game = IIf(Conv(s) > 0, StrReverse(s), "OFFL")
                
                '(STRING)         Channel
                s = KillNull(Mid$(Data, Position))
                .Location = s

                RaiseEvent FriendAdded(.DisplayName, .game, .LocationID, .Status, .Location)
            End With
            
            g_Friends.Add flTemp, flTemp.Name
            
            Set flTemp = Nothing
        
        Case SID_FRIENDREMOVED
            Exit Sub
        
            '(BYTE)       Entry Number
            bytTemp = Asc(Mid$(Data, 1, 1)) + 1
            
            If bytTemp > 0 And bytTemp <= g_Friends.Count Then
                RaiseEvent FriendRemoved(g_Friends.Item(bytTemp).DisplayName)
                
                g_Friends.Remove bytTemp
            End If
            
        
        Case SID_FRIENDMOVED
            Exit Sub
        
            '(BYTE) Old Position
            '(BYTE) New Position
            Set flTemp = Nothing
            RaiseEvent FriendMoved
            
    
    End Select
    
    Set flTemp = Nothing

    Exit Sub

ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error: " & Err.description & " in ParsePacket()."

    Exit Sub

    'debug.print "Error " & Err.Number & " (" & Err.Description & ") in procedure ParsePacket of Class Module clsFriendListHandler"
    
End Sub

Public Sub ResetList()
    ' ...
    'frmChat.lvFriendList.ListItems.Clear
    
    ' ...
    Set g_Friends = Nothing
    Set g_Friends = New Collection
End Sub

Public Function UsernameToFLIndex(ByVal sUsername As String) As Integer
    Dim I As Integer
    
    If g_Friends.Count > 0 Then
        For I = 1 To g_Friends.Count
            If StrComp(sUsername, g_Friends.Item(I).DisplayName, vbTextCompare) = 0 Then
                UsernameToFLIndex = I
                Exit Function
            End If
        Next I
    End If
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
