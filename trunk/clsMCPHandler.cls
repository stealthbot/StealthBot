VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMCPHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsMCPHandler
' DateTime  : 3/16/2004 23:00
' Author    : Stealth
' Purpose   : Handle and dispatch MCP packets
' Thanks    : Arta[vL]'s BNetDocs, DarkRaven
' Notes     : Goes with Stealth's modified DM DataBuffer class.
' Update    : 4/11/06 to fix realm connection issues; added CharListResponse event
'---------------------------------------------------------------------------------------

' Requires this function placed in a module:
'Public Function A2Hash(ByVal Password As String, ByVal ServerToken As Long) As String
'    Dim Hash As String
'
'    Hash = String(5 * 4, vbNullChar)
'    Call A2(Hash, ServerToken)
'
'    A2Hash = Hash
'End Function

Option Explicit

Private Const MCP_STARTUP& = &H1
Private Const MCP_CHARCREATE& = &H2
Private Const MCP_CHARLOGON& = &H7
Private Const MCP_CHARDELETE& = &HA
Private Const MCP_MOTD& = &H12
Private Const MCP_CHARLIST& = &H19
Private Const MCP_CHARUPGRADE& = &H18

Public Enum eCharacterFlags
    Expansion& = &H20
    ExpansionLadder& = &H60
    ExpansionHardcore& = &H24
    ExpansionHardcoreLadder& = &H64
    Classic& = &H0
    ClassicLadder& = &H40
    ClassicHardcore& = &H4
    ClassicHardcoreLadder& = &H48
End Enum

Private msCurrentChunk As String
Private msBNetUniqueUsername As String

Public Event CharListEntry(ByVal CharName As String, ByVal Statstring As String, ByVal ExpirationDate As Date)
Public Event CharUpgradeResponse(ByVal Success As Boolean)
Public Event CharLogonResponse(ByVal Status As Byte, ByVal Message As String)
Public Event CharDeleteResponse(ByVal Success As Boolean)
Public Event CharCreateResponse(ByVal Status As Byte, ByVal Message As String)
Public Event RealmMOTD(ByVal Message As String)
Public Event RealmStartup(ByVal Status As Byte, ByVal Message As String)
Public Event CharListResponse(ByVal NumCharacters As Integer)

Public Sub SendStartup()
    With PBuffer
        .InsertNonNTString msCurrentChunk
        .InsertNTString msBNetUniqueUsername
        .SendPacketMCP MCP_STARTUP
    End With
End Sub

Public Sub CreateMCPCharacter(ByVal CharType As Long, ByVal CharFlags As Long, ByVal CharName As String)
    With PBuffer
        .InsertDWord CharType
        .InsertWord CharFlags
        .InsertNTString CharName
        .SendPacketMCP MCP_CHARCREATE
    End With
End Sub

Public Sub LogonToCharacter(ByVal CharName As String)
    With PBuffer
        .InsertNTString CharName
        .SendPacketMCP MCP_CHARLOGON
    End With
End Sub

Public Sub DeleteCharacter(ByVal CharName As String, Optional ByVal Flags As Integer = &H1)
    With PBuffer
        .InsertWord Flags
        .InsertNTString CharName
        
        'Debug.Print "Sending: "
        'Debug.Print DebugOutput(PBuffer.GetPacket(MCP_CHARDELETE))
        
        .SendPacketMCP MCP_CHARDELETE
    End With
End Sub

Public Sub RequestRealmMOTD()
    PBuffer.SendPacketMCP MCP_MOTD
End Sub

Public Sub RequestCharacterList(Optional ByVal NumCharacters = 8)
    With PBuffer
        .InsertDWord NumCharacters
        .SendPacketMCP MCP_CHARLIST
    End With
End Sub

Public Sub UpgradeCharacter(ByVal CharName As String)
    With PBuffer
        .InsertNTString CharName
        .SendPacketMCP MCP_CHARUPGRADE
    End With
End Sub

'// note: this sub does not use the MCP socket and requires calchashbuf()
Public Sub LogonToRealm(ByVal cookie As Long, ByVal ServerToken As Long, ByVal RealmName As String) ', Optional ByVal RealmPass As String = "password")
    '(DWORD)      Cookie
    '(DWORD[5])   Hashed realm password
    '(STRING)     Realm title
    Dim Buffer As String * 20
    RealmError = False
    
    Buffer = String$(20, vbNullChar)
    Call calcHashBuf("password", 8, Buffer)
    Call calcHashBuf(MKL(cookie) & MKL(ServerToken) & Buffer, 28, Buffer)
    
    'frmChat.AddChat vbRed, RealmName
    
    With PBuffer
        .InsertDWord cookie
        .InsertNonNTString Buffer
        .InsertNTString RealmName
        .SendPacket &H3E
    End With
End Sub

Public Sub ParseMCPPacket(ByVal Data As String)

    On Error GoTo ERROR_HANDLER

    'Debug.Print "MCP Inbound:"
    'Debug.Print DebugOutput(Data)
    
    Static sBuffer As String
    
    Dim CharName As String, StringRaw As String, Statstring As String
    Dim Position As Long, n As Long, PacketLen As Long, seconds As Long
    Dim NullPos  As Integer
    Dim Expiry   As Date
    Dim IncompletePacketInBuffer As Boolean
    Dim veto     As Boolean
    Dim PacketID As Byte
    
    sBuffer = sBuffer & Data
    
    While Len(sBuffer) > 0 And Not IncompletePacketInBuffer
        PacketLen = StringToWord(Mid$(sBuffer, 1, 2))
        
        'Debug.Print "PacketLen: " & PacketLen
        
        If Len(sBuffer) >= PacketLen Then
            Data = Mid$(sBuffer, 1, PacketLen)
            PacketID = Asc(Mid$(Data, 3, 1))
            
            ' ...
            CachePacket StoC, stMCP, PacketID, Len(Data), Data
    
            ' ...
            WritePacketData stMCP, StoC, PacketID, Len(Data), Data
            
            If InStr(1, CommandLine, "-all") > 0 Then
                frmChat.AddChat COLOR_BLUE, "MCP RECV 0x" & ZeroOffset(PacketID, 2)
            End If
                
            ' ...
            If (RunInAll("Event_PacketReceived", "MCP", Asc(Mid$(Data, 3, 1)), Len(Data), Data)) Then
                Exit Sub
            End If
            
            Select Case Asc(Mid$(Data, 3, 1)) '
                Case MCP_STARTUP
                    
                    n = Conv(Mid$(Data, 4, 4))
                    '(DWORD)      Result
                    '0x00: Success
                    '0x0C: No Battle.net connection detected
                    'Debug.Print n
                    Select Case n
                        Case &H0: RaiseEvent RealmStartup(0, "Connected to Diablo II realm.")
                        Case &HC: RaiseEvent RealmStartup(1, "The server thinks you're not connected to Battle.net! Please try your connection again later.")
                        Case &H7E: RaiseEvent RealmStartup(2, "The Diablo II realm rejected your connection. Your CD-Key may have been banned from realm play. Please try your connection again later.")
                        Case &H7F: RaiseEvent RealmStartup(2, "You have been temporarily IP-banned from the Realm server. Please try connecting again later.")
                        Case Else: RaiseEvent RealmStartup(2, "Unknown response to MCP_STARTUP: 0x" & Right$("00000000" & Hex(n), 4))
                    End Select
            
                Case MCP_CHARCREATE
                    '(DWORD)      Result
                    n = Conv(Mid$(Data, 4, 4))
                    '0x00: Success
                    '0x14: Character already exists, or maximum number of characters (currently 8) reached.
                    '0x15: Invalid Name
                    Select Case n
                        Case &H0: RaiseEvent CharCreateResponse(0, "Character creation successful.")
                        Case &H14: RaiseEvent CharCreateResponse(1, "That character name already exists, or exceeds the maximum name Length.")
                        Case &H15: RaiseEvent CharCreateResponse(2, "That character name is invalid.")
                        Case Else: RaiseEvent CharCreateResponse(3, "Unknown response to MCP_CHARCREATE: 0x" & Right$("00000000" & Hex(n), 4))
                    End Select
            
                Case MCP_CHARLOGON
                    
                    n = Conv(Mid$(Data, 4, 4))
                    '(DWORD)      Result
                    '0 x00: Success
                    '0 x46: Player Not Found
                    Select Case n
                        Case &H0: RaiseEvent CharLogonResponse(0, "Realm login successful.")
                        Case &H46: RaiseEvent CharLogonResponse(1, "Player not found.")
                        Case &H7A: RaiseEvent CharLogonResponse(2, "Unable to log on to realm character!")
                        Case &H7B: RaiseEvent CharLogonResponse(3, "That character has expired.")
                        Case Else: RaiseEvent CharLogonResponse(4, "Unknown response to MCP_CHARLOGON: 0x" & Right$("00000000" & Hex(n), 4))
                    End Select
            
                Case MCP_CHARDELETE
                    '(DWORD)      Result
                    'Debug.Print "CharDelete response: " & DebugOutput(Data)
                    n = Conv(Mid$(Data, 5, 4))
                    
                    RaiseEvent CharDeleteResponse(n = 0)
                    '0 x00: Success
            
                Case MCP_MOTD
                    '(STRING)         Unknown - perhaps a headline?
                    'ignored
                    StringRaw = Mid$(Data, InStr(4, Data, Chr(0)) + 1)
                    StringRaw = Left$(StringRaw, InStr(StringRaw, Chr(0)) - 1)
                    '(STRING)         MOTD
                    RaiseEvent RealmMOTD(StringRaw)
            
                Case MCP_CHARLIST
                    '(WORD)       Unknown. Always 8.
                    'IGNORED
                    Position = 6
                    
                    '(DWORD)      Count
                    n = Conv(Mid$(Data, Position, 4))
                    
                    Position = Position + 6
                    '(WORD)       // Number of Characters in the list, again?
                    'skipped
                    
                    RaiseEvent CharListResponse(n)
                    
                    For n = 0 To (n - 1)
                    'For each character:
                    '  (DWORD) Expiration in seconds since 1970
                        StringRaw = Mid$(Data, Position, 4)
                        seconds = Conv(StringRaw)
                        Expiry = DateAdd("s", seconds, "1/1/1970")
                        Position = Position + 4
                        
                    '  (STRING)       Name
                        StringRaw = Mid$(Data, Position)
                        NullPos = InStr(StringRaw, Chr(0))
                        CharName = Left$(StringRaw, NullPos - 1)
                        
                        Position = Position + Len(CharName) + 3
                    '  (WORD)         Flags
                        'ignored
                    '  (STRING)       Modified statstring
                        StringRaw = Mid$(Data, Position)
                        NullPos = InStr(StringRaw, Chr(0))
                        Statstring = Left$(StringRaw, NullPos - 1)
                        
                        Position = Position + Len(Statstring) + 1
                        'Debug.Print "Statstring is " & Len(Statstring) & " bytes long."
                        
                        RaiseEvent CharListEntry(CharName, Statstring, Expiry)
                    Next n
        
                Case MCP_CHARUPGRADE
                    Select Case Conv(Mid$(Data, 4, 4))
                        '(DWORD)      Result
                        '0 x00: Success
                        Case &H0: RaiseEvent CharUpgradeResponse(True)
                        Case Else: RaiseEvent CharUpgradeResponse(False)
                    End Select
                    
            End Select
            
            sBuffer = Mid$(sBuffer, PacketLen + 1)
        Else
            IncompletePacketInBuffer = True
        End If
    Wend
    
    Exit Sub
    
ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & " in ParseMCPPacket()."
    
    Exit Sub
End Sub

Public Property Get CurrentChunk() As String

    CurrentChunk = msCurrentChunk

End Property

Public Property Let CurrentChunk(ByVal sCurrentChunk As String)

    msCurrentChunk = sCurrentChunk

End Property

Public Property Get BNetUniqueUsername() As String

    BNetUniqueUsername = msBNetUniqueUsername

End Property

Public Property Let BNetUniqueUsername(ByVal sBNetUniqueUsername As String)

    msBNetUniqueUsername = sBNetUniqueUsername

End Property


'    StatString << "PX2D";               // Product
'    StatString << "Moo,";               // Realm
'    StatString << "iagovL,";            // Character name, again, seems to do nothing
'    StatString << (BYTE) 0x84; // 0x84 = nothing?
'    StatString << (BYTE) 0x80; // 0x80 = nothing?
'    StatString << (BYTE) 40; // 0x01 = Helmet
'    StatString << (BYTE) 3; // Chest
'    StatString << (BYTE) 3; // Legs
'    StatString << (BYTE) 3; // Right Arm
'    StatString << (BYTE) 3; // left arm
'    StatString << (BYTE) 1; // weapon
'    StatString << (BYTE) 1; // bow?
'    StatString << (BYTE) 1; // shield
'    StatString << (BYTE) 1; // right shoulder
'    StatString << (BYTE) 1; // left shoulder
'    StatString << (BYTE) i; // nothing
'    StatString << (BYTE) 4; // Race - See notebook :)
'    StatString << (BYTE) 1; // helmet color
'    StatString << (BYTE) 1; // chest color
'    StatString << (BYTE) 1; // leg color
'    StatString << (BYTE) 1; // r arm color
'    StatString << (BYTE) 1; // l arm color
'    StatString << (BYTE) 1; // weapon color
'    StatString << (BYTE) 1; // bow color
'    StatString << (BYTE) 1; // shield color
'    StatString << (BYTE) 166; // right shoulder color
'    StatString << (BYTE) 169; // left shoulder color
'    StatString << (BYTE) i;
'    StatString << (BYTE) 0x63; // Level
'    StatString << (BYTE) (0xa0); // 0xa0 = bit 0 = ?, bit 1 = ?, bit 2 = Hardcore, bit 3 = Dead
'    StatString << (BYTE) 0x80; // 0x80 = Rank - see notebook :)
'    StatString << (BYTE) 0xff; // ?
'    StatString << (BYTE) 0xff; // ?
'    StatString << (BYTE) 0xff; // ?
'    StatString << (BYTE) 0x80; // Nothing?
'    StatString << (BYTE) 0x80; // Nothing?
'    StatString << (BYTE) 0x80; // Nothing?
'    StatString << (BYTE) 0; // Null-terminator.
                                '35 BYTES ABOVE

'   53 68 69 67 00      Shig.   '33 BYTES BELOW
'   [!OPEN] [] [BODY INFO      ]                                                                 [ IGNORE     ]
'   84  80  53 03 03 03 03 12 FF  51 02 02 FF  04 ED  02 02 02 02 23 FF  04 02 02 FF  35 A8  9A  FF FF FF FF FF
'   132 128 83 3  3  3  3  18 255 81 2  2  255 4  237 2  2  2  2  35 255  4  2  2 255 53 168 154 FF FF FF FF FF



'From UserLoser:


'0000:  50 58 32 44 55 53 57 65 73 74 2C 49 70 41 64 64   PX2DUSWest,IpAdd
'0010:  52 65 73 53 2C '84 80 01 01 01 01 01 FF FF FF 01   ResS,„€.....ÿÿÿ.
'0020:  01 FF 05 FF FF FF FF FF FF FF FF FF FF FF 01 A0   .ÿ.ÿÿÿÿÿÿÿÿÿÿÿ. 
'0030:  80 FF FF FF 80 80                                 €ÿÿÿ€€

'0000:  50 58 32 44 55 53 45 61 73 74 2C 53 68 61 76 6F   PX2DUSEast,Shavo
'0010:  2C 84 80 39 FF FF FF FF 18 FF 51 FF FF FF 05 4D   ,„€9ÿÿÿÿ.ÿQÿÿÿ.M
'0020:  FF FF FF FF FF FF FF FF FF FF 51 A8 9E FF FF FF   ÿÿÿÿÿÿÿÿÿÿQ¨žÿÿÿ
'0030:  FF FF                                             ÿÿ

'01 'Helmet -2
'01 'Chest  -3
'01 'Legs   -4
'01 'Right Arm  -5
'01 'Weapon -6
'FF 'Bow    -7
'FF 'Sheild -8
'FF 'Right Shoulder     9
'01 'Left Shoulder  10

'.....ÿÿÿ.

'01 'Nothing    11
'FF '?  12
'05 'Charclass - 13
'FF 'Helmet color   14
'FF 'Chest color    15
'FF 'Right arm color    16
'FF 'Left arm color 17
'FF 'Weapon color   18
'FF 'Bow color  -19
'FF 'Sheild color   -20
'FF 'Right shoulder color   -21
'FF 'Left shoulder color    -22
'FF 'Nothing    -23
'FF 'Nothing    -24
'01 'Level  -25
'A0 'bit 0 = ?, bit 1 = ?, bit 2 = Hardcore, bit 3 = Dead   26
'.ÿ.ÿÿÿÿÿÿÿÿÿÿÿ. 
'80 53 02 02 02 02 1C 1C FF 03 03 FF 05 ED 43 43 43 43 30 30 FF 43 43 FF 51 A4 9E FF FF FF FF FF 00

'80 'Rank   27
'FF '?  28
'FF '?  29
'FF '?  30
'80 'Nothing    31
'80 'Nothing    32


' SHIG
'       1   2  3 4  5  6  7  8  9  10 11 12 13 14 15 16
'0000:  53 03 03 03 03 12 FF 51 02 02 FF 04 ED 02 02 02   SÿQÿí

'       17 18 19 20 21 22 23 24 25 26 27 28 29 30 31
'0010:  02 23 FF 04 02 02 FF 35 A8 9A FF FF FF FF FF      #ÿÿ5¨šÿÿÿÿÿ.


'MULEORIFFIC
'       1   2  3 4  5  6  7  8  9  10 11 12 13 14 15 16
'0000:  FF FF FF FF FF FF FF FF FF FF FF 02 FF FF FF FF   ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ

'       17 18 19 20 21 22 23 24 25 26 27 28 29 30 31
'0010:  FF FF FF FF FF FF FF 01 A0 80 FF FF FF FF FF      ÿÿÿÿÿÿÿ €ÿÿÿÿÿ.


'A dumbed-down version of full d2 statstring parsing that returns Values()
' for use in adding people to the character list listview.
Public Sub GetD2CharStats(ByVal sIn As String, ByRef Class As String, ByRef ClassByte As Byte, _
                                ByRef Level As Byte, ByRef IsHardcore As Boolean, ByRef IsDead As Boolean, _
                                ByRef IsLadder As Boolean, ByRef IsExpansion As Boolean)
                                
    'Debug.Print "Statstring:"
    'Debug.Print DebugOutput(sIn)
                                
    Dim D2Classes(0 To 7) As String
        D2Classes(0) = "Unknown Class"
        D2Classes(1) = "Amazon"
        D2Classes(2) = "Sorceress"
        D2Classes(3) = "Necromancer"
        D2Classes(4) = "Paladin"
        D2Classes(5) = "Barbarian"
        D2Classes(6) = "Druid"
        D2Classes(7) = "Assassin"
        
    Dim Current As Byte
    
    If (LenB(sIn) > 0) Then
        Current = Asc(Mid$(sIn, 12, 1)) '// Class
            If Current > 7 Then
                'Debug.Print "Zeroing current. Old value: " & Hex(Current)
                Current = 0
            End If
                    
            ClassByte = Current
            'Debug.Print Current & "\"
            'Debug.Print ClassByte
            Class = D2Classes(Current)
        
        Current = Asc(Mid$(sIn, 24, 1)) '// Level
            Level = Current
        
        
        Current = Asc(Mid$(sIn, 25, 1)) '// Chartype
        
        IsHardcore = (Current And &H4)

        IsLadder = (Current And &H40)
        
        IsDead = (Current And &H8)
        
        IsExpansion = (Current And &H20)
    End If
End Sub
