Attribute VB_Name = "modParsing"
Option Explicit

Public Const COLOR_BLUE2 = 12092001

Public Function StrToHex(ByVal String1 As String, Optional ByVal NoSpaces As Boolean = False) As String
    Dim strTemp As String, strReturn As String, i As Long
    
    For i = 1 To Len(String1)
        strTemp = Hex(Asc(Mid(String1, i, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        
        strReturn = strReturn & IIf(NoSpaces, "", Space(1)) & strTemp
    Next i
        
    StrToHex = strReturn
End Function

Public Function RShift(ByVal pnValue As Long, ByVal pnShift As Long) As Double
    'on error resume next
    RShift = CDbl(pnValue \ (2 ^ pnShift))
End Function


Public Function GetHexValue(ByVal v As Long) As String

    v = v And &HF
    
    If v < 10 Then
    
        GetHexValue = Chr$(v + &H30)
        
    Else
    
        GetHexValue = Chr$(v + &H37)
        
    End If
    
End Function

Public Function GetNumValue(ByVal c As String) As Long
'on error resume next
    c = UCase(c)
    
    If StrictIsNumeric(c) Then
    
        GetNumValue = Asc(c) - &H30
        
    Else
    
        GetNumValue = Asc(c) - &H37
        
    End If
    
End Function

Public Sub NullTruncString(ByRef Text As String)
'on error resume next
    Dim i As Integer
    
    i = InStr(Text, vbNullChar)
    If i = 0 Then Exit Sub
    
    Text = Left$(Text, i - 1)
End Sub

Public Sub FullJoin(Channel As String, Optional ByVal i As Long = -1)
    Dim pBuf As clsDataBuffer
    Set pBuf = New clsDataBuffer
    With pBuf
        If i >= 0 Then
            .InsertDWord CLng(i)
        Else
            .InsertDWord &H2
        End If
        .InsertNTString Channel
        .SendPacket SID_JOINCHANNEL
    End With
    Set pBuf = Nothing
End Sub

Public Function HexToStr(ByVal Hex1 As String) As String
'on error resume next
    Dim strReturn As String, i As Long
    If Len(Hex1) Mod 2 <> 0 Then Exit Function
    For i = 1 To Len(Hex1) Step 2
    strReturn = strReturn & Chr(Val("&H" & Mid(Hex1, i, 2)))
    Next i
    HexToStr = strReturn
End Function

Public Sub RejoinChannel(Channel As String)
    'on error resume next
    Dim pBuf As clsDataBuffer
    Set pBuf = New clsDataBuffer
    With pBuf
        .SendPacket SID_LEAVECHAT
        .Clear
        .InsertDWord &H2
        .InsertNTString Channel
        .SendPacket SID_JOINCHANNEL
    End With
    Set pBuf = Nothing
End Sub

Public Sub RequestProfile(ByVal strUser As String, ByVal eType As enuServerRequestHandlerType, Optional ByRef oCommand As clsCommandObj)
    Dim aKeys(3) As String
    
    aKeys(0) = "Profile\Age"
    aKeys(1) = "Profile\Sex"
    aKeys(2) = "Profile\Location"
    aKeys(3) = "Profile\Description"

    Call RequestUserData(strUser, aKeys, eType, oCommand)
End Sub

Public Sub RequestUserData(ByVal sUsername As String, ByRef aKeys() As String, ByVal HandlerType As enuServerRequestHandlerType, Optional ByRef oCommand As clsCommandObj)
    Dim oRequest As udtServerRequest
    Dim pBuf     As clsDataBuffer
    Dim vArray() As String
    Dim i        As Integer

    With oRequest
        ' Attach handling info
        .ResponseReceived = False
        .HandlerType = HandlerType
        Set .Command = oCommand
        .PacketID = SID_READUSERDATA
        .PacketCommand = 0

        ' Add request data
        ReDim vArray(0 To UBound(aKeys) + 1)
        vArray(0) = sUsername
        For i = 0 To UBound(aKeys)
            vArray(i + 1) = aKeys(i)
        Next i
        .Tag = CVar(vArray)
    End With
    
    Call SaveServerRequest(oRequest)

    ' Build the packet
    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord 1
        .InsertDWord UBound(vArray)
        .InsertDWord oRequest.Cookie

        .InsertNTString CleanUsername(ReverseConvertUsernameGateway(sUsername))

        For i = 1 To UBound(vArray)
            .InsertNTString vArray(i)
        Next i

        .SendPacket SID_READUSERDATA
    End With
    Set pBuf = Nothing
End Sub

Public Sub SetProfile(ByVal Location As String, ByVal Description As String, Optional ByVal Sex As String = vbNullString)
    'Dim i As Byte
    Const MAX_DESCR As Long = 510
    Const MAX_SEX As Long = 200
    Const MAX_LOC As Long = 200
    Dim pBuf As clsDataBuffer
    Dim Encoding As STRINGENCODING

    If (Config.UseUTF8) Then
        Encoding = UTF8
    Else
        Encoding = ANSI
    End If
    
    '// Sanity checks
    If Len(Description) > MAX_DESCR Then
        Description = Left$(Description, MAX_DESCR)
    End If
    
    If Len(Sex) > MAX_SEX Then
        Sex = Left$(Sex, MAX_SEX)
    End If
    
    If Len(Location) > MAX_LOC Then
        Location = Left$(Location, MAX_LOC)
    End If

    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord &H1                    '// #accounts
        .InsertDWord 3                      '// #keys
        
        .InsertNTString CurrentUsername     '// account to update
                                            '// keys
        .InsertNTString "Profile\Location"
        .InsertNTString "Profile\Description"
        .InsertNTString "Profile\Sex"
                                            '// Values()
        .InsertNTString Location, Encoding
        .InsertNTString Description, Encoding
        .InsertNTString Sex, Encoding
        
        .SendPacket SID_WRITEUSERDATA
    End With
    Set pBuf = Nothing
End Sub

'// Extended version of this function for scripting use
'//  Will not ERASE if a field is left blank
'// 2007-06-07: SEX value is ignored because Blizzard removed that
'//     field from profiles
'// 2009-07-14: corrected a problem in this method, thanks Jack (t=42494) -andy
'//     method was erasing profile data
Public Sub SetProfileEx(ByVal Location As String, ByVal Description As String)
    'Dim i As Byte
    Const MAX_DESCR As Long = 510
    Const MAX_SEX As Long = 200
    Const MAX_LOC As Long = 200
    
    Dim nKeys As Integer, i As Integer
    Dim pKeys(1 To 3) As String
    Dim pData(1 To 3) As String
    Dim pBuf As clsDataBuffer
    
    If (LenB(Location) > 0) Then
        If (Len(Location) > MAX_LOC) Then
            Location = Left$(Location, MAX_LOC)
        End If
        
        nKeys = nKeys + 1
        pKeys(nKeys) = "Profile\Location"
        pData(nKeys) = Location
    End If
    
    '// Sanity checks
    If (LenB(Description) > 0) Then
        If (Len(Description) > MAX_DESCR) Then
            Description = Left$(Description, MAX_DESCR)
        End If
        
        nKeys = nKeys + 1
        pKeys(nKeys) = "Profile\Description"
        pData(nKeys) = Description
    End If
        
    If nKeys > 0 Then
        Set pBuf = New clsDataBuffer
        With pBuf
            .InsertDWord &H1                    '// #accounts
            .InsertDWord nKeys                  '// #keys
            .InsertNTString CurrentUsername     '// account to update
                                                '// keys
            For i = 1 To nKeys
                .InsertNTString pKeys(i)
            Next i
           
            '// Values()
            For i = 1 To nKeys
                .InsertNTString pData(i)
            Next i
            
            .SendPacket SID_WRITEUSERDATA
        End With
        Set pBuf = Nothing
    End If
End Sub

Public Function SaveServerRequest(ByRef oRequest As udtServerRequest) As Long
    Dim i          As Integer
    Dim bFoundSlot As Boolean

    ' Find an open slot in the request list
    bFoundSlot = False
    If UBound(ServerRequests) > LBound(ServerRequests) Then
        For i = 1 To UBound(ServerRequests)
            If ServerRequests(i).ResponseReceived Then
                bFoundSlot = True
                oRequest.Cookie = i
                ServerRequests(i) = oRequest
                Exit For
            End If
        Next
    End If

    ' If no slot was found, add the request to the end
    If Not bFoundSlot Then
        oRequest.Cookie = UBound(ServerRequests) + 1
        ReDim Preserve ServerRequests(oRequest.Cookie)
        ServerRequests(oRequest.Cookie) = oRequest
    End If

    'frmChat.AddChat vbWhite, StringFormat("Saved request (0x{0}/0x{1}): 0x{2}", ZeroOffset(oRequest.PacketID, 2), ZeroOffset(oRequest.PacketCommand, 2), ZeroOffset(oRequest.Cookie, 4))
    SaveServerRequest = oRequest.Cookie
End Function

Public Function FindServerRequest(ByRef oRequest As udtServerRequest, ByVal Cookie As Long, Optional ByVal PacketID As Byte = 0, Optional ByVal PacketCommand As Byte, Optional ByVal ResponseReceived As Boolean = True) As Boolean
    Dim i As Integer

    FindServerRequest = False

    ' we can search for any server request by cookie or by packet ID, but not by any of both
    If Cookie < LBound(ServerRequests) And PacketID = 0 Then
        Exit Function
    End If

    ' Find the request for this ID and hand it off to the event handler
    ' pass in Cookie = -1 or Cookie = 0 to find the first un-received request with that packet ID
    ' (i.e. with ResponseReceived = False to "peek" at a pending request)
    If Cookie <= 0 Then
        For i = 1 To UBound(ServerRequests)
            With ServerRequests(i)
                If Not .ResponseReceived And .PacketID = PacketID And .PacketCommand = PacketCommand Then
                    Cookie = i
                    Exit For
                End If
            End With
        Next i
        ' not found
        If Cookie <= 0 Then Exit Function
    End If

    If Cookie <= UBound(ServerRequests) Then
        ' Process the request
        With ServerRequests(Cookie)
            If PacketID > 0 Then
                If .PacketID <> PacketID Or .PacketCommand <> PacketCommand Then
                    frmChat.AddChat g_Color.ErrorMessageText, StringFormat("Error: Received data response for a different packet: 0x{0}/0x{1} instead of 0x{2}/0x{3}", ZeroOffset(PacketID, 2), ZeroOffset(PacketCommand, 2), ZeroOffset(.PacketID, 2), ZeroOffset(.PacketCommand, 2))
                    Exit Function
                End If
            End If

            If .ResponseReceived Then
                frmChat.AddChat g_Color.ErrorMessageText, StringFormat("Notice: Received extra data response for packet: 0x{0}/0x{1}", ZeroOffset(PacketID, 2), ZeroOffset(PacketCommand, 2))
            End If

            .ResponseReceived = ResponseReceived
            'If ResponseReceived Then
            '    frmChat.AddChat vbWhite, StringFormat("Found request (0x{0}/0x{1}): 0x{2}", ZeroOffset(.PacketID, 2), ZeroOffset(.PacketCommand, 2), ZeroOffset(.Cookie, 4))
            'End If

            oRequest.ResponseReceived = ResponseReceived
            oRequest.HandlerType = .HandlerType
            Set oRequest.Command = .Command
            oRequest.PacketID = .PacketID
            oRequest.PacketCommand = .PacketCommand
            oRequest.Cookie = .Cookie
            oRequest.Tag = .Tag
        End With

        FindServerRequest = True

        ' Shrink the array if we are able (remove all where ResponseReceived is True from the end of the array)
        If Cookie > 1 Then
            For i = Cookie To 2 Step -1
                If Not ServerRequests(i).ResponseReceived Then
                    i = -1
                    Exit For
                End If
            Next i
            If i > 1 Then ReDim Preserve ServerRequests(i - 1)
        End If
    End If
End Function

