Attribute VB_Name = "modParsing"
Option Explicit

Public Const COLOR_BLUE2 = 12092001

Public Sub SendHeader()
    frmChat.sckBNet.SendData ChrW(1)
End Sub

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
    
    i = InStr(Text, Chr(0))
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

Public Sub RequestProfile(strUser As String, ByVal eType As enuUserDataRequestType, Optional ByRef oCommand As clsCommandObj)
    Dim aKeys(3) As String
    
    aKeys(0) = "Profile\Age"
    aKeys(1) = "Profile\Sex"
    aKeys(2) = "Profile\Location"
    aKeys(3) = "Profile\Description"
    
    Call RequestUserData(strUser, aKeys, eType, oCommand)
End Sub

Public Sub RequestUserData(ByVal sUsername As String, ByRef aKeys() As String, Optional ByVal eType As enuUserDataRequestType, Optional ByRef oCommand As clsCommandObj)
    Dim oRequest As udtUserDataRequest
    Dim i As Integer
    Dim bFoundSlot As Boolean
    Dim pBuf As clsDataBuffer

    With oRequest
        ' Attach handling info
        .RequestType = eType
        Set .Command = oCommand
        .ResponseReceived = False
    
        ' Add request data
        .Account = sUsername
        .keys = aKeys
    End With
    
    ' Find an open slot in the request list
    bFoundSlot = False
    If UBound(UserDataRequests) > 0 Then
        For i = 1 To UBound(UserDataRequests)
            If UserDataRequests(i).ResponseReceived Then
                bFoundSlot = True
            
                oRequest.RequestID = i
                UserDataRequests(i) = oRequest
                Exit For
            End If
        Next
    End If
    
    ' If no slot was found, add the request to the end
    If Not bFoundSlot Then
        oRequest.RequestID = UBound(UserDataRequests) + 1
        
        ReDim Preserve UserDataRequests(oRequest.RequestID)
        UserDataRequests(oRequest.RequestID) = oRequest
    End If

    ' Build the packet
    Set pBuf = New clsDataBuffer
    With pBuf
        .InsertDWord 1
        .InsertDWord UBound(oRequest.keys) + 1
        .InsertDWord oRequest.RequestID
        
        .InsertNTString CleanUsername(ReverseConvertUsernameGateway(oRequest.Account))
        
        For i = 0 To UBound(aKeys)
            .InsertNTString oRequest.keys(i)
        Next

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
        .InsertNTString Location
        .InsertNTString Description
        .InsertNTString Sex
        
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

Public Function Conv(ByVal RawString As String) As Long
    Dim lReturn As Long
    
    If Len(RawString) = 4 Then
        Call CopyMemory(lReturn, ByVal RawString, 4)
    Else
        Debug.Print "---------- WARNING! Invalid string Length in Conv()!"
        Debug.Print "---------- Length: " & Len(RawString)
        Debug.Print DebugOutput(RawString)
    End If
    
    Conv = lReturn
End Function

'// COLORMODIFY - where L is passed as the start position of the text to be checked
Public Sub ColorModify(ByRef rtb As RichTextBox, ByRef L As Long)
    Dim i As Long
    Dim s As String
    Dim temp As Long
    Dim SelStart As Long
    Dim SelLength As Long
    
    If L = 0 Then L = 1
    
    temp = L
    
    With rtb
        ' store previous selstart and len
        SelStart = .SelStart
        SelLength = .SelLength
        
        If InStr(temp, .Text, "ÿc", vbTextCompare) > 0 Then
            .Visible = False
            Do
                i = InStr(temp, .Text, "ÿc", vbTextCompare)
                
                If StrictIsNumeric(Mid$(.Text, i + 2, 1)) Then
                    s = GetColorVal(Mid$(.Text, i + 2, 1))
                    .SelStart = i - 1
                    .SelLength = 3
                    .SelText = vbNullString
                    .SelStart = i - 1
                    .SelLength = Len(.Text) + 1 - i
                    .SelColor = s
                Else
                    Select Case Mid$(.Text, i + 2, 1)
                        Case "i"
                            .SelStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .SelStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            If .SelItalic = True Then
                                .SelItalic = False
                            Else
                                .SelItalic = True
                            End If
                            
                        Case "b", "."       'BOLD
                            .SelStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .SelStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            If .SelBold = True Then
                                .SelBold = False
                            Else
                                .SelBold = True
                            End If
                            
                        Case "u", "."       'underline
                            .SelStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .SelStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            If .SelUnderline = True Then
                                .SelUnderline = False
                            Else
                                .SelUnderline = True
                            End If
                            
                        Case ";"
                            .SelStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .SelStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            .SelColor = HTMLToRGBColor("8D00CE")    'Purple
                            
                        Case ":"
                            .SelStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .SelStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            .SelColor = 186408      '// Lighter green
                            
                        Case "<"
                            .SelStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .SelStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            .SelColor = HTMLToRGBColor("00A200")    'Dark green
                        'Case Else: Debug.Print s
                    End Select
                End If
                temp = temp + 1
                
            Loop While InStr(temp, .Text, "ÿc", vbTextCompare) > 0
            .Visible = True
        End If
        
        '// Check for SC color codes
        temp = L
        
        If InStr(temp, .Text, "Á", vbBinaryCompare) > 0 Then
            .Visible = False
            Do
                i = InStr(temp, .Text, "Á", vbBinaryCompare)
                s = GetScriptColorString(Mid$(.Text, i + 1, 1))
                
                If Len(s) > 0 Then
                    .Visible = False
                    .SelStart = i - 1
                    .SelLength = 2
                    .SelText = vbNullString
                    .SelStart = i - 1
                    .SelLength = Len(.Text) + 1 - 1
                    .SelColor = s
                    .Visible = True
                End If
                
                temp = temp + 1
                
            Loop While InStr(temp, .Text, "Á", vbBinaryCompare) > 0
            .Visible = True
        End If
        
        ' restore previous selstart and len
        .SelStart = SelStart
        .SelLength = SelLength
    End With
End Sub

Public Function GetScriptColorString(ByVal scCC As String) As String
    Select Case Asc(scCC)
        Case Asc("Q"): GetScriptColorString = RGB(93, 93, 93)       'Grey
        Case Asc("R"): GetScriptColorString = RGB(30, 224, 54)      'Green
        Case Asc("Z"), Asc("X"), Asc("S"): GetScriptColorString = RGB(160, 169, 116)    'Yellow
        Case Asc("Y"), Asc("["), Asc("Y"): GetScriptColorString = RGB(231, 38, 82)      'Red
        Case Asc("V"), Asc("@"): GetScriptColorString = RGB(98, 77, 232)      'Blue
        Case Asc("W"), Asc("P"): GetScriptColorString = vbWhite               'White
        Case Asc("T"), Asc("U"), Asc("V"): GetScriptColorString = HTMLToRGBColor("00CCCC") 'cyan/teal
        Case Else: GetScriptColorString = vbNullString
    End Select
End Function

Public Function GetColorVal(ByVal d2CC As String) As String
    Select Case CInt(d2CC)
        Case 1: GetColorVal = HTMLToRGBColor("CE3E3E")  'Red
        Case 2: GetColorVal = HTMLToRGBColor("00CE00")  'Green
        Case 3: GetColorVal = HTMLToRGBColor("44409C")  'Blue
        Case 4: GetColorVal = HTMLToRGBColor("A19160")  'Gold
        Case 5: GetColorVal = HTMLToRGBColor("555555")  'Grey
        Case 6: GetColorVal = HTMLToRGBColor("080808")  'Black
        Case 7: GetColorVal = HTMLToRGBColor("A89D65")  'Gold
        Case 8: GetColorVal = HTMLToRGBColor("CE8800")  'Gold-Orange
        Case 9: GetColorVal = HTMLToRGBColor("CECE51")  'Light Yellow
        Case 0: GetColorVal = HTMLToRGBColor("FFFFFF")  'White
    End Select
End Function


