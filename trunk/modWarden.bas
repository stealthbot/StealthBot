Attribute VB_Name = "modWarden"
'// thx to ringo <3

'//Does the job of Maiev.mod
Private Type RANDOMDATA
    pos                     As Long
    data                    As String * 20
    Sorc1                   As String * 20
    Sorc2                   As String * 20
End Type
Private m_Parse(5)          As Long
Private m_CallBack(7)       As Long 'callback function list, for warden
Private m_Func(2)           As Long 'wardens exports
Private m_KeyOut(257)       As Byte
Private m_KeyIn(257)        As Byte
Private m_Seed              As Long
Private m_Mod               As Long 'pointer to the module
Private m_ModMem            As Long 'pointer to wardens memory block
Private m_ModState          As Byte '0=idle,1=downloading,2=hackyhacky
Private m_RC4               As Long
Private m_PKT               As String
'//Warden download stuff
Private m_ModName           As String * 16 'the modules name
Private m_ModFolder         As String 'the warden folder
Private m_ModKey(15)        As Byte 'key to crypt module with
Private m_ModLen            As Long 'lengh of downloading module
Private m_ModPos            As Long 'position in write data for downloading
Private m_ModData()         As Byte 'module download buffer

Public Sub WardenInit(ByVal lngSeed As Long)
    Dim bOut(15)        As Byte
    Dim bIn(15)         As Byte
    Dim uRan            As RANDOMDATA
    
    If (Not CanHandleWarden()) Then
        Call frmChat.AddChat(vbRed, "[Warden] Warden support has not been initialized because zlib.dll could not be found.")
        Exit Sub
    End If
    
    Call WardenCleanUp
    '//Create new RC4 Keys
    m_Seed = lngSeed
    Call Data_Init(uRan, lngSeed)
    Call Data_Get_Bytes(uRan, bOut(), 16)
    Call Data_Get_Bytes(uRan, bIn(), 16)
    Call RC4Key(bOut(), m_KeyOut(), 16)
    Call RC4Key(bIn(), m_KeyIn(), 16)
    m_Parse(0) = Addr2Ptr(AddressOf HW0x00)
    m_Parse(1) = Addr2Ptr(AddressOf HW0x01)
    m_Parse(2) = Addr2Ptr(AddressOf HW0x02)
    m_Parse(3) = Addr2Ptr(AddressOf HW0x03)
    m_Parse(4) = Addr2Ptr(AddressOf HW0x04)
    m_Parse(5) = Addr2Ptr(AddressOf HW0x05)
    m_ModFolder = App.Path & "\Warden\"
    
    Call frmChat.AddChat(vbGreen, "[Warden] Initialized!")
End Sub
Public Sub WardenCleanUp()
    '//Unload any existing module
    If m_Mod Then
        Call UnloadModule
        Call free(m_Mod)
        Call ZeroMemory(m_Func(0), 12)
        m_Mod = 0
        m_ModMem = 0
    End If
    m_ModState = 0
    '//Clear download variables
    Call ZeroMemory(ByVal m_ModName, 16)
    Call ZeroMemory(m_ModKey(0), 16)
    m_ModLen = 0
    m_ModPos = 0
    Erase m_ModData()
End Sub
Private Function LoadModule(ByVal lngMod As Long, ByRef strPath As String) As Long
    Dim bData()         As Byte
    Dim I               As Long
    If (lngMod = 0) Then
        I = FreeFile
        Open strPath For Binary Lock Read As #I
            If (LOF(I) < 1) Then
                Close #I
                Exit Function
            End If
            ReDim bData(LOF(I))
            Get #I, 1, bData()
        Close #I
        lngMod = VarPtr(bData(0))
    End If
    If m_Mod Then
        Call UnloadModule
        Call free(m_Mod)
        Call ZeroMemory(m_Func(0), 12)
    End If
    m_ModMem = 0
    m_Mod = PrepareModule(lngMod)
    If (m_Mod = 0) Then Exit Function
    Call InitModule
    If (m_ModMem = 0) Then
        Call free(m_Mod)
        Exit Function
    End If
    Call CopyMemory(I, ByVal m_ModMem, 4)
    Call CopyMemory(m_Func(0), ByVal I, 12)
    Call frmChat.AddChat(vbGreen, "[Warden] Loaded module: ", vbWhite, StrToHex(m_ModName, True) & ".bin")
    LoadModule = 1
End Function
Public Sub WardenOnData(ByRef s As String)
    Dim lngData         As Long
    Dim lngLengh        As Long
    Dim lngID           As Long
    
    If (Not CanHandleWarden()) Then
        Exit Sub
    End If
    
    lngLengh = (Len(s) - 4)
    If (lngLengh < 1) Then Exit Sub
    lngData = malloc(lngLengh)
    Call RC4CryptStr(s, m_KeyIn(), 5)
    lngID = Asc(Mid$(s, 5, 1))
    If (lngID < 6) Then
        Call CopyMemory(ByVal lngData, ByVal Mid$(s, 5, lngLengh), lngLengh)
        Call CallWindowProcA(m_Parse(lngID), lngData, lngID, lngLengh, 0)
    End If
    Call free(lngData)
End Sub
Private Function HW0x00(ByVal hData As Long, ByVal uMsg As Long, ByVal wLen As Long, ByVal lParam As Long) As Long
    Dim s           As String
    Call WardenCleanUp
    If (wLen < 37) Then Exit Function
    Call CopyMemory(ByVal m_ModName, ByVal hData + 1, 16)
    Call CopyMemory(m_ModKey(0), ByVal hData + 17, 16)
    Call CopyMemory(m_ModLen, ByVal hData + 33, 4)
    s = m_ModFolder & StrToHex(m_ModName, True) & ".bin"
    If (Len(Dir$(s)) = 0) Then
        If (m_ModLen < 50) Or (m_ModLen > 5000000) Then
            m_ModLen = 0
            Exit Function
        End If
        s = vbNullChar
        ReDim m_ModData(m_ModLen - 1)
        m_ModState = 1
        Call frmChat.AddChat(vbYellow, "[Warden] Downloading module: ", vbWhite, StrToHex(m_ModName, True) & ".bin")
    Else
        If (LoadModule(0, s) = 0) Then Exit Function
        s = Chr$(&H1)
        m_ModLen = 0
        m_ModState = 2
    End If
    m_ModPos = 0
    Call RC4CryptStr(s, m_KeyOut(), 1)
    Call modOutboundPackets.Send0x5E(s)
    HW0x00 = 1
End Function
Private Function HW0x01(ByVal hData As Long, ByVal uMsg As Long, ByVal wLen As Long, ByVal lParam As Long) As Long
    If (Not m_ModState = 1) Then Exit Function
    If (m_ModLen = 0) Then Exit Function
    If (wLen < 4) Then Exit Function
    Call CopyMemory(m_ModData(m_ModPos), ByVal hData + 3, wLen - 3)
    m_ModPos = m_ModPos + (wLen - 3)
    'Debug.Print m_ModPos & " Of " & m_ModLen
    If (m_ModPos >= m_ModLen) Then
        m_ModState = 2
        HW0x01 = HW0x01Ex()
    Else
        HW0x01 = 1
    End If
End Function
Private Function HW0x01Ex() As Long
    On Error GoTo HW0x01ExErr
    Dim bData()         As Byte
    Dim I               As Long
    Dim s               As String
    ReDim bData(257)
    Call RC4Key(m_ModKey(), bData(), 16)
    Call RC4Crypt(m_ModData(), bData(), m_ModLen)
    Call CopyMemory(I, m_ModData(0), 4)
    If (I < &H120) Or (I > 5000000) Then GoTo HW0x01ExErr
    ReDim bData(I - 1)
    If (Not uncompress(bData(0), I, m_ModData(4), CLng(m_ModLen - &H108)) = 0) Then GoTo HW0x01ExErr
    m_ModLen = 0
    m_ModPos = 0
    Erase m_ModData()
    s = m_ModFolder & StrToHex(m_ModName, True) & ".bin"
    I = FreeFile
    Open s For Binary Lock Write As #I
        Put #I, 1, bData()
    Close #I
    If (LoadModule(VarPtr(bData(0)), s) = 0) Then GoTo HW0x01ExErr
    m_ModState = 2
    bData(0) = 1
    Call RC4Crypt(bData(), m_KeyOut(), 1)
    Call modOutboundPackets.Send0x5E(Chr$(bData(0)))
    Erase bData()
    HW0x01Ex = 1
    Exit Function
HW0x01ExErr:
    Erase m_ModData()
    m_ModLen = 0
    m_ModPos = 0
    m_ModState = 0
    Debug.Print "HW0x01Ex() Error: " & Err.description
End Function
Private Function HW0x02(ByVal hData As Long, ByVal uMsg As Long, ByVal wLen As Long, ByVal lParam As Long) As Long
    On Error GoTo HW0x02Err
    'eww, yep
    Dim s               As String
    Dim strData         As String
    Dim P               As Long
    Dim strOut          As String
    Dim PosOut          As Long
    Dim bHeader(6)      As Byte
    
    If (Not m_ModState = 2) Then
        Debug.Print "HW0x02() Error: m_ModState == " & m_ModState & " (needs to be 2)"
        Exit Function
    End If
    
    If (wLen < 2) Then
        Debug.Print "HW0x02() Error: wLen == " & wLen & " (needs to be >2)"
        Exit Function
    End If
    
    s = Space(wLen)
    Call CopyMemory(ByVal s, ByVal hData, wLen)
    
    If (Not Asc(Mid$(s, 2, 1)) = 0) Then
        Debug.Print "HW0x02() Error: Asc(Mid$(s, 2, 1)) == " & Asc(Mid$(s, 2, 1)) & " (needs to be 0)"
        Exit Function
    End If
    
    P = 3
    PosOut = 8
    strOut = Space(255) 'max size are are send buffer
    Do Until (P >= wLen)
        strData = Get0x02Data(s, P)
        
        If (Len(strData) = 0) Then
            Debug.Print "HW0x02() Error: Len(strData) == 0"
            Exit Function
        End If
        
        Mid$(strOut, PosOut, Len(strData)) = strData
        PosOut = PosOut + Len(strData)
    Loop
    strOut = Left$(strOut, (PosOut - 1))
    bHeader(0) = &H2
    Call CopyMemory(bHeader(1), CInt(PosOut - 8), 2)
    Call CopyMemory(bHeader(3), WardenChecksum(Mid$(strOut, 8)), 4)
    Call CopyMemory(ByVal strOut, bHeader(0), 7)
    Call RC4CryptStr(strOut, m_KeyOut(), 1)
    Call modOutboundPackets.Send0x5E(strOut)
    HW0x02 = 1
    Exit Function
HW0x02Err:
    Debug.Print "HW0x02() Error: " & Err.description
End Function
Private Function HW0x03(ByVal hData As Long, ByVal uMsg As Long, ByVal wLen As Long, ByVal lParam As Long) As Long
    '//Ignore this
End Function
Private Function HW0x04(ByVal hData As Long, ByVal uMsg As Long, ByVal wLen As Long, ByVal lParam As Long) As Long
    '//Ignore this
End Function
Private Function HW0x05(ByVal hData As Long, ByVal uMsg As Long, ByVal wLen As Long, ByVal lParam As Long) As Long
    Dim ASM             As clsASM
    Dim bKey(257)       As Byte
    Dim bData()         As Byte
    Dim lngRecv         As Long
    If (Not m_ModState = 2) Then Exit Function
    Set ASM = New clsASM
    m_RC4 = 0
    m_ModState = 0
    With ASM
        .push_v32 (4)
        .push_v32 (VarPtr(m_Seed))
        .mov__ecx_v32 (m_ModMem)
        .xor__edx_edx
        .mov__eax_v32 (m_Func(0))
        .call_eax
        .retn 8
        .Execute
    End With
    Set ASM = Nothing
    If (m_RC4 = 0) Then Exit Function
    ReDim bData(wLen - 1)
    Call CopyMemory(bData(0), ByVal hData, wLen)
    Call CopyMemory(bKey(0), ByVal m_RC4 + 258, 258)
    Call RC4Crypt(bData(), bKey(), wLen)
    Call CopyMemory(bKey(0), ByVal m_RC4, 258)
    m_PKT = vbNullString
    Set ASM = New clsASM
    With ASM
        .push_v32 (VarPtr(lngRecv))
        .push_v32 (wLen)
        .push_v32 (VarPtr(bData(0)))
        .mov__ecx_v32 (m_ModMem)
        .xor__edx_edx
        .mov__eax_v32 (m_Func(2))
        .call_eax
        .retn 8
        .Execute
    End With
    Set ASM = Nothing
    If (Len(m_PKT) = 0) Then Exit Function
    Call RC4CryptStr(m_PKT, bKey(), 1)
    Call RC4CryptStr(m_PKT, m_KeyOut(), 1)
    Call CopyMemory(m_KeyOut(0), ByVal m_RC4, 258)
    Call CopyMemory(m_KeyIn(0), ByVal m_RC4 + 258, 258)
    m_ModState = 2
    Call modOutboundPackets.Send0x5E(m_PKT)
    m_RC4 = 0
    m_PKT = vbNullString
End Function


Private Function PrepareModule(ByRef pModule As Long) As Long
    '//carbon copy port from iagos code
    Debug.Print "PrepareModule()"
    Dim dwModuleSize        As Long
    Dim pNewModule          As Long
    dwModuleSize = getInteger(pModule, &H0)
    pNewModule = malloc(dwModuleSize)
    Call ZeroMemory(ByVal pNewModule, dwModuleSize)
    Debug.Print "   Allocated " & dwModuleSize & " (0x" & hex(dwModuleSize) & ") bytes for new module"
    Call CopyMemory(ByVal pNewModule, ByVal pModule, 40)
    Dim dwSrcLocation       As Long
    Dim dwDestLocation      As Long
    Dim dwLimit             As Long
    dwSrcLocation = &H28 + (getInteger(pNewModule, &H24) * 12)
    dwDestLocation = getInteger(pModule, &H28)
    dwLimit = getInteger(pModule, &H0)
    Dim bSkip               As Boolean
    Debug.Print "   Copying code sections to module."
    While (dwDestLocation < dwLimit)
        Dim dwCount         As Long
        Call CopyMemory(ByVal VarPtr(dwCount), ByVal pModule + dwSrcLocation, 1)
        Call CopyMemory(ByVal VarPtr(dwCount) + 1, ByVal pModule + dwSrcLocation + 1, 1)
        dwSrcLocation = dwSrcLocation + 2
        If (bSkip = False) Then
            Call CopyMemory(ByVal pNewModule + dwDestLocation, ByVal pModule + dwSrcLocation, dwCount)
            dwSrcLocation = dwSrcLocation + dwCount
        End If
        bSkip = Not bSkip
        dwDestLocation = dwDestLocation + dwCount
    Wend
    Debug.Print "   Adjusting references to global variables..."
    dwSrcLocation = getInteger(pModule, 8)
    dwDestLocation = 0
    Dim I                       As Long
    Dim lng0x0C                 As Long
    Dim lngTest                 As Long
    Call CopyMemory(lng0x0C, ByVal pNewModule + &HC, 4)
    While (I < lng0x0C)
        Call CopyMemory(lngTest, ByVal pNewModule + dwSrcLocation, 1)
        lngTest = lngTest And &HFF&
        Call CopyMemory(ByVal VarPtr(lngTest) + 0, ByVal pNewModule + dwSrcLocation + 1, 1)
        Call CopyMemory(ByVal VarPtr(lngTest) + 1, ByVal pNewModule + dwSrcLocation, 1)
        dwDestLocation = dwDestLocation + lngTest
        dwSrcLocation = dwSrcLocation + 2
        Call insertInteger(pNewModule, dwDestLocation, getInteger(pNewModule, dwDestLocation) + pNewModule)
        I = I + 1
    Wend
    Debug.Print "   Updating API library references.."
    dwLimit = getInteger(pNewModule, &H20)
    Dim dwProcStart             As Long
    Dim szLib                   As String
    Dim dwProcOffset            As Long
    Dim hModule                 As Long
    Dim dwProc                  As Long
    Dim szFunc                  As String
    For I = 0 To dwLimit - 1
        dwProcStart = getInteger(pNewModule, &H1C) + (I * 8)
        szLib = GetSTR(pNewModule + getInteger(pNewModule, dwProcStart))
        dwProcOffset = getInteger(pNewModule, dwProcStart + 4)
        Debug.Print "   Lib: " & szLib
        hModule = LoadLibraryA(szLib)
        dwProc = getInteger(pNewModule, dwProcOffset)
        While dwProc
            If (dwProc > 0) Then
                szFunc = GetSTR(pNewModule + dwProc)
                Debug.Print "       Function: " & szFunc
                Call insertInteger(pNewModule, dwProcOffset, GetProcAddress(hModule, szFunc))
            Else
                dwProc = dwProc And &H7FFFFFFF
                Debug.Print "       Ordinary: 0x" & hex(dwProc)
            End If
            dwProcOffset = dwProcOffset + 4
            dwProc = getInteger(pNewModule, dwProcOffset)
        Wend
    Next I
    Debug.Print "   Successfully mapped Warden Module to 0x" & hex(pNewModule)
    PrepareModule = pNewModule
End Function
Private Sub InitModule()
    Debug.Print "InitModule()"
    Dim A               As Long
    Dim B               As Long
    Dim C               As Long
    C = getInteger(m_Mod, &H18)
    B = 1 - C
    If (B > getInteger(m_Mod, &H14)) Then Exit Sub
    A = getInteger(m_Mod, &H10)
    A = getInteger(m_Mod, A + (B * 4)) + m_Mod
    Debug.Print "   Initialize Function is mapped at 0x" & hex(A)
    m_CallBack(0) = Addr2Ptr(AddressOf SendPacket)
    m_CallBack(1) = Addr2Ptr(AddressOf CheckModule)
    m_CallBack(2) = Addr2Ptr(AddressOf ModuleLoad)
    m_CallBack(3) = Addr2Ptr(AddressOf AllocateMem)
    m_CallBack(4) = Addr2Ptr(AddressOf FreeMemory)
    m_CallBack(5) = Addr2Ptr(AddressOf SetRC4Data)
    m_CallBack(6) = Addr2Ptr(AddressOf GetRC4Data)
    m_CallBack(7) = VarPtr(m_CallBack(0))
    Dim ASM         As New clsASM
    With ASM
        .mov__ecx_v32 (VarPtr(m_CallBack(7)))
        .call_ptr (VarPtr(A))
        .mov__ptr_eax (VarPtr(m_ModMem))
        .retn 8
        .Execute
    End With
    Set ASM = Nothing
End Sub
Private Sub UnloadModule()
    Dim ASM         As New clsASM
    With ASM
        .mov__ecx_v32 (m_ModMem)
        .call_ptr (VarPtr(m_Func(1)))
        .retn 8
        .Execute
    End With
    Set ASM = Nothing
End Sub



Private Sub SendPacket(ByVal ptrPacket As Long, ByVal dwSize As Long)
    If (dwSize < 1) Then Exit Sub
    If (dwSize > 5000) Then Exit Sub
    m_PKT = Space(dwSize)
    Call CopyMemory(ByVal m_PKT, ByVal ptrPacket, dwSize)
    'Debug.Print "Warden.SendPacket() pkt=0x" & Hex(ptrPacket) & ", size=" & dwSize & vbCrLf & GetLog(m_PKT)
End Sub
Private Function CheckModule(ByVal ptrMod As Long, ByVal ptrKey As Long) As Long
    'Debug.Print "Warden.CheckModule() " & ptrMod & "/" & ptrKey
    'CheckModule = 0 '//Need to download
    'CheckModule = 1 '//Don't need to download
    CheckModule = 1
End Function
Private Function ModuleLoad(ByVal ptrRC4Key As Long, ByVal pModule As Long, ByVal dwModSize As Long) As Long
    'Debug.Print "Warden.ModuleLoad() " & ptrMod & "/" & ptrKey
    'ModuleLoad = 0 '//Need to download
    'ModuleLoad = 1 '//Don't need to download
    ModuleLoad = 1
End Function
Private Function AllocateMem(ByVal dwSize As Long) As Long
    AllocateMem = malloc(dwSize)
End Function
Private Sub FreeMemory(ByVal dwMemory As Long)
    Call free(dwMemory)
    'Debug.Print "Warden.FreeMemory() 0x" & Hex(dwMemory)
End Sub
Private Function SetRC4Data(ByVal lpKeys As Long, ByVal dwSize As Long) As Long
    'Debug.Print "Warden.SetRC4Data() 0x" & Hex(lpKeys) & "/0x" & Hex(dwSize)
End Function
Private Function GetRC4Data(ByVal lpBuffer As Long, ByRef dwSize As Long) As Long
    'Debug.Print "Warden.GetRC4Data() 0x" & Hex(lpBuffer) & "/0x" & Hex(dwSize)
    'GetRC4Data = 1 'got the keys already
    'GetRC4Data = 0 'generate new keys
    GetRC4Data = m_RC4
    m_RC4 = lpBuffer
End Function



Private Function getInteger(ByRef bArray As Long, ByVal dwLocation As Long) As Long
    Call CopyMemory(getInteger, ByVal bArray + dwLocation, 4)
End Function
Private Sub insertInteger(ByRef bArray As Long, ByVal dwLocation As Long, ByVal dwValue As Long)
    Call CopyMemory(ByVal bArray + dwLocation, dwValue, 4)
End Sub
Private Function GetSTR(ByRef bArray As Long) As String
    Dim bTest           As Byte
    Dim I               As Long
    Do
        Call CopyMemory(bTest, ByVal bArray + I, 1)
        If (bTest = 0) Then
            If (I = 0) Then Exit Function
            GetSTR = String(I, 0)
            Call CopyMemory(ByVal GetSTR, ByVal bArray, I)
            Exit Function
        End If
        I = I + 1
    Loop
End Function
Private Function Addr2Ptr(ByVal lngAddr As Long) As Long
    Addr2Ptr = lngAddr
End Function




Private Sub Data_Init(ByRef R As RANDOMDATA, ByVal lngSeed As Long)
    Dim s           As String * 4
    Call CopyMemory(ByVal s, lngSeed, 4)
    R.Sorc1 = modSHA1.Warden_SHA1(Left$(s, 2))
    'R.Sorc1 = BSHA1(Left$(s, 2), True, True)
    R.Sorc2 = modSHA1.Warden_SHA1(Right$(s, 2))
    'R.Sorc2 = BSHA1(Right$(s, 2), True, True)
    R.data = String$(20, 0)
    R.data = modSHA1.Warden_SHA1(R.Sorc1 & R.data & R.Sorc2)
    'R.Data = BSHA1(R.Sorc1 & R.Data & R.Sorc2, True, True)
    R.pos = 1
End Sub
Private Sub Data_Get_Bytes(ByRef R As RANDOMDATA, ByRef bData() As Byte, ByVal lngBytes As Long)
    Dim I           As Long
    For I = 0 To (lngBytes - 1)
        bData(I) = Asc(Mid$(R.data, R.pos, 1))
        R.pos = R.pos + 1
        If (R.pos > 20) Then
            R.pos = 1
            R.data = modSHA1.Warden_SHA1(R.Sorc1 & R.data & R.Sorc2)
            'R.Data = BSHA1(R.Sorc1 & R.Data & R.Sorc2, True, True)
        End If
    Next I
End Sub
Private Sub RC4Key(ByRef bData() As Byte, ByRef B() As Byte, ByVal lngLengh As Long)
    Dim I           As Long
    Dim A           As Long
    Dim C           As Byte
    Dim bR(255)     As Byte
    B(256) = 0
    B(257) = 0
    For I = 0 To 255
        bR(I) = bData(I Mod lngLengh)
        B(I) = I
    Next I
    A = 0
    For I = 0 To 255
        A = (A + B(I) + bR(I)) Mod 256
        C = B(I)
        B(I) = B(A)
        B(A) = C
    Next I
End Sub
Private Sub RC4CryptStr(ByRef s As String, ByRef bK() As Byte, ByVal pos As Long)
    Dim A           As Long
    Dim B           As Long
    Dim C           As Byte
    Dim I           As Long
    A = bK(256)
    B = bK(257)
    For I = pos To Len(s)
        A = (A + 1) Mod 256
        B = (B + bK(A)) Mod 256
        C = bK(A)
        bK(A) = bK(B)
        bK(B) = C
        Mid(s, I, 1) = Chr$(Asc(Mid$(s, I, 1)) Xor bK((CInt(bK(A)) + bK(B)) Mod 256))
    Next I
    bK(256) = A
    bK(257) = B
End Sub
Private Sub RC4Crypt(ByRef bData() As Byte, ByRef bK() As Byte, ByVal lngLengh As Long)
    Dim A           As Long
    Dim B           As Long
    Dim C           As Byte
    Dim I           As Long
    A = bK(256)
    B = bK(257)
    For I = 0 To (lngLengh - 1)
        A = (A + 1) Mod 256
        B = (B + bK(A)) Mod 256
        C = bK(A)
        bK(A) = bK(B)
        bK(B) = C
        bData(I) = bData(I) Xor bK((CInt(bK(A)) + bK(B)) Mod 256)
    Next I
    bK(256) = A
    bK(257) = B
End Sub
Private Function WardenChecksum(ByRef s As String) As Long
    Dim lngData(4)  As Long
    Call CopyMemory(lngData(0), ByVal modSHA1.Warden_SHA1(s), 20)
    WardenChecksum = lngData(0) Xor lngData(1) Xor lngData(2) Xor lngData(3) Xor lngData(4)
End Function
Private Function Get0x02Data(ByRef s As String, ByRef P As Long) As String
    Dim R           As String
    Dim bTest       As Boolean
    Dim A           As Long
    Dim L           As Byte
    If ((P + 6) >= Len(s)) Then Exit Function
    bTest = (Asc(Mid(s, P + 1, 1)) = 0)
    bTest = bTest And (Asc(Mid(s, P + 6, 1)) < &H40)
    If bTest Then
        Call CopyMemory(A, ByVal Mid$(s, P + 2, 4), 4)
        L = Asc(Mid$(s, P + 6, 1))
        If (A = &H4A3357) And (L = 8) Then R = RingoHexToStr("A3 80 CC 59 00 E8 3F 24")
        If (A = &H46F428) And (L = 9) Then R = RingoHexToStr("84 C8 0F 84 05 01 00 00 8B")
        If (A = &H4512E8) And (L = 5) Then R = RingoHexToStr("74 07 8A 43 46")
        If (A = &H41E237) And (L = 4) Then R = RingoHexToStr("74 38 A0 51")
        If (A = &H41E23E) And (L = 16) Then R = RingoHexToStr("0F BF 0D 54 EF 6C 00 0F BF 15 58 EF 6C 00 0C 01")
        If (A = &H41E24F) And (L = 9) Then R = RingoHexToStr("0F BF 35 56 EF 6C 00 A2 51")
        If (A = &H4BD60F) And (L = 8) Then R = RingoHexToStr("E8 CC 32 FC FF E8 47 F9")
        If (A = &H46F42A) And (L = 9) Then R = RingoHexToStr("0F 84 05 01 00 00 8B 8E DC")
        If (A = &H41E25B) And (L = 10) Then R = RingoHexToStr("0F BF 05 52 EF 6C 00 8D 74 06")
        If (A = &H450240) And (L = 6) Then R = RingoHexToStr("74 72 85 C0 74 6E")
        If Len(R) Then
            P = P + 7
            Get0x02Data = vbNullChar & R
            Exit Function
        End If
    End If
    If ((P + 29) >= Len(s)) Then Exit Function
    bTest = (Asc(Mid$(s, P + 29, 1)) < &H80)
    bTest = bTest And (Asc(Mid$(s, P + 28, 1)) = 0)
    bTest = bTest And (Asc(Mid$(s, P + 27, 1)) < &H40)
    If (bTest = False) Then Exit Function
    Call CopyMemory(A, ByVal Mid$(s, P + 26, 4), 4)
    Select Case A
        Case &H10000021
        Case &H10000050
        Case &H10000070
        Case &H100000A1
        Case &H11000020
        Case &H11000021
        Case &H16000030
        Case &H1700007C
        Case &H170001E9
        Case &H19000059
        Case &H1A0000C3
        Case &H1F000219
        Case &H1F000234
        Case &H20000022
        Case &H20000049
        Case &H23000048
        Case &H24000032
        Case &H250001EE
        Case &H250001FE
        Case &H28000091
        Case &H2A0000E1
        Case &H2A0000F1
        Case &H3000069C
        Case &H300006D4
        Case &H300006D7
        Case &H300007A8
        Case &H32000121
        Case &H3700008E
        Case &H40000081
        Case &H58000092
        Case &HC0002D0
        Case &HD0000E8
        Case &HE0001FD
        Case &HE000622
        Case Else: Exit Function
    End Select
    P = P + 30
    Get0x02Data = vbNullChar
End Function

Public Function CanHandleWarden() As Boolean
    CanHandleWarden = (Dir$(App.Path & "\zlib.dll") <> vbNullString)
End Function

Private Function RingoHexToStr(ByVal hex As String) As String
    RingoHexToStr = String(Len(hex) / 3, 0)
    Dim iPos As Long
    For I = 1 To Len(hex) Step 3
        iPos = iPos + 1
        Mid$(RingoHexToStr, iPos, 1) = Chr("&H" & Mid$(hex, I, 2))
    Next I
End Function
