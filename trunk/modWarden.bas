Attribute VB_Name = "modWarden"
Option Explicit

Public Type WARDENCONTEXT
  l_Debug As Long
  l_Module As Long
  s_MD5 As String * 16
  s_Key As String * 16
  l_Product As Long
  l_Callbacks(0 To 7) As Long
  l_SocketHandle As Long
  s_OutKey As String * &H102
  s_InKey As String * &H102
  l_ModuleLen As Long
  s_Module As String
  l_InitReturn As Long
  s_RC4Seed As String
  b_PAGE_CHECK_A As Byte
  b_MEM_CHECK As Byte
End Type

Private Enum SHA1Versions
  Sha1 = 0
  BrokenSHA1 = 1
  LockdownSHA1 = 2
  WardenSHA1 = 3
  max = &HFFFFFFFF
End Enum

Private Type SHA1Context
  IntermediateHash(0 To 4) As Long
  LengthLow As Long
  LengthHigh As Long
  MessageBlockIndex As Integer
  MessageBlock(0 To 63) As Byte
  Computed As Byte
  Corrupted As Byte
  Version As SHA1Versions
End Type

Private Type MD5Context
  IntermediateHash(0 To 3) As Long
  LengthLow As Long
  LengthHigh As Long
  MessageBlockIndex As Integer
  MessageBlock(0 To 63) As Byte
  Computed As Byte
  Corrupted As Byte
End Type

Private Type MedivRandomContext
  Index As Long
  Data(0 To 19) As Byte
  Source1(0 To 19) As Byte
  Source2(0 To 19) As Byte
End Type

Private Declare Sub rc4_init Lib "Warden.dll" (ByVal key As String, ByVal Base As String, ByVal length As Long)
Private Declare Sub rc4_crypt Lib "Warden.dll" (ByVal key As String, ByVal Data As String, ByVal length As Long)
Private Declare Sub rc4_crypt_data Lib "Warden.dll" (ByVal Data As String, ByVal DataLength As Long, ByVal Base As String, ByVal BaseLength As Long)

Private Declare Function sha1_reset Lib "Warden.dll" (ByRef Context As SHA1Context) As Long
Private Declare Function sha1_input Lib "Warden.dll" (ByRef Context As SHA1Context, ByVal Data As String, ByVal length As Long) As Long
Private Declare Function sha1_digest Lib "Warden.dll" (ByRef Context As SHA1Context, ByVal digest As String) As Long
Private Declare Function sha1_checksum Lib "Warden.dll" (ByVal Data As String, ByVal length As Long, ByVal Version As Long) As Long

Private Declare Function md5_reset Lib "Warden.dll" (ByRef Context As MD5Context) As Long
Private Declare Function md5_input Lib "Warden.dll" (ByRef Context As MD5Context, ByVal Data As String, ByVal length As Long) As Long
Private Declare Function md5_digest Lib "Warden.dll" (ByRef Context As MD5Context, ByVal digest As String) As Long
Private Declare Function md5_verify_data Lib "Warden.dll" (ByVal Data As String, ByVal length As Long, ByVal CorrectMD5 As String) As Boolean

Private Declare Sub mediv_random_init Lib "Warden.dll" (ByRef Context As MedivRandomContext, ByVal seed As String, ByVal length As Long)
Private Declare Sub mediv_random_get_bytes Lib "Warden.dll" (ByRef Context As MedivRandomContext, ByVal Buffer As String, ByVal length As Long)

Private Declare Function module_get_uncompressed_size Lib "Warden.dll" (ByVal Data As String) As Long
Private Declare Function module_get_prep_size Lib "Warden.dll" (ByVal Data As String) As Long
Private Declare Function module_decompress Lib "Warden.dll" (ByVal Source As String, ByVal SourceLength As Long, ByVal Destination As String, ByVal DestinationLength As Long) As Long

Private Declare Function module_prep Lib "Warden.dll" (ByVal Source As String, ByVal Callback As Long) As Long
Private Declare Function module_init Lib "Warden.dll" (ByVal address As Long, ByVal Callbacks As Long) As Long
Private Declare Function module_get_init_address Lib "Warden.dll" (ByVal module As Long) As Long
Private Declare Sub module_init_rc4 Lib "Warden.dll" (ByVal Callback As Long, ByVal InitData As Long, ByVal Data As String, ByVal length As Long)
Private Declare Function module_handle_packet Lib "Warden.dll" (ByVal InitData As Long, ByVal Data As String, ByVal length As Long) As Long


Private m_CallBack(7)       As Long 'callback function list, for warden
'//Warden download stuff
Private m_ModFolder         As String 'the warden folder
Public warden_context       As WARDENCONTEXT
Private Const MEDIV_MODULE_INFORMATION As Byte = &H0
Private Const MEDIV_MODULE_TRANSFER    As Byte = &H1
Private Const WARDEN_CHEAT_CHECKS      As Byte = &H2
Private Const WARDEN_NEW_CRYPT_KEYS    As Byte = &H5

Public Sub WardenInitRC4(ByRef war_ctx As WARDENCONTEXT)
    Dim ctx             As MedivRandomContext
    Dim out_seed        As String
    Dim in_seed         As String
    
    If (Not CanHandleWarden()) Then
        Call frmChat.AddChat(vbRed, "[Warden] Warden support has not been initialized because zlib1.dll or Warden.dll could not be found.")
        Exit Sub
    End If
    
    Call WardenCleanUp(war_ctx)
    '//Create new RC4 Keys
  
    out_seed = String(16, vbNull)
    in_seed = String(16, vbNull)
    war_ctx.s_OutKey = String(&H102, vbNull)
    war_ctx.s_InKey = String(&H102, vbNull)
  
    On Error GoTo DLL_ERROR:
    
    Call mediv_random_init(ctx, war_ctx.s_RC4Seed, Len(war_ctx.s_RC4Seed))
    Call mediv_random_get_bytes(ctx, out_seed, 16)
    Call mediv_random_get_bytes(ctx, in_seed, 16)
    If (MDebug("warden")) Then
      Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Random Seed: " & StrToHex(war_ctx.s_RC4Seed, True))
      Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Out Seed: " & StrToHex(out_seed))
      Call frmChat.AddChat(RTBColors.InformationText, "[Warden] In Seed:  " & StrToHex(in_seed))
    End If
    Call rc4_init(war_ctx.s_OutKey, out_seed, 16)
    Call rc4_init(war_ctx.s_InKey, in_seed, 16)

    
On Error GoTo handler 'This is eww I know, but if the folder is empty it'll try and make and kill us all.
    m_ModFolder = App.Path & "\Warden\"
    If (Dir$(m_ModFolder) = vbNullString) Then
        MkDir m_ModFolder
    End If
handler:
    Call frmChat.AddChat(vbGreen, "[Warden] Initialized!")
    Exit Sub
    
DLL_ERROR:
    Call frmChat.AddChat(vbRed, "[Warden] Error #" & Err.Number & ": " & Err.description)
End Sub
Public Sub WardenCleanUp(Context As WARDENCONTEXT)
  If (Len(Context.s_Module) > 0) Then Context.s_Module = vbNullString
  Context.b_MEM_CHECK = 0
  Context.b_PAGE_CHECK_A = 0
  Context.s_InKey = String$(&H102, Chr$(0))
  Context.s_OutKey = String$(&H102, Chr$(0))
End Sub
Public Function WardenClientData(ByRef Context As WARDENCONTEXT, sData As String) As Boolean
  Dim ID As Integer
  
  ID = Asc(Mid(sData, 2, 1))
  WardenClientData = False
    
  If (ID = &H50) Then
    WardenCleanup Context
    Context.l_Product = GetDWORD(Mid$(sData, 13, 4))
  ElseIf (ID = &H51) Then
    Context.s_RC4Seed = Mid$(sData, 41, 4)
  End If
End Function

Public Function WardenServerData(ByRef Context As WARDENCONTEXT, sData As String) As Boolean
  Dim ID As Integer
  
  ID = Asc(Mid(sData, 2, 1))
  WardenServerData = False
  If (ID = &H5E) Then
      'frmChat.AddChat vbGreen, "Received Warden packet from Server"
      
      Dim sPacket As String
      Dim opcode As Integer
      
      If (Context.s_InKey = String(&H102, Chr$(0))) Then
        Call WardenInitRC4(Context)
      End If
      
      sPacket = Mid$(sData, 5)
      Call rc4_crypt(Context.s_InKey, sPacket, Len(sPacket))
      opcode = Asc(Left$(sPacket, 1))
      
      'frmChat.AddChat vbWhite, "Packet Data: " & vbNewLine & DebugOutput(sPacket)
      
      Select Case opcode
        Case MEDIV_MODULE_INFORMATION: Call WardenModuleInfo(Context, Mid$(sPacket, 2))
        Case MEDIV_MODULE_TRANSFER:    Call WardenModuleTransfer(Context, Mid$(sPacket, 2))
        Case WARDEN_CHEAT_CHECKS:      Call WardenCheatChecks(Context, Mid$(sPacket, 2))
        Case WARDEN_NEW_CRYPT_KEYS:    Call WardenHandleGeneric(Context, Mid$(sData, 5), opcode)
        Case Else:                     Call WardenUnknown(Context, sPacket, opcode)
      End Select
      WardenServerData = True
  End If
End Function

Private Function Addr2Ptr(ByVal lngAddr As Long) As Long
    Addr2Ptr = lngAddr
End Function

Public Function CanHandleWarden() As Boolean
    CanHandleWarden = (Dir$(App.Path & "\zlib1.dll") <> vbNullString) And _
                      (Dir$(App.Path & "\Warden.dll") <> vbNullString)
End Function
Private Function GetDWORD(ByVal Value As String) As Long
    Dim Result As Long
    CopyMemory Result, ByVal Value, 4
    GetDWORD = Result
End Function
Private Function CreateDWORD(ByVal Value As Long) As String
    Dim Result As String * 4
    CopyMemory ByVal Result, Value, 4
    CreateDWORD = Result
End Function


Private Sub WardenModuleInfo(ByRef Context As WARDENCONTEXT, sData As String)
  'MEDIV_MODULE_INFORMATION
  '(BYTE[16]) MD5 Checksum
  '(BYTE[16]) Module RC4 Seed
  '(DWORD)    Module Compressed Length
  
  Dim pBuf As New clsDataBuffer
  pBuf.Data = sData
  
  Context.s_MD5 = pBuf.GetRaw(16)
  Context.s_Key = pBuf.GetRaw(16)
  Context.l_ModuleLen = pBuf.GetDWORD
  
  If MDebug("warden") Then
    Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Received Warden 0x00")
    Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Name:     ", vbWhite, StrToHex(Context.s_MD5, True))
    Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Key Seed: ", vbWhite, StrToHex(Context.s_Key, True))
    Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Length:   ", vbWhite, Context.l_ModuleLen)
  End If
    
  If (WardenLoadModule(Context, True)) Then
    If (WardenInitModule(Context)) Then
      Call WardenSendData(Context, Chr$(1))
    Else
      If (MDebug("warden")) Then
        Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Corrupted module, Requesting download")
      End If
      Call WardenSendData(Context, Chr$(1))
    End If
  Else
    If (MDebug("warden")) Then
      Call frmChat.AddChat(RTBColors.InformationText, "[Warden] New Module, Requesting download")
    End If
    Call WardenSendData(Context, Chr$(0))
  End If
End Sub

Private Sub WardenModuleTransfer(ByRef Context As WARDENCONTEXT, sData As String)
  'MEDIV_MODULE_TRANSFER
  '(DWORD) Payload Length
  '(Void)  Payload

  Dim pBuf As New clsDataBuffer
  Dim data_length As Long
  
  pBuf.Data = sData
  data_length = pBuf.GetWord
  Context.s_Module = Context.s_Module & pBuf.GetRaw(data_length)
      
  If (Len(Context.s_Module) = Context.l_ModuleLen) Then
    If (MDebug("warden")) Then
      Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Module Download complete")
    End If
    
    If (WardenLoadModule(Context, False)) Then
      Call WriteFile(m_ModFolder & StrToHex(Context.s_MD5, True) & ".bin", Context.s_Module)
      If (WardenInitModule(Context)) Then
        Call WardenSendData(Context, Chr$(1))
      Else
        Call WardenSendData(Context, Chr$(0))
      End If
    Else
      Call WardenSendData(Context, Chr$(0))
    End If
  End If
End Sub

Private Sub WardenCheatChecks(ByRef Context As WARDENCONTEXT, sData As String)
  'WARDEN_CHEAT_CHECKS
  '(PString[]) Libraries
  '[void] Commands
  '  PAGE_CHECK_A
  '    (DWORD)    SHA1 Seed
  '    (Byte[20]) SHA1
  '    (DWORD)    Address
  '    (Byte)     Length
  '  MEM_CHECK
  '    (Byte)  Lib ID
  '    (DWORD) Address
  '    (Byte)  Length
  
  Dim sNames As String
  Dim key As Byte
  Dim lib_length As Long
  Dim X As Integer
  Dim Offset As Long
  Dim mem As String
  Dim opcode As Byte
  Dim lib_id As Byte
  Dim lib_count As Integer
  Dim sTemp As String
  
  Dim tBuffer As New clsDataBuffer
  Dim pBuffer As New clsDataBuffer
  pBuffer.Data = sData
  
  key = Asc(Right(sData, 1))
  If (MDebug("warden")) Then
    Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Received Cheat Check")
    Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Key: 0x" & ZeroOffset(key, 2))
  End If
  
  lib_length = pBuffer.GetByte
  X = 0
  
  Do While (lib_length > 0)
    X = X + 1
    sTemp = pBuffer.GetRaw(lib_length)
    sNames = sNames & Chr$(0) & sTemp
    If (MDebug("warden")) Then
      Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Library: (" & X & ") " & sTemp)
    End If
    lib_length = pBuffer.GetByte
  Loop
  lib_count = X
  sTemp = vbNullString
 
  Do While (pBuffer.length() >= pBuffer.Position() + 2)
    opcode = pBuffer.GetByte Xor key
    
    If (Context.b_MEM_CHECK = 0 Or Context.b_PAGE_CHECK_A = 0) Then
      tBuffer.Clear
      tBuffer.Data = pBuffer.GetRaw(6, True)
      lib_id = tBuffer.GetByte
      Offset = tBuffer.GetDWORD
      X = tBuffer.GetByte
      
      If (lib_id > lib_count) Then
        Context.b_PAGE_CHECK_A = opcode
      Else
        'Not perfect but works for all know versions of Warden and there data
        
        If ((Offset < &H40000000 And lib_id = 0) And (Offset > &H60000000 And lib_id = 0)) _
           Or (Offset > &H10000000 And lib_id > 0) Then
          Context.b_PAGE_CHECK_A = opcode
        Else
          Context.b_MEM_CHECK = opcode
        End If
      End If
    End If
    
    If (opcode = Context.b_PAGE_CHECK_A) Then
      If (MDebug("warden")) Then
        Dim page_seed As Long
        Dim page_sha1 As String
        Dim page_address As Long
        page_seed = pBuffer.GetDWORD
        page_sha1 = pBuffer.GetRaw(20)
        page_address = pBuffer.GetDWORD
        X = pBuffer.GetByte
        Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Opcode: 0x" & ZeroOffset(opcode, 2) & _
        " Page: " & Right$("   " & X, 3) & " @ 0x" & ZeroOffset(page_address, 8) & _
        " Seed: 0x" & ZeroOffset(page_seed, 8) & " Hash: " & StrToHex(page_sha1, True))
      Else
        Call pBuffer.GetRaw(29) 'remove from the buffer
      End If
      sTemp = sTemp & Chr$(0)
    ElseIf (opcode = Context.b_MEM_CHECK) Then
      lib_id = pBuffer.GetByte
      Offset = pBuffer.GetDWORD
      X = pBuffer.GetByte
      mem = WardenGetMemorySegment(Context, Offset, CByte(X), opcode, lib_id, sNames)
      If (mem = vbNullString) Then
        'sTemp = sTemp & Chr$(1)
        Exit Sub 'No point in sending it without data
      Else
        sTemp = sTemp & Chr$(0) & mem
      End If
    Else
      Call frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Unknown Mem Check Opcode, You will be disconnected soon.")
      Exit Sub
    End If
  Loop
  
  tBuffer.Clear
  tBuffer.InsertByte WARDEN_CHEAT_CHECKS
  tBuffer.InsertWord Len(sTemp)
  tBuffer.InsertDWord sha1_checksum(sTemp, Len(sTemp), 0)
  tBuffer.InsertNonNTString sTemp
  
  sTemp = tBuffer.Data
  
  Call WardenSendData(Context, sTemp)
End Sub


Private Sub WardenHandleGeneric(ByRef Context As WARDENCONTEXT, sData As String, ID As Integer)
  Dim length As Long
  Dim tmp As String
  tmp = sData
    
  Call CopyMemory(ByVal Context.l_InitReturn + &H20, ByVal Context.s_OutKey, &H102)
  length = module_handle_packet(ByVal Context.l_InitReturn, tmp, Len(sData))
      
  If (length = Len(sData)) Then
    If (MDebug("warden")) Then Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Handled packet successfully")
    Call CopyMemory(ByVal Context.s_OutKey, ByVal Context.l_InitReturn + &H20, &H102)
    Call CopyMemory(ByVal Context.s_InKey, ByVal Context.l_InitReturn + &H122, &H102)
  Else
    Call frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Failed to handle packet, you will be disconnected soon!")
  End If

End Sub

Private Sub WardenUnknown(ByRef Context As WARDENCONTEXT, sData As String, ID As Integer)
  Call frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Unknown Opcode, You may be disconnected soon.")
  If MDebug("warden") Then
    Call frmChat.AddChat(RTBColors.ErrorMessageText, "Unhandled Warden Opcode 0x" & ZeroOffset(ID, 2))
    Call frmChat.AddChat(RTBColors.ErrorMessageText, "Packet data: " & vbCrLf & DebugOutput(sData))
  End If
End Sub

Private Sub WardenSendData(ByRef Context As WARDENCONTEXT, sData As String)
  Dim pBuffer As New clsDataBuffer
  Call rc4_crypt(Context.s_OutKey, sData, Len(sData))
  pBuffer.InsertNonNTString sData
  pBuffer.SendPacket &H5E
End Sub

Public Function WardenLoadModule(ByRef Context As WARDENCONTEXT, Optional FromFile As Boolean = False) As Boolean
  Dim X As Long
  Dim temp As String
  Dim Data As String
  
  If (FromFile) Then
    If (Dir$(m_ModFolder & StrToHex(Context.s_MD5, True) & ".bin") <> vbNullString) Then
      Open m_ModFolder & StrToHex(Context.s_MD5, True) & ".bin" For Binary Access Read As #1
        Data = String$(LOF(1), Chr$(0))
        Get 1, 1, Data
      Close #1
    Else
      WardenLoadModule = False
      Exit Function
    End If
  Else
    Data = Context.s_Module
  End If
  
  If md5_verify_data(Data, Len(Data), Context.s_MD5) Then
    If (MDebug("warden")) Then Call frmChat.AddChat(RTBColors.InformationText, "[Warden] MD5 Passed")
    Dim Base As String
    rc4_crypt_data Data, Len(Data), Context.s_Key, Len(Context.s_Key)
    
    If (Mid(Data, Len(Data) - &H103, 4) = "NGIS") Then
      If (MDebug("warden")) Then Call frmChat.AddChat(RTBColors.InformationText, "[Warden] RC4 Passed")
      X = module_get_uncompressed_size(Data)
      temp = String$(X, Chr$(0))
      If (Not module_decompress(Data, Len(Data), temp, X)) Then
        If (MDebug("warden")) Then Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Decompressions Successful")
        X = module_get_prep_size(temp)
        Context.l_Module = module_prep(temp, AddressOf WardenDebugCallback)
        WardenLoadModule = (Context.l_Module > 0)
        
        Exit Function
      Else
        If (MDebug("warden")) Then Call frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Decompression Failed")
      End If
    Else
      If (MDebug("warden")) Then Call frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] RC4 Failed")
    End If
  Else
    If (MDebug("warden")) Then Call frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] MD5 Failed")
  End If
  WardenLoadModule = False
End Function

Public Function WardenInitModule(ByRef Context As WARDENCONTEXT) As Boolean
  Dim address As Long
  address = module_get_init_address(Context.l_Module) + Context.l_Module
  
  If (MDebug("warden")) Then Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Attempting to call init at 0x" & ZeroOffset(address, 8))
  
  Context.l_Callbacks(0) = Addr2Ptr(AddressOf WardenSendPacket)
  Context.l_Callbacks(1) = Addr2Ptr(AddressOf WardenCheckModule)
  Context.l_Callbacks(2) = Addr2Ptr(AddressOf WardenModuleLoad)
  Context.l_Callbacks(3) = Addr2Ptr(AddressOf WardenAllocateMem)
  Context.l_Callbacks(4) = Addr2Ptr(AddressOf WardenFreeMemory)
  Context.l_Callbacks(5) = Addr2Ptr(AddressOf WardenSetRC4Data)
  Context.l_Callbacks(6) = Addr2Ptr(AddressOf WardenGetRC4Data)
  Context.l_Callbacks(7) = VarPtr(Context.l_Callbacks(0))
  
  Context.l_InitReturn = module_init(address, VarPtr(Context.l_Callbacks(7)))
  
  If (Context.l_InitReturn <> 0) Then
    If (MDebug("warden")) Then Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Init() = 0x" & ZeroOffset(Context.l_InitReturn, 8))
    Call module_init_rc4(AddressOf WardenDebugCallback, ByVal Context.l_InitReturn, ByVal Context.s_RC4Seed, Len(Context.s_RC4Seed))
    
    Call CopyMemory(ByVal Context.l_InitReturn + &H20, ByVal Context.s_OutKey, &H102)
    Call CopyMemory(ByVal Context.l_InitReturn + &H122, ByVal Context.s_InKey, &H102)
  Else
    Call frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Init failed, You will be disconnected soon!")
  End If
  WardenInitModule = (Context.l_InitReturn <> 0)
End Function



Public Sub WardenDebugCallback(ByVal color As Long, ByVal addr As Long, ByVal length As Long)
  If (Not MDebug("warden")) Then Exit Sub
  Dim msg As String
  msg = Space(length)
  CopyMemory ByVal msg, ByVal addr, length
  
  Call frmChat.AddChat(RTBColors.InformationText, "[Warden] ", color, msg)
End Sub
Public Sub WriteFile(sPath As String, sData As String)
  Dim I As Integer
  I = FreeFile
  Open sPath For Binary Access Write As #I
    Put #I, 1, sData
  Close #1
End Sub



Private Sub WardenSendPacket(ByVal ptrPacket As Long, ByVal dwSize As Long)
  Dim m_PKT As String
  Dim pBuffer As New clsDataBuffer
  
  m_PKT = Space(dwSize)
  Call CopyMemory(ByVal m_PKT, ByVal ptrPacket, dwSize)
  If (MDebug("warden")) Then
    Call frmChat.AddChat(RTBColors.InformationText, "[Warden] SendPacket(0x" & ZeroOffset(ptrPacket, 8) & _
         ", " & dwSize & ")" & vbNewLine & DebugOutput(m_PKT))
  End If
  pBuffer.Data = m_PKT
  pBuffer.SendPacket &H5E
End Sub
Private Function WardenCheckModule(ByVal ptrMod As Long, ByVal ptrKey As Long) As Long
  If (MDebug("warden")) Then
    Call frmChat.AddChat(RTBColors.InformationText, "[Warden] CheckModule(0x" & ZeroOffset(ptrMod, 8) & _
         ", 0x" & ZeroOffset(ptrKey, 8) & ")")
  End If
  'CheckModule = 0 '//Need to download
  'CheckModule = 1 '//Don't need to download
  WardenCheckModule = 1
End Function
Private Function WardenModuleLoad(ByVal ptrRC4Key As Long, ByVal pModule As Long, ByVal dwModSize As Long) As Long
  If (MDebug("warden")) Then
    Call frmChat.AddChat(RTBColors.InformationText, "[Warden] ModuleLoad(0x" & ZeroOffset(ptrRC4Key, 8) & _
         ", 0x" & ZeroOffset(pModule, 8) & ", 0x" & ZeroOffset(dwModSize, 8) & ")")
  End If
  'ModuleLoad = 0 '//Need to download
  'ModuleLoad = 1 '//Don't need to download
  WardenModuleLoad = 1
End Function
Private Function WardenAllocateMem(ByVal dwSize As Long) As Long
  WardenAllocateMem = malloc(dwSize)
  If (MDebug("warden")) Then
    Call frmChat.AddChat(RTBColors.InformationText, "[Warden] AllocateMem(" & dwSize & ") = 0x" & ZeroOffset(WardenAllocateMem, 8))
  End If
End Function
Private Sub WardenFreeMemory(ByVal dwMemory As Long)
  Call free(dwMemory)
    
  If (MDebug("warden")) Then
    Call frmChat.AddChat(RTBColors.InformationText, "[Warden] FreeMem(0x" & ZeroOffset(dwMemory, 8) & ")")
  End If
End Sub
Private Function WardenSetRC4Data(ByVal lpKeys As Long, ByVal dwSize As Long) As Long
  If (MDebug("warden")) Then
    Call frmChat.AddChat(RTBColors.InformationText, "[Warden] SetRC4Data(0x" & ZeroOffset(lpKeys, 8) & _
                         ", 0x" & ZeroOffset(dwSize, 8) & ")")
  End If
End Function
Private Function WardenGetRC4Data(ByVal lpBuffer As Long, ByRef dwSize As Long) As Long
  If (MDebug("warden")) Then
    Call frmChat.AddChat(RTBColors.InformationText, "[Warden] GetRC4Data(0x" & ZeroOffset(lpBuffer, 8) & _
                         ", 0x" & ZeroOffset(dwSize, 8) & ")")
  End If
  'GetRC4Data = 1 'got the keys already
  'GetRC4Data = 0 'generate new keys
  WardenGetRC4Data = 1
End Function

Private Function WardenGetMemorySegment(ByRef Context As WARDENCONTEXT, address As Long, length As Long, opcode As Byte, lib_id As Byte, names As String) As String
  Dim Data As String
  Dim game_client As String
  Dim lib_name As String
  If (lib_id > 0) Then lib_name = Split(names & String$(lib_id, Chr$(0)), Chr$(0))(lib_id)
  
  game_client = CreateDWORD(Context.l_Product)
  If (game_client = "PX3W") Then
    game_client = "3RAW"
  ElseIf (game_client = "PXES") Then
    game_client = "RATS"
  End If
  
  If (lib_id > 0) Then
    Data = ReadINI(game_client & "_" & lib_name, ZeroOffset(length, 2) & "_" & ZeroOffset(address, 8), "./Warden.ini")
  Else
    Data = ReadINI(game_client & "_Mem_Check", ZeroOffset(length, 2) & "_" & ZeroOffset(address, 8), "./Warden.ini")
  End If
  
  WardenGetMemorySegment = vbNullString
  
  If Data = vbNullString Then
    Call frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Could not find memory segment " & length & _
    " at 0x" & ZeroOffset(address, 8) & " for " & IIf(Len(lib_name) > 0, lib_name, "MEM_CHECK") & ". You will be disconnected soon.")
    Call frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Make sure you have the latest Warden.ini from http://www.stealthbot.net/board/index.php?showtopic=41491")
  Else
    WardenGetMemorySegment = HexToStr(Data)
    If (MDebug("warden")) Then
      Call frmChat.AddChat(RTBColors.InformationText, "[Warden] Opcode: 0x" & ZeroOffset(opcode, 2) & _
      " Read: " & Right$("   " & length, 3) & " @ (" & lib_id & ") 0x" & ZeroOffset(address, 8) & _
      " Data: " & Data)
    End If
  End If
End Function


