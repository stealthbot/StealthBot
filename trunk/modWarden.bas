Attribute VB_Name = "modWarden"
Option Explicit

Private Enum SHA1Versions
  Sha1 = 0
  BrokenSHA1 = 1
  LockdownSHA1 = 2
  WardenSHA1 = 3
  Max = &HFFFFFFFF
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

Private Declare Sub rc4_init Lib "Warden.dll" (ByVal Key As String, ByVal Base As String, ByVal length As Long)
Private Declare Sub rc4_crypt Lib "Warden.dll" (ByVal Key As String, ByVal Data As String, ByVal length As Long)
Private Declare Sub rc4_crypt_data Lib "Warden.dll" (ByVal Data As String, ByVal DataLength As Long, ByVal Base As String, ByVal BaseLength As Long)

Private Declare Function sha1_reset Lib "Warden.dll" (ByRef Context As SHA1Context) As Long
Private Declare Function sha1_input Lib "Warden.dll" (ByRef Context As SHA1Context, ByVal Data As String, ByVal length As Long) As Long
Private Declare Function sha1_digest Lib "Warden.dll" (ByRef Context As SHA1Context, ByVal digest As String) As Long
Private Declare Function sha1_checksum Lib "Warden.dll" (ByVal Data As String, ByVal length As Long, ByVal Version As Long) As Long

Private Declare Function md5_reset Lib "Warden.dll" (ByRef Context As MD5Context) As Long
Private Declare Function md5_input Lib "Warden.dll" (ByRef Context As MD5Context, ByVal Data As String, ByVal length As Long) As Long
Private Declare Function md5_digest Lib "Warden.dll" (ByRef Context As MD5Context, ByVal digest As String) As Long
Private Declare Function md5_verify_data Lib "Warden.dll" (ByVal Data As String, ByVal length As Long, ByVal CorrectMD5 As String) As Boolean

Private Declare Sub mediv_random_init Lib "Warden.dll" (ByRef Context As MedivRandomContext, ByVal Seed As String, ByVal length As Long)
Private Declare Sub mediv_random_get_bytes Lib "Warden.dll" (ByRef Context As MedivRandomContext, ByVal Buffer As String, ByVal length As Long)

Private Declare Function warden_init Lib "Warden.dll" (ByVal SocketHandle As Long) As Long
Private Declare Function warden_data Lib "Warden.dll" (ByVal Instance As Long, ByVal Direction As Long, ByVal PacketID As Long, ByVal Data As String, ByVal length As Long) As Long
Private Declare Function warden_cleanup Lib "Warden.dll" (ByVal Instance As Long) As Long
Private Declare Function warden_set_data_file Lib "Warden.dll" (ByVal Instance As Long, ByVal File As String, ByVal length As Long) As Long
Private Declare Function warden_config Lib "Warden.dll" (ByVal Instance As Long, ByVal ConfigBit As Long, ByVal Enabled As Byte) As Long

Public Const WARDEN_CONFIG_SAVE_CHECKS    As Long = 1  '//Save Information about cheat checks (Opcode 0x02) to Data File
Public Const WARDEN_CONFIG_SAVE_UNKNOWN   As Long = 2  '//Save Unknown information (use in conjunction with Debug mode to get new Warden offsets)
Public Const WARDEN_CONFIG_LOG_CHECKS     As Long = 4  '//Log ALL information about checks that happen, in real time
Public Const WARDEN_CONFIG_LOG_PACKETS    As Long = 8  '//Log ALL decoded Warden packet data
Public Const WARDEN_CONFIG_DEBUG_MODE     As Long = 16 '//Debug mode, does a lot of shit u.u
Public Const WARDEN_CONFIG_USE_GAME_FILES As Long = 32 '//Will attempt to grab unknown Mem Check offsets from the game file specified
                                                       '//  Will try to load library the file, using the path specified in the INI EXA:
                                                       '//[Files_WAR3]
                                                       '//Default=C:\Program Files\Warcraft III\WAR3.exe
                                                       '//Game.dll=C:\Program Files\Warcraft III\Game.dll
Private Const WARDEN_SEND              As Long = &H0
Private Const WARDEN_RECV              As Long = &H1
Private Const WARDEN_BNCS              As Long = &H2

Private Const WARDEN_IGNORE                  As Long = &H0  '//Not a warden packet, Handle internally
Private Const WARDEN_SUCCESS                 As Long = &H1  '//All Went Well, Don't handle the packet Internally
Private Const WARDEN_UNKNOWN_PROTOCOL        As Long = &H2  '//Not used, will be when adding support for MCP/UDP
Private Const WARDEN_UNKNOWN_SUBID           As Long = &H3  '//Unknown Sub-ID [Not 0x00, 0x01, 0x02, or 0x05]
Private Const WARDEN_RAW_FAILURE             As Long = &H4  '//The module was not able to handle the packet itself
Private Const WARDEN_PACKET_FAILURE          As Long = &H5  '//Something went HORRIBLY wrong in warden_packet, should NEVER happen.
Private Const WARDEN_INIT_FAILURE            As Long = &H6  '//Calling Init() in the module failed
Private Const WARDEN_LOAD_FILE_FAILURE       As Long = &H7  '//Could not load module from file [Not to bad, prolly just dosen't exist]
Private Const WARDEN_LOAD_MD5_FAILURE        As Long = &H8  '//Failed MD5 checksum when loading module [Either Bad tranfer or HD file corrupt]
Private Const WARDEN_LOAD_INVALID_SIGNATURE  As Long = &H9  '//Module failed RSA verification
Private Const WARDEN_LOAD_DECOMPRESS_FAILURE As Long = &HA  '//Module failed to decompress properly
Private Const WARDEN_LOAD_PREP_FAILURE       As Long = &HB  '//Module prepare failed, Usually if module is corrupt
Private Const WARDEN_CHECK_UNKNOWN_COMMAND   As Long = &HC  '//Unknown sub-command in CHEAT_CHECKS
Private Const WARDEN_CHECK_TO_MANY_LIBS      As Long = &HD  '//There were more then 4 libraries in a single 0x02 packet [this is eww yes, but I'll figure out a beter way later]
Private Const WARDEN_MEM_UNKNOWN_PRODUCT     As Long = &HE  '//The product from 0x50 != WC3, SC, or D2
Private Const WARDEN_MEM_UNKNOWN_SEGMENT     As Long = &HF  '//Could not read segment from ini file
Private Const WARDEN_INVALID_INSTANCE        As Long = &H10 '//Instance passed to this function was invalid

Public WardenInstance As Long
'====================================================================================================
'^^^^Warden Stuff
'vvvvCrev Stuff
'====================================================================================================
Private Declare Function check_revision Lib "Warden.dll" (ByVal ArchiveTime As String, ByVal ArchiveName As String, ByVal Seed As String, ByVal INIFile As String, ByVal INIHeader As String, ByRef Version As Long, ByRef Checksum As Long, ByVal Result As String) As Long
Private Declare Function crev_max_result Lib "Warden.dll" () As Long
Private Declare Function crev_error_description Lib "Warden.dll" (ByVal ErrorCode As Long, ByVal description As String, ByVal Size As Long) As Long

Private Const CREV_SUCCESS             As Long = 0  '//If everything went ok
Private Const CREV_UNKNOWN_VERSION     As Long = 1  '//Unknown version, Not lockdown, Or Ver
Private Const CREV_UNKNOWN_REVISION    As Long = 2  '//Unknown Revision (0-7 for old, 0-19 for lockdown)
Private Const CREV_MALFORMED_SEED      As Long = 3  '//If the Seed passed in wasn't able to be translated properly
Private Const CREV_MISSING_FILENAME    As Long = 4  '//We were not able to get the file path information from the INI file, Result holds more info.
Private Const CREV_MISSING_FILE        As Long = 5  '//Was not able to open a file, Result has the File Path
Private Const CREV_FILE_INFO_ERROR     As Long = 6  '//And error while trying to get the file info string, Result holds the path of the file
Private Const CREV_TOFEW_RVAS          As Long = 7  '//Less then 14 RVAs in hash file, probably corrupt
Private Const CREV_UNKNOWN_RELOC_TYPE  As Long = 8  '//Unknown Reloc Type, only 16, 32, and 64 bit are known
Private Const CREV_OUT_OF_MEMORY       As Long = 9  '//Out of memory
Private Const CREV_CORRUPT_IMPORT_DATA As Long = 10 '//Corrupt IAT data, File may be Corrupt
'====================================================================================================
Public Function Warden_CheckRevision(sArchiveName As String, sArchiveFileTime As String, sSeed As String, sHeader As String, ByRef lVersion As Long, ByRef lChecksum As Long, ByRef sResult As String) As Boolean
On Error GoTo trap:
    Dim lRet       As Long
    Dim ltVersion  As Long
    Dim ltChecksum As Long
    Dim stResult   As String
    Dim sError     As String
    Dim i          As Long
    Dim ft         As FILETIME
    Dim st         As SYSTEMTIME
    Dim sFileTime  As String
    
    Warden_CheckRevision = False
    stResult = String$(crev_max_result, Chr$(0))
    
    lRet = check_revision(sArchiveFileTime, sArchiveName, sSeed, _
        GetFilePath("CheckRevision.ini", StringFormat("{0}\", App.Path)), sHeader, _
        ltVersion, ltChecksum, stResult)
    
    i = InStr(1, stResult, Chr$(0))
    If (i > 0) Then stResult = Left$(stResult, i - 1)
    
    
    If (Not Len(sArchiveFileTime) = 8) Then sArchiveFileTime = Left$(StringFormat("{0}{1}", sArchiveFileTime, String$(8, Chr$(0))), 8)
    CopyMemory ft, ByVal sArchiveFileTime, 8
    
    FileTimeToSystemTime ft, st
    sFileTime = StringFormat("{0}/{1}/{2} {3}:{4}:{5}", st.wMonth, st.wDay, st.wYear, st.wHour, st.wMinute, st.wSecond)
    
    Select Case lRet
        Case CREV_SUCCESS:
            Warden_CheckRevision = True
            sResult = stResult
            lChecksum = ltChecksum
            lVersion = ltVersion
            Exit Function
            
        Case CREV_UNKNOWN_VERSION, CREV_UNKNOWN_REVISION:
            frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("[BNCS] Warden.dll does not support checkrevision for {0} {1}", sArchiveName, sFileTime)
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Make sure you have the latest Warden.dll from http://www.stealthbot.net/sb/warden/"
        
        Case CREV_MALFORMED_SEED:
            frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("[BNCS] The seed value string was malformed: {0}", IIf(InStr(1, sSeed, "A=", vbTextCompare) > 0, sSeed, StrToHex(sSeed, True)))
            frmChat.AddChat RTBColors.ErrorMessageText, "[BNCS] Make sure you have the latest Warden.dll from http://www.stealthbot.net/sb/warden/"
            
        Case CREV_MISSING_FILENAME:
            frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("[BNCS] Could not read key {0}{1}{0} under [{2}] from CheckRevision.ini, Update your configuration", Chr$(34), stResult, sHeader)
            
        Case CREV_MISSING_FILE:
            frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("[BNCS] Could not open file {0}{1}{0}", Chr$(34), stResult)
            
        Case CREV_FILE_INFO_ERROR:
            frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("[BNCS] Error retrieving information from {0}{1}{0}", Chr$(34), stResult)
        
        Case Else:
            sResult = String$(crev_max_result, Chr$(0))
            i = crev_error_description(lRet, sResult, Len(sResult))
            If (i > 0) Then
                sResult = String$(i, Chr$(0))
                i = crev_error_description(lRet, sResult, Len(sResult))
            End If
            
            i = InStr(1, sResult, Chr$(0))
            If (i > 0) Then stResult = Left$(stResult, i - 1)
            frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("[BNCS] Unknown CheckRevision error: {0}", stResult)
        
    End Select
  
trap:
  If (Err.Number = 53) Then
    frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] Warden.dll was not found, Local Hashing will not work."
    frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] To fix this, make sure you have the latest Warden.dll from http://www.stealthbot.net/sb/warden/"
    Err.Clear
  End If
  Warden_CheckRevision = False
End Function

Public Sub WardenCleanup(Instance As Long)
  On Error GoTo trap
  If (Not Instance = 0) Then Call warden_cleanup(Instance)
  Exit Sub
  
trap:
  If (Err.Number = 53) Then
    frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] Warden.dll was not found, Warden support will not work."
    frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] To fix this, make sure you have the latest Warden data from http://www.stealthbot.net/sb/warden/"
    Err.Clear
  End If
End Sub

Public Function WardenInitilize(ByVal SocketHandle As Long) As Long
  On Error GoTo trap
  Dim INIPath As String
  Dim Instance As Long
  Dim DebugString As String
  
  Instance = warden_init(SocketHandle)
  
  If (Instance > 0) Then
  
    INIPath = GetFilePath("Warden.ini")
  
    warden_set_data_file Instance, INIPath, Len(INIPath)
    
    DebugString = ReadCfg("Override", "WardenDebug")
    If StrictIsNumeric(DebugString) Then Call warden_config(Instance, CLng(DebugString), 2)
    
    WardenInitilize = Instance
  End If
  Exit Function
  
trap:
  If (Err.Number = 53) Then
    frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] Warden.dll was not found, Warden support will not work."
    frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] To fix this, make sure you have the latest Warden data from http://www.stealthbot.net/sb/warden/ ."
    Err.Clear
  End If
End Function

Public Function WardenData(Instance As Long, sData As String, Send As Boolean) As Boolean
  Dim ID As Long
  Dim Result As Long
  Dim Data As String

  ID = Asc(Mid(sData, 2, 1))
  Data = Mid$(sData, 5)
  
  If (Instance = 0) Then
    If (MDebug("warden")) Then
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] Attempted to call Data() with an invalid instance."
    End If
    WardenData = False
    Exit Function
  End If
  
  Result = warden_data(Instance, WARDEN_BNCS Or IIf(Send, WARDEN_SEND, WARDEN_RECV), ID, Data, Len(Data))
  
  Select Case Result
    Case WARDEN_SUCCESS: '//All Went Well, Don't handle the packet Internally
        If (MDebug("warden")) Then
            Select Case Asc(Left$(Data, 1))
                Case 0:    frmChat.AddChat RTBColors.InformationText, "[Warden] Handled Module Information"
                Case 1:    frmChat.AddChat RTBColors.InformationText, "[Warden] Handled Module Transfer"
                Case 2:    frmChat.AddChat RTBColors.InformationText, "[Warden] Handled Cheat Check"
                Case 5:    frmChat.AddChat RTBColors.InformationText, "[Warden] Handled New Crypt Keys"
                Case Else: frmChat.AddChat RTBColors.InformationText, "[Warden] Handled Unknown 0x" & ZeroOffset(Asc(Left(Data, 1)), 2)
            End Select
        End If
    'case WARDEN_UNKNOWN_PROTOCOL '//Not used, will be when adding support for MCP/UDP
    Case WARDEN_UNKNOWN_SUBID: '//Unknown Sub-ID [Not 0x00, 0x01, 0x02, or 0x05]
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] Unknown sub-command 0x" & ZeroOffset(Asc(Left$(Data, 1)), 2) & ", you will be disconnected soon."
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] For information on this, please visit http://www.stealthbot.net/sb/warden-issues/?unknown&id=0x" & ZeroOffset(Asc(Left$(Data, 1)), 2) & " ."
        
        If (MDebug("warden")) Then
            frmChat.AddChat RTBColors.InformationText, "[Warden] Packet Data:" & vbNewLine & DebugOutput(Data)
        End If
    
    Case WARDEN_RAW_FAILURE: '//The module was not able to handle the packet itself (most likely 0x05)
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] The Warden module was unable to handle a packet, you will be disconnected soon."
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] For information on this, please visit http://www.stealthbot.net/sb/warden-issues/?handlefailed ."
        
        If (MDebug("warden")) Then
            frmChat.AddChat RTBColors.InformationText, "[Warden] Packet Data:" & vbNewLine & DebugOutput(Data)
        End If
        
    Case WARDEN_PACKET_FAILURE: '//Something went HORRIBLY wrong in warden_packet, should NEVER happen.
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] Something went horribly wrong in Warden_Packet(), you will be disconnected soon."
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] For information on this, please visit http://www.stealthbot.net/sb/warden-issues/?horrible ."
        
        If (MDebug("warden")) Then
            frmChat.AddChat RTBColors.InformationText, "[Warden] Packet Data:" & vbNewLine & DebugOutput(Data)
        End If
        
    Case WARDEN_INIT_FAILURE: '//Calling Init() in the module failed
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] Unable to initalize the Warden module, you will be disconnected soon."
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] For information on this, please visit http://www.stealthbot.net/sb/warden-issues/?init ."
    
    'case WARDEN_LOAD_FILE_FAILURE '//Could not load module from file [Not to bad, prolly just dosen't exist] This should never come up
    
    Case WARDEN_LOAD_MD5_FAILURE: '//Failed MD5 checksum when loading module [Either Bad tranfer or HD file corrupt]
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] Transfer failed because the MD5 checksum incorrect, you will be disconnected soon."
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] For information on this, please visit http://www.stealthbot.net/sb/warden-issues/md5 ."
        
    Case WARDEN_LOAD_INVALID_SIGNATURE: '//Module failed RSA verification
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] Transfer failed because the RSA signature is invalid, you will be disconnected soon."
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] For information on this, please visit http://www.stealthbot.net/sb/warden-issues/?rsa ."
        
    Case WARDEN_LOAD_DECOMPRESS_FAILURE: '//Module failed to decompress properly
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] Failed to decompress the Warden module, you will be disconnected soon."
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] For information on this, please visit http://www.stealthbot.net/sb/warden-issues/?decompress ."
        
    Case WARDEN_LOAD_PREP_FAILURE: '//Module prepare failed, Usually if module is corrupt
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] Failed to prep the Warden module, you will be disconnected soon."
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] For information on this, please visit http://www.stealthbot.net/sb/warden-issues/?prep ."
        
    Case WARDEN_CHECK_UNKNOWN_COMMAND: '//Unknown sub-command in CHEAT_CHECKS
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] The Warden has asked us to perform an unknown cheat-check, you will be disconnected soon."
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] For information on this, please visit " _
         & "http://www.stealthbot.net/sb/warden-issues/?unknown-cheatcheck ."
        
        If (MDebug("warden")) Then
            frmChat.AddChat RTBColors.InformationText, "[Warden] Packet Data: " & vbNewLine & DebugOutput(Data)
        End If
        
    Case WARDEN_CHECK_TO_MANY_LIBS: '//There were more then 4 libraries in a single 0x02 packet [this is eww yes, but I'll figure out a beter way later]
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] To many libraries in Cheat Check, you will be disconnected soon"
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] For information on this, please visit " _
         & "http://www.stealthbot.net/sb/warden-issues/?toomanylibs ."
         
        If (MDebug("warden")) Then
            frmChat.AddChat RTBColors.InformationText, "[Warden] Packet Data: " & vbNewLine & DebugOutput(Data)
        End If
    
    Case WARDEN_MEM_UNKNOWN_PRODUCT: '//The product from 0x50 != WC3, SC, or D2
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] Unknown product code form SID_AUTH_INFO, you will be diconnected soon"
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] For information on this, please visit " _
         & "http://www.stealthbot.net/sb/warden-issues/?unknown-prodcode ."
        
    Case WARDEN_MEM_UNKNOWN_SEGMENT: '//Could not read segment from ini file
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] Could not read a segment from Warden.ini, you will be disconnected soon."
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] Make sure you have the latest Warden data from http://www.stealthbot.net/sb/warden/"
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] For more information on this, please visit " _
         & "http://www.stealthbot.net/sb/warden-issues/?unknown-segment ."
        
        If (MDebug("warden")) Then
            frmChat.AddChat RTBColors.InformationText, "[Warden] Packet Data: " & vbNewLine & DebugOutput(Data)
        End If
        
    Case WARDEN_INVALID_INSTANCE: '//The instance passed to this function was invalid
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] An Invalid instance was passed to Data, Did Init() fail?"
        frmChat.AddChat RTBColors.ErrorMessageText, "[Warden] For information on this, please visit " _
         & "http://www.stealthbot.net/sb/warden-issues/?invalid-instance ."
        
  End Select
    
  WardenData = (Result <> WARDEN_IGNORE)
End Function
    
