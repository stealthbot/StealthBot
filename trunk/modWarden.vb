Option Strict Off
Option Explicit On
Module modWarden
	
	Private Enum SHA1Versions
		Sha1 = 0
		BrokenSHA1 = 1
		LockdownSHA1 = 2
		WardenSHA1 = 3
		Max = &HFFFFFFFF
	End Enum
	
	Private Structure SHA1Context
		<VBFixedArray(4)> Dim IntermediateHash() As Integer
		Dim LengthLow As Integer
		Dim LengthHigh As Integer
		Dim MessageBlockIndex As Short
		<VBFixedArray(63)> Dim MessageBlock() As Byte
		Dim Computed As Byte
		Dim Corrupted As Byte
		Dim Version As SHA1Versions
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim IntermediateHash(4)
			ReDim MessageBlock(63)
		End Sub
	End Structure
	
	Private Structure MD5Context
		<VBFixedArray(3)> Dim IntermediateHash() As Integer
		Dim LengthLow As Integer
		Dim LengthHigh As Integer
		Dim MessageBlockIndex As Short
		<VBFixedArray(63)> Dim MessageBlock() As Byte
		Dim Computed As Byte
		Dim Corrupted As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim IntermediateHash(3)
			ReDim MessageBlock(63)
		End Sub
	End Structure
	
	Private Structure MedivRandomContext
		Dim Index As Integer
		<VBFixedArray(19)> Dim Data() As Byte
		<VBFixedArray(19)> Dim Source1() As Byte
		<VBFixedArray(19)> Dim Source2() As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim Data(19)
			ReDim Source1(19)
			ReDim Source2(19)
		End Sub
	End Structure
	
	Private Declare Sub rc4_init Lib "Warden.dll" (ByVal Key As String, ByVal Base As String, ByVal length As Integer)
	Private Declare Sub rc4_crypt Lib "Warden.dll" (ByVal Key As String, ByVal Data As String, ByVal length As Integer)
	Private Declare Sub rc4_crypt_data Lib "Warden.dll" (ByVal Data As String, ByVal DataLength As Integer, ByVal Base As String, ByVal BaseLength As Integer)
	
	'UPGRADE_WARNING: Structure SHA1Context may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function sha1_reset Lib "Warden.dll" (ByRef Context As SHA1Context) As Integer
	'UPGRADE_WARNING: Structure SHA1Context may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function sha1_input Lib "Warden.dll" (ByRef Context As SHA1Context, ByVal Data As String, ByVal length As Integer) As Integer
	'UPGRADE_WARNING: Structure SHA1Context may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function sha1_digest Lib "Warden.dll" (ByRef Context As SHA1Context, ByVal digest As String) As Integer
	Private Declare Function sha1_checksum Lib "Warden.dll" (ByVal Data As String, ByVal length As Integer, ByVal Version As Integer) As Integer
	
	'UPGRADE_WARNING: Structure MD5Context may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function md5_reset Lib "Warden.dll" (ByRef Context As MD5Context) As Integer
	'UPGRADE_WARNING: Structure MD5Context may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function md5_input Lib "Warden.dll" (ByRef Context As MD5Context, ByVal Data As String, ByVal length As Integer) As Integer
	'UPGRADE_WARNING: Structure MD5Context may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function md5_digest Lib "Warden.dll" (ByRef Context As MD5Context, ByVal digest As String) As Integer
	Private Declare Function md5_verify_data Lib "Warden.dll" (ByVal Data As String, ByVal length As Integer, ByVal CorrectMD5 As String) As Boolean
	
	'UPGRADE_WARNING: Structure MedivRandomContext may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Sub mediv_random_init Lib "Warden.dll" (ByRef Context As MedivRandomContext, ByVal Seed As String, ByVal length As Integer)
	'UPGRADE_WARNING: Structure MedivRandomContext may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Sub mediv_random_get_bytes Lib "Warden.dll" (ByRef Context As MedivRandomContext, ByVal Buffer As String, ByVal length As Integer)
	
	Private Declare Function warden_init Lib "Warden.dll" (ByVal SocketHandle As Integer) As Integer
    Private Declare Function warden_data Lib "Warden.dll" (ByVal Instance As Integer, ByVal Direction As Integer, ByVal PacketID As Integer, ByVal Data() As Byte, ByVal length As Integer) As Integer
	Private Declare Function warden_cleanup Lib "Warden.dll" (ByVal Instance As Integer) As Integer
	Private Declare Function warden_set_data_file Lib "Warden.dll" (ByVal Instance As Integer, ByVal File As String, ByVal length As Integer) As Integer
	Private Declare Function warden_config Lib "Warden.dll" (ByVal Instance As Integer, ByVal ConfigBit As Integer, ByVal Enabled As Byte) As Integer
	
	Public Const WARDEN_CONFIG_SAVE_CHECKS As Integer = 1 '//Save Information about cheat checks (Opcode 0x02) to Data File
	Public Const WARDEN_CONFIG_SAVE_UNKNOWN As Integer = 2 '//Save Unknown information (use in conjunction with Debug mode to get new Warden offsets)
	Public Const WARDEN_CONFIG_LOG_CHECKS As Integer = 4 '//Log ALL information about checks that happen, in real time
	Public Const WARDEN_CONFIG_LOG_PACKETS As Integer = 8 '//Log ALL decoded Warden packet data
	Public Const WARDEN_CONFIG_DEBUG_MODE As Integer = 16 '//Debug mode, does a lot of shit u.u
	Public Const WARDEN_CONFIG_USE_GAME_FILES As Integer = 32 '//Will attempt to grab unknown Mem Check offsets from the game file specified
	'//  Will try to load library the file, using the path specified in the INI EXA:
	'//[Files_WAR3]
	'//Default=C:\Program Files\Warcraft III\WAR3.exe
	'//Game.dll=C:\Program Files\Warcraft III\Game.dll
	Private Const WARDEN_SEND As Integer = &H0
	Private Const WARDEN_RECV As Integer = &H1
	Private Const WARDEN_BNCS As Integer = &H2
	
	Private Const WARDEN_IGNORE As Integer = &H0 '//Not a warden packet, Handle internally
	Private Const WARDEN_SUCCESS As Integer = &H1 '//All Went Well, Don't handle the packet Internally
	Private Const WARDEN_UNKNOWN_PROTOCOL As Integer = &H2 '//Not used, will be when adding support for MCP/UDP
	Private Const WARDEN_UNKNOWN_SUBID As Integer = &H3 '//Unknown Sub-ID [Not 0x00, 0x01, 0x02, or 0x05]
	Private Const WARDEN_RAW_FAILURE As Integer = &H4 '//The module was not able to handle the packet itself
	Private Const WARDEN_PACKET_FAILURE As Integer = &H5 '//Something went HORRIBLY wrong in warden_packet, should NEVER happen.
	Private Const WARDEN_INIT_FAILURE As Integer = &H6 '//Calling Init() in the module failed
	Private Const WARDEN_LOAD_FILE_FAILURE As Integer = &H7 '//Could not load module from file [Not to bad, prolly just dosen't exist]
	Private Const WARDEN_LOAD_MD5_FAILURE As Integer = &H8 '//Failed MD5 checksum when loading module [Either Bad tranfer or HD file corrupt]
	Private Const WARDEN_LOAD_INVALID_SIGNATURE As Integer = &H9 '//Module failed RSA verification
	Private Const WARDEN_LOAD_DECOMPRESS_FAILURE As Integer = &HA '//Module failed to decompress properly
	Private Const WARDEN_LOAD_PREP_FAILURE As Integer = &HB '//Module prepare failed, Usually if module is corrupt
	Private Const WARDEN_CHECK_UNKNOWN_COMMAND As Integer = &HC '//Unknown sub-command in CHEAT_CHECKS
	Private Const WARDEN_CHECK_TO_MANY_LIBS As Integer = &HD '//There were more then 4 libraries in a single 0x02 packet [this is eww yes, but I'll figure out a beter way later]
	Private Const WARDEN_MEM_UNKNOWN_PRODUCT As Integer = &HE '//The product from 0x50 != WC3, SC, or D2
	Private Const WARDEN_MEM_UNKNOWN_SEGMENT As Integer = &HF '//Could not read segment from ini file
	Private Const WARDEN_INVALID_INSTANCE As Integer = &H10 '//Instance passed to this function was invalid
	
	Public WardenInstance As Integer
	'====================================================================================================
	'^^^^Warden Stuff
	'vvvvCrev Stuff
	'====================================================================================================
	Private Declare Function check_revision Lib "Warden.dll" (ByVal ArchiveTime As String, ByVal ArchiveName As String, ByVal Seed As String, ByVal INIFile As String, ByVal INIHeader As String, ByRef Version As Integer, ByRef Checksum As Integer, ByVal Result As String) As Integer
	Private Declare Function crev_max_result Lib "Warden.dll" () As Integer
	Private Declare Function crev_error_description Lib "Warden.dll" (ByVal ErrorCode As Integer, ByVal description As String, ByVal Size As Integer) As Integer
	
	Private Const CREV_SUCCESS As Integer = 0 '//If everything went ok
	Private Const CREV_UNKNOWN_VERSION As Integer = 1 '//Unknown version, Not lockdown, Or Ver
	Private Const CREV_UNKNOWN_REVISION As Integer = 2 '//Unknown Revision (0-7 for old, 0-19 for lockdown)
	Private Const CREV_MALFORMED_SEED As Integer = 3 '//If the Seed passed in wasn't able to be translated properly
	Private Const CREV_MISSING_FILENAME As Integer = 4 '//We were not able to get the file path information from the INI file, Result holds more info.
	Private Const CREV_MISSING_FILE As Integer = 5 '//Was not able to open a file, Result has the File Path
	Private Const CREV_FILE_INFO_ERROR As Integer = 6 '//And error while trying to get the file info string, Result holds the path of the file
	Private Const CREV_TOFEW_RVAS As Integer = 7 '//Less then 14 RVAs in hash file, probably corrupt
	Private Const CREV_UNKNOWN_RELOC_TYPE As Integer = 8 '//Unknown Reloc Type, only 16, 32, and 64 bit are known
	Private Const CREV_OUT_OF_MEMORY As Integer = 9 '//Out of memory
	Private Const CREV_CORRUPT_IMPORT_DATA As Integer = 10 '//Corrupt IAT data, File may be Corrupt
	'====================================================================================================
	Public Function Warden_CheckRevision(ByRef sArchiveName As String, ByRef sArchiveFileTime As String, ByRef sSeed As String, ByRef sHeader As String, ByRef lVersion As Integer, ByRef lChecksum As Integer, ByRef sResult As String) As Boolean
		On Error GoTo trap
		Dim lRet As Integer
		Dim ltVersion As Integer
		Dim ltChecksum As Integer
		Dim stResult As String
		Dim sError As String
		Dim i As Integer
		Dim ft As FILETIME
		Dim st As SYSTEMTIME
		Dim sFileTime As String
		
		Warden_CheckRevision = False
		stResult = New String(Chr(0), crev_max_result)
		
		lRet = check_revision(sArchiveFileTime, sArchiveName, sSeed, GetFilePath(FILE_CREV_INI), sHeader, ltVersion, ltChecksum, stResult)
		
		i = InStr(1, stResult, Chr(0))
		If (i > 0) Then stResult = Left(stResult, i - 1)
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Not Len(sArchiveFileTime) = 8) Then sArchiveFileTime = Left(StringFormat("{0}{1}", sArchiveFileTime, New String(Chr(0), 8)), 8)
		'UPGRADE_WARNING: Couldn't resolve default property of object ft. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(ft, sArchiveFileTime, 8)
		
		FileTimeToSystemTime(ft, st)
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sFileTime = StringFormat("{0}/{1}/{2} {3}:{4}:{5}", st.wMonth, st.wDay, st.wYear, st.wHour, st.wMinute, st.wSecond)
		
		Select Case lRet
			Case CREV_SUCCESS
				Warden_CheckRevision = True
				sResult = stResult
				lChecksum = ltChecksum
				lVersion = ltVersion
				Exit Function
				
			Case CREV_UNKNOWN_VERSION, CREV_UNKNOWN_REVISION
				frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNCS] Warden.dll does not support checkrevision for {0} {1}", sArchiveName, sFileTime))
				frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Make sure you have the latest Warden.dll from http://www.stealthbot.net/sb/warden/")
				
			Case CREV_MALFORMED_SEED
				frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNCS] The seed value string was malformed: {0}", IIf(InStr(1, sSeed, "A=", CompareMethod.Text) > 0, sSeed, StrToHex(sSeed, True))))
				frmChat.AddChat(RTBColors.ErrorMessageText, "[BNCS] Make sure you have the latest Warden.dll from http://www.stealthbot.net/sb/warden/")
				
			Case CREV_MISSING_FILENAME
				frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNCS] Could not read key {0}{1}{0} under [{2}] from CheckRevision.ini, Update your configuration", Chr(34), stResult, sHeader))
				
			Case CREV_MISSING_FILE
				frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNCS] Could not open file {0}{1}{0}", Chr(34), stResult))
				
			Case CREV_FILE_INFO_ERROR
				frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNCS] Error retrieving information from {0}{1}{0}", Chr(34), stResult))
				
			Case Else
				sResult = New String(Chr(0), crev_max_result)
				i = crev_error_description(lRet, sResult, Len(sResult))
				If (i > 0) Then
					sResult = New String(Chr(0), i)
					i = crev_error_description(lRet, sResult, Len(sResult))
				End If
				
				i = InStr(1, sResult, Chr(0))
				If (i > 0) Then stResult = Left(stResult, i - 1)
				frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("[BNCS] Unknown CheckRevision error: {0}", stResult))
				
		End Select
		
trap: 
		If (Err.Number = 53) Then
			frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Warden.dll was not found, Local Hashing will not work.")
			frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] To fix this, make sure you have the latest Warden.dll from http://www.stealthbot.net/sb/warden/")
			Err.Clear()
		End If
		Warden_CheckRevision = False
	End Function
	
	Public Sub WardenCleanup(ByRef Instance As Integer)
		On Error GoTo trap
		If (Not Instance = 0) Then Call warden_cleanup(Instance)
		Exit Sub
		
trap: 
		If (Err.Number = 53) Then
			frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Warden.dll was not found, Warden support will not work.")
			frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] To fix this, make sure you have the latest Warden data from http://www.stealthbot.net/sb/warden/")
			Err.Clear()
		End If
	End Sub
	
	Public Function WardenInitilize(ByVal SocketHandle As Integer) As Integer
		On Error GoTo trap
		Dim INIPath As String
		Dim Instance As Integer
		Dim DebugString As String
		
		Instance = warden_init(SocketHandle)
		
		If (Instance > 0) Then
			
			INIPath = GetFilePath(FILE_WARDEN_INI)
			
			warden_set_data_file(Instance, INIPath, Len(INIPath))
			
			DebugString = CStr(Config.DebugWarden)
			If StrictIsNumeric(DebugString) Then Call warden_config(Instance, CInt(DebugString), 2)
			
			WardenInitilize = Instance
		End If
		Exit Function
		
trap: 
		If (Err.Number = 53) Then
			frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Warden.dll was not found, Warden support will not work.")
			frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] To fix this, make sure you have the latest Warden data from http://www.stealthbot.net/sb/warden/ .")
			Err.Clear()
		End If
	End Function
	
    Public Function WardenData(ByRef Instance As Integer, ByRef bData() As Byte, ByRef Send As Boolean) As Boolean
        Dim ID As Integer
        Dim Result As Integer
        Dim Data() As Byte

        ' Packet ID
        ID = bData(1)

        ' Packet Data
        ReDim Data(bData.Length - 4)
        Buffer.BlockCopy(bData, 4, Data, 0, Data.Length)

        If (Instance = 0) Then
            If (MDebug("warden")) Then
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Attempted to call Data() with an invalid instance.")
            End If
            WardenData = False
            Exit Function
        End If

        Result = warden_data(Instance, WARDEN_BNCS Or IIf(Send, WARDEN_SEND, WARDEN_RECV), ID, Data, Len(Data))

        Select Case Result
            Case WARDEN_SUCCESS '//All Went Well, Don't handle the packet Internally
                If (MDebug("warden")) Then
                    Select Case Data(0)
                        Case 0 : frmChat.AddChat(RTBColors.InformationText, "[Warden] Handled Module Information")
                        Case 1 : frmChat.AddChat(RTBColors.InformationText, "[Warden] Handled Module Transfer")
                        Case 2 : frmChat.AddChat(RTBColors.InformationText, "[Warden] Handled Cheat Check")
                        Case 5 : frmChat.AddChat(RTBColors.InformationText, "[Warden] Handled New Crypt Keys")
                        Case Else : frmChat.AddChat(RTBColors.InformationText, "[Warden] Handled Unknown 0x" & ZeroOffset(Data(0), 2))
                    End Select
                End If
                'case WARDEN_UNKNOWN_PROTOCOL '//Not used, will be when adding support for MCP/UDP
            Case WARDEN_UNKNOWN_SUBID '//Unknown Sub-ID [Not 0x00, 0x01, 0x02, or 0x05]
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Unknown sub-command 0x" & ZeroOffset(Data(0), 2) & ", you will be disconnected soon.")
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] For information on this, please visit http://www.stealthbot.net/sb/issues/warden/?unknown&id=0x" & ZeroOffset(Data(0), 2) & " .")

                If (MDebug("warden")) Then
                    frmChat.AddChat(RTBColors.InformationText, "[Warden] Packet Data:" & vbNewLine & DebugOutput(Data))
                End If

            Case WARDEN_RAW_FAILURE '//The module was not able to handle the packet itself (most likely 0x05)
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] The Warden module was unable to handle a packet, you will be disconnected soon.")
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] For information on this, please visit http://www.stealthbot.net/sb/issues/warden/?handlefailed .")

                If (MDebug("warden")) Then
                    frmChat.AddChat(RTBColors.InformationText, "[Warden] Packet Data:" & vbNewLine & DebugOutput(Data))
                End If

            Case WARDEN_PACKET_FAILURE '//Something went HORRIBLY wrong in warden_packet, should NEVER happen.
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Something went horribly wrong in Warden_Packet(), you will be disconnected soon.")
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] For information on this, please visit http://www.stealthbot.net/sb/issues/warden/?horrible .")

                If (MDebug("warden")) Then
                    frmChat.AddChat(RTBColors.InformationText, "[Warden] Packet Data:" & vbNewLine & DebugOutput(Data))
                End If

            Case WARDEN_INIT_FAILURE '//Calling Init() in the module failed
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Unable to initalize the Warden module, you will be disconnected soon.")
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] For information on this, please visit http://www.stealthbot.net/sb/issues/warden/?init .")

                'case WARDEN_LOAD_FILE_FAILURE '//Could not load module from file [Not to bad, prolly just dosen't exist] This should never come up

            Case WARDEN_LOAD_MD5_FAILURE '//Failed MD5 checksum when loading module [Either Bad tranfer or HD file corrupt]
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Transfer failed because the MD5 checksum incorrect, you will be disconnected soon.")
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] For information on this, please visit http://www.stealthbot.net/sb/issues/warden/md5 .")

            Case WARDEN_LOAD_INVALID_SIGNATURE '//Module failed RSA verification
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Transfer failed because the RSA signature is invalid, you will be disconnected soon.")
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] For information on this, please visit http://www.stealthbot.net/sb/issues/warden/?rsa .")

            Case WARDEN_LOAD_DECOMPRESS_FAILURE '//Module failed to decompress properly
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Failed to decompress the Warden module, you will be disconnected soon.")
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] For information on this, please visit http://www.stealthbot.net/sb/issues/warden/?decompress .")

            Case WARDEN_LOAD_PREP_FAILURE '//Module prepare failed, Usually if module is corrupt
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Failed to prep the Warden module, you will be disconnected soon.")
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] For information on this, please visit http://www.stealthbot.net/sb/issues/warden/?prep .")

            Case WARDEN_CHECK_UNKNOWN_COMMAND '//Unknown sub-command in CHEAT_CHECKS
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] The Warden has asked us to perform an unknown cheat-check, you will be disconnected soon.")
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] For information on this, please visit " & "http://www.stealthbot.net/sb/issues/warden/?unknownCheatCheck .")

                If (MDebug("warden")) Then
                    frmChat.AddChat(RTBColors.InformationText, "[Warden] Packet Data: " & vbNewLine & DebugOutput(Data))
                End If

            Case WARDEN_CHECK_TO_MANY_LIBS '//There were more then 4 libraries in a single 0x02 packet [this is eww yes, but I'll figure out a beter way later]
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] To many libraries in Cheat Check, you will be disconnected soon")
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] For information on this, please visit " & "http://www.stealthbot.net/sb/issues/warden/?toomanylibs .")

                If (MDebug("warden")) Then
                    frmChat.AddChat(RTBColors.InformationText, "[Warden] Packet Data: " & vbNewLine & DebugOutput(Data))
                End If

            Case WARDEN_MEM_UNKNOWN_PRODUCT '//The product from 0x50 != WC3, SC, or D2
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Unknown product code form SID_AUTH_INFO, you will be diconnected soon")
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] For information on this, please visit " & "http://www.stealthbot.net/sb/issues/warden/?unknownProdCode .")

            Case WARDEN_MEM_UNKNOWN_SEGMENT '//Could not read segment from ini file
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Could not read a segment from Warden.ini, you will be disconnected soon.")
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] Make sure you have the latest Warden data from http://www.stealthbot.net/sb/warden/")
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] For more information on this, please visit " & "http://www.stealthbot.net/sb/issues/warden/?unknownSegment .")

                If (MDebug("warden")) Then
                    frmChat.AddChat(RTBColors.InformationText, "[Warden] Packet Data: " & vbNewLine & DebugOutput(Data))
                End If

            Case WARDEN_INVALID_INSTANCE '//The instance passed to this function was invalid
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] An Invalid instance was passed to Data, Did Init() fail?")
                frmChat.AddChat(RTBColors.ErrorMessageText, "[Warden] For information on this, please visit " & "http://www.stealthbot.net/sb/issues/warden/?invalidInstance .")

        End Select

        WardenData = (Result <> WARDEN_IGNORE)
    End Function
End Module