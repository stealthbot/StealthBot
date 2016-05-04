Option Strict Off
Option Explicit On
Friend Class clsKeyDecoder
	' clsKeyDecoder.cls
	' Copyright (C) 2016
	' Provides access to the BNCSutil key decoding functions
	
	
	
	' BNCSutil.dll functions
	Private Declare Function kd_quick Lib "BNCSutil.dll" (ByVal CDKey As String, ByVal ClientToken As Integer, ByVal ServerToken As Integer, ByRef PublicValue As Integer, ByRef Product As Integer, ByVal HashBuffer As String, ByVal BufferLen As Integer) As Integer
	
	Private Declare Function kd_init Lib "BNCSutil.dll" () As Integer
	
	Private Declare Function kd_create Lib "BNCSutil.dll" (ByVal CDKey As String, ByVal KeyLength As Integer) As Integer
	
	Private Declare Function kd_free Lib "BNCSutil.dll" (ByVal decoder As Integer) As Integer
	
	Private Declare Function kd_val2Length Lib "BNCSutil.dll" (ByVal decoder As Integer) As Integer
	
	Private Declare Function kd_product Lib "BNCSutil.dll" (ByVal decoder As Integer) As Integer
	
	Private Declare Function kd_val1 Lib "BNCSutil.dll" (ByVal decoder As Integer) As Integer
	
	Private Declare Function kd_val2 Lib "BNCSutil.dll" (ByVal decoder As Integer) As Integer
	
    Private Declare Function kd_longVal2 Lib "BNCSutil.dll" (ByVal decoder As Integer, ByVal Out() As Byte) As Integer
	
	Private Declare Function kd_calculateHash Lib "BNCSutil.dll" (ByVal decoder As Integer, ByVal ClientToken As Integer, ByVal ServerToken As Integer) As Integer
	
    Private Declare Function kd_getHash Lib "BNCSutil.dll" (ByVal decoder As Integer, ByVal Out() As Byte) As Integer
	
	Private Declare Function kd_isValid Lib "BNCSutil.dll" (ByVal decoder As Integer) As Integer
	
    Private Declare Sub calcHashBuf Lib "BNCSutil.dll" (ByVal Data() As Byte, ByVal length As Integer, ByVal Hash() As Byte)
	
	
	Private m_ProductLookup As Scripting.Dictionary
	
	Private m_Result As Integer ' The result of the decoder initialization
	Private m_Handle As Integer ' A handle to this decoder
	Private m_Key As String ' The key supplied during initialization
	Private m_HashSize As Integer ' The size of the returned hash.
    Private m_Hash() As Byte ' The celculated keyhash
	
	' Performs initial key analysis
	Public Function Initialize(ByVal strCdKey As String) As Boolean
		strCdKey = UCase(CDKeyReplacements(strCdKey))
		
		If m_Result = -1 Then
			m_Result = kd_init()
		End If
		
		If m_Result > 0 Then
			If m_Handle > 0 Then
				Call kd_free(m_Handle)
			End If
			m_Handle = kd_create(strCdKey, Len(strCdKey))
			Initialize = True
		Else
			Initialize = False
		End If
		
		m_HashSize = 0
		m_Key = strCdKey
	End Function
	
	' Returns true if the key was successfully validated
	Public ReadOnly Property IsValid() As Boolean
		Get
			IsValid = CBool(kd_isValid(m_Handle) = 1)
		End Get
	End Property
	
	' Returns the key provided in Initialize()
	Public ReadOnly Property Key() As String
		Get
			Key = m_Key
		End Get
	End Property
	
	' Return the length of the key
	Public ReadOnly Property KeyLength() As Short
		Get
			KeyLength = Len(m_Key)
		End Get
	End Property
	
	' Returns the key's product value
	Public ReadOnly Property ProductValue() As Integer
		Get
			ProductValue = kd_product(m_Handle)
		End Get
	End Property
	
	' Returns the key's public value
	Public ReadOnly Property PublicValue() As Integer
		Get
			PublicValue = kd_val1(m_Handle)
		End Get
	End Property
	
	' Returns the key's private value
    Public ReadOnly Property PrivateValue() As Byte()
        Get
            ReDim PrivateValue(kd_val2Length(m_Handle) - 1)

            If (kd_longVal2(m_Handle, PrivateValue) <= 0) Then
                PrivateValue = BitConverter.GetBytes(kd_val2(m_Handle))
            End If
        End Get
    End Property
	
	' Returns the calculated hash of the key's product, public, and private values.
    Public ReadOnly Property Hash() As Byte()
        Get
            If Not (m_HashSize > 0) Then
                Hash = Nothing
                Exit Property
            End If

            Hash = m_Hash
        End Get
    End Property
	
	' Calculates the key's hash
	Public Function CalculateHash(ByVal ClientToken As Integer, ByVal ServerToken As Integer, Optional ByVal LogonSystem As Integer = BNCS_NLS) As Boolean
		Dim HashContents As New clsDataBuffer
		
		' if private value is 0
		If Not IsValid Then
			CalculateHash = False
		End If
		
		With HashContents
			.InsertDWord(ClientToken)
			.InsertDWord(ServerToken)
			.InsertDWord(ProductValue)
			.InsertDWord(PublicValue)
			If LogonSystem = BNCS_NLS Then
				.InsertDWord(0)
			End If
            .InsertByteArr(PrivateValue)

            ReDim m_Hash(20)
			If KeyLength = 26 Then
				' fuck, this shit doesn't even give correct sha1s:
				'm_Hash = SHA1b(.Data)
				' just use standard bncsutil kd_calculatehash
				m_HashSize = kd_calculateHash(m_Handle, ClientToken, ServerToken)
				If m_HashSize <= 0 Then
					CalculateHash = False
				Else
                    ReDim m_Hash(m_HashSize)
					Call kd_getHash(m_Handle, m_Hash)
				End If
			Else
                Call calcHashBuf(.Data, Len(.Data), m_Hash)
			End If
		End With
		
		'UPGRADE_NOTE: Object HashContents may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		HashContents = Nothing
		
		'm_HashSize = kd_calculateHash(m_Handle, ClientToken, ServerToken)
		m_HashSize = Len(m_Hash)
		CalculateHash = CBool(Len(m_Hash) = 20)
	End Function
	
	' Returns the product to use with this key (if known).
	Public Function GetProduct() As String
		Dim prodData() As Object
		
		GetProduct = vbNullString
		
		If ProductValue < 1 Or Not m_ProductLookup.Exists(ProductValue) Then
			Exit Function
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_ProductLookup.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prodData = m_ProductLookup.Item(ProductValue)
		'UPGRADE_WARNING: Couldn't resolve default property of object prodData(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetProduct = CStr(prodData(0))
	End Function
	
	' Returns a human-friendly version of the key's product value.
	Public Function GetProductName() As String
		Dim prodData() As Object
		
		If ProductValue < 1 Then
			GetProductName = "Invalid"
			Exit Function
		End If
		
		If Not m_ProductLookup.Exists(ProductValue) Then
			GetProductName = "Unrecognized product"
			Exit Function
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_ProductLookup.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prodData = m_ProductLookup.Item(ProductValue)
		'UPGRADE_WARNING: Couldn't resolve default property of object prodData(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetProductName = CStr(prodData(1))
	End Function
	
	' Returns the key for display (with "-"'s)
	Public Function GetKeyForDisplay() As String
		Select Case KeyLength
			Case 13 : GetKeyForDisplay = Mid(Key, 1, 4) & "-" & Mid(Key, 5, 5) & "-" & Mid(Key, 10, 4)
			Case 16 : GetKeyForDisplay = Mid(Key, 1, 4) & "-" & Mid(Key, 5, 4) & "-" & Mid(Key, 9, 4) & "-" & Mid(Key, 13, 4)
			Case 26 : GetKeyForDisplay = Mid(Key, 1, 6) & "-" & Mid(Key, 7, 4) & "-" & Mid(Key, 11, 6) & "-" & Mid(Key, 17, 4) & "-" & Mid(Key, 21, 6)
			Case Else : GetKeyForDisplay = Key
		End Select
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'Default values
		m_Result = -1
		m_Handle = -1
		m_Key = vbNullString
		m_HashSize = 0
		
		'Create product name lookup dictionary
		m_ProductLookup = New Scripting.Dictionary
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&H1, New Object(){"STAR", "StarCraft"}) '13
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&H2, New Object(){"STAR", "StarCraft"}) '13
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&H4, New Object(){"W2BN", "WarCraft II"}) '16
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&H5, New Object(){"D2DV", "Diablo II Beta"}) '16
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&H6, New Object(){"D2DV", "Diablo II"}) '16
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&H7, New Object(){"D2DV", "Diablo II"}) '16
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&H9, New Object(){"D2DV", "Diablo II Stress Test"}) '16
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&HA, New Object(){"D2XP", "Diablo II: Lord of Destruction"}) '16
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&HC, New Object(){"D2XP", "Diablo II: Lord of Destruction"}) '16
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&HD, New Object(){"WAR3", "WarCraft III: Reign of Chaos Beta"}) '26
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&HE, New Object(){"WAR3", "WarCraft III: Reign of Chaos"}) '26
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&HF, New Object(){"WAR3", "WarCraft III: Reign of Chaos"}) '26
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&H11, New Object(){"W3XP", "WarCraft III: The Frozen Throne Beta"}) '26
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&H12, New Object(){"W3XP", "WarCraft III: The Frozen Throne"}) '26
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&H13, New Object(){"W3XP", "WarCraft III: The Frozen Throne Retail"}) '26
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&H17, New Object(){"STAR", "StarCraft Anthology"}) '26
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&H18, New Object(){"D2DV", "Diablo II Digital Download"}) '26
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		m_ProductLookup.Add(&H19, New Object(){"D2XP", "Diablo II: Lord of Destruction Digital Download"}) '26
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Private Function CDKeyReplacements(ByVal inString As String) As String
		inString = Replace(inString, "-", "")
		inString = Replace(inString, " ", "")
		CDKeyReplacements = Trim(inString)
	End Function
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		Call kd_free(m_Handle)
		
		'UPGRADE_NOTE: Object m_ProductLookup may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_ProductLookup = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class