Option Strict Off
Option Explicit On
Friend Class clsNLS
	' clsNLS.cls
	' Copyright (C) 2009 Nate Book
	' this class provides scripters the ability to use the features of nls.c/h in BNCSutil:
	' NLS/SRP handling functions
	
	
	
	' BNCSutil.dll functions
	Private Declare Function nls_init Lib "BNCSutil.dll" (ByVal Username As String, ByVal Password As String) As Integer ' returns a pointer
	
	Private Declare Sub nls_free Lib "BNCSutil.dll" (ByVal NLS As Integer)
	
	Private Declare Sub nls_get_A Lib "BNCSutil.dll" (ByVal NLS As Integer, ByVal Out As String)
	
	Private Declare Sub nls_get_M1 Lib "BNCSutil.dll" (ByVal NLS As Integer, ByVal Out As String, ByVal B As String, ByVal Salt As String)
	
	Private Declare Sub nls_get_v Lib "BNCSutil.dll" (ByVal NLS As Integer, ByVal Out As String, ByVal Salt As String)
	
	Private Declare Function nls_check_M2 Lib "BNCSutil.dll" (ByVal NLS As Integer, ByVal M2 As String, ByVal B As String, ByVal Salt As String) As Integer
	
	Private Declare Function nls_check_signature Lib "BNCSutil.dll" (ByVal Address As Integer, ByVal Signature As String) As Integer
	
	Private Declare Sub nls_get_S Lib "BNCSutil.dll" (ByVal NLS As Integer, ByVal Out As String, ByVal B As String, ByVal Salt As String)
	
	Private Declare Sub nls_get_K Lib "BNCSutil.dll" (ByVal NLS As Integer, ByVal Out As String, ByVal s As String)
	
	Private Declare Function nls_account_change_proof Lib "BNCSutil.dll" (ByVal NLS As Integer, ByVal Buffer As String, ByVal NewPassword As String, ByVal B As String, ByVal Salt As String) As Integer 'returns a new NLS pointer for the new password
	
	
	Private m_NlsHandle As Integer
	Private m_NewNlsHandle As Integer
	Private m_OldNlsHandle As Integer
	
	Private m_Salt As New VB6.FixedLengthString(32)
	Private m_B As New VB6.FixedLengthString(32)
	Private m_v As New VB6.FixedLengthString(32)
	
	Private m_Username As String
	Private m_Password As String
	Private m_NewPassword As String
	Private m_Initialized As Boolean
	
	' make sure all possible handles have been freed
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		If Not m_NlsHandle = 0 Then
			nls_free(m_NlsHandle)
			m_NlsHandle = 0
		End If
		
		If Not m_NewNlsHandle = 0 Then
			nls_free(m_NewNlsHandle)
			m_NewNlsHandle = 0
		End If
		
		If Not m_OldNlsHandle = 0 Then
			nls_free(m_OldNlsHandle)
			m_OldNlsHandle = 0
		End If
		
		m_Initialized = False
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Function Initialize(ByVal Username As String, ByVal Password As String) As Boolean
		
		' default to return false
		Initialize = False
		
		' dispose of all previous NLS objects
		Class_Terminate_Renamed()
		
		' save username and password
		m_Username = Username
		m_Password = Password
		
		m_NlsHandle = nls_init(Username, Password)
		
		' return true if nls_init succeeded
		If Not m_NlsHandle = 0 Then
			Initialize = True
			m_Initialized = True
			SrpGetSaltAndVerifier(m_Salt.Value, m_v.Value)
		End If
		
	End Function
	
	Public ReadOnly Property Username() As String
		Get
			Username = m_Username
		End Get
	End Property
	
	
	' SRP-level functions (use these if you know what you're doing)
	
	' get the A value
	' get this value when building SID_AUTH_ACCOUNTLOGON->S
	' length will be 32 bytes
	Public ReadOnly Property SrpA() As String
		Get
			If (Not m_Initialized) Then Initialize(BotVars.Username, BotVars.Password)
			Dim a As New VB6.FixedLengthString(32)
			
			nls_get_A(m_NlsHandle, a.Value)
			
			SrpA = a.Value
			
		End Get
	End Property
	
	' store the Salt value
	' store the value when parsing SID_AUTH_ACCOUNTLOGON->C
	' length should be 32 bytes
	
	' gets the stored Salt value
	' this just gets the value you stored (or created in AccountCreate())
	Public Property SrpSalt() As String
		Get
			If (Not m_Initialized) Then Initialize(BotVars.Username, BotVars.Password)
			If (StrComp(m_Salt.Value, New String(Chr(0), 32)) = 0) Then Initialize(BotVars.Username, BotVars.Password)
			SrpSalt = m_Salt.Value
			
		End Get
		Set(ByVal Value As String)
			
			m_Salt.Value = Value
			
		End Set
	End Property
	
	' store the B value
	' store this value when parsing SID_AUTH_ACCOUNTLOGON->C
	' length should be 32 bytes
	
	' gets the stored B value
	' this just gets the value you stored
	Public Property SrpB() As String
		Get
			If (Not m_Initialized) Then Initialize(BotVars.Username, BotVars.Password)
			SrpB = m_B.Value
			
		End Get
		Set(ByVal Value As String)
			
			m_B.Value = Value
			
		End Set
	End Property
	
	' store the verifier value
	' length should be 32 bytes
	
	' gets the stored verifier value
	' this just gets the value you stored
	Public Property Srpv() As String
		Get
			If (Not m_Initialized) Then Initialize(BotVars.Username, BotVars.Password)
			Srpv = m_v.Value
			
		End Get
		Set(ByVal Value As String)
			
			m_v.Value = Value
			
		End Set
	End Property
	
	' get the M[1] value
	' get this value when building SID_AUTH_ACCOUNTLOGONPROOF->S
	' length will be 20 bytes
	Public ReadOnly Property SrpM1() As String
		Get
			If (Not m_Initialized) Then Initialize(BotVars.Username, BotVars.Password)
			
			Dim M1 As New VB6.FixedLengthString(20)
			
			nls_get_M1(m_NlsHandle, M1.Value, m_B.Value, m_Salt.Value)
			
			SrpM1 = M1.Value
			
		End Get
	End Property
	
	' get the S value (the secret value)
	' length will be 32 bytes
	Public ReadOnly Property SrpS() As String
		Get
			If (Not m_Initialized) Then Initialize(BotVars.Username, BotVars.Password)
			Dim s As New VB6.FixedLengthString(32)
			
			nls_get_S(m_NlsHandle, s.Value, m_B.Value, m_Salt.Value)
			
			SrpS = s.Value
			
		End Get
	End Property
	
	' get the K value (a value based on the secret)
	' length will be 40 bytes
	Public ReadOnly Property SrpK() As String
		Get
			If (Not m_Initialized) Then Initialize(BotVars.Username, BotVars.Password)
			Dim K As New VB6.FixedLengthString(40)
			
			nls_get_K(m_NlsHandle, K.Value, m_Salt.Value)
			
			SrpK = K.Value
			
		End Get
	End Property
	
	' check the M[2] value
	' optionally check this value when parsing SID_AUTH_ACCOUNTLOGONPROOF->C
	' M[2] length should be 20 bytes
	Public Function SrpVerifyM2(ByVal M2 As String) As Boolean
		If (Not m_Initialized) Then Initialize(BotVars.Username, BotVars.Password)
		SrpVerifyM2 = nls_check_M2(m_NlsHandle, M2, m_B.Value, m_Salt.Value)
		
	End Function
	
	' check the M[2] value
	' optionally check this value when parsing SID_AUTH_ACCOUNTCHANGEPROOF->C
	' M[2] length should be 20 bytes
	' must have set PersistOld in .AccountChangeProof() before calling this, or the handle was lost!
	Public Function SrpVerifyOldM2(ByVal M2 As String) As Boolean
		If (Not m_Initialized) Then Initialize(BotVars.Username, BotVars.Password)
		SrpVerifyOldM2 = nls_check_M2(m_OldNlsHandle, M2, m_B.Value, m_Salt.Value)
		
		' they shouldn't need to use this handle anymore-- free it
		nls_free(m_OldNlsHandle)
		
		m_OldNlsHandle = 0
		
	End Function
	
	' create the Salt and Verifier
	' create these values when building SID_AUTH_ACCOUNTCREATE->S
	' Salt length will be 32 bytes
	' Verifier length will be 32 bytes
	Public Sub SrpGetSaltAndVerifier(ByRef Salt As String, ByRef Verifier As String)
		If (Not m_Initialized) Then Initialize(BotVars.Username, BotVars.Password)
		Dim s As New VB6.FixedLengthString(32)
		Dim v As New VB6.FixedLengthString(32)
		Dim i As Short
		
		Randomize()
		
		s.Value = New String(Chr(0), 32)
		For i = 1 To 32
			Mid(s.Value, i, 1) = Chr(CShort(Rnd() * 255))
		Next i
		
		m_Salt.Value = s.Value
		
		nls_get_v(m_NlsHandle, v.Value, s.Value)
		
		Salt = s.Value
		Verifier = v.Value
		
	End Sub
	
	' Battle.net packet-level functions (use these to populate a DataBuffer automatically)
	' this is more for scripts-- they must pass a clsDataBuffer into the Buffer As Variant arguments
	' (defining them As clsDataBuffer resulted in scripting type mismatch errors)
	
	' populates your databuffer for SID_AUTH_ACCOUNTCREATE->S
	Public Sub AccountCreate(ByRef Buffer As Object)
		
		Dim s As New VB6.FixedLengthString(32)
		Dim v As New VB6.FixedLengthString(32)
		
		' create an s and v
		SrpGetSaltAndVerifier(s.Value, v.Value)
		
		' insert s
		'UPGRADE_WARNING: Couldn't resolve default property of object Buffer.InsertNonNTString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Buffer.InsertNonNTString(s.Value)
		
		' insert v
		'UPGRADE_WARNING: Couldn't resolve default property of object Buffer.InsertNonNTString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Buffer.InsertNonNTString(v.Value)
		
		' insert username
		'UPGRADE_WARNING: Couldn't resolve default property of object Buffer.InsertNTString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Buffer.InsertNTString(m_Username)
		
	End Sub
	
	' populates your databuffer for SID_AUTH_ACCOUNTLOGON->S
	Public Sub AccountLogon(ByRef Buffer As Object)
		
		Dim a As New VB6.FixedLengthString(32)
		
		' get A
		a.Value = SrpA()
		
		' insert A
		'UPGRADE_WARNING: Couldn't resolve default property of object Buffer.InsertNonNTString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Buffer.InsertNonNTString(a.Value)
		
		' insert username
		'UPGRADE_WARNING: Couldn't resolve default property of object Buffer.InsertNTString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Buffer.InsertNTString(m_Username)
		
	End Sub
	
	' populates your databuffer for SID_AUTH_ACCOUNTLOGONPROOF->S
	Public Sub AccountLogonProof(ByRef Buffer As Object, ByVal Salt As String, ByVal B As String)
		
		Dim M1 As New VB6.FixedLengthString(20)
		
		' let salt
		SrpSalt = Salt
		
		' let B
		SrpB = B
		
		' get M[1]
		M1.Value = SrpM1()
		
		' insert M[1]
		'UPGRADE_WARNING: Couldn't resolve default property of object Buffer.InsertNonNTString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Buffer.InsertNonNTString(M1.Value)
		
	End Sub
	
	' populates your databuffer for SID_AUTH_ACCOUNTCHANGE->S
	Public Sub AccountChange(ByRef Buffer As Object, ByVal NewPassword As String)
		
		' store new password
		m_NewPassword = NewPassword
		
		' create the new NLS handle
		m_NewNlsHandle = nls_init(m_Username, m_NewPassword)
		If m_NewNlsHandle = 0 Then
			Exit Sub
		End If
		
		' do the same as SID_AUTH_ACCOUNTLOGON->S
		AccountLogon(Buffer)
		
	End Sub
	
	' populates your databuffer for SID_AUTH_ACCOUNTCHANGEPROOF->S
	' pass true to PersistOld here to keep a copy of the old NLS handle in order
	' to check the old password's M[2] value with .SrpVerifyOldM2(M2)
	Public Sub AccountChangeProof(ByRef Buffer As Object, ByVal Salt As String, ByVal B As String, Optional ByVal PersistOld As Boolean = False)
		
		Dim s As New VB6.FixedLengthString(32)
		Dim v As New VB6.FixedLengthString(32)
		
		' do the same as SID_AUTH_ACCOUNTLOGONPROOF->S
		AccountLogonProof(Buffer, Salt, B)
		
		' if we are keeping the "old" handle in m_OldNlsHandle for .VerifyOldM2()...
		If PersistOld Then
			' move current handle to "old" handle-- for use with .VerifyOldM2()
			m_OldNlsHandle = m_NlsHandle
		Else
			' free handle
			nls_free(m_NlsHandle)
		End If
		
		' move "new" handle to current handle
		m_NlsHandle = m_NewNlsHandle
		
		' zero "new" handle
		m_NewNlsHandle = 0
		
		' create an s and v
		SrpGetSaltAndVerifier(s.Value, v.Value)
		
		' insert s
		'UPGRADE_WARNING: Couldn't resolve default property of object Buffer.InsertNonNTString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Buffer.InsertNonNTString(s.Value)
		
		' insert v
		'UPGRADE_WARNING: Couldn't resolve default property of object Buffer.InsertNonNTString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Buffer.InsertNonNTString(v.Value)
	End Sub
	
	
	' verifies a WC3 server signature, no .Initialize required
	' pass IPAddress as "#.#.#.#"
	Public Function VerifyServerSignature(ByVal IPAddress As String, ByVal Signature As String) As Boolean
		
		Dim lngAddr As Integer
		
		VerifyServerSignature = nls_check_signature(aton(IPAddress), Signature)
		
	End Function
End Class