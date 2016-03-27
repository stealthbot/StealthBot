VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsNLS.cls
' Copyright (C) 2009 Nate Book
' this class provides scripters the ability to use the features of nls.c/h in BNCSutil:
' NLS/SRP handling functions

Option Explicit


' BNCSutil.dll functions
Private Declare Function nls_init Lib "BNCSutil.dll" _
    (ByVal Username As String, ByVal Password As String) As Long    ' returns a pointer
    
Private Declare Sub nls_free Lib "BNCSutil.dll" _
    (ByVal NLS As Long)
    
Private Declare Sub nls_get_A Lib "BNCSutil.dll" _
    (ByVal NLS As Long, ByVal Out As String)
    
Private Declare Sub nls_get_M1 Lib "BNCSutil.dll" _
    (ByVal NLS As Long, ByVal Out As String, ByVal B As String, ByVal Salt As String)
    
Private Declare Sub nls_get_v Lib "BNCSutil.dll" _
    (ByVal NLS As Long, ByVal Out As String, ByVal Salt As String)
    
Private Declare Function nls_check_M2 Lib "BNCSutil.dll" _
    (ByVal NLS As Long, ByVal M2 As String, ByVal B As String, ByVal Salt As String) As Long

Private Declare Function nls_check_signature Lib "BNCSutil.dll" _
    (ByVal Address As Long, ByVal Signature As String) As Long
    
Private Declare Sub nls_get_S Lib "BNCSutil.dll" _
    (ByVal NLS As Long, ByVal Out As String, ByVal B As String, ByVal Salt As String)
    
Private Declare Sub nls_get_K Lib "BNCSutil.dll" _
    (ByVal NLS As Long, ByVal Out As String, ByVal s As String)
    
Private Declare Function nls_account_change_proof Lib "BNCSutil.dll" _
    (ByVal NLS As Long, ByVal Buffer As String, ByVal NewPassword As String, _
    ByVal B As String, ByVal Salt As String) As Long 'returns a new NLS pointer for the new password
    

Private m_NlsHandle As Long
Private m_NewNlsHandle As Long
Private m_OldNlsHandle As Long

Private m_Salt As String * 32
Private m_B As String * 32
Private m_v As String * 32

Private m_Username As String
Private m_Password As String
Private m_NewPassword As String
Private m_Initialized As Boolean

' make sure all possible handles have been freed
Private Sub Class_Terminate()

    If Not m_NlsHandle = 0 Then
        nls_free m_NlsHandle
        m_NlsHandle = 0
    End If

    If Not m_NewNlsHandle = 0 Then
        nls_free m_NewNlsHandle
        m_NewNlsHandle = 0
    End If

    If Not m_OldNlsHandle = 0 Then
        nls_free m_OldNlsHandle
        m_OldNlsHandle = 0
    End If
    
    m_Initialized = False

End Sub

Public Function Initialize(ByVal Username As String, ByVal Password As String) As Boolean

    ' default to return false
    Initialize = False
    
    ' dispose of all previous NLS objects
    Class_Terminate
    
    ' save username and password
    m_Username = Username
    m_Password = Password
    
    m_NlsHandle = nls_init(Username, Password)
    
    ' return true if nls_init succeeded
    If Not m_NlsHandle = 0 Then
        Initialize = True
        m_Initialized = True
        SrpGetSaltAndVerifier m_Salt, m_v
    End If

End Function

Public Property Get Username() As String
    Username = m_Username
End Property


' SRP-level functions (use these if you know what you're doing)

' get the A value
' get this value when building SID_AUTH_ACCOUNTLOGON->S
' length will be 32 bytes
Public Property Get SrpA() As String
    If (Not m_Initialized) Then Initialize BotVars.Username, BotVars.Password
    Dim a As String * 32

    nls_get_A m_NlsHandle, a

    SrpA = a

End Property

' store the Salt value
' store the value when parsing SID_AUTH_ACCOUNTLOGON->C
' length should be 32 bytes
Public Property Let SrpSalt(ByVal Salt As String)

    m_Salt = Salt

End Property

' gets the stored Salt value
' this just gets the value you stored (or created in AccountCreate())
Public Property Get SrpSalt() As String
    If (Not m_Initialized) Then Initialize BotVars.Username, BotVars.Password
    If (StrComp(m_Salt, String$(32, Chr$(0))) = 0) Then Initialize BotVars.Username, BotVars.Password
    SrpSalt = m_Salt
    
End Property

' store the B value
' store this value when parsing SID_AUTH_ACCOUNTLOGON->C
' length should be 32 bytes
Public Property Let SrpB(ByVal B As String)

    m_B = B

End Property

' gets the stored B value
' this just gets the value you stored
Public Property Get SrpB() As String
    If (Not m_Initialized) Then Initialize BotVars.Username, BotVars.Password
    SrpB = m_B

End Property

' store the verifier value
' length should be 32 bytes
Public Property Let Srpv(ByVal v As String)

    m_v = v

End Property

' gets the stored verifier value
' this just gets the value you stored
Public Property Get Srpv() As String
    If (Not m_Initialized) Then Initialize BotVars.Username, BotVars.Password
    Srpv = m_v

End Property

' get the M[1] value
' get this value when building SID_AUTH_ACCOUNTLOGONPROOF->S
' length will be 20 bytes
Public Property Get SrpM1() As String
    If (Not m_Initialized) Then Initialize BotVars.Username, BotVars.Password

    Dim M1 As String * 20

    nls_get_M1 m_NlsHandle, M1, m_B, m_Salt
    
    SrpM1 = M1

End Property

' get the S value (the secret value)
' length will be 32 bytes
Public Property Get SrpS() As String
    If (Not m_Initialized) Then Initialize BotVars.Username, BotVars.Password
    Dim s As String * 32
    
    nls_get_S m_NlsHandle, s, m_B, m_Salt

    SrpS = s

End Property

' get the K value (a value based on the secret)
' length will be 40 bytes
Public Property Get SrpK() As String
    If (Not m_Initialized) Then Initialize BotVars.Username, BotVars.Password
    Dim K As String * 40
    
    nls_get_K m_NlsHandle, K, m_Salt
    
    SrpK = K

End Property

' check the M[2] value
' optionally check this value when parsing SID_AUTH_ACCOUNTLOGONPROOF->C
' M[2] length should be 20 bytes
Public Function SrpVerifyM2(ByVal M2 As String) As Boolean
    If (Not m_Initialized) Then Initialize BotVars.Username, BotVars.Password
    SrpVerifyM2 = nls_check_M2(m_NlsHandle, M2, m_B, m_Salt)

End Function

' check the M[2] value
' optionally check this value when parsing SID_AUTH_ACCOUNTCHANGEPROOF->C
' M[2] length should be 20 bytes
' must have set PersistOld in .AccountChangeProof() before calling this, or the handle was lost!
Public Function SrpVerifyOldM2(ByVal M2 As String) As Boolean
    If (Not m_Initialized) Then Initialize BotVars.Username, BotVars.Password
    SrpVerifyOldM2 = nls_check_M2(m_OldNlsHandle, M2, m_B, m_Salt)
    
    ' they shouldn't need to use this handle anymore-- free it
    nls_free m_OldNlsHandle
    
    m_OldNlsHandle = 0

End Function

' create the Salt and Verifier
' create these values when building SID_AUTH_ACCOUNTCREATE->S
' Salt length will be 32 bytes
' Verifier length will be 32 bytes
Public Sub SrpGetSaltAndVerifier(ByRef Salt As String, ByRef Verifier As String)
    If (Not m_Initialized) Then Initialize BotVars.Username, BotVars.Password
    Dim s As String * 32
    Dim v As String * 32
    Dim i As Integer
    
    Randomize
    
    s = String$(32, 0)
    For i = 1 To 32
        Mid$(s, i, 1) = Chr$(CInt(Rnd() * 255))
    Next i
    
    m_Salt = s
    
    nls_get_v m_NlsHandle, v, s
    
    Salt = s
    Verifier = v

End Sub

' Battle.net packet-level functions (use these to populate a DataBuffer automatically)
' this is more for scripts-- they must pass a clsDataBuffer into the Buffer As Variant arguments
' (defining them As clsDataBuffer resulted in scripting type mismatch errors)

' populates your databuffer for SID_AUTH_ACCOUNTCREATE->S
Public Sub AccountCreate(ByRef Buffer As Variant)

    Dim s As String * 32
    Dim v As String * 32
    
    ' create an s and v
    SrpGetSaltAndVerifier s, v
    
    ' insert s
    Buffer.InsertNonNTString s
    
    ' insert v
    Buffer.InsertNonNTString v
    
    ' insert username
    Buffer.InsertNTString m_Username

End Sub

' populates your databuffer for SID_AUTH_ACCOUNTLOGON->S
Public Sub AccountLogon(ByRef Buffer As Variant)

    Dim a As String * 32
    
    ' get A
    a = SrpA()
    
    ' insert A
    Buffer.InsertNonNTString a
    
    ' insert username
    Buffer.InsertNTString m_Username

End Sub

' populates your databuffer for SID_AUTH_ACCOUNTLOGONPROOF->S
Public Sub AccountLogonProof(ByRef Buffer As Variant, ByVal Salt As String, ByVal B As String)

    Dim M1 As String * 20
    
    ' let salt
    SrpSalt = Salt
    
    ' let B
    SrpB = B
    
    ' get M[1]
    M1 = SrpM1()
    
    ' insert M[1]
    Buffer.InsertNonNTString M1

End Sub

' populates your databuffer for SID_AUTH_ACCOUNTCHANGE->S
Public Sub AccountChange(ByRef Buffer As Variant, ByVal NewPassword As String)

    ' store new password
    m_NewPassword = NewPassword
    
    ' create the new NLS handle
    m_NewNlsHandle = nls_init(m_Username, m_NewPassword)
    If m_NewNlsHandle = 0 Then
        Exit Sub
    End If

    ' do the same as SID_AUTH_ACCOUNTLOGON->S
    AccountLogon Buffer

End Sub

' populates your databuffer for SID_AUTH_ACCOUNTCHANGEPROOF->S
' pass true to PersistOld here to keep a copy of the old NLS handle in order
' to check the old password's M[2] value with .SrpVerifyOldM2(M2)
Public Sub AccountChangeProof(ByRef Buffer As Variant, ByVal Salt As String, ByVal B As String, Optional ByVal PersistOld As Boolean = False)
    
    Dim s As String * 32
    Dim v As String * 32

    ' do the same as SID_AUTH_ACCOUNTLOGONPROOF->S
    AccountLogonProof Buffer, Salt, B
    
    ' if we are keeping the "old" handle in m_OldNlsHandle for .VerifyOldM2()...
    If PersistOld Then
        ' move current handle to "old" handle-- for use with .VerifyOldM2()
        m_OldNlsHandle = m_NlsHandle
    Else
        ' free handle
        nls_free m_NlsHandle
    End If
    
    ' move "new" handle to current handle
    m_NlsHandle = m_NewNlsHandle
    
    ' zero "new" handle
    m_NewNlsHandle = 0
    
    ' create an s and v
    SrpGetSaltAndVerifier s, v
    
    ' insert s
    Buffer.InsertNonNTString s
    
    ' insert v
    Buffer.InsertNonNTString v
End Sub


' verifies a WC3 server signature, no .Initialize required
' pass IPAddress as "#.#.#.#"
Public Function VerifyServerSignature(ByVal IPAddress As String, ByVal Signature As String) As Boolean

    Dim lngAddr As Long

    VerifyServerSignature = nls_check_signature(aton(IPAddress), Signature)

End Function

