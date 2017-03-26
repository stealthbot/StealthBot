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
    (ByVal NLS As Long, ByVal Out As Long)

Private Declare Sub nls_get_M1 Lib "BNCSutil.dll" _
    (ByVal NLS As Long, ByVal Out As Long, ByVal B_Value As Long, ByVal Salt As Long)

Private Declare Sub nls_get_v Lib "BNCSutil.dll" _
    (ByVal NLS As Long, ByVal Out As Long, ByVal Salt As Long)

Private Declare Function nls_check_M2 Lib "BNCSutil.dll" _
    (ByVal NLS As Long, ByVal M2 As Long, ByVal B_Value As Long, ByVal Salt As Long) As Long

Private Declare Function nls_check_signature Lib "BNCSutil.dll" _
    (ByVal Address As Long, ByVal Signature As Long) As Long

Private Declare Sub nls_get_S Lib "BNCSutil.dll" _
    (ByVal NLS As Long, ByVal Out As Long, ByVal B_Value As Long, ByVal Salt As Long)

Private Declare Sub nls_get_K Lib "BNCSutil.dll" _
    (ByVal NLS As Long, ByVal Out As Long, ByVal Salt As Long)

Private Declare Function nls_account_change_proof Lib "BNCSutil.dll" _
    (ByVal NLS As Long, ByVal Buffer As Long, ByVal NewPassword As Long, _
    ByVal B_Value As Long, ByVal Salt As Long) As Long 'returns a new NLS pointer for the new password
    

Private m_NlsHandle As Long

Private m_Salt As String * 32
Private m_v As String * 32
Private m_B As String * 32

Private m_Username As String
Private m_Password As String
Private m_Initialized As Boolean
Private m_Generated As Boolean

' make sure all possible handles have been freed
Private Sub Class_Terminate()

    If Not m_NlsHandle = 0 Then
        nls_free m_NlsHandle
        m_NlsHandle = 0
    End If

    m_Initialized = False
    m_Generated = False

End Sub

Public Sub Terminate()

    Call Class_Terminate

End Sub

Public Function Initialize(ByVal Username As String, ByVal Password As String) As Boolean
    
    Dim arrU() As Byte
    Dim arrP() As Byte

    ' default to return false
    Initialize = False

    ' dispose of all previous NLS objects
    Class_Terminate

    ' save username and password
    m_Username = Username
    m_Password = Password
    m_NlsHandle = nls_init(Username, Password)

    ' return true if nls_init succeeded
    If m_NlsHandle <> 0 Then
        Initialize = True
        m_Initialized = True
    End If

End Function

Public Property Get Username() As String

    Username = m_Username

End Property

Public Property Get Password() As String

    Password = m_Password

End Property

Public Property Get GeneratedSaltAndVerifier() As Boolean

    GeneratedSaltAndVerifier = m_Generated

End Property


' SRP-level functions (use these if you know what you're doing)

' store the Salt value
' store the value when parsing SID_AUTH_ACCOUNTLOGON->C
' length should be 32 bytes
Public Property Let Srp_Salt(ByVal Salt As String)

    m_Salt = Salt

End Property

' gets the stored Salt value
' this just gets the value you stored (or created in AccountCreate())
Public Property Get Srp_Salt() As String

    Srp_Salt = m_Salt

End Property

' store the verifier value
' length should be 32 bytes
Public Property Let Srp_v(ByVal Value As String)

    m_v = Value

End Property

' gets the stored verifier value
' this just gets the value you stored (or created in AccountCreate())
Public Property Get Srp_v() As String

    Srp_v = m_v

End Property

' get the A value
' get this value when building SID_AUTH_ACCOUNTLOGON->S
' length will be 32 bytes
Public Property Get Srp_A() As String

    Dim arrA(0 To 31) As Byte

    If (Not m_Initialized) Then
        Srp_A = String$(32, vbNullChar)
        Exit Property
    End If

    nls_get_A m_NlsHandle, VarPtr(arrA(0))

    Srp_A = ByteArrToString(arrA())

End Property

' store the B value
' store this value when parsing SID_AUTH_ACCOUNTLOGON->C
' length should be 32 bytes
Public Property Let Srp_B(ByVal B_Value As String)

    m_B = B_Value

End Property

' gets the stored B value
' this just gets the value you stored
Public Property Get Srp_B() As String

    Srp_B = m_B

End Property

' get the M[1] value
' get this value when building SID_AUTH_ACCOUNTLOGONPROOF->S
' length will be 20 bytes
Public Property Get Srp_M1() As String

    Dim arrM1(0 To 19) As Byte
    Dim arrB() As Byte
    Dim arrSalt() As Byte

    If (Not m_Initialized) Then
        Srp_M1 = String$(20, vbNullChar)
        Exit Property
    End If

    arrB() = StringToByteArr(m_B)
    arrSalt() = StringToByteArr(m_Salt)

    nls_get_M1 m_NlsHandle, VarPtr(arrM1(0)), VarPtr(arrB(0)), VarPtr(arrSalt(0))

    Srp_M1 = ByteArrToString(arrM1())

End Property

' get the S value (the secret value)
' length will be 32 bytes
Public Property Get Srp_S() As String

    Dim arrS(0 To 31) As Byte
    Dim arrB() As Byte
    Dim arrSalt() As Byte

    If (Not m_Initialized) Then
        Srp_S = String$(20, vbNullChar)
        Exit Property
    End If

    arrB() = StringToByteArr(m_B)
    arrSalt() = StringToByteArr(m_Salt)

    nls_get_S m_NlsHandle, VarPtr(arrS(0)), VarPtr(arrB(0)), VarPtr(arrSalt(0))

    Srp_S = ByteArrToString(arrS())

End Property

' get the K value (a value based on the secret)
' length will be 40 bytes
Public Property Get Srp_K() As String

    Dim arrK(0 To 39) As Byte
    Dim arrSalt() As Byte

    If (Not m_Initialized) Then
        Srp_K = String$(40, vbNullChar)
        Exit Property
    End If

    arrSalt() = StringToByteArr(m_Salt)

    nls_get_K m_NlsHandle, VarPtr(arrK(0)), VarPtr(arrSalt(0))

    Srp_K = ByteArrToString(arrK())

End Property

' check the M[2] value
' optionally check this value when parsing SID_AUTH_ACCOUNTLOGONPROOF->C
' M[2] length should be 20 bytes
Public Function SrpVerifyM2(ByVal M2 As String) As Boolean

    Dim arrM2() As Byte
    Dim arrB() As Byte
    Dim arrSalt() As Byte

    If (Not m_Initialized) Then
        SrpVerifyM2 = False
        Exit Function
    End If

    arrM2() = StringToByteArr(M2)
    arrB() = StringToByteArr(m_B)
    arrSalt() = StringToByteArr(m_Salt)

    SrpVerifyM2 = nls_check_M2(m_NlsHandle, VarPtr(arrM2(0)), VarPtr(arrB(0)), VarPtr(arrSalt(0)))

End Function

' create the Salt and Verifier
' create these values when building SID_AUTH_ACCOUNTCREATE->S or SID_AUTH_CHANGEPROOF->S
' Salt length will be 32 bytes
' Verifier length will be 32 bytes
Public Sub GenerateSaltAndVerifier()

    Dim arrSalt(0 To 31) As Byte
    Dim arrV(0 To 31) As Byte
    Dim i As Integer

    If (Not m_Initialized) Then
        Exit Sub
    End If

    Randomize

    For i = LBound(arrSalt) To UBound(arrSalt)
        arrSalt(i) = CByte(Rnd() * 255)
    Next i

    nls_get_v m_NlsHandle, VarPtr(arrV(0)), VarPtr(arrSalt(0))

    m_Salt = ByteArrToString(arrSalt())
    m_v = ByteArrToString(arrV())

    m_Generated = True

End Sub

' Battle.net packet-level functions (use these to populate a DataBuffer automatically)
' this is more for scripts-- they must pass a clsDataBuffer into the Buffer As Variant arguments
' (defining them As clsDataBuffer resulted in scripting type mismatch errors)

' populates your databuffer for SID_AUTH_ACCOUNTCREATE->S
Public Sub AccountCreate(ByRef Buffer As Variant)

    ' create an s and v
    Call GenerateSaltAndVerifier

    ' insert s
    Buffer.InsertNonNTString m_Salt

    ' insert v
    Buffer.InsertNonNTString m_v

    ' insert username
    Buffer.InsertNTString m_Username

End Sub

' populates your databuffer for SID_AUTH_ACCOUNTLOGON->S
Public Sub AccountLogon(ByRef Buffer As Variant)

    Dim A_Value As String * 32

    ' get A
    A_Value = Srp_A()

    ' insert A
    Buffer.InsertNonNTString A_Value

    ' insert username
    Buffer.InsertNTString m_Username

End Sub

' populates your databuffer for SID_AUTH_ACCOUNTLOGONPROOF->S
Public Sub AccountLogonProof(ByRef Buffer As Variant, ByVal Salt As String, ByVal B_Value As String)

    Dim M1_Value As String * 20

    ' let salt
    Srp_Salt = Salt

    ' let B
    Srp_B = B_Value

    ' get M[1]
    M1_Value = Srp_M1()

    ' insert M[1]
    Buffer.InsertNonNTString M1_Value

End Sub

' populates your databuffer for SID_AUTH_ACCOUNTCHANGE->S
Public Sub AccountChange(ByRef Buffer As Variant)

    ' do the same as SID_AUTH_ACCOUNTLOGON->S
    AccountLogon Buffer

End Sub

' populates your databuffer for SID_AUTH_ACCOUNTCHANGEPROOF->S
Public Sub AccountChangeProof(ByRef Buffer As Variant, ByVal Salt As String, ByVal B_Value As String, ByVal NewPassword As String)

    Dim oNLS As clsNLS
    Set oNLS = New clsNLS

    Call oNLS.Initialize(m_Username, NewPassword)
    ' create an s and v
    Call oNLS.GenerateSaltAndVerifier

    ' do the same as SID_AUTH_ACCOUNTLOGONPROOF->S
    AccountLogonProof Buffer, Salt, B_Value

    ' insert s
    Buffer.InsertNonNTString oNLS.Srp_Salt

    ' insert v
    Buffer.InsertNonNTString oNLS.Srp_v

    Call oNLS.Terminate
    Set oNLS = Nothing

End Sub


' verifies a WC3 server signature, no .Initialize required
' pass IPAddress as "#.#.#.#"
Public Function VerifyServerSignature(ByVal IPAddress As String, ByVal Signature As String) As Boolean

    Dim lngAddr As Long
    Dim arrSig() As Byte

    lngAddr = inet_addr(IPAddress)
    arrSig() = StringToByteArr(Signature)

    VerifyServerSignature = nls_check_signature(lngAddr, VarPtr(arrSig(0)))

End Function

