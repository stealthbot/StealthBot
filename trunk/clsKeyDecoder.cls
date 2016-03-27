VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyDecoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsKeyDecoder.cls
' Copyright (C) 2016
' Provides access to the BNCSutil key decoding functions

Option Explicit


' BNCSutil.dll functions
Private Declare Function kd_quick Lib "BNCSutil.dll" _
    (ByVal CDKey As String, ByVal clientToken As Long, ByVal serverToken As Long, _
    PublicValue As Long, Product As Long, ByVal HashBuffer As String, ByVal BufferLen As Long) As Long

Private Declare Function kd_init Lib "BNCSutil.dll" () As Long

Private Declare Function kd_create Lib "BNCSutil.dll" _
    (ByVal CDKey As String, ByVal KeyLength As Long) As Long
    
Private Declare Function kd_free Lib "BNCSutil.dll" _
    (ByVal decoder As Long) As Long
    
Private Declare Function kd_val2Length Lib "BNCSutil.dll" _
    (ByVal decoder As Long) As Long
    
Private Declare Function kd_product Lib "BNCSutil.dll" _
    (ByVal decoder As Long) As Long
    
Private Declare Function kd_val1 Lib "BNCSutil.dll" _
    (ByVal decoder As Long) As Long
    
Private Declare Function kd_val2 Lib "BNCSutil.dll" _
    (ByVal decoder As Long) As Long
    
Private Declare Function kd_longVal2 Lib "BNCSutil.dll" _
    (ByVal decoder As Long, ByVal Out As String) As Long
    
Private Declare Function kd_calculateHash Lib "BNCSutil.dll" _
    (ByVal decoder As Long, ByVal clientToken As Long, ByVal serverToken As Long) As Long
    
Private Declare Function kd_getHash Lib "BNCSutil.dll" _
    (ByVal decoder As Long, ByVal Out As String) As Long
    
Private Declare Function kd_isValid Lib "BNCSutil.dll" _
    (ByVal decoder As Long) As Long
    
    
    
Private m_ProductLookup As Dictionary

Private m_Result As Long    ' The result of the decoder initialization
Private m_Handle As Long    ' A handle to this decoder
Private m_Key As String     ' The key supplied during initialization
Private m_HashSize As Long  ' The size of the returned hash.

' Performs initial key analysis
Public Sub Initialize(ByVal strCdKey As String)
    If m_Result = -1 Then
        m_Result = kd_init()
    End If
    
    If m_Result > 0 Then
        If m_Handle > 0 Then
            Call kd_free(m_Handle)
        End If
        m_Handle = kd_create(strCdKey, Len(strCdKey))
    Else
        Call frmChat.AddChat(RTBColors.ErrorMessageText, "Unable to initialize CD key decoder.")
    End If

    m_HashSize = 0
    m_Key = strCdKey
End Sub

' Returns true if the key was successfully validated
Public Property Get IsValid() As Boolean
    IsValid = Not CBool(kd_isValid(m_Handle) = 0)
End Property

' Returns the key provided in Initialize()
Public Property Get Key() As String
    Key = m_Key
End Property

' Return the length of the key
Public Property Get KeyLength() As Integer
    KeyLength = Len(m_Key)
End Property

' Returns the key's product value
Public Property Get Product() As Long
    Product = kd_product(m_Handle)
End Property

' Returns the key's public value
Public Property Get PublicValue() As Long
    PublicValue = kd_val1(m_Handle)
End Property

' Returns the key's private value
Public Property Get PrivateValue() As Long
    PrivateValue = kd_val2(m_Handle)
End Property

' Returns the calculated hash of the key's product, public, and private values.
Public Property Get Hash() As String
    If Not (m_HashSize > 0) Then
        Hash = vbNullString
        Exit Property
    End If
    
    Hash = String$(m_HashSize, vbNullChar)
    Call kd_getHash(m_Handle, Hash)
End Property

' Calculates the key's hash
Public Function CalculateHash(ByVal clientToken As Long, ByVal serverToken As Long) As Boolean
    m_HashSize = kd_calculateHash(m_Handle, clientToken, serverToken)
    CalculateHash = CBool(m_HashSize > 0)
End Function

' Returns a human-friendly version of the key's product value.
Public Function GetProductName() As String
    If Product < 1 Then
        GetProductName = "Invalid"
        Exit Function
    End If
    If Not m_ProductLookup.Exists(Product) Then
        GetProductName = "Unrecognized product"
        Exit Function
    End If
    
    GetProductName = m_ProductLookup.Item(Product)
End Function


Private Sub Class_Initialize()
    'Default values
    m_Result = -1
    m_Handle = -1
    m_Key = vbNullString
    m_HashSize = 0
    
    'Create product name lookup dictionary
    Set m_ProductLookup = New Dictionary
    m_ProductLookup.Add &H1, "StarCraft"
    m_ProductLookup.Add &H2, "StarCraft"
    m_ProductLookup.Add &H4, "WarCraft II"
    m_ProductLookup.Add &H5, "Diablo II Beta"
    m_ProductLookup.Add &H6, "Diablo II"
    m_ProductLookup.Add &H7, "Diablo II"
    m_ProductLookup.Add &H9, "Diablo II Stress Test"
    m_ProductLookup.Add &HA, "Diablo II: Lord of Destruction"
    m_ProductLookup.Add &HC, "Diablo II: Lord of Destruction"
    m_ProductLookup.Add &HD, "WarCraft III: Reign of Chaos Beta"
    m_ProductLookup.Add &HE, "WarCraft III: Reign of Chaos"
    m_ProductLookup.Add &HF, "WarCraft III: Reign of Chaos"
    m_ProductLookup.Add &H11, "WarCraft III: The Frozen Throne Beta"
    m_ProductLookup.Add &H12, "WarCraft III: The Frozen Throne"
    m_ProductLookup.Add &H13, "WarCraft III: The Frozen Throne Retail"
    m_ProductLookup.Add &H17, "StarCraft Anthology"
    m_ProductLookup.Add &H18, "Diablo II Digital Download"
    m_ProductLookup.Add &H19, "Diablo II: Lord of Destruction Digital Download"
End Sub


Private Sub Class_Terminate()
    Call kd_free(m_Handle)

    Set m_ProductLookup = Nothing
End Sub
