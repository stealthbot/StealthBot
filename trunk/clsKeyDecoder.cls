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
Public Function Initialize(ByVal strCdKey As String) As Boolean
    strCdKey = UCase$(CDKeyReplacements(strCdKey))

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
Public Property Get IsValid() As Boolean
    IsValid = CBool(kd_isValid(m_Handle) = 1)
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
Public Property Get ProductValue() As Long
    ProductValue = kd_product(m_Handle)
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

' Returns the product to use with this key (if known).
Public Function GetProduct() As String
    Dim prodData() As Variant
    
    GetProduct = vbNullString
    
    If ProductValue < 1 Or Not m_ProductLookup.Exists(ProductValue) Then
        Exit Function
    End If
    
    prodData = m_ProductLookup.Item(ProductValue)
    GetProduct = CStr(prodData(0))
End Function

' Returns a human-friendly version of the key's product value.
Public Function GetProductName() As String
    Dim prodData() As Variant
    
    If ProductValue < 1 Then
        GetProductName = "Invalid"
        Exit Function
    End If
    
    If Not m_ProductLookup.Exists(ProductValue) Then
        GetProductName = "Unrecognized product"
        Exit Function
    End If
    
    prodData = m_ProductLookup.Item(ProductValue)
    GetProductName = CStr(prodData(1))
End Function

' Returns the key for display (with "-"'s)
Public Function GetKeyForDisplay() As String
    Select Case KeyLength
        Case 13: GetKeyForDisplay = Mid$(Key, 1, 4) & "-" & Mid$(Key, 5, 5) & "-" & Mid$(Key, 10, 4)
        Case 16: GetKeyForDisplay = Mid$(Key, 1, 4) & "-" & Mid$(Key, 5, 4) & "-" & Mid$(Key, 9, 4) & "-" & Mid$(Key, 13, 4)
        Case 26: GetKeyForDisplay = Mid$(Key, 1, 6) & "-" & Mid$(Key, 7, 4) & "-" & Mid$(Key, 11, 6) & "-" & Mid$(Key, 17, 4) & "-" & Mid$(Key, 21, 6)
        Case Else: GetKeyForDisplay = Key
    End Select
End Function

Private Sub Class_Initialize()
    'Default values
    m_Result = -1
    m_Handle = -1
    m_Key = vbNullString
    m_HashSize = 0
    
    'Create product name lookup dictionary
    Set m_ProductLookup = New Dictionary
    m_ProductLookup.Add &H1, Array("STAR", "StarCraft") '13
    m_ProductLookup.Add &H2, Array("STAR", "StarCraft") '13
    m_ProductLookup.Add &H4, Array("W2BN", "WarCraft II") '16
    m_ProductLookup.Add &H5, Array("D2DV", "Diablo II Beta") '16
    m_ProductLookup.Add &H6, Array("D2DV", "Diablo II") '16
    m_ProductLookup.Add &H7, Array("D2DV", "Diablo II") '16
    m_ProductLookup.Add &H9, Array("D2DV", "Diablo II Stress Test") '16
    m_ProductLookup.Add &HA, Array("D2XP", "Diablo II: Lord of Destruction") '16
    m_ProductLookup.Add &HC, Array("D2XP", "Diablo II: Lord of Destruction") '16
    m_ProductLookup.Add &HD, Array("WAR3", "WarCraft III: Reign of Chaos Beta") '26
    m_ProductLookup.Add &HE, Array("WAR3", "WarCraft III: Reign of Chaos") '26
    m_ProductLookup.Add &HF, Array("WAR3", "WarCraft III: Reign of Chaos") '26
    m_ProductLookup.Add &H11, Array("W3XP", "WarCraft III: The Frozen Throne Beta") '26
    m_ProductLookup.Add &H12, Array("W3XP", "WarCraft III: The Frozen Throne") '26
    m_ProductLookup.Add &H13, Array("W3XP", "WarCraft III: The Frozen Throne Retail") '26
    m_ProductLookup.Add &H17, Array("STAR", "StarCraft Anthology") '26
    m_ProductLookup.Add &H18, Array("D2DV", "Diablo II Digital Download") '26
    m_ProductLookup.Add &H19, Array("D2XP", "Diablo II: Lord of Destruction Digital Download") '26
End Sub

Private Function CDKeyReplacements(ByVal inString As String) As String
    inString = Replace(inString, "-", "")
    inString = Replace(inString, " ", "")
    CDKeyReplacements = Trim$(inString)
End Function

Private Sub Class_Terminate()
    Call kd_free(m_Handle)

    Set m_ProductLookup = Nothing
End Sub
