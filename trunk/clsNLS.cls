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
Option Explicit

Private m_id As Long            'This instances indentifier.

Private m_username As String    'The username associated with this instance.
Private m_password As String    'The password associated with this instance.

Private Sub Class_Initialize()
    m_id = -1
End Sub


Public Property Get ID() As Long
    ID = m_id
End Property

Public Property Let ID(ByVal Value As Long)
    m_id = Value
End Property

Public Property Get Username() As String
    Username = m_username
End Property

Public Property Let Username(ByVal Value As String)
    m_username = Value
End Property

Public Property Get Password() As String
    Password = m_password
End Property

Public Property Let Password(ByVal Value As String)
    m_password = Value
End Property


Public Function Initialize(ByVal strUsername As String, ByVal strPassword As String) As Long
    'Have we already initialized?
    If m_id = -1 Then
        'No.
        m_username = strUsername
        m_password = strPassword
    
        m_id = nls_init(strUsername, strPassword)
    Else
        'Yes, re-initialize.
        m_id = nls_reinit(m_id, strUsername, strPassword)
    End If
    
    Initialize = m_id
End Function

Public Function CreateAccount(ByVal Buffer As String) As Long
    'Have we initialied?
    If m_id = -1 Then
        'No.
        CreateAccount = -1
        Exit Function
    Else
        CreateAccount = nls_account_create(m_id, Buffer, Len(Buffer))
    End If
End Function

Public Function LogonAccount(ByVal Buffer As String) As Long
    'Have we initialized?
    If m_id = -1 Then
        'No.
        LogonAccount = -1
        Exit Function
    Else
        LogonAccount = nls_account_logon(m_id, Buffer, Len(Buffer))
    End If
End Function

Public Function GetClientKey() As String
    Dim strBuf As String * 32
    
    'Have we initialized?
    If m_id = -1 Then
        'No.
        GetClientKey = vbNullString
        Exit Function
    Else
        Call nls_get_A(m_id, strBuf)
        GetClientKey = strBuf
    End If
End Function

Public Function GetClientPasswordProof(ByVal b As String, ByVal salt As String) As String
    Dim strBuf As String * 20
    
    'Getting tired of writing the same comment...
    If m_id = -1 Then
        'I think you know what that means.
        GetPasswordProof = vbNullString
        Exit Function
    Else
        Call nls_get_M1(m_id, strBuf, b, salt)
        GetPasswordProof = strBuf
    End If
End Function

Public Function CheckPasswordProof(ByVal M2 As String, ByVal b As String, ByVal salt As String) As Long
    'I'm just gonna stop now.
    If m_id = -1 Then
        CheckPasswordProof = -1
        Exit Function
    Else
        CheckPasswordProof = nls_check_M2(m_id, M2, b, salt)
    End If
End Function

Public Function CheckSignature(ByVal Signature As String) As Long
    If m_id = -1 Then
        CheckSignature = -1
        Exit Function
    Else
        CheckSignature = nls_check_signature(m_id, Signature)
    End If
End Function

Public Function ChangeAccount(ByVal Buffer As String, ByVal NewPassword As String, ByVal b As String, ByVal salt As String) As clsNLS
    If m_id = -1 Then
        ChangeAccount = -1
        Exit Function
    Else
        Set ChangeAccount = New clsNLS
        ChangeAccount.ID = nls_account_change_proof(m_id, Buffer, NewPassword, b, salt)
    End If
End Function

Public Function GetSecret(ByVal b As String, ByVal salt As String) As String
    Dim strBuf As String * 32
    If m_id = -1 Then
        GetSecret = vbNullString
        Exit Function
    Else
        Call nls_get_S(m_id, strBuf, b, salt)
        GetSecret = strBuf
    End If
End Function

Public Function GetPasswordProof(ByVal s As String) As String
    Dim strBuf As String * 20
    If m_id = -1 Then
        GetPasswordProof = -1
        Exit Function
    Else
        Call nls_get_K(m_id, strBuf, s)
        GetPasswordProof = strBuf
    End If
End Function
