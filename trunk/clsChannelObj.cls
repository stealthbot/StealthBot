VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChannelObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsChannelUserObj.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Private Const CHAN_PUBLIC     As Long = &H1
Private Const CHAN_MODERATED  As Long = &H2
Private Const CHAN_RESTRICTED As Long = &H4
Private Const CHAN_SILENT     As Long = &H8
Private Const CHAN_SYSTEM     As Long = &H10
Private Const CHAN_PRODUCT    As Long = &H20
Private Const CHAN_GLOBAL     As Long = &H1000

Private m_name            As String
Private m_flags           As Long
Private m_designated_heir As Integer
Private m_num_users       As Integer
Private m_num_joins       As Long
Private m_num_bans        As Long
Private m_num_kicks       As Long
Private m_users           As Collection
Private m_chat_queue      As Collection

Private Sub Class_Initialize()
    
    ' ...
    Set m_users = New Collection
    
End Sub

Public Property Get Name() As String

    ' ...
    Name = m_name

End Property

Public Property Let Name(str As String)

    ' ...
    m_name = str

End Property

Public Property Get Flags() As Long

    ' ...
    Flags = m_flags

End Property

Public Property Let Flags(lng As Long)

    ' ...
    m_flags = lng

End Property

Public Property Get IsSilent() As Boolean

    ' ...
    IsSilent = ((m_flags And CHAN_SILENT) = CHAN_SILENT)

End Property

Public Property Get UserCount() As Integer

    ' ...
    UserCount = m_num_users

End Property

Public Property Let UserCount(i As Integer)

    ' ...
    m_num_users = i

End Property

Public Property Get TotalJoinCount() As Long

    ' ...

End Property

Public Property Get TotalBanCount() As Long

    ' ...

End Property

Public Property Get TotalKickCount() As Long

    ' ...

End Property

Public Property Get Users() As Collection

    ' ...
    Set Users = m_users

End Property

Public Sub ClearUsers()

    ' ...
    Set m_users = New Collection

End Sub

Public Property Get Self() As clsChannelUserObj

    Dim i As Integer ' ...

    ' ...
    Set Self = New clsChannelUserObj
    
    ' ...
    For i = 1 To m_users.Count
        ' ...
        If (StrComp(m_users(i).DisplayName, CurrentUsername, vbBinaryCompare) = 0) Then
            ' ...
            Set Self = m_users(i)
            
            ' ...
            Exit Function
        End If
    Next i

End Property

Public Function GetUserIndexByName(ByVal Username As String) As Integer

    Dim i As Integer ' ...

    ' ...
    For i = 1 To m_users.Count
        ' ...
        If (StrComp(m_users(i).Name, Username, vbBinaryCompare) = 0) Then
            ' ...
            GetUserIndexByName = i
            
            ' ...
            Exit Function
        End If
    Next i

End Function

Public Function GetUserIndexByDisplayName(ByVal Username As String) As Integer

    Dim i As Integer ' ...

    ' ...
    For i = 1 To m_users.Count
        ' ...
        If (StrComp(m_users(i).DisplayName, Username, vbBinaryCompare) = 0) Then
            ' ...
            GetUserIndexByDisplayName = i
            
            ' ...
            Exit Function
        End If
    Next i

End Function
