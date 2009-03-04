VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBanlistUserObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsBanlistUserObj.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Private m_name          As String
Private m_banned_date   As Date
Private m_ban_reason    As String
Private m_operator      As String
Private m_duplicate_ban As Boolean
Private m_active_ban    As Boolean

Private Sub Class_Initialize()

    ' ...
    m_active_ban = True
    
End Sub

Public Property Get Name() As String

    ' ...
    Name = modEvents.CleanUsername(m_name)

End Property

Public Property Let Name(ByVal str As String)

    ' ...
    m_name = str

End Property

' ...
Public Property Get DisplayName() As String

    ' ...
    DisplayName = convertUsername(Name)

End Property

Public Property Get Operator() As String

    ' ...
    Operator = m_operator

End Property

Public Property Let Operator(ByVal str As String)

    ' ...
    m_operator = str

End Property

Public Property Get DateOfBan() As Date

    ' ...
    DateOfBan = m_banned_date

End Property

Public Property Let DateOfBan(ByVal BanDate As Date)

    ' ...
    m_banned_date = BanDate

End Property

Public Property Get Reason() As String

    ' ...
    Reason = m_ban_reason

End Property

Public Property Let Reason(ByVal str As String)

    ' ...
    m_ban_reason = str

End Property

Public Property Let IsDuplicateBan(ByVal bln As Boolean)

    ' ...
    m_duplicate_ban = bln

End Property

Public Property Get IsDuplicateBan() As Boolean

    ' ...
    IsDuplicateBan = m_duplicate_ban

End Property

Public Property Let IsActive(ByVal bln As Boolean)

    ' ...
    m_active_ban = bln

End Property

Public Property Get IsActive() As Boolean

    ' ...
    IsActive = m_active_ban

End Property

Public Function Clone() As Object

    ' ...
    Set Clone = New clsBanlistUserObj
    
    ' ...
    Clone.Name = Name
    Clone.DateOfBan = DateOfBan
    Clone.IsActive = IsActive
    Clone.IsDuplicateBan = IsDuplicateBan
    Clone.Operator = Operator
    Clone.Reason = Reason

End Function


