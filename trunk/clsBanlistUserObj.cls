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

Option Explicit

Private m_Name         As String
Private m_BannedDate   As Date
Private m_BanReason    As String
Private m_Operator     As String
Private m_DuplicateBan As Boolean
Private m_ActiveBan    As Boolean

Private Sub Class_Initialize()
    m_ActiveBan = True
End Sub

Public Property Get Name() As String
    Name = modEvents.CleanUsername(m_Name)
End Property

Public Property Let Name(ByVal str As String)
    m_Name = str
End Property

Public Property Get DisplayName() As String
    DisplayName = ConvertUsername(Name)
End Property

Public Property Get Operator() As String
    Operator = m_Operator
End Property

Public Property Let Operator(ByVal str As String)
    m_Operator = str
End Property

Public Property Get DateOfBan() As Date
    DateOfBan = m_BannedDate
End Property

Public Property Let DateOfBan(ByVal BanDate As Date)
    m_BannedDate = BanDate
End Property

Public Property Get Reason() As String
    Reason = m_BanReason
End Property

Public Property Let Reason(ByVal str As String)
    m_BanReason = str
End Property

Public Property Let IsDuplicateBan(ByVal bln As Boolean)
    m_DuplicateBan = bln
End Property

Public Property Get IsDuplicateBan() As Boolean
    IsDuplicateBan = m_DuplicateBan
End Property

Public Property Let IsActive(ByVal bln As Boolean)
    m_ActiveBan = bln
End Property

Public Property Get IsActive() As Boolean
    IsActive = m_ActiveBan
End Property

Public Function Clone() As Object
    Set Clone = New clsBanlistUserObj
    
    Clone.Name = Name
    Clone.DateOfBan = DateOfBan
    Clone.IsActive = IsActive
    Clone.IsDuplicateBan = IsDuplicateBan
    Clone.Operator = Operator
    Clone.Reason = Reason
End Function
