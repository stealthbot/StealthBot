VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBannedUserObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsBannedUserObj.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Private m_name          As String
Private m_banned_date   As Date
Private m_ban_reason    As String
Private m_operator      As String
Private m_duplicate_ban As Boolean

Public Property Get Name() As String

    ' ...
    Name = m_name

End Property

Public Property Let Name(str As String)

    ' ...
    m_name = str

End Property

Public Property Get Operator() As String

    ' ...
    Operator = m_operator

End Property

Public Property Let Operator(str As String)

    ' ...
    m_operator = str

End Property

Public Property Get DateOfBan() As Date

    ' ...
    DateOfBan = m_banned_date

End Property

Public Property Let DateOfBan(BanDate As Date)

    ' ...
    m_banned_date = BanDate

End Property

Public Property Get Reason() As String

    ' ...
    Reason = m_ban_reason

End Property

Public Property Let Reason(str As String)

    ' ...
    m_ban_reason = str

End Property

Public Property Let IsDuplicateBan(bln As Boolean)

    ' ...
    m_duplicate_ban = bln

End Property

Public Property Get IsDuplicateBan() As Boolean

    ' ...
    IsDuplicateBan = m_duplicate_ban

End Property


