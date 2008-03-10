VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommandRestrictionObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsCommandRestrictionObj.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Private m_name      As String
Private m_message   As String
Private m_req_rank  As Integer
Private m_req_flags As String

Public Property Get Name() As String

    ' ...
    Name = m_name

End Property

Public Property Let Name(strName As String)

    ' ...
    m_name = strName

End Property

Public Property Get RequiredRank() As Integer

    ' ...
    RequiredRank = m_req_rank

End Property

Public Property Let RequiredRank(rank As Integer)

    ' ...
    m_req_rank = rank

End Property

Public Property Get RequiredFlags() As String

    ' ...
    RequiredFlags = m_req_flags

End Property

Public Property Let RequiredFlags(flags As String)

    ' ...
    m_req_flags = flags

End Property

Public Property Get MatchMessage() As String

    ' ...
    MatchMessage = m_message

End Property

Public Property Let MatchMessage(strMessage As String)

    ' ...
    m_message = strMessage

End Property


