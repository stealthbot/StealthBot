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

Option Explicit

Private m_name           As String
Private m_req_rank       As Integer
Private m_req_flags      As String
Private m_matchmessage   As String
Private m_casesensitive  As Boolean
Private m_error          As String
Private m_description    As String
Private m_special_notes  As String
Private m_fatal          As Boolean

Public Property Get Name() As String
    Name = m_name
End Property

Public Property Let Name(strName As String)
    m_name = strName
End Property

Public Property Get RequiredRank() As Integer
    RequiredRank = m_req_rank
End Property

Public Property Let RequiredRank(Rank As Integer)
    m_req_rank = Rank
End Property

Public Property Get RequiredFlags() As String
    RequiredFlags = m_req_flags
End Property

Public Property Let RequiredFlags(Flags As String)
    m_req_flags = Flags
End Property

Public Property Get MatchMessage() As String
    MatchMessage = m_matchmessage
End Property

Public Property Let MatchMessage(str As String)
    m_matchmessage = str
End Property

Public Property Get MatchError() As String
    MatchError = m_error
End Property

Public Property Let MatchError(str As String)
    m_error = str
End Property

Public Property Get MatchCaseSensitive() As Boolean
    MatchCaseSensitive = m_casesensitive
End Property

Public Property Let MatchCaseSensitive(bln As Boolean)
    m_casesensitive = bln
End Property

Public Property Get Description() As String
    Description = m_description
End Property

Public Property Let Description(str As String)
    m_description = str
End Property

Public Property Get SpecialNotes() As String
    SpecialNotes = m_special_notes
End Property

Public Property Let SpecialNotes(str As String)
    m_special_notes = str
End Property

Public Property Let Fatal(bFatal As Boolean)
    m_fatal = bFatal
End Property

Public Property Get Fatal() As Boolean
    Fatal = m_fatal
End Property

Public Function ToString() As String
Attribute ToString.VB_UserMemId = 0
    ToString = "Restriction: " & Name
End Function
