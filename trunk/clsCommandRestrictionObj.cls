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

Private m_Name           As String
Private m_RequiredRank   As Integer
Private m_RequiredFlags  As String
Private m_MatchMessage   As String
Private m_CaseSensitive  As Boolean
Private m_Error          As String
Private m_Description    As String
Private m_SpecialNotes   As String
Private m_Fatal          As Boolean

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(strName As String)
    m_Name = strName
End Property

Public Property Get RequiredRank() As Integer
    RequiredRank = m_RequiredRank
End Property

Public Property Let RequiredRank(Rank As Integer)
    m_RequiredRank = Rank
End Property

Public Property Get RequiredFlags() As String
    RequiredFlags = m_RequiredFlags
End Property

Public Property Let RequiredFlags(Flags As String)
    m_RequiredFlags = Flags
End Property

Public Property Get MatchMessage() As String
    MatchMessage = m_MatchMessage
End Property

Public Property Let MatchMessage(str As String)
    m_MatchMessage = str
End Property

Public Property Get MatchError() As String
    MatchError = m_Error
End Property

Public Property Let MatchError(str As String)
    m_Error = str
End Property

Public Property Get MatchCaseSensitive() As Boolean
    MatchCaseSensitive = m_CaseSensitive
End Property

Public Property Let MatchCaseSensitive(bln As Boolean)
    m_CaseSensitive = bln
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Description(str As String)
    m_Description = str
End Property

Public Property Get SpecialNotes() As String
    SpecialNotes = m_SpecialNotes
End Property

Public Property Let SpecialNotes(str As String)
    m_SpecialNotes = str
End Property

Public Property Let Fatal(bFatal As Boolean)
    m_Fatal = bFatal
End Property

Public Property Get Fatal() As Boolean
    Fatal = m_Fatal
End Property

Public Function ToString() As String
Attribute ToString.VB_UserMemId = 0
    ToString = "Restriction: " & Name
End Function
