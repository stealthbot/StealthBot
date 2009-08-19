VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommandParamsObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsCommandParamsObj.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Private m_name           As String
Private m_optional       As Boolean
Private m_required_rank  As Integer
Private m_required_flags As String
Private m_description    As String
Private m_special_notes  As String
Private m_restrictions   As Collection
Private m_data_type      As String
Private m_matchmessage   As String
Private m_casesensitive  As Boolean
Private m_error          As String

Private Sub Class_Initialize()
    
    Set m_restrictions = New Collection
    
End Sub

Private Sub Class_Terminate()
    
    Set m_restrictions = Nothing

End Sub

Public Property Get Restrictions() As Collection
    Set Restrictions = m_restrictions
End Property
Public Property Let Restrictions(col As Collection)
    Set m_restrictions = col
End Property
Public Property Get DataType() As String
    DataType = m_data_type
End Property
Public Property Let DataType(strType As String)
    Select Case LCase(strType)
        Case "string"
        Case "number", "numeric"
        Case "word"
        Case Else
            '// default to string
            strType = "string"
    End Select
    m_data_type = strType
End Property
Public Property Get Name() As String
    Name = m_name
End Property
Public Property Let Name(str As String)
    m_name = str
End Property
Public Property Get description() As String
    description = m_description
End Property
Public Property Let description(str As String)
    m_description = str
End Property
Public Property Get SpecialNotes() As String
    SpecialNotes = m_special_notes
End Property
Public Property Let SpecialNotes(str As String)
    m_special_notes = str
End Property
Public Property Get IsOptional() As Boolean
    IsOptional = m_optional
End Property
Public Property Let IsOptional(bln As Boolean)
    m_optional = bln
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


'Public Property Get Pattern() As String
'    Pattern = m_data_pattern
'End Property
'Public Property Let Pattern(strPattern As String)
'    m_data_pattern = strPattern
'End Property
'Public Property Get min() As Long
'    min = m_data_min
'End Property
'Public Property Let min(Val As Long)
'    m_data_min = Val
'End Property
'Public Property Get Max() As Long
'    Max = m_data_max
'End Property
'Public Property Let Max(Val As Long)
'    m_data_max = Val
'End Property
