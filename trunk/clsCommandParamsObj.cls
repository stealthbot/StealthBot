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

Option Explicit

Private m_Name           As String
Private m_Optional       As Boolean
Private m_RequiredRank   As Integer
Private m_RequiredFlags  As String
Private m_Description    As String
Private m_SpecialNotes   As String
Private m_Restrictions   As Collection
Private m_DataType       As String
Private m_MatchMessage   As String
Private m_CaseSensitive  As Boolean
Private m_Error          As String
Private m_RSError        As String

Private Sub Class_Initialize()
    Set m_Restrictions = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Restrictions = Nothing
End Sub


Public Function GetRestrictionByName(ByVal sRestrictionName As String) As clsCommandRestrictionObj
    Dim r As clsCommandRestrictionObj
    Dim col As Collection
    Dim i As Integer
    
    Set col = Me.Restrictions
        
    For i = 1 To col.Count
        Set r = col(i)
        If StrComp(sRestrictionName, r.Name, vbTextCompare) = 0 Then
            Set GetRestrictionByName = r
            Exit Function
        End If
    Next i
End Function

Public Property Get Restrictions() As Collection
    Set Restrictions = m_Restrictions
End Property

Public Property Let Restrictions(col As Collection)
    Set m_Restrictions = col
End Property

Public Property Get DataType() As String
    DataType = m_DataType
End Property

Public Property Let DataType(strType As String)
    Select Case LCase$(strType)
        Case "string"
        Case "number", "numeric"
        Case "word"
        Case Else
            '// default to string
            strType = "string"
    End Select
    m_DataType = strType
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(str As String)
    m_Name = str
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

Public Property Get IsOptional() As Boolean
    IsOptional = m_Optional
End Property

Public Property Let IsOptional(bln As Boolean)
    m_Optional = bln
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

Public Property Get RestrictionsSharedError() As String
    RestrictionsSharedError = m_RSError
End Property

Public Property Let RestrictionsSharedError(str As String)
    m_RSError = str
End Property

Public Property Get MatchCaseSensitive() As Boolean
    MatchCaseSensitive = m_CaseSensitive
End Property

Public Property Let MatchCaseSensitive(bln As Boolean)
    m_CaseSensitive = bln
End Property

Public Function ToString(Optional ByVal ShowLong As Boolean = False) As String
Attribute ToString.VB_UserMemId = 0

    Dim BracketS As String, BracketE As String
    BracketS = "<": BracketE = ">"
    If Me.IsOptional Then
        BracketS = "[": BracketE = "]"
    End If
    If ShowLong Then
        ToString = StringFormat("{0} ({1} {2})", Me.Name, IIf(Me.IsOptional, "optional", "required"), LCase$(Me.DataType))
    Else
        ToString = StringFormat("{0}{1}{2}{3}", BracketS, Me.Name, BracketE, IIf(LCase$(Me.DataType) = "string", "*", vbNullString))
    End If

End Function

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
