VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBEntryObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsDBEntryObj.cls
' Copyright (C) 2008 Eric Evans

Option Explicit

Private m_Type        As String
Private m_Name        As String
Private m_Rank        As Long
Private m_Flags       As String
Private m_CreatedOn  As Date
Private m_CreatedBy  As String
Private m_ModifiedOn As Date
Private m_ModifiedBy As String
Private m_Groups      As Collection
Private m_LastSeen    As Date
Private m_BanMessage As String

Private Sub Class_Initialize()
    Set m_Groups = New Collection
End Sub

Public Property Get EntryType() As String
    EntryType = m_Type
End Property

Public Property Let EntryType(ByVal strType As String)
    m_Type = strType
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(ByVal strName As String)
    m_Name = strName
End Property

Public Property Get Rank() As Long
    Rank = m_Rank
End Property

Public Property Let Rank(ByVal lngRank As Long)
    m_Rank = lngRank
End Property

Public Property Get Flags() As String
    Flags = m_Flags
End Property

Public Property Let Flags(ByVal strFlags As String)
    m_Flags = strFlags
End Property

Public Function IsInGroup(ByVal GroupName As String) As Boolean
    Dim i   As Integer
    Dim pos As Integer
    
    For i = 1 To Groups.Count
        pos = InStr(1, Groups(i).Name, Space$(1), vbBinaryCompare)
        
        If (pos > 0) Then
            Groups(i).Name = Mid$(Groups(i).Name, 1, pos - 1)
        End If
    
        If (StrComp(GroupName, Groups(i).Name, vbTextCompare) = 0) Then
            IsInGroup = True
            
            Exit Function
        End If
    Next i
End Function

Public Function HasFlag(ByVal strFlag As String, Optional ByVal CaseSensitive As Boolean = True) As Boolean
    If (CaseSensitive) Then
        HasFlag = (InStr(1, m_flags, strFlag, vbBinaryCompare) <> 0)
    Else
        HasFlag = (InStr(1, m_flags, strFlag, vbTextCompare) <> 0)
    End If
End Function

Public Function HasAnyFlag(ByVal strFlags As String, Optional ByVal CaseSensitive As Boolean = True) As Boolean
    Dim i As Integer
    
    For i = 1 To Len(strFlags)
        If (CaseSensitive) Then
            HasAnyFlag = (InStr(1, m_flags, Mid$(strFlags, i, 1), vbBinaryCompare) <> 0)
        Else
            HasAnyFlag = (InStr(1, m_flags, Mid$(strFlags, i, 1), vbTextCompare) <> 0)
        End If
        
        If (HasAnyFlag) Then
            Exit Function
        End If
    Next i
End Function

Public Function HasFlags(ByVal strFlags As String, Optional ByVal CaseSensitive As Boolean = True) As Boolean
    Dim i As Integer
    
    For i = 1 To Len(strFlags)
        If (CaseSensitive) Then
            HasFlags = (InStr(1, m_flags, Mid$(strFlags, i, 1), vbBinaryCompare) <> 0)
        Else
            HasFlags = (InStr(1, m_flags, Mid$(strFlags, i, 1), vbTextCompare) <> 0)
        End If
        
        If (HasFlags = False) Then
            Exit Function
        End If
    Next i
End Function

Public Property Get CreatedOn() As Date
    CreatedOn = m_CreatedOn
End Property

Public Property Let CreatedOn(ByRef dCreatedOn As Date)
    m_CreatedOn = dCreatedOn
End Property

Public Property Get CreatedBy() As String
    CreatedBy = m_CreatedBy
End Property

Public Property Let CreatedBy(ByVal sCreatedBy As String)
    m_CreatedBy = sCreatedBy
End Property

Public Property Get ModifiedOn() As Date
    ModifiedOn = m_ModifiedOn
End Property

Public Property Let ModifiedOn(ByRef dModifiedOn As Date)
    m_ModifiedOn = dModifiedOn
End Property

Public Property Get ModifiedBy() As String
    ModifiedBy = m_ModifiedBy
End Property

Public Property Let ModifiedBy(ByVal sModifiedBy As String)
    m_ModifiedBy = sModifiedBy
End Property

Public Property Get LastSeen() As Date
    LastSeen = m_LastSeen
End Property

Public Property Let LastSeen(ByRef dLastSeen As Date)
    m_LastSeen = dLastSeen
End Property

Public Property Get Groups() As Collection
    Set Groups = m_Groups
End Property

'What exactly is this function suposto do?
'To-Do: Complete function
Public Property Get MembersOf() As Collection
    Set MembersOf = New Collection
    
    If (StrComp(m_type, "Group", vbTextCompare) = 0) Then
        'Do something..
    End If
End Property

Public Property Get BanMessage() As String
    BanMessage = m_BanMessage
End Property

Public Property Let BanMessage(ByVal sBanMessage As String)
  m_BanMessage = sBanMessage
End Property

Public Sub AddGroup(ByVal sGroup As String)
    m_Groups.Add sGroup, sGroup
End Sub
