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
' ...

Option Explicit

Private m_type        As String
Private m_name        As String
Private m_rank        As Long
Private m_flags       As String
Private m_created_on  As Date
Private m_created_by  As String
Private m_modified_on As Date
Private m_modified_by As String
Private m_groups      As Collection
Private m_lastseen    As Date

' ...
Private Sub Class_Initialize()

    Set m_groups = New Collection

End Sub

' ...
Public Property Get entryType() As String

    entryType = m_type

End Property

' ...
Public Property Let entryType(ByVal strType As String)
    
    m_type = strType

End Property

' ...
Public Property Get Name() As String
    
    Name = m_name

End Property

' ...
Public Property Let Name(ByVal strName As String)
    
    m_name = strName

End Property

' ...
Public Property Get Rank() As Long
    
    Rank = m_rank

End Property

' ...
Public Property Let Rank(ByVal lngRank As Long)
    
    m_rank = lngRank

End Property

' ...
Public Property Get Flags() As String
    
    Flags = m_flags

End Property

' ...
Public Property Let Flags(ByVal strFlags As String)
    
    m_flags = strFlags

End Property

' ...
Public Function IsInGroup(ByVal GroupName As String) As Boolean

    Dim i   As Integer ' ...
    Dim pos As Integer ' ...
    
    ' ...
    For i = 1 To Groups.Count
        ' ...
        pos = InStr(1, Groups(i).Name, Space$(1), vbBinaryCompare)
        
        ' ...
        If (pos > 0) Then
            Groups(i).Name = Mid$(Groups(i).Name, 1, pos - 1)
        End If
    
        ' ...
        If (StrComp(GroupName, Groups(i).Name, vbTextCompare) = 0) Then
            ' ...
            IsInGroup = True
            
            ' ...
            Exit Function
        End If
    Next i

End Function

' ...
Public Function HasFlag(ByVal strFlag As String, Optional ByVal CaseSensitive As Boolean = True) As Boolean
    
    ' ...
    If (CaseSensitive) Then
        HasFlag = (InStr(1, m_flags, strFlag, vbBinaryCompare) <> 0)
    Else
        HasFlag = (InStr(1, m_flags, strFlag, vbTextCompare) <> 0)
    End If

End Function ' end function HasFlag

' ...
Public Function HasAnyFlag(ByVal strFlags As String, Optional ByVal CaseSensitive As Boolean = True) As Boolean
    
    Dim i As Integer ' ...
    
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
    
End Function ' end function HasAnyFlag

' ...
Public Function HasFlags(ByVal strFlags As String, Optional ByVal CaseSensitive As Boolean = True) As Boolean
    
    Dim i As Integer ' ...
    
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
    
End Function ' end function HasFlags

' ...
Public Property Get CreatedOn() As Date
    
    CreatedOn = m_created_on

End Property

' ...
Public Property Let CreatedOn(ByRef dCreatedOn As Date)
    
    m_created_on = dCreatedOn

End Property

' ...
Public Property Get CreatedBy() As String
    
    CreatedBy = m_created_by

End Property

' ...
Public Property Let CreatedBy(ByVal sCreatedBy As String)
    
    m_created_by = sCreatedBy

End Property

' ...
Public Property Get ModifiedOn() As Date
    
    ModifiedOn = m_modified_on

End Property

' ...
Public Property Let ModifiedOn(ByRef dModifiedOn As Date)
    
    m_modified_on = dModifiedOn

End Property

' ...
Public Property Get ModifiedBy() As String
    
    ModifiedBy = m_modified_by

End Property

' ...
Public Property Let ModifiedBy(ByVal sModifiedBy As String)
    
    m_modified_by = sModifiedBy

End Property

' ...
Public Property Get LastSeen() As Date

    ' ...
    
End Property

' ...
Public Property Let LastSeen(ByRef dLastSeen As Date)

    ' ...

End Property

' ...
Public Property Get Groups() As Collection
    
    Set Groups = m_groups
    
End Property

' ...
Public Property Get MembersOf() As Collection
    ' ...
    Set MembersOf = New Collection
    
    ' ...
    If (StrComp(m_type, "Group", vbTextCompare) = 0) Then

        ' ...
        
    End If
End Property
