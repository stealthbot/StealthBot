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
Private m_name        As String
Private m_rank        As Long
Private m_flags       As String
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
    Name = m_name
End Property

Public Property Let Name(ByVal strName As String)
    m_name = strName
End Property

Public Property Get Rank() As Long
    Rank = m_rank
End Property

Public Property Let Rank(ByVal lngRank As Long)
    m_rank = lngRank
End Property

Public Property Get Flags() As String
    Flags = m_flags
End Property

Public Property Let Flags(ByVal strFlags As String)
    m_flags = strFlags
End Property

Public Function IsInGroup(ByVal sGroupName As String) As Boolean
    Dim i           As Integer
    Dim sCurGroup   As String
    
    IsInGroup = False
    
    For i = 1 To Groups.Count
        ' Internally all of the group objects are strings,
        '   but when used in scripting they are instances of clsDBEntryObj.
        If TypeOf Groups.Item(i) Is clsDBEntryObj Then
            sCurGroup = Groups.Item(i).Name
        Else
            sCurGroup = Split(Groups.Item(i), Space(1))(0)
        End If
        
        If (StrComp(sGroupName, sCurGroup, vbTextCompare) = 0) Then
            IsInGroup = True
            Exit Function
        End If
    Next i
End Function

Public Function HasFlag(ByVal strFlag As String, Optional ByVal bCaseSensitive As Boolean = True) As Boolean
    Dim cmCompare As CompareMethod
    cmCompare = IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare)
    
    HasFlag = (InStr(1, m_flags, strFlag, cmCompare) > 0)
End Function

Public Function HasAnyFlag(ByVal strFlags As String, Optional ByVal bCaseSensitive As Boolean = True) As Boolean
    Dim i As Integer
    
    HasAnyFlag = False
    
    For i = 1 To Len(strFlags)
        If HasFlag(Mid(strFlags, i, 1), bCaseSensitive) Then
            HasAnyFlag = True
            Exit Function
        End If
    Next i
End Function

Public Function HasFlags(ByVal strFlags As String, Optional ByVal bCaseSensitive As Boolean = True) As Boolean
    Dim i As Integer
    
    HasFlags = True
    
    For i = 1 To Len(strFlags)
        If Not HasFlag(Mid(strFlags, i, 1), bCaseSensitive) Then
            HasFlags = False
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
    
    If (StrComp(m_Type, DB_TYPE_GROUP, vbTextCompare) = 0) Then
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

Public Sub ClearGroups()
    Set m_Groups = New Collection
End Sub

' Returns a display string describing the entry.
Public Function ToString() As String
Attribute ToString.VB_UserMemId = 0
    ToString = Me.Name
    
    If Len(Me.EntryType) > 0 Then
        If StrComp(Me.EntryType, DB_TYPE_USER, vbTextCompare) <> 0 Then
            ToString = ToString & " (" & LCase(Me.EntryType) & ")"
        End If
    End If
End Function

' Creates a deep copy of the entry.
Public Function CreateCopy() As clsDBEntryObj
    Dim i As Integer
    
    Set CreateCopy = New clsDBEntryObj
    With CreateCopy
        .Name = Me.Name
        .Rank = Me.Rank
        .Flags = Me.Flags
        .EntryType = Me.EntryType
        .BanMessage = Me.BanMessage
        
        .CreatedBy = Me.CreatedBy
        .CreatedOn = Me.CreatedOn
        .ModifiedBy = Me.ModifiedBy
        .ModifiedOn = Me.ModifiedOn
        
        If Me.Groups.Count > 0 Then
            If TypeOf Me.Groups.Item(1) Is clsDBEntryObj Then
                For i = 1 To Me.Groups.Count
                    .Groups.Add Me.Groups.Item(i).CreateCopy(), Me.Groups.Item(i).Name
                Next i
            Else
                For i = 1 To Me.Groups.Count
                    .Groups.Add Me.Groups.Item(i), Me.Groups.Item(i)
                Next i
            End If
        End If
        
        .LastSeen = Me.LastSeen
    End With
End Function

