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

Private m_type        As Integer
Private m_name        As String
Private m_rank        As Long
Private m_flags       As String
Private m_created_on  As Date
Private m_created_by  As String
Private m_modified_on As Date
Private m_modified_by As String
Private m_groups()    As clsDBEntryObj
Private m_lastseen    As Date

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

End Property

' ...
Public Property Let LastSeen(ByRef dLastSeen As Date)

End Property

' ...
Public Function Group(ByVal grp As Variant) As Object
    
End Function
