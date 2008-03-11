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
Private m_data_min       As Integer
Private m_data_max       As Integer
Private m_data_pattern   As String

Private Sub Class_Initialize()

    ' ...
    Set m_restrictions = New Collection

End Sub

Private Sub Class_Terminate()

    ' ...
    Set m_restrictions = Nothing

End Sub

Public Property Get restrictions() As Collection

    ' ...
    Set restrictions = m_restrictions

End Property

Public Property Get dataType() As String

    ' ...
    dataType = m_data_type

End Property

Public Property Let dataType(strType As String)

    ' ...
    m_data_type = strType

End Property

Public Property Get pattern() As String

    ' ...
    pattern = m_data_pattern

End Property

Public Property Let pattern(strPattern As String)

    ' ...
    m_data_pattern = strPattern

End Property

Public Property Get min() As Long

    ' ...
    min = m_data_min

End Property

Public Property Let min(val As Long)

    ' ...
    m_data_min = val

End Property

Public Property Get max() As Long

    ' ...
    max = m_data_max

End Property

Public Property Let max(val As Long)

    ' ...
    m_data_max = val

End Property

Public Property Get minLength() As Long

    ' ...

End Property

Public Property Let minLength(val As Long)

    ' ...

End Property

Public Property Get maxLength() As Long

    ' ...

End Property

Public Property Let maxLength(val As Long)

    ' ...

End Property

Public Property Get IsOptional() As Boolean

    ' ...
    IsOptional = m_optional

End Property

Public Property Let IsOptional(bln As Boolean)

    ' ...
    m_optional = bln

End Property

