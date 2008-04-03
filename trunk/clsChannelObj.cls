VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChannelObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsChannelUserObj.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Private Const CHAN_PUBLIC     As Long = &H1
Private Const CHAN_MODERATED  As Long = &H2
Private Const CHAN_RESTRICTED As Long = &H4
Private Const CHAN_SILENT     As Long = &H8
Private Const CHAN_SYSTEM     As Long = &H10
Private Const CHAN_PRODUCT    As Long = &H20
Private Const CHAN_GLOBAL     As Long = &H1000

Private m_name            As String
Private m_flags           As Long
Private m_designated_heir As Integer
Private m_num_joins       As Long
Private m_num_bans        As Long
Private m_num_kicks       As Long
Private m_join_date       As Date
Private m_users           As Collection
Private m_banlist         As Collection
Private m_chat_queue      As Collection

Private Sub Class_Initialize()
    
    ' ...
    Set m_users = New Collection
    Set m_banlist = New Collection
    
End Sub

Public Property Get Name() As String

    ' ...
    Name = m_name

End Property

Public Property Let Name(str As String)

    ' ...
    m_name = str

End Property

Public Property Get Flags() As Long

    ' ...
    Flags = m_flags

End Property

Public Property Let Flags(lng As Long)

    ' ...
    m_flags = lng

End Property

Public Property Let JoinDate(jdate As Date)

    ' ...
    m_join_date = jdate

End Property

Public Property Get JoinDate() As Date

    ' ...
    JoinDate = m_join_date

End Property

Public Property Get IsSilent() As Boolean

    ' ...
    IsSilent = ((m_flags And CHAN_SILENT) = CHAN_SILENT)

End Property

Public Property Get TotalJoinCount() As Long

    ' ...

End Property

Public Property Get TotalBanCount() As Long

    ' ...
    TotalBanCount = m_num_bans

End Property

Public Property Let TotalBanCount(lng As Long)

    ' ...
    m_num_bans = lng
    
End Property

Public Property Get TotalKickCount() As Long

    ' ...

End Property

Public Property Get Users() As Collection

    ' ...
    Set Users = m_users

End Property

Public Sub ClearUsers()

    ' ...
    Set m_users = New Collection

End Sub

Public Property Get Self() As clsChannelUserObj

    Dim i As Integer ' ...

    ' ...
    Set Self = New clsChannelUserObj
    
    ' ...
    For i = 1 To m_users.Count
        ' ...
        If (StrComp(m_users(i).DisplayName, CurrentUsername, vbBinaryCompare) = 0) Then
            ' ...
            Set Self = m_users(i)
            
            ' ...
            Exit Function
        End If
    Next i

End Property

Public Function GetUserIndexByName(ByVal Username As String) As Integer

    Dim i As Integer ' ...

    ' ...
    For i = 1 To m_users.Count
        ' ...
        If (StrComp(m_users(i).Name, Username, vbBinaryCompare) = 0) Then
            ' ...
            GetUserIndexByName = i
            
            ' ...
            Exit Function
        End If
    Next i

End Function

Public Function GetUserIndexByDisplayName(ByVal Username As String) As Integer

    Dim i As Integer ' ...

    ' ...
    For i = 1 To m_users.Count
        ' ...
        If (StrComp(m_users(i).DisplayName, Username, vbBinaryCompare) = 0) Then
            ' ...
            GetUserIndexByDisplayName = i
            
            ' ...
            Exit Function
        End If
    Next i

End Function

Public Property Get Banlist() As Collection
    
    ' ...
    Set Banlist = m_banlist

End Property

Public Function IsOnBanList(ByVal Username As String, Optional Operator As String = vbNullString) As Integer

    ' ...
    Dim i   As Integer ' ...
    Dim bln As Boolean ' ...
    
    ' ...
    For i = 1 To m_banlist.Count
        ' ...
        If (StrComp(m_banlist(i).Name, Username, vbTextCompare) = 0) Then
            ' ...
            If (Operator <> vbNullString) Then
                ' ...
                If (StrComp(m_banlist(i).Operator, Operator, vbTextCompare) = 0) Then
                    bln = True
                End If
            Else
                bln = True
            End If
        
            ' ...
            If (bln) Then
                IsOnBanList = i
                
                Exit Function
            End If
        End If
    Next i

    IsOnBanList = 0
End Function

Public Sub RemoveBansFromOperator(ByVal Username As String)

    Dim i   As Integer ' ...
    Dim bln As Boolean ' ...
    Dim pos As Integer ' ...

    ' ...
    Do
        ' ...
        bln = False
    
        ' ...
        For i = 1 To m_banlist.Count
            ' ...
            If (StrComp(m_banlist(i).Operator, Username, vbTextCompare) = 0) Then
                ' ...
                m_banlist.Remove i
                
                ' ...
                pos = IsOnBanList(Username)
                
                ' ...
                If (pos > 0) Then
                    Banlist(pos).IsDuplicateBan = False
                End If
                
                ' ...
                bln = True
                
                ' ...
                Exit For
            End If
        Next i
    Loop While (bln = True)

End Sub
