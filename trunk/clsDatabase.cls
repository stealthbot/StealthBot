VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_DB()    As DATABASE ' ...
Private m_dbCount As Integer  ' ...

' ...
Public Function load() As Boolean
    ' ...
End Function

' ...
Public Function create_entry(ByVal dbeType As DB_ENTRY_TYPE) As Integer
    ' ...
    
    Select Case (dbeType)
        Case DB_ENTRY_TYPE.TYPE_USER:  ' User
        Case DB_ENTRY_TYPE.TYPE_CLAN:  ' Clan
        Case DB_ENTRY_TYPE.TYPE_GAME:  ' Game
        Case DB_ENTRY_TYPE.TYPE_GROUP: ' Group
    End Select
End Function

' ...
Public Function modify_entry(ByVal dbeType As DB_ENTRY_TYPE) As Boolean
    ' ...
    
    Select Case (dbeType)
        Case DB_ENTRY_TYPE.TYPE_USER:  ' User
        Case DB_ENTRY_TYPE.TYPE_CLAN:  ' Clan
        Case DB_ENTRY_TYPE.TYPE_GAME:  ' Game
        Case DB_ENTRY_TYPE.TYPE_GROUP: ' Group
    End Select
End Function

' ...
Public Function delete_entry(ByVal dbeType As DB_ENTRY_TYPE) As Boolean
    ' ...
    
    Select Case (dbeType)
        Case DB_ENTRY_TYPE.TYPE_USER:  ' User
        Case DB_ENTRY_TYPE.TYPE_CLAN:  ' Clan
        Case DB_ENTRY_TYPE.TYPE_GAME:  ' Game
        Case DB_ENTRY_TYPE.TYPE_GROUP: ' Group
    End Select
End Function
