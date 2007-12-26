Attribute VB_Name = "modDatabase"
Option Explicit

Public Type DATABASE
    Username    As String
    rank        As Integer
    Flags       As String
    AddedBy     As String
    AddedOn     As Date
    ModifiedBy  As String
    ModifiedOn  As Date
    Type        As String
    Groups      As String
    BanMessage  As String
End Type

Public Enum DB_ENTRY_TYPE
    TYPE_USER = 1
    TYPE_CLAN = 2
    TYPE_GAME = 3
    TYPE_GROUP = 4
End Enum
