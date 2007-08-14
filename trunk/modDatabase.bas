Attribute VB_Name = "Module1"
'project StealthBot - modDatabase
' Version: 2
' December 2005 by Stealth

Option Explicit

Public gDB As Collection

Public Type udtExtendedUsername
    Username    As String * 40
    Occurrence  As Date
End Type

Public Type udtDatabaseEntry
    Username    As String * 40
    Flags       As String * 26
    Added       As udtExtendedUsername
    Modified    As udtExtendedUsername
End Type

Public Sub LoadDatabase(Optional ByVal sPath As String)
    ' 0: CHECK FOR OLD-STYLE DATABASES AND IMPORT
    
    ' 1: VERIFY THE DATABASE FILE EXISTS, ELSE CREATE IT
    If LenB(sPath) = 0 Then
    
    End If
    
    ' 2: PULL DATA OUT OF THE DATABASE
    
End Sub

Public Sub ImportOld()

End Sub
