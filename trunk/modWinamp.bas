Attribute VB_Name = "modWinamp"
' modWinamp.bas
' Copyright (C) 2008 Eric Evans
' ...

Public g_find_file      As String  ' ...
Public g_find_file_done As Boolean ' ...

Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As String
End Type

Private Type FILEINFO
    file  As String
    index As Long
End Type

Option Explicit

Public Function WndProc()

End Function
