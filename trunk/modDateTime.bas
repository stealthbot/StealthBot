Attribute VB_Name = "modDateTime"
' modDateTime.bas
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

' ...
Private Declare Function GetSystemTime Lib "Kernel32.dll" () As SYSTEMTIME

' ...
Function UtcNow() As Date

    Dim SysTime As SYSTEMTIME ' ...
    Dim tempDate As Date      ' ...
    Dim tempTime As Date      ' ...
    
    ' ...
    SysTime = GetSystemTime()
    
    ' ...
    tempDate = DateSerial(SysTime.wYear, SysTime.wMonth, SysTime.wDay)
    tempTime = TimeSerial(SysTime.wHour, SysTime.wMinute, SysTime.wSecond)
    
    ' ...
    UtcNow = (tempDate + tempTime)
 
End Function

' ...
Public Function UtcToLocal(ByRef UtcDate As Date) As Date

    ' ...

End Function

' ...
Public Function FileTimeToDate(ByRef FTime As FILETIME)

    ' ...

End Function

' ...
Public Function SystemTimeToDate(ByRef STime As SYSTEMTIME)

    ' ...

End Function

