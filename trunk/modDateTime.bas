Attribute VB_Name = "modDateTime"
' modDateTime.bas
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

' ...
Private Declare Function GetSystemTime Lib "Kernel32.dll" () As SYSTEMTIME
Private Declare Function FileTimeToSystemTime Lib "Kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "Kernel32.dll" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "Kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long

' ...
Public Function UtcNow() As Date

    ' ...
    UtcNow = SystemTimeToDate(GetSystemTime())
 
End Function

' ...
Public Function UtcToLocal(ByRef UtcDate As Date) As Date

    Dim FTime As FILETIME ' ...
    
    ' ...
    FTime = DateToFileTime(UtcDate)
    
    ' ...
    FileTimeToLocalFileTime FTime, FTime
    
    ' ...
    UtcToLocal = FileTimeToDate(FTime)

End Function

' ...
Public Function FileTimeToDate(ByRef FTime As FILETIME) As Date

    Dim STime As SYSTEMTIME ' ...

    ' ...
    FileTimeToSystemTime FTime, STime
    
    ' ...
    FileTimeToDate = SystemTimeToDate(STime)

End Function

' ...
Public Function DateToFileTime(ByRef DDate As Date) As FILETIME

    Dim STime As SYSTEMTIME ' ...
    
    ' ...
    STime = DateToSystemTime(DDate)

    ' ...
    SystemTimeToFileTime STime, DateToFileTime

End Function

' ...
Public Function SystemTimeToDate(ByRef STime As SYSTEMTIME) As Date

    Dim tempDate As Date ' ...
    Dim tempTime As Date ' ...

    ' ...
    tempDate = DateSerial(STime.wYear, STime.wMonth, STime.wDay)
    tempTime = TimeSerial(STime.wHour, STime.wMinute, STime.wSecond)
    
    ' ...
    SystemTimeToDate = (tempDate + tempTime)

End Function

' ...
Public Function DateToSystemTime(ByRef DDate As Date) As SYSTEMTIME

    ' ...
    With DateToSystemTime
        .wYear = DatePart("yyyy", DDate)
        .wMonth = DatePart("m", DDate)
        .wDay = DatePart("d", DDate)
        .wDayOfWeek = DatePart("w", DDate)
        .wHour = DatePart("h", DDate)
        .wMinute = DatePart("n", DDate)
        .wSecond = DatePart("s", DDate)
    End With

End Function

