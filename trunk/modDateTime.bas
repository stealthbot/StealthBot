Attribute VB_Name = "modDateTime"
' modDateTime.bas
' Copyright (C) 2008 Eric Evans

Option Explicit

Public Declare Function GetSystemTime Lib "Kernel32.dll" () As SYSTEMTIME
Public Declare Function FileTimeToSystemTime Lib "Kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function SystemTimeToFileTime Lib "Kernel32.dll" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Public Declare Function FileTimeToLocalFileTime Lib "Kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Public Declare Function GetTimeZoneInformation Lib "Kernel32.dll" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public Declare Function timeGetSystemTime Lib "winmm.dll" (lpTime As MMTIME, ByVal uSize As Long) As Long
Public Declare Function GetTickCount Lib "Kernel32.dll" () As Long
Public Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Private Const TIME_ZONE_ID_UNKNOWN = 0
Private Const TIME_ZONE_ID_STANDARD = 1
Private Const TIME_ZONE_ID_DAYLIGHT = 2

Public Type FILETIME
    dwLowDateTime   As Long
    dwHighDateTime  As Long
End Type

Public Type SYSTEMTIME
    wYear             As Integer
    wMonth            As Integer
    wDayOfWeek        As Integer
    wDay              As Integer
    wHour             As Integer
    wMinute           As Integer
    wSecond           As Integer
    wMilliseconds     As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias                   As Long
    StandardName(0 To 31)  As Integer
    StandardDate           As SYSTEMTIME
    StandardBias           As Long
    DaylightName(0 To 31)  As Integer
    DaylightDate           As SYSTEMTIME
    DaylightBias           As Long
End Type

Public Type SMPTE
    hour        As Byte
    min         As Byte
    sec         As Byte
    frame       As Byte
    fps         As Byte
    dummy       As Byte
    pad(2)      As Byte
End Type

Public Type MMTIME
    wType       As Long
    units       As Long
    smpteVal    As SMPTE
    songPtrPos  As Long
End Type

Public Function UtcNow() As Date
    UtcNow = SystemTimeToDate(GetSystemTime())
End Function

Public Function UtcToLocal(ByRef UtcDate As Date) As Date
    Dim FTime As FILETIME
    
    FTime = DateToFileTime(UtcDate)
    
    FileTimeToLocalFileTime FTime, FTime
    
    UtcToLocal = FileTimeToDate(FTime)
End Function

Public Function FileTimeToDate(ByRef FTime As FILETIME) As Date
    Dim STime As SYSTEMTIME
    FileTimeToSystemTime FTime, STime
    
    FileTimeToDate = SystemTimeToDate(STime)
End Function

Public Function DateToFileTime(ByRef DDate As Date) As FILETIME
    Dim STime As SYSTEMTIME
    
    STime = DateToSystemTime(DDate)
    SystemTimeToFileTime STime, DateToFileTime
End Function

Public Function SystemTimeToDate(ByRef STime As SYSTEMTIME) As Date
    Dim tempDate As Date
    Dim tempTime As Date 
    tempDate = DateSerial(STime.wYear, STime.wMonth, STime.wDay)
    tempTime = TimeSerial(STime.wHour, STime.wMinute, STime.wSecond)
    
    SystemTimeToDate = (tempDate + tempTime)
End Function

Public Function DateToSystemTime(ByRef DDate As Date) As SYSTEMTIME
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

Public Function GetTimeZoneBias() As Long
    Dim TZinfo As TIME_ZONE_INFORMATION
    Dim lngL   As Long
    
    lngL = GetTimeZoneInformation(TZinfo)

    Select Case (lngL)
        Case TIME_ZONE_ID_STANDARD
            GetTimeZoneBias = (TZinfo.Bias + TZinfo.StandardBias)
            
        Case TIME_ZONE_ID_DAYLIGHT
            GetTimeZoneBias = (TZinfo.Bias + TZinfo.DaylightBias)
            
        Case Else
            GetTimeZoneBias = TZinfo.Bias
    End Select    
End Function

