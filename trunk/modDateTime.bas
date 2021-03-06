Attribute VB_Name = "modDateTime"
' modDateTime.bas
' Copyright (C) 2008 Eric Evans

Option Explicit

Public Declare Function GetSystemTime Lib "Kernel32.dll" () As SYSTEMTIME
Public Declare Function FileTimeToSystemTime Lib "Kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function SystemTimeToFileTime Lib "Kernel32.dll" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Public Declare Function FileTimeToLocalFileTime Lib "Kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Public Declare Function GetTimeZoneInformation Lib "Kernel32.dll" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public Declare Function GetTickCount Lib "Kernel32.dll" () As Long
Public Declare Function GetTickCount64 Lib "Kernel32.dll" () As Currency
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

Public Function GetTimeZoneName() As String
    Dim TZinfo   As TIME_ZONE_INFORMATION   ' holds time zone info
    Dim lngL     As Long                    ' time zone info result
    Dim i        As Long                    ' counter
    Dim b        As Integer                 ' a single character from the time zone name
    Dim str      As String                  ' return string

    ' Get the time zone info
    lngL = GetTimeZoneInformation(TZinfo)
    
    ' Convert the name
    For i = 0 To 31
    
        ' Standard or daylight savings?
        If lngL = TIME_ZONE_ID_DAYLIGHT Then
            b = TZinfo.DaylightName(i)
        Else
            b = TZinfo.StandardName(i)
        End If
        
        ' Name is null-padded
        If b = 0 Then
            Exit For
        Else
            str = str & Chr(b)
        End If
    Next
    
    GetTimeZoneName = str
End Function

Public Function GetTickCountMS() As Currency
    GetTickCountMS = (GetTickCount64() * 10000)
End Function

Public Function GetTickCountS() As Long
    GetTickCountS = CLng(GetTickCount64() * 10)
End Function

Public Function Mod64(ByVal Value1 As Currency, ByVal Value2 As Currency) As Currency
    Dim x As Currency, y As Currency
    x = Abs(Value1)
    y = Abs(Value2)
    Mod64 = x - Int(x / y) * y
    If Value1 < 0 Then
        Mod64 = Mod64 * -1
    End If
End Function

'// Converts a millisecond or second time value to humanspeak.. modified to support BNet's Time
Public Function ConvertTimeInterval(ByVal MS As Currency, Optional ByVal IsSeconds As Boolean = False, Optional ByVal PrintMS As Boolean = False) As String
    Dim Seconds   As Currency
    Dim Minutes   As Currency
    Dim Hours     As Currency
    Dim Days      As Currency
    Dim Years     As Currency

    Dim Parts(6)  As String
    Dim sPlural   As String
    Dim sAnd      As String
    Dim sComma    As String

    Dim PartCount As Integer
    Dim i         As Integer
    Dim PartLB    As Integer

    If (IsSeconds) Then
        Seconds = MS
        MS = 0
    ElseIf PrintMS Then
        Seconds = Int(MS / 1000)
        MS = Mod64(MS, 1000)
    Else
        Seconds = Round(MS / 1000)
        MS = 0
    End If

    Years = Int(Seconds / 31536000)
    Seconds = Mod64(Seconds, 31536000)

    If Years > 0 Then
        sPlural = "s"
        If Years = 1 Then sPlural = vbNullString
        Parts(0) = StringFormat("{0} year{1}", Years, sPlural)
    End If

    Days = Int(Seconds / 86400)
    Seconds = Mod64(Seconds, 86400)

    If Days > 0 Then
        sPlural = "s"
        If Days = 1 Then sPlural = vbNullString
        Parts(1) = StringFormat("{0} day{1}", Days, sPlural)
    End If

    Hours = Int(Seconds / 3600)
    Seconds = Mod64(Seconds, 3600)

    If Hours > 0 Then
        sPlural = "s"
        If Hours = 1 Then sPlural = vbNullString
        Parts(2) = StringFormat("{0} hour{1}", Hours, sPlural)
    End If

    Minutes = Int(Seconds / 60)
    Seconds = Mod64(Seconds, 60)

    If Minutes > 0 Then
        sPlural = "s"
        If Minutes = 1 Then sPlural = vbNullString
        Parts(3) = StringFormat("{0} minute{1}", Minutes, sPlural)
    End If

    If Seconds > 0 Or (Years = 0 And Days = 0 And Hours = 0 And Minutes = 0 And MS = 0 And Not PrintMS) Then
        sPlural = "s"
        If Seconds = 1 Then sPlural = vbNullString
        Parts(4) = StringFormat("{0} second{1}", Seconds, sPlural)
    End If

    If MS > 0 Or (Years = 0 And Days = 0 And Hours = 0 And Minutes = 0 And Seconds = 0 And PrintMS) Then
        sPlural = "s"
        If MS = 1 Then sPlural = vbNullString
        Parts(5) = StringFormat("{0} millisecond{1}", MS, sPlural)
    End If

    PartCount = 0
    PartLB = -1
    For i = LBound(Parts) To UBound(Parts)
        If Len(Parts(i)) > 0 Then
            If PartLB < 0 Then PartLB = i
            PartCount = PartCount + 1
        End If
    Next i
    If PartCount = 2 Then
        sComma = " "
        sAnd = "and "
    ElseIf PartCount > 2 Then
        sComma = ", "
        sAnd = "and "
    End If
    For i = UBound(Parts) To LBound(Parts) Step -1
        If Len(Parts(i)) > 0 And i <> PartLB Then
            Parts(i) = sComma & sAnd & Parts(i)
            sAnd = vbNullString
        End If
    Next i

    ConvertTimeInterval = StringFormat("{0}{1}{2}{3}{4}{5}", _
            Parts(0), Parts(1), Parts(2), Parts(3), Parts(4), Parts(5))
End Function

