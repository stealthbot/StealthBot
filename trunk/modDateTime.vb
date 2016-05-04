Option Strict Off
Option Explicit On
Module modDateTime
	' modDateTime.bas
	' Copyright (C) 2008 Eric Evans
	
	
	Public Declare Function GetSystemTime Lib "Kernel32.dll" () As SYSTEMTIME
	'UPGRADE_WARNING: Structure SYSTEMTIME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure FILETIME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function FileTimeToSystemTime Lib "Kernel32.dll" (ByRef lpFileTime As FILETIME, ByRef lpSystemTime As SYSTEMTIME) As Integer
	'UPGRADE_WARNING: Structure FILETIME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure SYSTEMTIME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function SystemTimeToFileTime Lib "Kernel32.dll" (ByRef lpSystemTime As SYSTEMTIME, ByRef lpFileTime As FILETIME) As Integer
	'UPGRADE_WARNING: Structure FILETIME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure FILETIME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function FileTimeToLocalFileTime Lib "Kernel32.dll" (ByRef lpFileTime As FILETIME, ByRef lpLocalFileTime As FILETIME) As Integer
	'UPGRADE_WARNING: Structure TIME_ZONE_INFORMATION may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function GetTimeZoneInformation Lib "Kernel32.dll" (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Integer
	'UPGRADE_WARNING: Structure MMTIME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function timeGetSystemTime Lib "winmm.dll" (ByRef lpTime As MMTIME, ByVal uSize As Integer) As Integer
	Public Declare Function GetTickCount Lib "Kernel32.dll" () As Integer
	'UPGRADE_WARNING: Structure SYSTEMTIME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Sub GetLocalTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)
	
	Private Const TIME_ZONE_ID_UNKNOWN As Short = 0
	Private Const TIME_ZONE_ID_STANDARD As Short = 1
	Private Const TIME_ZONE_ID_DAYLIGHT As Short = 2
	
	Public Structure FILETIME
		Dim dwLowDateTime As Integer
		Dim dwHighDateTime As Integer
	End Structure
	
	Public Structure SYSTEMTIME
		Dim wYear As Short
		Dim wMonth As Short
		Dim wDayOfWeek As Short
		Dim wDay As Short
		Dim wHour As Short
		Dim wMinute As Short
		Dim wSecond As Short
		Dim wMilliseconds As Short
	End Structure
	
	Private Structure TIME_ZONE_INFORMATION
		Dim Bias As Integer
		<VBFixedArray(31)> Dim StandardName() As Short
		Dim StandardDate As SYSTEMTIME
		Dim StandardBias As Integer
		<VBFixedArray(31)> Dim DaylightName() As Short
		Dim DaylightDate As SYSTEMTIME
		Dim DaylightBias As Integer
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim StandardName(31)
			ReDim DaylightName(31)
		End Sub
	End Structure
	
	Public Structure SMPTE
		'UPGRADE_NOTE: hour was upgraded to hour_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim hour_Renamed As Byte
		Dim min As Byte
		Dim sec As Byte
		Dim frame As Byte
		Dim fps As Byte
		Dim dummy As Byte
		<VBFixedArray(2)> Dim pad() As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim pad(2)
		End Sub
	End Structure
	
	Public Structure MMTIME
		Dim wType As Integer
		Dim units As Integer
		'UPGRADE_WARNING: Arrays in structure smpteVal may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim smpteVal As SMPTE
		Dim songPtrPos As Integer
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			smpteVal.Initialize()
		End Sub
	End Structure
	
	Public Function UtcNow() As Date
		UtcNow = SystemTimeToDate(GetSystemTime())
	End Function
	
	Public Function UtcToLocal(ByRef UtcDate As Date) As Date
		Dim FTime As FILETIME
		
		'UPGRADE_WARNING: Couldn't resolve default property of object FTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FTime = DateToFileTime(UtcDate)
		
		FileTimeToLocalFileTime(FTime, FTime)
		
		UtcToLocal = FileTimeToDate(FTime)
	End Function
	
	Public Function FileTimeToDate(ByRef FTime As FILETIME) As Date
		Dim STime As SYSTEMTIME
		FileTimeToSystemTime(FTime, STime)
		
		FileTimeToDate = SystemTimeToDate(STime)
	End Function
	
	Public Function DateToFileTime(ByRef DDate As Date) As FILETIME
		Dim STime As SYSTEMTIME
		
		'UPGRADE_WARNING: Couldn't resolve default property of object STime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		STime = DateToSystemTime(DDate)
		SystemTimeToFileTime(STime, DateToFileTime)
	End Function
	
	Public Function SystemTimeToDate(ByRef STime As SYSTEMTIME) As Date
		Dim tempDate As Date
		Dim tempTime As Date
		tempDate = DateSerial(STime.wYear, STime.wMonth, STime.wDay)
		tempTime = TimeSerial(STime.wHour, STime.wMinute, STime.wSecond)
		
		SystemTimeToDate = (System.Date.FromOADate(tempDate.ToOADate + tempTime.ToOADate))
	End Function
	
	Public Function DateToSystemTime(ByRef DDate As Date) As SYSTEMTIME
		With DateToSystemTime
			.wYear = DatePart(Microsoft.VisualBasic.DateInterval.Year, DDate)
			.wMonth = DatePart(Microsoft.VisualBasic.DateInterval.Month, DDate)
			.wDay = DatePart(Microsoft.VisualBasic.DateInterval.Day, DDate)
			.wDayOfWeek = DatePart(Microsoft.VisualBasic.DateInterval.WeekDay, DDate)
			.wHour = DatePart(Microsoft.VisualBasic.DateInterval.Hour, DDate)
			.wMinute = DatePart(Microsoft.VisualBasic.DateInterval.Minute, DDate)
			.wSecond = DatePart(Microsoft.VisualBasic.DateInterval.Second, DDate)
		End With
	End Function
	
	Public Function GetTimeZoneBias() As Integer
		'UPGRADE_WARNING: Arrays in structure TZinfo may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim TZinfo As TIME_ZONE_INFORMATION
		Dim lngL As Integer
		
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
		'UPGRADE_WARNING: Arrays in structure TZinfo may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim TZinfo As TIME_ZONE_INFORMATION ' holds time zone info
		Dim lngL As Integer ' time zone info result
		Dim i As Integer ' counter
		Dim b As Short ' a single character from the time zone name
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String ' return string
		
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
				str_Renamed = str_Renamed & Chr(b)
			End If
		Next 
		
		GetTimeZoneName = str_Renamed
	End Function
End Module