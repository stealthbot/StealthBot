Option Strict Off
Option Explicit On
Module modUTF8
	'StealthBot modUTF8 -- UTF-8 Conversion Module
	'   Thanks to Skywing and Camel for much of the code in this module.
	
	'http://forum.valhallalegends.com/phpbbs/index.php?board=18;action=display;threadid=1027&start=0
	
	Private Declare Function GetLastError Lib "Kernel32.dll" () As Integer
	
	Private Declare Function MultiByteToWideChar Lib "Kernel32.dll" (ByVal Codepage As Integer, ByVal dwFlags As Integer, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Integer, ByVal lpWideCharStr As String, ByVal cchWideChar As Integer) As Integer
	Private Declare Function WideCharToMultiByte Lib "Kernel32.dll" (ByVal Codepage As Integer, ByVal dwFlags As Integer, ByVal lpWideCharStr As Integer, ByVal cchWideChar As Integer, ByVal lpMultiByteStr As Integer, ByVal cchMultiByte As Integer, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Integer) As Integer
	
	Private Const MB_ERR_INVALID_CHARS As Integer = &H8
	
	Private Const CP_ACP As Integer = 0
	Private Const CP_UTF8 As Integer = 65001
	
	'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function UTF8Encode(ByRef str_Renamed As String) As Byte()
		Dim UTF8Buffer() As Byte
		Dim UTF8Chars As Integer
		Dim lstr As String
		
		lstr = str_Renamed
		
		' grab Length of string after conversion
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		UTF8Chars = WideCharToMultiByte(CP_UTF8, 0, StrPtr(lstr), Len(lstr), 0, 0, vbNullString, 0)
		
		If (UTF8Chars = 0) Then
			Exit Function
		End If
		
		' initialize buffer
		ReDim UTF8Buffer(UTF8Chars - 1)
		
		' translate from unicode to utf-8
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(lstr), Len(lstr), VarPtr(UTF8Buffer(0)), UTF8Chars, vbNullString, 0)
		
		' return unicode buffer
		UTF8Encode = VB6.CopyArray(UTF8Buffer)
	End Function
	
	'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function UTF8Decode(ByRef str_Renamed As String, Optional ByRef LocaleID As Integer = 1252) As String
		Dim UnicodeBuffer As String
		Dim UnicodeChars As Integer
		Dim lstr As String
		
		lstr = str_Renamed
		
		' grab Length of string after conversion
		UnicodeChars = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, lstr, Len(lstr), vbNullString, 0)
		
		If (UnicodeChars = 0) Then
			Exit Function
		End If
		
		' initialize buffer
		UnicodeBuffer = New String(vbNullChar, UnicodeChars * 2)
		
		' translate utf-8 string to unicode
		Call MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, lstr, Len(lstr), UnicodeBuffer, UnicodeChars)
		
		' translate from unicode to ansi
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		UTF8Decode = StrConv(UnicodeBuffer, vbFromUnicode)
	End Function
End Module