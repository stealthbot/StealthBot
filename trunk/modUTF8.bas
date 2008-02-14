Attribute VB_Name = "modUTF8"
'StealthBot modUTF8 -- UTF-8 Conversion Module
'   Thanks to Skywing and Camel for much of the code in this module.

'http://forum.valhallalegends.com/phpbbs/index.php?board=18;action=display;threadid=1027&start=0
Option Explicit

Private Declare Function GetLastError Lib "Kernel32" () As Long

Public Declare Function MultiByteToWideChar Lib "Kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long
Public Declare Function WideCharToMultiByte Lib "Kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long

Private Const MB_ERR_INVALID_CHARS As Long = &H8

Private Const CP_ACP               As Long = 0
Private Const CP_UTF8              As Long = 65001

' ...
Public Function UTF8Encode(ByRef str As String) As String
    Dim UnicodeBuffer As String ' ...
    Dim UTF8Buffer    As String ' ...
    Dim UTF8Chars     As Long   ' ...

    ' translate string to unicode
    UnicodeBuffer = StrConv(str, vbUnicode)

    ' grab length of string after conversion
    UTF8Chars = WideCharToMultiByte(CP_UTF8, 0, UnicodeBuffer, Len(str), _
        0, 0, vbNullString, 0)
    
    ' initialize buffer
    UTF8Buffer = String$(UTF8Chars, vbNullChar)
    
    ' translate from unicode to utf-8
    WideCharToMultiByte CP_UTF8, 0, UnicodeBuffer, Len(str), UTF8Buffer, _
        UTF8Chars, vbNullString, 0
    
    ' return unicode buffer
    UTF8Encode = UTF8Buffer
End Function

' ...
Public Function UTF8Decode(ByRef str As String) As String
    Dim UnicodeBuffer As String ' ...
    Dim UnicodeChars  As Long   ' ...
    
    ' grab length of string after conversion
    UnicodeChars = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, str, _
        Len(str), vbNullString, 0)
    
    ' initialize buffer
    UnicodeBuffer = String$(UnicodeChars * 2, vbNullChar)
    
    ' translate utf-8 string to unicode
    MultiByteToWideChar CP_UTF8, 0, str, Len(str), UnicodeBuffer, UnicodeChars
   
    ' translate from unicode to ansi
    UTF8Decode = StrConv(UnicodeBuffer, vbFromUnicode)
End Function
