Attribute VB_Name = "modUTF8"
'StealthBot modUTF8 -- UTF-8 Conversion Module
'   Thanks to Skywing and Camel for much of the code in this module.

'http://forum.valhallalegends.com/phpbbs/index.php?board=18;action=display;threadid=1027&start=0
Option Explicit

Private Declare Function GetLastError Lib "Kernel32.dll" () As Long

Private Declare Function MultiByteToWideChar Lib "Kernel32.dll" (ByVal Codepage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "Kernel32.dll" (ByVal Codepage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long

Private Const MB_ERR_INVALID_CHARS As Long = &H8

Private Const CP_ACP               As Long = 0
Private Const CP_UTF8              As Long = 65001

' ...
Public Function UTF8Encode(ByRef str As String) As Byte()
    Dim UTF8Buffer() As Byte   ' ...
    Dim UTF8Chars    As Long   ' ...
    Dim lstr         As String ' ...
    
    ' ...
    lstr = str
    
    ' grab Length of string after conversion
    UTF8Chars = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(lstr), Len(lstr), 0, 0, _
        vbNullString, 0)
    
    ' ...
    If (UTF8Chars = 0) Then
        ' ...
        Exit Function
    End If

    ' initialize buffer
    ReDim UTF8Buffer(0 To UTF8Chars - 1)
    
    ' translate from unicode to utf-8
    Call WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(lstr), Len(lstr), ByVal VarPtr(UTF8Buffer(0)), _
        UTF8Chars, vbNullString, 0)
    
    ' return unicode buffer
    UTF8Encode = UTF8Buffer
End Function

' ...
Public Function UTF8Decode(ByRef str As String, Optional LocaleID As Long = 1252) As String
    Dim UnicodeBuffer As String ' ...
    Dim UnicodeChars  As Long   ' ...
    Dim lstr          As String ' ...
    
    ' ...
    lstr = str
    
    ' grab Length of string after conversion
    UnicodeChars = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, lstr, Len(lstr), _
        vbNullString, 0)
            
    ' ...
    If (UnicodeChars = 0) Then
        ' ...
        Exit Function
    End If
    
    ' initialize buffer
    UnicodeBuffer = String$(UnicodeChars * 2, vbNullChar)
    
    ' translate utf-8 string to unicode
    Call MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, lstr, Len(lstr), UnicodeBuffer, _
            UnicodeChars)
   
    ' translate from unicode to ansi
    UTF8Decode = StrConv(UnicodeBuffer, vbFromUnicode)
End Function
