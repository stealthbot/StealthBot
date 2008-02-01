Attribute VB_Name = "modUTF8"
'StealthBot modUTF8 -- UTF-8 Conversion Module
'   Thanks to Skywing and Camel for much of the code in this module.

'http://forum.valhallalegends.com/phpbbs/index.php?board=18;action=display;threadid=1027&start=0
Option Explicit

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long

Private Const CP_ACP = 0
Private Const CP_UTF8 = 65001

Public Function UTF8Encode(str As String) As String
    Dim InputChars As Long
    InputChars = Len(str)
    
    'We need to first convert the ASCII input to Unicode before we can convert it to UTF-8...
    Dim UnicodeChars As Long, UnicodeBuffer As String
    UnicodeChars = MultiByteToWideChar(CP_ACP, 0, str, InputChars, vbNullString, 0)
    UnicodeBuffer = Space(UnicodeChars * 2)
    MultiByteToWideChar CP_ACP, 0, str, InputChars, UnicodeBuffer, UnicodeChars
    
    'Now that we've got everything translated to Unicode, we can (finally) convert it to UTF-8.
    Dim UTF8Chars As Long, UTF8Buffer As String
    UTF8Chars = WideCharToMultiByte(CP_UTF8, 0, UnicodeBuffer, UnicodeChars, 0, 0, vbNullString, 0)
    UTF8Buffer = Space(UTF8Chars)
    WideCharToMultiByte CP_UTF8, 0, UnicodeBuffer, UnicodeChars, UTF8Buffer, UTF8Chars, vbNullString, 0
    
    UTF8Encode = UTF8Buffer
End Function

Public Function UTF8Decode(str As String) As String
    Dim InputBytes As Long
    InputBytes = Len(str)
    
    'Again, we need to convert the UTF-8 string to Unicode before we can convert it to 8-bit.
    Dim UnicodeChars As Long, UnicodeBuffer As String
    UnicodeChars = MultiByteToWideChar(CP_UTF8, 0, str, InputBytes, vbNullString, 0)
    UnicodeBuffer = Space(UnicodeChars * 2)
    MultiByteToWideChar CP_UTF8, 0, str, InputBytes, UnicodeBuffer, UnicodeChars
   
    'Now that we've got everything translated to Unicode, we can convert it to 8-bit characters.
    'Dim SingleByteChars As Long, SingleByteBuffer As String
    'SingleByteChars = WideCharToMultiByte(CP_ACP, 0, UnicodeBuffer, UnicodeChars, vbNullString, 0, vbNullString, 0)
    'SingleByteBuffer = Space(SingleByteChars)
    'WideCharToMultiByte CP_ACP, 0, UnicodeBuffer, UnicodeChars, SingleByteBuffer, SingleByteChars, vbNullString, 0
    
    'UTF8Decode = SingleByteBuffer
    
    UTF8Decode = Replace(UnicodeBuffer, Chr$(0), vbNullString)
End Function
