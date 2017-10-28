Attribute VB_Name = "modUTF8"
'StealthBot modUTF8 -- UTF-8 Conversion Module
'   Thanks to Skywing and Camel for much of the code in this module.

'http://forum.valhallalegends.com/phpbbs/index.php?board=18;action=display;threadid=1027&start=0
Option Explicit

Private Declare Function GetLastError Lib "Kernel32.dll" () As Long

Private Declare Function MultiByteToWideChar Lib "Kernel32.dll" (ByVal Codepage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "Kernel32.dll" (ByVal Codepage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Private Const MB_ERR_INVALID_CHARS As Long = &H8

Private Const CP_ACP               As Long = 0
Private Const CP_UTF8              As Long = 65001

Public Function UTF8Encode(ByRef str As String) As Byte()
    Dim UTF8Buffer() As Byte
    Dim UTF8Chars    As Long
    Dim lstr         As String
    
    lstr = str
    
    ' grab Length of string after conversion
    UTF8Chars = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(lstr), Len(lstr), 0, 0, 0, 0)
    
    If (UTF8Chars = 0) Then
        ReDim UTF8Encode(0)
        Exit Function
    End If

    ' initialize buffer
    ReDim UTF8Buffer(0 To UTF8Chars)
    
    ' translate from unicode to utf-8
    Call WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(lstr), Len(lstr), VarPtr(UTF8Buffer(0)), UTF8Chars, 0, 0)
    
    ' return unicode buffer
    UTF8Encode = UTF8Buffer
End Function

Public Function UTF8Decode(ByRef buf() As Byte) As String
    Dim UnicodeBuffer() As Byte
    Dim UnicodeChars    As Long
    
    ' grab Length of string after conversion
    UnicodeChars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(buf(0)), UBound(buf) + 1, 0, 0)
            
    If (UnicodeChars = 0) Then
        Exit Function
    End If
    
    ' initialize buffer
    ReDim UnicodeBuffer(0 To (UnicodeChars * 2 - 1))
    
    ' translate utf-8 string to unicode
    Call MultiByteToWideChar(CP_UTF8, 0, VarPtr(buf(0)), UBound(buf) + 1, VarPtr(UnicodeBuffer(0)), UnicodeChars)
   
    ' translate from unicode to ansi
    UTF8Decode = StrConv(ByteArrToString(UnicodeBuffer()), vbFromUnicode)
End Function

' given a string, returns a string up to, and not including, the first vbNullChar found (null terminator)
Public Function KillNull(ByVal Text As String) As String

    Dim i As Integer
    i = InStr(1, Text, vbNullChar)
    If (i = 0) Then
        KillNull = Text
        Exit Function
    End If
    KillNull = Left$(Text, i - 1)

End Function

' converts a string to a byte array
' may have undefined behavior on strings of 0 length: use this on fixed-length strings (NLS values, hashes, etc)
Public Function StringToByteArr(ByVal Text As String) As Byte()

    StringToByteArr = StrConv(Text, vbFromUnicode)

End Function

' converts a string to a byte array with a null-terminating element
' empty strings return an array with one element, 0
Public Function StringToNTByteArr(ByVal Text As String) As Byte()

    If LenB(Text) = 0 Then
        ReDim StringToNTByteArr(0)
        Exit Function
    End If

    StringToNTByteArr = StrConv(Text, vbFromUnicode)
    ReDim Preserve StringToNTByteArr(LBound(StringToNTByteArr) To (UBound(StringToNTByteArr) + 1))

End Function

' converts a byte array to a string
' use this for fixed-length strings (NLS values, hashes, etc)
Public Function ByteArrToString(ByRef arr() As Byte) As String

    ByteArrToString = StrConv(arr(), vbUnicode)

End Function

' converts a byte array to a string
' the array will be attempted to be resized to trim after the first vbNullChar
Public Function NTByteArrToString(ByRef arr() As Byte) As String

    Dim i As Integer

    If UBound(arr) < 0 Then
        NTByteArrToString = vbNullString
        Exit Function
    End If

    If arr(0) = &H0 Then
        NTByteArrToString = vbNullString
        Exit Function
    End If

    For i = LBound(arr) To UBound(arr)
        If arr(i) = &H0 Then
            Exit For ' preserve i
        End If
    Next i

    ReDim Preserve arr(LBound(arr) To (i - 1))
    NTByteArrToString = StrConv(arr(), vbUnicode)

End Function

Public Function ByteArrToWCArr(ByRef bArr() As Byte) As Integer()

    Dim iArr() As Integer
    Dim i As Integer

    ReDim iArr(0 To UBound(bArr) \ 2)
    For i = 0 To UBound(bArr) - 1 Step 2
        iArr(i \ 2) = bArr(i) Or (bArr(i + 1) * &H100)
    Next i
    ByteArrToWCArr = iArr

End Function

Public Function DWordToString(ByVal Value As Long) As String
    
    Dim Buffer As String * 4
    CopyMemory ByVal Buffer, Value, 4
    DWordToString = KillNull(StrReverse$(Buffer))

End Function

Public Function StringToDWord(ByVal Value As String) As Long
    
    Dim Buffer As String * 4
    Buffer = String$(4, vbNullChar)
    Mid$(Buffer, 1, Len(Value)) = Value
    Buffer = StrReverse$(Buffer)
    CopyMemory StringToDWord, ByVal Buffer, 4

End Function
