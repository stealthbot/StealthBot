Attribute VB_Name = "modUTF8"
'StealthBot modUTF8 -- UTF-8 Conversion Module
'   Thanks to Skywing and Camel for much of the code in this module.

'http://forum.valhallalegends.com/phpbbs/index.php?board=18;action=display;threadid=1027&start=0
Option Explicit

Private Declare Function MultiByteToWideChar Lib "Kernel32.dll" (ByVal Codepage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "Kernel32.dll" (ByVal Codepage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Private Enum RTBC_FLAGS ' CharFormat (SCF_) flags for EM_SETCHARFORMAT message.
    RTBC_DEFAULT = 0
    RTBC_SELECTION = 1
    RTBC_WORD = 2 'Combine with RTBC_SELECTION!
    RTBC_ALL = 4
End Enum

Public Enum RTBW_FLAGS ' Flags for the SETEXTEX data structure.
    RTBW_DEFAULT = 0   ' Deletes undo stack, discards RTF formatting, replaces all text.
    RTBW_KEEPUNDO = 1  ' Keeps undo stack.
    RTBW_SELECTION = 2 ' Replaces selection and keeps RTF formatting.
End Enum

Private Type SETTEXTEX
    Flags As RTBW_FLAGS
    Codepage As Long
End Type

Private Type GETTEXTLENGTHEX
    Flags As Long
    Codepage As Long
End Type

Private Type GETTEXTEX
    cb As Long
    Flags As Long
    Codepage As Long
    lpDefaultChar As Long
    lpUsedDefChar As Long
End Type

Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Private Const WM_USER As Long = &H400
Private Const EM_EXGETSEL As Long = WM_USER + 52
Private Const EM_EXSETSEL As Long = WM_USER + 55
Private Const EM_SETTEXTEX As Long = WM_USER + 97
Private Const EM_GETTEXTEX As Long = WM_USER + 94
Private Const EM_GETTEXTLENGTHEX As Long = WM_USER + 95

Private Const GT_USECRLF As Long = 1&

Private Const GTL_USECRLF As Long = 1&
Private Const GTL_PRECISE As Long = 2&
Private Const GTL_NUMCHARS As Long = 8&

Private Const MB_ERR_INVALID_CHARS As Long = &H8

Private Const CP_ACP     As Long = 0
Private Const CP_1252    As Long = 1252
Private Const CP_UTF8    As Long = 65001
Private Const CP_UNICODE As Long = 1200&

Private Const ERROR_SUCCESS                 As Long = 0
Private Const ERROR_NOT_ENOUGH_MEMORY       As Long = 8
Private Const ERROR_INVALID_PARAMETER       As Long = 87
Private Const ERROR_INSUFFICIENT_BUFFER     As Long = 122
Private Const ERROR_NO_UNICODE_TRANSLATION  As Long = 1113

Public Function UTF8Encode(ByRef str As String, Optional ByVal Codepage As Long = CP_UTF8) As Byte()
    Dim UTF8Buffer() As Byte
    Dim UTF8Chars    As Long
    Dim lstr         As String
    
    lstr = str
    
    ' grab Length of string after conversion
    UTF8Chars = WideCharToMultiByte(Codepage, 0, ByVal StrPtr(lstr), Len(lstr), 0, 0, 0, 0)
    
    If (UTF8Chars = 0) Then
        ReDim UTF8Encode(0)
        Exit Function
    End If

    ' initialize buffer
    ReDim UTF8Buffer(0 To UTF8Chars)
    
    ' translate from unicode to utf-8
    Call WideCharToMultiByte(Codepage, 0, ByVal StrPtr(lstr), Len(lstr), VarPtr(UTF8Buffer(0)), UTF8Chars, 0, 0)
    
    ' return unicode buffer
    UTF8Encode = UTF8Buffer
End Function

Public Function UTF8Decode(ByRef buf() As Byte, Optional ByVal Codepage As Long = CP_UTF8) As String
    Dim UnicodeBuffer() As Byte
    Dim UnicodeChars    As Long
    
    ' grab Length of string after conversion
    UnicodeChars = MultiByteToWideChar(Codepage, MB_ERR_INVALID_CHARS, VarPtr(buf(0)), UBound(buf) + 1, 0, 0)
            
    If (UnicodeChars = 0) Then
        Select Case Err.LastDllError
            Case ERROR_NO_UNICODE_TRANSLATION
                ' try again with CP_1252
                Codepage = CP_1252
                UnicodeChars = MultiByteToWideChar(Codepage, 0, VarPtr(buf(0)), UBound(buf) + 1, 0, 0)
                
                If (UnicodeChars = 0) Then
                    Exit Function
                End If
            Case Else
                ' another error
                Exit Function
        End Select
    End If
    
    ' initialize buffer
    ReDim UnicodeBuffer(0 To (UnicodeChars * 2 - 1))
    
    ' translate utf-8 string to unicode
    Call MultiByteToWideChar(Codepage, 0, VarPtr(buf(0)), UBound(buf) + 1, VarPtr(UnicodeBuffer(0)), UnicodeChars)
   
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

Public Function GetRTBLength(rtb As RichTextBox) As Long

    Dim GetTextLengthObj As GETTEXTLENGTHEX
    
    GetTextLengthObj.Flags = GTL_USECRLF Or GTL_PRECISE Or GTL_NUMCHARS
    GetTextLengthObj.Codepage = CP_UNICODE
    GetRTBLength = SendMessageW(rtb.hWnd, EM_GETTEXTLENGTHEX, VarPtr(GetTextLengthObj), 0)

End Function

Public Function GetRTBText(rtb As RichTextBox) As String
    
    Dim GetTextObj As GETTEXTEX
    Dim iChars As Long
    
    iChars = GetRTBLength(rtb)
    
    If iChars > 0 Then
        GetTextObj.cb = (iChars + 1) * 2
        GetTextObj.Flags = GT_USECRLF
        GetTextObj.Codepage = CP_UNICODE
        GetRTBText = String$(iChars, vbNullChar)
        SendMessageW rtb.hWnd, EM_GETTEXTEX, VarPtr(GetTextObj), StrPtr(GetRTBText)
    Else
        GetRTBText = vbNullString
    End If

End Function

Public Sub RTBSetSelectedText(rtb As RichTextBox, ByVal sText As String)
    
    Dim SetTextObj As SETTEXTEX
    
    SetTextObj.Flags = RTBW_SELECTION
    SetTextObj.Codepage = CP_UNICODE
    SendMessageW rtb.hWnd, EM_SETTEXTEX, ByVal VarPtr(SetTextObj), ByVal StrPtr(sText)

End Sub

Public Sub GetTextSelection(cnt As Control, sParam As Long, eParam As Long)

    Dim CharRangeObj As CHARRANGE

    SendMessageW cnt.hWnd, EM_EXGETSEL, 0, ByVal VarPtr(CharRangeObj)
    sParam = CharRangeObj.cpMin
    eParam = CharRangeObj.cpMax

End Sub

Public Sub SetTextSelection(cnt As Control, ByVal sParam As Long, ByVal eParam As Long)

    Dim CharRangeObj As CHARRANGE
    
    CharRangeObj.cpMin = sParam
    CharRangeObj.cpMax = eParam
    SendMessageW cnt.hWnd, EM_EXSETSEL, 0, ByVal VarPtr(CharRangeObj)

End Sub
