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

Public Function ApplyGameColors(saElements() As Variant, arr() As Variant) As Boolean

    Dim i, j        As Long
    Dim CodePos     As Long
    Dim IsColor     As Boolean
    Dim Color       As Long
    Dim StyleSpec   As String
    Dim CodeLength  As Long
    Dim TextBefore  As String
    Dim TextAfter   As String

    ApplyGameColors = False

    ' start off with the same number of elements
    ReDim arr(LBound(saElements) To UBound(saElements))

    j = LBound(arr)
    For i = LBound(saElements) To UBound(saElements) Step 3
        arr(j) = saElements(i)
        arr(j + 1) = saElements(i + 1)
        arr(j + 2) = saElements(i + 2)
        Do
            CodePos = IsColorCode(saElements(i + 2), IsColor, Color, StyleSpec, CodeLength)
            If CodePos > 0 Then
                ApplyGameColors = True
                TextBefore = Left$(saElements(i + 2), CodePos - 1)
                TextAfter = Mid$(saElements(i + 2), CodePos + CodeLength)
                If LenB(TextBefore) > 0 And LenB(TextAfter) > 0 Then
                    ' color code mid string, split required, add element to arr
                    ReDim Preserve arr(LBound(arr) To UBound(arr) + 3)
                    arr(j + 2) = TextBefore
                    If Not IsColor Then
                        arr(j + 3) = CombineStyle(saElements(i), StyleSpec)
                        arr(j + 4) = arr(j + 1) ' continue color
                    Else
                        arr(j + 3) = vbNullString ' continue font
                        arr(j + 4) = Color
                    End If
                    arr(j + 5) = TextAfter
                    j = j + 3
                    saElements(i + 2) = TextAfter
                ElseIf LenB(TextBefore) = 0 Then
                    ' color code starts element
                    If Not IsColor Then
                        arr(j) = CombineStyle(arr(j), StyleSpec)
                        'arr(j + 1) = saElements(i + 1) ' continue color
                    Else
                        'arr(j) = vbNullString  ' continue font
                        arr(j + 1) = Color
                    End If
                    arr(j + 2) = TextAfter
                    saElements(i + 2) = TextAfter
                ElseIf LenB(TextAfter) = 0 Then
                    ' color code ends element
                    arr(j + 2) = TextBefore
                    saElements(i + 2) = vbNullString
                End If
            End If
        Loop While CodePos > 0
        j = j + 3
    Next i

End Function

Private Function CombineStyle(ByVal StyleFont As String, ByVal StyleSpec As String) As String

    Dim ColPos As Long
    ColPos = InStr(1, StyleFont, ":", vbBinaryCompare)

    If ColPos > 0 Then
        CombineStyle = Left$(StyleFont, ColPos - 1) & StyleSpec & Mid$(StyleFont, ColPos)
    Else
        CombineStyle = StyleSpec & ":" & StyleFont
    End If

End Function

' given a string, returns the string index of the first game color code inside
' if none, returns 0
' if one is found, sets IsColor, Color, and StyleSpec
Private Function IsColorCode(ByVal sInput As String, ByRef IsColor As Boolean, ByRef Color As Long, ByRef StyleSpec As String, ByRef CodeLength As Long) As Long

    Dim i      As Long
    Dim c      As String * 1
    Dim CodeID As String
    Dim IsCode As Boolean

    ' default byrefs
    IsColor = False
    Color = 0
    StyleSpec = vbNullString
    CodeLength = 0

    ' search by char
    For i = 1 To Len(sInput) - 1
        c = Mid$(sInput, i, 1)
        If AscW(c) = &HC1 Or c = Chr$(&HC1) Then
            ' DIABLO/WARCRAFT II/STARCRAFT
            CodeID = Mid$(sInput, i + 1, 1)
            IsCode = True
            Select Case CodeID
                Case "Q":           IsColor = True: Color = &H5D5D5D  ' grey
                Case "R":           IsColor = True: Color = &H36E01E  ' green
                Case "Z", "X", "S": IsColor = True: Color = &H74A9A0  ' yellow
                Case "Y", "[":      IsColor = True: Color = &H5226E7  ' red
                Case "V", "@":      IsColor = True: Color = &HE84D62  ' blue
                Case "W", "P":      IsColor = True: Color = &HFFFFFF  ' white
                Case "T", "U":      IsColor = True: Color = &HCCCC00  ' cyan/teal
                Case Else:          IsCode = False
            End Select

            If IsCode Then
                IsColorCode = i
                CodeLength = 2
                Exit Function
            End If
        End If

        If AscW(c) = &HFF Or c = Chr$(&HFF) Then
            ' DIABLO II
            ' ycX where X = b, i, u, ., ;, :, <, or 0-9
            If StrComp(Mid$(sInput, i + 1, 1), "c", vbTextCompare) = 0 Then
                CodeID = Mid$(sInput, i + 2, 1)
                IsCode = True
                Select Case CodeID
                    Case "b":  StyleSpec = "B"
                    Case "i":  StyleSpec = "I"
                    Case "u":  StyleSpec = "U"
                    Case ".":  StyleSpec = "BU"
                    Case ";":  IsColor = True: Color = &HCE008D ' purple
                    Case ":":  IsColor = True: Color = &H2D828  ' ligher green
                    Case "<":  IsColor = True: Color = &HA200&  ' dark green
                    Case "0":  IsColor = True: Color = &HFFFFFF ' white
                    Case "1":  IsColor = True: Color = &H3E3ECE ' red
                    Case "2":  IsColor = True: Color = &HCE00&  ' green
                    Case "3":  IsColor = True: Color = &H9C4044 ' blue
                    Case "4":  IsColor = True: Color = &H6091A1 ' gold
                    Case "5":  IsColor = True: Color = &H555555 ' grey
                    Case "6":  IsColor = True: Color = &H80808  ' black
                    Case "7":  IsColor = True: Color = &H659DA8 ' gold
                    Case "8":  IsColor = True: Color = &H88CE&  ' gold orange
                    Case "9":  IsColor = True: Color = &H51CECE ' light yellow
                    Case Else: IsCode = False
                End Select

                If IsCode Then
                    IsColorCode = i
                    CodeLength = 3
                    Exit Function
                End If
            End If
        End If
    Next i

    ' nothing found
    IsColorCode = 0

End Function

