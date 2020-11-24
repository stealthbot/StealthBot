Attribute VB_Name = "modChat"
'StealthBot modChat -- RichTextBox, Chat, and UTF-8/String Manipulation Module
'   Thanks to Skywing and Camel for much of the UTF-8 conversion code.

'http://forum.valhallalegends.com/phpbbs/index.php?board=18;action=display;threadid=1027&start=0
Option Explicit

Private Declare Function MultiByteToWideChar Lib "Kernel32.dll" (ByVal Codepage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "Kernel32.dll" (ByVal Codepage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

' Unicode interaction
' https://stackoverflow.com/questions/540361/whats-the-best-option-to-display-unicode-text-hebrew-etc-in-vb6
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal OleStr As Long, ByVal bLen As Long) As Long
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "Kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "Kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "Kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "Kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long

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

Private Const WM_SETTEXT         As Long = &HC
Private Const WM_GETTEXT         As Long = &HD
Private Const WM_GETTEXTLENGTH   As Long = &HE
Private Const WM_VSCROLL         As Long = &H115
Private Const WM_USER            As Long = &H400

Private Const EM_GETSEL          As Long = &HB0
Private Const EM_SETSEL          As Long = &HB1
Private Const EM_REPLACESEL      As Long = &HC2
Private Const EM_GETSELTEXT      As Long = WM_USER + 62
Private Const EM_SETTEXTEX       As Long = WM_USER + 97
Private Const EM_GETTEXTEX       As Long = WM_USER + 94
Private Const EM_GETTEXTLENGTHEX As Long = WM_USER + 95

Private Const EM_SCROLL      As Long = &HB5
Private Const EM_SCROLLCARET As Long = &HB7
Private Const EM_GETTHUMB    As Long = &HBE

Private Const SB_VERT As Long = 1
Private Const SB_HORZ As Long = 0
Private Const SB_BOTH As Long = 3

Private Const SB_THUMBPOSITION As Long = &H4

Private Const SB_TOP       As Long = 8
Private Const SB_BOTTOM    As Long = 7

Private Const WS_HSCROLL   As Long = &H100000
Private Const WS_VSCROLL   As Long = &H200000
Private Const GWL_STYLE    As Long = (-16)

Private Const GT_DEFAULT      As Long = 0&
Private Const GT_USECRLF      As Long = 1&
Private Const GT_SELECTION    As Long = 2&
Private Const GT_RAWTEXT      As Long = 4&
Private Const GT_NOHIDDENTEXT As Long = 8&

Private Const GTL_USECRLF  As Long = 1&
Private Const GTL_PRECISE  As Long = 2&
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

' DISPLAYRICHTEXT

' Fixed font issue when an element was only 1 character long -Pyro (9/28/08)
' Fixed issue with displaying null text.

' I changed the location of where fontStr was being declared. I can't figure out why it makes any difference, but _
    when I have it declared above I get memory errors, my IDE crashes, I runtime erro 380 and _
  - Error (#-2147417848): Method 'SelFontName' of object 'IRichText' failed in DisplayRichText().
' I believe it's something to do with the subclassing overwriting the memory, but it only occurs when run from the IDE. - FrOzeN
' Overhauls to DisplayRichText functionality, allowing Unicode and better color applying. - Ribose 2017-10
Public Sub DisplayRichText(ByRef rtb As RichTextBox, ByRef saElements() As Variant)
    On Error GoTo ERROR_HANDLER

    ' index for iterating through elements
    Dim i              As Long
    ' for pre-modifying the array (CTCT... > FCTFCT... and applying game colors)
    Dim arr()          As Variant
    Dim j              As Long
    ' values to save and restore after print
    Dim lngVerticalPos As Long
    Dim blnCanVScroll  As Boolean
    Dim blnCaratAtEnd  As Boolean
    Dim blnScrollAtEnd As Boolean
    Dim SelStart       As Long
    Dim SelLength      As Long
    ' don't draw while printing
    Dim blnUnlock      As Boolean
    Dim blnVisible     As Boolean
    ' logging
    Dim LineLength     As Long
    Dim LineText       As String
    Dim LogThis        As Boolean
    ' for backlog removal
    Dim Length         As Long
    Dim RemoveLength   As Long
    ' line etyle state
    Dim StyleBold      As Boolean
    Dim StyleItal      As Boolean
    Dim StyleUndl      As Boolean
    Dim StyleStri      As Boolean
    ' string to print
    Dim ElementText    As String
    
    Static RichTextErrorCounter As Integer

    ' *****************************************
    '              SANITY CHECKS
    ' *****************************************

    ' empty array
    If LBound(saElements) > UBound(saElements) Then
        Exit Sub
    End If

    ' if first element is numeric
    If (IsNumeric(saElements(0))) Then
        j = 2

        ' convert AddChat Color, Text, Color, Text, ...
        '      to AddChat DefaultFont, Color, Text, DefaultFont, Color, Text, ...
        For i = LBound(saElements) To UBound(saElements) Step 2
            ReDim Preserve arr(0 To j) As Variant

            arr(j) = saElements(i + 1)
            arr(j - 1) = saElements(i)
            arr(j - 2) = vbNullString

            j = j + 3
        Next i

        saElements() = arr()
    End If

    ' verify arguments
    For i = LBound(saElements) To UBound(saElements) Step 3
        ' element count not a multiple of 3
        If ((i + 2) > UBound(saElements)) Then
            Exit Sub
        End If

        ' color is not positive Integer or Long value
        If (IsNumeric(saElements(i + 1)) = False) Then
            Exit Sub
        End If

        ' convert negative Integer values to Long (for example &H99CC is negative, convert to &H99CC&)
        If (saElements(i + 1) < 0) Then
            saElements(i + 1) = CLng(saElements(i + 1) + &H10000)
        Else
            saElements(i + 1) = CLng(saElements(i + 1))
        End If

        ' out of color range
        If (saElements(i + 1) < 0 Or saElements(i + 1) > &HFFFFFF) Then
            Exit Sub
        End If

        ' store combined length of input
        LineLength = LineLength + Len(KillNull(saElements(i + 2)))
    Next i
    
    If ApplyGameColors(saElements(), arr()) Then
        saElements() = arr()
    End If

    ' input must have non-zero length
    If (LineLength = 0) Then
        Exit Sub
    End If

    If ((BotVars.LockChat = False) Or (rtb <> frmChat.rtbChat)) Then
        ' store rtb carat and whether rtb has focus
        GetTextSelection rtb.hWnd, SelStart, SelLength

        ' whether carat is at the end or within one vbCrLf of the end
        blnCaratAtEnd = (SelStart >= GetRTBLength(rtb.hWnd) - 2)

        ' is the RTB at the bottom?
        lngVerticalPos = GetVScrollPosition(rtb)
        blnCanVScroll = CanVScroll(rtb.hWnd)
        blnScrollAtEnd = (Not blnCanVScroll) Or (lngVerticalPos = 0)

        If (rtb.Visible) Then
            ' disallow redraw
            DisableWindowRedraw rtb.hWnd

            blnUnlock = True
        End If

        ' how to log this event
        If (rtb = frmChat.rtbChat) Or (rtb = frmChat.rtbWhispers) Then
            LogThis = (BotVars.Logging > 0)
        End If

        ' remove from backlog if overflow
        Length = GetRTBLength(rtb.hWnd)
        If ((BotVars.MaxBacklogSize) And (Length > BotVars.MaxBacklogSize)) Then
            With rtb
                'Debug.Print "S " & SelStart & ", L " & SelLength
                RemoveLength = InStr(Length - BotVars.MaxBacklogSize, GetRTBText(rtb.hWnd), vbLf, vbBinaryCompare)
                SetTextSelection rtb.hWnd, 0, RemoveLength
                ' remove line from stored selection
                SelStart = SelStart - RemoveLength
                SelLength = SelLength - RemoveLength
                ' if selection included part of what was removed, add negative start point
                ' to length to get difference in length, and set start selection at 0
                If SelStart < 0 Then SelStart = 0
                If SelLength < 0 Then SelLength = 0
                'Debug.Print "S " & SelStart & ", L " & SelLength
                SetSelectedRTBText rtb.hWnd, vbNullString
            End With
        End If

        ' place timestamp
        ElementText = GetTimeStamp()
        If LenB(ElementText) > 0 Then
            With rtb
                SetTextSelection rtb.hWnd, -1, -1
                .SelFontName = rtb.Font.Name
                .SelFontSize = rtb.Font.Size
                .SelBold = False
                .SelItalic = False
                .SelUnderline = False
                .SelStrikeThru = False
                .SelColor = g_Color.TimeStamps
                SetSelectedRTBText rtb.hWnd, ElementText
            End With
        End If

        ' place each element
        For i = LBound(saElements) To UBound(saElements) Step 3
            DisplayRichTextElement rtb, saElements(), i, StyleBold, StyleItal, StyleUndl, StyleStri

            ElementText = KillNull(saElements(i + 2))
            LineText = LineText & ElementText
        Next i

        With rtb
            SetTextSelection rtb.hWnd, -1, -1
            SetSelectedRTBText rtb.hWnd, vbCrLf
        End With

        ' set scrollbar
        If blnScrollAtEnd Then
            ' scroll to end
            'Debug.Print "SCROLL TO BOTTOM"
            ScrollToCaret rtb.hWnd
        Else
            ' set to scroll specific position
            'Debug.Print "SCROLL TO " & lngVerticalPos
            ScrollToPosition rtb, lngVerticalPos
        End If

        ' set carat
        If blnCaratAtEnd Then
            ' carat was at the end before
            SelStart = -1
            SelLength = -1
        End If
        'Debug.Print "SET CARAT " & SelStart & "," & SelLength
        SetTextSelection rtb.hWnd, SelStart, SelLength

        If (blnUnlock) Then
            ' allow redraw
            EnableWindowRedraw rtb.hWnd
            PerformWindowRedraw rtb.hWnd
        End If

        If (LogThis) Then
            If (rtb = frmChat.rtbChat) Then
                g_Logger.WriteChat LineText
            ElseIf (rtb = frmChat.rtbWhispers) Then
                g_Logger.WriteWhisper LineText
            End If
        End If

        'rtb.Visible = blnVisible
    End If

    RichTextErrorCounter = 0

    Exit Sub
    
ERROR_HANDLER:

    If (blnUnlock) Then
        ' allow redraw
        EnableWindowRedraw rtb.hWnd
        PerformWindowRedraw rtb.hWnd
    End If

    RichTextErrorCounter = RichTextErrorCounter + 1
    If RichTextErrorCounter > 2 Then
        RichTextErrorCounter = 0
        Exit Sub
    End If
    
    If (Err.Number = 13 Or Err.Number = 91) Then
        Exit Sub
    End If

    frmChat.AddChat g_Color.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in DisplayRichText()."
    
    Exit Sub
    
End Sub

' DISPLAYRICHTEXTELEMENT
' Displays a single "3-part element" of a DisplayRichText() call.
' saElements() must have elements X to X + 2 where X = the passed in i value
' saElements(X    ) = Font (string name, or "StyleSpec:Font")
' saElements(X + 1) = Color (integer or long)
' saElements(X + 2) = Text (string)
'
' Font value:
'   A "StyleSpec" is used to specify Bold, Italic, Underline, or Strikethrough style along with the font.
'   Value should be one or more of these characters followed by a ":" followed by a Font name
'     "B": toggle bold
'     "I": toggle italic
'     "U": toggle underline
'     "S": toggle strikethrough
'     "R": turn off styles
'   Font names otherwise can be vbNullString (continue previous font), "%" (go back to rtb.Font.Name), or a valid Font (change to this)
'   Font values such as these are examples:
'     ":" (the same as vbNullstring)
'     "...:" (where ... is a StyleSpec; sets style but not font)
'     ":Font with : in it" (sets font to "Font with : in it")
' Color value:
'   Any integer/long in the range &H000000 and &HFFFFFF, inclusive.
' Text value:
'   The string to write. Assumes Unicode ("UTF-16").
'
' The optional four ByRef booleans allow the DisplayRichText() function to "continue" toggling font styles throughout the line.
Public Sub DisplayRichTextElement(rtb As RichTextBox, saElements() As Variant, ByVal i As Long, _
        Optional StyleBold As Boolean, Optional StyleItal As Boolean, Optional StyleUndl As Boolean, Optional StyleStri As Boolean)
    Dim s        As String
    Dim j        As Long
    Dim StyleSep As Long
    Dim FontName As String

    If ((StrictIsNumeric(saElements(i + 1))) And (Len(saElements(i + 2)) > 0)) Then
        s = KillNull(saElements(i + 2))

        With rtb
            SetTextSelection rtb.hWnd, -1, -1
            StyleSep = InStr(1, saElements(i), ":", vbBinaryCompare)
            If StyleSep > 1 Then
                For j = 1 To StyleSep - 1
                    Select Case UCase$(Mid$(saElements(i), j, 1))
                        Case "B": StyleBold = Not StyleBold ' bold
                        Case "I": StyleItal = Not StyleItal ' italic
                        Case "U": StyleUndl = Not StyleUndl ' underline
                        Case "S": StyleStri = Not StyleStri ' strikethrough
                        Case "R":                           ' reset style
                            StyleBold = False
                            StyleItal = False
                            StyleUndl = False
                            StyleStri = False
                    End Select
                Next j
            End If
            .SelBold = StyleBold
            .SelItalic = StyleItal
            .SelUnderline = StyleUndl
            .SelStrikeThru = StyleStri
            FontName = Mid$(saElements(i), StyleSep + 1)
            If StrComp(FontName, "%", vbBinaryCompare) = 0 Then
                ' default font
                .SelFontName = rtb.Font.Name
            ElseIf LenB(FontName) > 0 Then
                ' change font
                .SelFontName = FontName
            End If
            .SelColor = saElements(i + 1)
            SetSelectedRTBText rtb.hWnd, s
        End With
    End If
End Sub

' APPLYGAMECOLORS
' Given a saElements() array [to pass to DisplayRichText()], modify this array to add colors,
' to allow in-chat colors as the official clients used to support, and now bots widely support.
'
' Builds the arr() array as the result.
'
' Returns True if there are changes, and False if saElements() should continue to be used (no color codes found).
'
' Supported color codes:
' * DIABLO/STARCRAFT/WARCRAFT II Style
'   Length: 2 bytes
'   Format: 0xC1 __
'                ^Alphabet character
'   Results: color only
' * DIABLO II Style
'   Length 3 bytes
'   Format: 0xFF 'c' __
'                    ^Alphanumeric character
'   Results: font style (bold, italic, underline) or color
'
' Codes can be seen in IsColorCode()
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

' ISCOLORCODE
' given a string, returns the string index of the first game color code inside
' if none, returns 0
' if one is found, sets IsColor, Color, StyleSpec, and CodeLength
Private Function IsColorCode(ByVal sInput As String, ByRef IsColor As Boolean, ByRef Color As Long, ByRef StyleSpec As String, ByRef CodeLength As Long) As Long

    Dim i      As Long
    Dim j      As Long
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
            ' DIABLO/STARCRAFT/WARCRAFT II
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

        If c = "|" Then
            ' WARCRAFT III
            ' |c######## to set color
            IsCode = True
            Select Case LCase$(Mid$(sInput, i + 1, 1))
                Case "c"
                    CodeID = Mid$(sInput, i + 2, 8)
                    If Len(CodeID) <> 8 Then
                        IsCode = False
                    Else
                        IsCode = True
                        For j = 1 To 8
                            Select Case LCase$(Mid$(CodeID, j, 1))
                                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"
                                Case Else
                                    IsCode = False
                                    Exit For
                            End Select
                        Next j
                        If IsCode Then
                            IsColor = True
                            Color = g_Color.FromHex(Mid$(CodeID, 3, 6))
                            CodeLength = 10
                        End If
                    End If

                Case "b":  StyleSpec = "B": CodeLength = 2
                Case "i":  StyleSpec = "I": CodeLength = 2
                Case "u":  StyleSpec = "U": CodeLength = 2
                Case "r":  IsColor = True: Color = g_Color.White: StyleSpec = "R": CodeLength = 2
                Case Else: IsCode = False

            End Select

            If IsCode Then
                IsColorCode = i
                Exit Function
            End If
        End If
    Next i

    ' nothing found
    IsColorCode = 0

End Function

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

Public Function DWordToStringR(ByVal Value As Long) As String

    Dim Buffer As String * 4
    CopyMemory ByVal Buffer, Value, 4
    DWordToStringR = Right(Buffer, 4 - InStrRev(Buffer, Chr$(0)))

End Function

Public Function StringToDWord(ByVal Value As String) As Long

    Dim Buffer As String * 4
    Buffer = String$(4, vbNullChar)
    Mid$(Buffer, 1, Len(Value)) = Value
    Buffer = StrReverse$(Buffer)
    CopyMemory StringToDWord, ByVal Buffer, 4

End Function

Public Function StringRToDWord(ByVal Value As String) As Long

    Dim Buffer As String * 4
    Buffer = String$(4, vbNullChar)
    Mid$(Buffer, 1, Len(Value)) = Value
    CopyMemory StringRToDWord, ByVal Buffer, 4

End Function

Public Function GetRTBLength(ByVal hWnd As Long) As Long

    Dim GetTextLengthObj As GETTEXTLENGTHEX
    
    GetTextLengthObj.Flags = GTL_USECRLF Or GTL_PRECISE Or GTL_NUMCHARS
    GetTextLengthObj.Codepage = CP_UNICODE
    GetRTBLength = SendMessageW(hWnd, EM_GETTEXTLENGTHEX, VarPtr(GetTextLengthObj), 0)

End Function

Public Function GetRTBText(ByVal hWnd As Long, Optional ByVal OnlySelection As Boolean = False) As String

    Dim iChars As Long

    iChars = GetRTBLength(hWnd)

    If iChars > 0 Then
        Dim GetTextObj As GETTEXTEX
        GetTextObj.cb = (iChars + 1) * 2
        GetTextObj.Flags = GT_USECRLF
        If OnlySelection Then
            Dim sParam As Long, eParam As Long
            GetTextObj.Flags = GT_USECRLF Or GT_SELECTION
            GetTextSelection hWnd, sParam, eParam
            iChars = (eParam - sParam)
            GetTextObj.cb = (iChars + 1) * 2
        End If
        If iChars > 0 Then
            GetTextObj.Codepage = CP_UNICODE
            GetRTBText = String$(iChars, vbNullChar)
            iChars = SendMessageW(hWnd, EM_GETTEXTEX, VarPtr(GetTextObj), StrPtr(GetRTBText))
        Else
            GetRTBText = vbNullString
        End If
    Else
        GetRTBText = vbNullString
    End If

End Function

Public Sub SetSelectedRTBText(ByVal hWnd As Long, ByVal sText As String)

    Dim SetTextObj As SETTEXTEX

    SetTextObj.Flags = RTBW_SELECTION
    SetTextObj.Codepage = CP_UNICODE
    SendMessageW hWnd, EM_SETTEXTEX, ByVal VarPtr(SetTextObj), ByVal StrPtr(sText)

End Sub

Public Sub GetTextSelection(ByVal hWnd As Long, ByRef sParam As Long, ByRef eParam As Long)

    SendMessageW hWnd, EM_GETSEL, VarPtr(sParam), VarPtr(eParam)

End Sub

Public Sub SetTextSelection(ByVal hWnd As Long, ByVal sParam As Long, ByVal eParam As Long)

    SendMessageW hWnd, EM_SETSEL, sParam, eParam

End Sub

Public Sub SetClipboardText(ByVal sUniText As String, Optional ByVal hWnd As Long = 0&)
    ' Puts a VB string in the clipboard without converting it to ASCII.
    On Error GoTo ERROR_HANDLER

    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
    Const GMEM_MOVEABLE As Long = &H2
    Const GMEM_ZEROINIT As Long = &H40
    Const CF_UNICODETEXT As Long = &HD

    OpenClipboard hWnd
    EmptyClipboard
    iLen = LenB(sUniText) + 2&
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(iStrPtr)
    lstrcpy iLock, StrPtr(sUniText)
    GlobalUnlock iStrPtr
    SetClipboardData CF_UNICODETEXT, iStrPtr

ERROR_HANDLER:
    If Err.Number Then
        Call frmChat.AddChat(g_Color.ConsoleText, "Error: #" & Err.Number & ": " & Err.Description & _
                " in modClipboard.SetClipboardText().")
    End If

    CloseClipboard
End Sub

Public Function GetClipboardText(Optional ByVal hWnd As Long = 0&, Optional ByRef iLen As Long = 0&) As String
    ' Gets a UNICODE string from the clipboard and puts it in a standard VB string (which is UNICODE).
    On Error GoTo ERROR_HANDLER

    Dim iStrPtr As Long
    Dim iLock As Long
    Const CF_UNICODETEXT As Long = 13&

    GetClipboardText = vbNullString
    OpenClipboard hWnd
    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        iStrPtr = GetClipboardData(CF_UNICODETEXT)
        If iStrPtr Then
            iLock = GlobalLock(iStrPtr)
            iLen = GlobalSize(iStrPtr)
            iLen = iLen \ 2&
            GetClipboardText = String$(iLen - 1&, vbNullChar)
            lstrcpy StrPtr(GetClipboardText), iLock
            GlobalUnlock iStrPtr
        End If
    End If

ERROR_HANDLER:
    If Err.Number Then
        Call frmChat.AddChat(g_Color.ConsoleText, "Error: #" & Err.Number & ": " & Err.Description & _
                " in modClipboard.GetClipboardText().")
        GetClipboardText = vbNullString
    End If

    CloseClipboard
End Function

Public Function GetVScrollPosition(cnt As Control) As Long

    Dim lngVerticalPos As Long
    Dim Difference     As Long
    Dim Range          As Integer

    If (g_OSVersion.IsWin2000Plus()) Then

        GetScrollRange cnt.hWnd, SB_VERT, 0, Range

        lngVerticalPos = SendMessage(cnt.hWnd, EM_GETTHUMB, 0&, 0&)

        If ((lngVerticalPos = 0) And (Range > 0)) Then
            lngVerticalPos = 1
        End If

        Difference = ((lngVerticalPos + (cnt.Height / Screen.TwipsPerPixelY)) - _
            Range)

        ' In testing it appears that if the value I calcuate as Diff is negative,
        ' the scrollbar is not at the bottom.
        If (Difference < 0) Then
            GetVScrollPosition = lngVerticalPos
        End If

    End If

End Function

Public Function CanVScroll(ByVal hWnd As Long) As Boolean

    Dim Style As Long
    Style = GetWindowLong(hWnd, GWL_STYLE)
    CanVScroll = CBool((Style And WS_VSCROLL) = WS_VSCROLL)

End Function

Public Sub ScrollToBottom(ByVal hWnd As Long)

    SendMessage hWnd, EM_SCROLL, SB_BOTTOM, &H0

End Sub

Public Sub ScrollToTop(ByVal hWnd As Long)
    
    SendMessage hWnd, EM_SCROLL, SB_TOP, &H0

End Sub

Public Sub ScrollToCaret(ByVal hWnd As Long)

    SendMessageW hWnd, EM_SCROLLCARET, 0&, 0&

End Sub

Public Sub ScrollToPosition(ByVal hWnd As Long, ByVal lngVerticalPos As Long)

    SendMessage hWnd, WM_VSCROLL, _
            SB_THUMBPOSITION + &H10000 * lngVerticalPos, 0&

End Sub

