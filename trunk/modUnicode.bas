Attribute VB_Name = "modUnicode"
'StealthBot modUTF8 -- UTF-8 Conversion Module
'   Thanks to Skywing and Camel for much of the code in this module.

'http://forum.valhallalegends.com/phpbbs/index.php?board=18;action=display;threadid=1027&start=0
Option Explicit

Private Const WM_GETTEXT As Long = &HD
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_SETTEXT As Long = &HC

Private Const WM_USER As Long = &H400
Private Const EM_SETTEXTMODE As Long = WM_USER + 89
Private Const EM_SETTEXTEX As Long = WM_USER + 97
Private Const EM_GETTEXTEX As Long = WM_USER + 94
Private Const EM_GETTEXTLENGTHEX As Long = WM_USER + 95
Private Const CB_INSERTSTRING As Long = &H14A
Private Const CB_GETEDITSEL As Long = &H140
Private Const CB_SETEDITSEL As Long = &H142
Private Const TM_PLAINTEXT As Long = &H1
Private Const TM_MULTICODEPAGE As Long = &H32
' CodePage constant for current
Private Const CP_ACP     As Long = 0
' CodePage constant for Unicode
Private Const CP_UNICODE As Long = 1200&
' CodePage constant for UTF-8
Private Const CP_UTF8    As Long = 65001
Private Const GT_USECRLF As Long = 1&
Private Const GTL_USECRLF As Long = 1&
Private Const GTL_PRECISE As Long = 2&
Private Const GTL_NUMCHARS As Long = 8&
Private Const MB_ERR_INVALID_CHARS As Long = &H8

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

Private Declare Function SetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function GetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowTextLengthW Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal OleStr As Long, ByVal bLen As Long) As Long

Private Declare Sub PutMem4 Lib "msvbvm60" (Destination As Any, Value As Any)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetLastError Lib "Kernel32.dll" () As Long
Private Declare Function MultiByteToWideChar Lib "Kernel32.dll" (ByVal Codepage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "Kernel32.dll" (ByVal Codepage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long

Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long

Public Enum ECustomClipboardErrorConstant
   eccErrorBase = vbObjectError + 1048 + 521
   eccClipboardNotOpen
   eccCantOpenClipboard
End Enum

' Memory functions:
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_DISCARDED = &H4000
Private Const GMEM_FIXED = &H0
Private Const GMEM_INVALID_HANDLE = &H8000
Private Const GMEM_LOCKCOUNT = &HFF
Private Const GMEM_MODIFY = &H80
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_NOCOMPACT = &H10
Private Const GMEM_NODISCARD = &H20
Private Const GMEM_NOT_BANKED = &H1000
Private Const GMEM_NOTIFY = &H4000
Private Const GMEM_SHARE = &H2000
Private Const GMEM_VALID_FLAGS = &H7F72
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const GMEM_LOWER = GMEM_NOT_BANKED
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDest As String, pvSource As Any, ByVal cbCopy As Long)
Private Declare Function lstrlenW Lib "Kernel32.dll" (lpString As Any) As Long

'/*
' * Predefined Clipboard Formats
' */
Public Enum EPredefinedClipboardFormatConstants
   CF_TEXT = 1
   CF_BITMAP = 2
   CF_METAFILEPICT = 3
   CF_SYLK = 4
   CF_DIF = 5
   CF_TIFF = 6
   CF_OEMTEXT = 7
   CF_DIB = 8
   CF_PALETTE = 9
   CF_PENDATA = 10
   CF_RIFF = 11
   CF_WAVE = 12
   CF_UNICODETEXT = 13
   CF_ENHMETAFILE = 14
   ''#if(WINVER >= 0x0400)
   CF_HDROP = 15
   CF_LOCALE = 16
   CF_MAX = 17
   '#endif /* WINVER >= 0x0400 */
   CF_OWNERDISPLAY = &H80
   CF_DSPTEXT = &H81
   CF_DSPBITMAP = &H82
   CF_DSPMETAFILEPICT = &H83
   CF_DSPENHMETAFILE = &H8E
   '/*
   ' * "Private" formats don't get GlobalFree()'d
   ' */
   CF_PRIVATEFIRST = &H200
   CF_PRIVATELAST = &H2FF
   '/*
   ' * "GDIOBJ" formats do get DeleteObject()'d
   ' */
   CF_GDIOBJFIRST = &H300
   CF_GDIOBJLAST = &H3FF

End Enum

' Members:
Private m_lId()         As Long
Private m_sName()       As String
Private m_iCount        As Long
Private m_bClipboardIsOpen As Boolean
Private m_hWnd          As Long

' UTF8Encode
' STRING -> BYTE()
' expects a standard VB6 Unicode string
' returns a BYTE() for packet write (no null terminator in returned array)
Public Function UTF8Encode(ByRef str As String) As Byte()
    Dim UTF8Buffer() As Byte
    Dim UTF8Chars    As Long
    Dim lstr         As String
    Dim i            As Integer
    
    lstr = str
    
    ' grab Length of string after conversion
    UTF8Chars = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(lstr), Len(lstr), 0, 0, _
        vbNullString, 0)
    
    If (UTF8Chars = 0) Then
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

' UTF8EncodeS
' STRING -> STRING
' expects a standard VB6 Unicode string
' returns a VB6 Unicode string conversion of UTF-8 suitable for file and script output (UTF-8)
Public Function UTF8EncodeS(ByVal str As String) As String
    Dim arrStr() As Byte
    arrStr() = modUnicode.UTF8Encode(str)
    UTF8EncodeS = StrConv(arrStr(), vbUnicode)
End Function

' UTF8Decode
' BYTE() -> STRING
' expects a buffer of bytes in UTF-8 (null terminator(s) are stripped)
' returns a VB6 Unicode string
Public Function UTF8Decode(ByRef arrStr() As Byte, Optional LocaleID As Long = 1252) As String
    Dim sStr As String
    
    sStr = StrConv(arrStr, vbUnicode)
    sStr = KillNull(sStr)
    
    UTF8Decode = UTF8DecodeS(sStr, LocaleID)
End Function

' UTF8DecodeS
' STRING -> STRING
' expects a VB6 Unicode string that was not UTF-8 decoded [i.e. StrConv(byte[], vbUnicode)]
' returns a VB6 Unicode string, decoded
Public Function UTF8DecodeS(ByRef str As String, Optional LocaleID As Long = 1252) As String
    Dim UnicodeBuffer As String
    Dim UnicodeChars  As Long
    Dim lstr          As String
    
    lstr = str
    
    ' grab Length of string after conversion
    UnicodeChars = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, lstr, Len(lstr), _
        vbNullString, 0)
            
    If (UnicodeChars = 0) Then
        Exit Function
    End If
    
    ' initialize buffer
    UnicodeBuffer = String$(UnicodeChars * 2, vbNullChar)
    
    ' translate utf-8 string to unicode
    Call MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, lstr, Len(lstr), UnicodeBuffer, _
            UnicodeChars)
   
    ' translate from unicode to ansi
    UTF8DecodeS = StrConv(UnicodeBuffer, vbFromUnicode)
End Function

Public Sub UniTextCaptionSetText(ctrl As Object, sUniCaption As String)
    ' USAGE: UniCaption(SomeControl) = s
    
    ' This is known to work on Form, MDIForm, Checkbox, CommandButton, Frame, & OptionButton.
    ' Other controls are not known.
    
    ' As a tip, build your Unicode caption using ChrW.
    ' Also note the careful way we pass the string to the unicode API call to circumvent VB6's auto-ASCII-conversion.
    SetWindowTextW ctrl.hWnd, ByVal StrPtr(sUniCaption)
End Sub

Public Function UniTextCaptionGetLength(ctrl As Object) As Long
    UniTextCaptionGetLength = GetWindowTextLengthW(ctrl.hWnd)
End Function

Public Function UniTextCaptionGetText(ctrl As Object) As String
    ' USAGE: s = UniCaption(SomeControl)
    
    ' This is known to work on Form, MDIForm, Checkbox, CommandButton, Frame, & OptionButton.
    ' Other controls are not known.
    Dim lLen As Long
    Dim lPtr As Long
    
    lLen = UniTextCaptionGetLength(ctrl) ' Get length of caption.
    If lLen Then ' Must have length.
        'lPtr = SysAllocStringLen(0, lLen) ' Create a BSTR of that length.
        'UniTextCaptionGetText = String$(lLen + 2, vbNullChar)
        'PutMem4 ByVal VarPtr(UniTextCaptionGetText), ByVal lPtr ' Make the property return the BSTR.
        UniTextCaptionGetText = String$(lLen + 1, vbNullChar)
        GetWindowTextW ctrl.hWnd, StrPtr(UniTextCaptionGetText), lLen + 1
        'DefWindowProcW ctrl.hWnd, WM_GETTEXT, lLen + 1, ByVal lPtr ' Call the default Unicode window procedure to fill the BSTR.
    End If
End Function

Public Sub UniTextComboBoxAppendText(ctrl As ComboBox, sUniCaption As String)
    Dim Value As String
    
    Value = UniTextCaptionGetText(ctrl)
    
    UniTextCaptionSetText ctrl, Value & sUniCaption
    
    ctrl.selStart = UniTextCaptionGetLength(ctrl)
End Sub

Public Sub UniTextComboBoxSetSelectedText(ctrl As ComboBox, sUniCaption As String)
    Dim Value As String
    Dim lStart As Long
    Dim lEnd As Long
    
    Value = UniTextCaptionGetText(ctrl)
    
    'SendMessageW ctrl.hWnd, CB_GETEDITSEL, ByVal VarPtr(lStart), ByVal VarPtr(lEnd)
    
    lStart = ctrl.selStart
    lEnd = ctrl.SelLength
    
    Mid$(Value, lStart, lEnd) = sUniCaption
    
    UniTextCaptionSetText ctrl, Value
    
    ctrl.selStart = lStart
    ctrl.SelLength = Len(sUniCaption)
    
    'SendMessageW ctrl.hWnd, CB_SETEDITSEL, 0&, 0
End Sub

Public Function UniTextComboBoxGetSelectedText(ctrl As ComboBox) As String
    Dim Value As String

    Dim lStart As Long
    Dim lEnd As Long
    
    lStart = ctrl.selStart
    lEnd = ctrl.SelLength
    
    Value = UniTextCaptionGetText(ctrl)
    
    UniTextComboBoxGetSelectedText = Mid$(Value, lStart, lEnd)
End Function

Public Sub UniTextComboBoxAddItem(ctrl As ComboBox, sUniCaption As String, Index As Long)
    SendMessageW ctrl.hWnd, CB_INSERTSTRING, Index, ByVal StrPtr(sUniCaption)
End Sub

'Public Sub SetupRichTextboxForUnicode(rtb As RichTextBox)
'    SendMessage rtb.hWnd, EM_SETTEXTMODE, TM_PLAINTEXT, 0 ' Set the control to use "plain text" mode so RTF isn't interpreted.
'End Sub

Public Sub UniTextRichTextAppendText(rtb As RichTextBox, sUniText As String)
    Dim stUnicode As SETTEXTEX
    
    stUnicode.Flags = RTBW_SELECTION
    stUnicode.Codepage = CP_UNICODE
    SendMessageW rtb.hWnd, EM_SETTEXTEX, ByVal VarPtr(stUnicode), ByVal StrPtr(sUniText)
End Sub

Public Function UniTextRichTextGetLength(rtb As RichTextBox) As Long
    Dim gtlUnicode As GETTEXTLENGTHEX
    
    gtlUnicode.Flags = GTL_USECRLF Or GTL_PRECISE 'Or GTL_NUMCHARS
    gtlUnicode.Codepage = CP_UNICODE
    UniTextRichTextGetLength = SendMessageW(rtb.hWnd, EM_GETTEXTLENGTHEX, VarPtr(gtlUnicode), 0)
End Function

Public Function UniTextRichTextGetText(rtb As RichTextBox) As String
    Dim gtUnicode As GETTEXTEX
    Dim iChars As Long
    
    iChars = UniTextRichTextGetLength(rtb)
    
    If iChars > 0 Then
        gtUnicode.cb = (iChars + 1) * 2
        gtUnicode.Flags = GT_USECRLF
        gtUnicode.Codepage = CP_UNICODE
        UniTextRichTextGetText = String$(iChars, vbNullChar)
        SendMessageW rtb.hWnd, EM_GETTEXTEX, VarPtr(gtUnicode), StrPtr(UniTextRichTextGetText)
    Else
        UniTextRichTextGetText = vbNullString
    End If
End Function

' Fixed font issue when an element was only 1 character long -Pyro (9/28/08)
' Fixed issue with displaying null text.

' I changed the location of where fontStr was being declared. I can't figure out why it makes any difference, but _
    when I have it declared above I get memory errors, my IDE crashes, I runtime erro 380 and _
  - Error (#-2147417848): Method 'SelFontName' of object 'IRichText' failed in DisplayRichText().
'I believe it's something to do with the subclassing overwriting the memory, but it only occurs when run from the IDE. - FrOzeN

Public Sub DisplayRichText(ByRef rtb As RichTextBox, ByRef saElements() As Variant)
    On Error GoTo ERROR_HANDLER
   
    Dim arr()          As Variant
    Dim s              As String
    Dim L              As Long
    Dim lngVerticalPos As Long
    Dim Diff           As Long
    Dim i              As Long
    Dim intRange       As Long
    Dim blUnlock       As Boolean
    Dim LogThis        As Boolean
    Dim Length         As Long
    Dim ConvertIndex   As Long
    Dim str            As String
    Dim arrCount       As Long
    Dim selStart       As Long
    Dim SelLength      As Long
    Dim blnHasFocus    As Boolean
    Dim blnAtEnd       As Boolean
    
    Static RichTextErrorCounter As Integer

    ' *****************************************
    '              SANITY CHECKS
    ' *****************************************
    
    ' empty array
    If LBound(saElements) > UBound(saElements) Then
        Exit Sub
    End If
    
    ' convert AddChat Text
    '      to AddChat DefaultFont, DefaultColor, Text
    If LBound(saElements) = UBound(saElements) Then
        ReDim Preserve arr(0 To 2) As Variant
        
        arr(0) = rtb.Font.Name
        arr(1) = vbWhite
        arr(2) = saElements(LBound(saElements))
        
        saElements() = arr()
    End If
    
    ' if first element is numeric
    If (IsNumeric(saElements(0))) Then
        ConvertIndex = 2
        
        ' convert AddChat Color, Text, Color, Text, ...
        '      to AddChat DefaultFont, Color, Text, DefaultFont, Color, Text, ...
        For i = LBound(saElements) To UBound(saElements) Step 2
            ReDim Preserve arr(0 To ConvertIndex) As Variant
            
            arr(ConvertIndex) = saElements(i + 1)
            arr(ConvertIndex - 1) = saElements(i)
            arr(ConvertIndex - 2) = rtb.Font.Name
            
            ConvertIndex = ConvertIndex + 3
        Next i
        
        saElements() = arr()
    End If
    
    ' store full text length
    rtbChatLength = UniTextRichTextGetLength(rtb)
    
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
        Length = Length + Len(KillNull(saElements(i + 2)))
    Next i
    
    ' input must have non-zero length
    If (Length = 0) Then
        Exit Sub
    End If
    
    If ((BotVars.LockChat = False) Or (rtb <> frmChat.rtbChat)) Then
        
        ' store rtb carat and whether rtb has focus
        With rtb
            selStart = .selStart
            SelLength = .SelLength
            blnHasFocus = (rtb.Parent.ActiveControl Is rtb And rtb.Parent.WindowState <> vbMinimized)
            ' whether it's at the end or within one vbCrLf of the end
            blnAtEnd = (selStart >= rtbChatLength - 2)
        End With
        
        ' is the RTB at the bottom?
        lngVerticalPos = IsScrolling(rtb)
    
        If (lngVerticalPos) Then
            rtb.Visible = False
        
            ' below causes smooth scrolling, but also screen flickers :(
            'LockWindowUpdate rtb.hWnd
        
            blUnlock = True
        End If
        
        ' how to log this event
        If (rtb = frmChat.rtbChat) Then
            LogThis = (BotVars.Logging > 0)
        ElseIf (rtb = frmChat.rtbWhispers) Then
            LogThis = (BotVars.Logging > 0)
        End If
        
        ' remove from backlog if overflow
        If ((BotVars.MaxBacklogSize) And (rtbChatLength >= BotVars.MaxBacklogSize)) Then
            If (blUnlock = False) Then
                rtb.Visible = False
            
                ' below causes smooth scrolling, but also screen flickers :(
                'LockWindowUpdate rtb.hWnd
            End If
        
            With rtb
                .selStart = 0
                .SelLength = InStr(1, UniTextRichTextGetText(rtb), vbLf, vbBinaryCompare)
                ' remove line from stored selection
                selStart = selStart - .SelLength
                ' if selection included part of what was removed, add negative start point
                ' to length to get difference length and start selection at 0
                If selStart < 0 Then
                    SelLength = SelLength + selStart
                    selStart = 0
                    ' if new length is negative, then the selection is now gone, so selection
                    ' length should be 0
                    If SelLength < 0 Then SelLength = 0
                End If
                .SelFontName = rtb.Font.Name
                .SelFontSize = rtb.Font.Size
                .SelText = ""
            End With
            
            If (blUnlock = False) Then
                rtb.Visible = True
            
                ' below causes smooth scrolling, but also screen flickers :(
                'LockWindowUpdate &H0
            End If
        End If
        
        ' place timestamp
        s = GetTimeStamp()
        
        If LenB(s) > 0 Then
            With rtb
                .selStart = UniTextRichTextGetLength(rtb)
                .SelLength = 0
                .SelFontName = rtb.Font.Name
                .SelFontSize = rtb.Font.Size
                .SelBold = False
                .SelItalic = False
                .SelUnderline = False
                .SelColor = RTBColors.TimeStamps
                UniTextRichTextAppendText rtb, s
                '.SelText = s
                '.SelLength = Len(s)
            End With
        End If
        
        ' place each element
        For i = LBound(saElements) To UBound(saElements) Step 3
            If (InStr(1, saElements(i + 2), ChrW$(0), vbBinaryCompare) > 0) Then
                KillNull saElements(i + 2)
            End If
        
            If ((StrictIsNumeric(saElements(i + 1))) And (Len(saElements(i + 2)) > 0)) Then
                's = EscapeRichText(saElements(i + 2))
                
            
                L = InStr(1, saElements(i + 2), "{\rtf", vbTextCompare)
                
                While (L > 0)
                    Mid$(saElements(i + 2), L + 1, 1) = "/"
                    
                    L = InStr(1, saElements(i + 2), "{\rtf", vbTextCompare)
                Wend
            
                L = Len(rtb.Text)
                
                s = saElements(i + 2) ' & Left$(vbCrLf, -2 * CLng((i + 2) = UBound(saElements)))
                
                With rtb
                    .selStart = UniTextRichTextGetLength(rtb)
                    .SelLength = 0
                    .SelFontName = saElements(i)
                    .SelColor = saElements(i + 1)
                    UniTextRichTextAppendText rtb, s
                    '.SelText = s
                    str = str & s
                End With
            End If
        Next i
        
        With rtb
            .selStart = UniTextRichTextGetLength(rtb)
            .SelLength = 0
            UniTextRichTextAppendText rtb, vbCrLf
        End With
        
        If (LogThis) Then
            If (rtb = frmChat.rtbChat) Then
                g_Logger.WriteChat str
            ElseIf (rtb = frmChat.rtbWhispers) Then
                g_Logger.WriteWhisper str
            End If
        End If

        ColorModify rtb, L

        If (blUnlock) Then
            SendMessage rtb.hWnd, WM_VSCROLL, SB_THUMBPOSITION + &H10000 * lngVerticalPos, 0&
                
            rtb.Visible = True
                
            ' below causes smooth scrolling, but also screen flickers :(
            'LockWindowUpdate &H0
        End If
        
        With rtb
            ' if has focus
            If blnHasFocus Then
                ' restore carat location and selection if not previously at end
                If Not blnAtEnd Then
                    .selStart = selStart
                    .SelLength = SelLength
                End If
                
                ' restore focus
                '.SetFocus
            End If
        End With
    End If
    
    RichTextErrorCounter = 0
    
    Exit Sub
    
ERROR_HANDLER:

    RichTextErrorCounter = RichTextErrorCounter + 1
    If RichTextErrorCounter > 2 Then
        RichTextErrorCounter = 0
        Exit Sub
    End If
    
    If (Err.Number = 13 Or Err.Number = 91) Then
        Exit Sub
    End If

    frmChat.AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.description & " in DisplayRichText()."
    
    Exit Sub
    
End Sub

Public Function IsScrolling(ByRef rtb As RichTextBox) As Long

    Dim lngVerticalPos As Long
    Dim difference     As Long
    Dim range          As Integer

    If (g_OSVersion.IsWin2000Plus()) Then

        GetScrollRange rtb.hWnd, SB_VERT, 0, range
        
        lngVerticalPos = SendMessage(rtb.hWnd, EM_GETTHUMB, 0&, 0&)
        
        If ((lngVerticalPos = 0) And (range > 0)) Then
            lngVerticalPos = 1
        End If

        difference = ((lngVerticalPos + (rtb.Height / Screen.TwipsPerPixelY)) - _
            range)

        ' In testing it appears that if the value I calcuate as Diff is negative,
        ' the scrollbar is not at the bottom.
        If (difference < 0) Then
            IsScrolling = lngVerticalPos
        End If
        
    End If

End Function


'// COLORMODIFY - where L is passed as the start position of the text to be checked
Public Sub ColorModify(ByRef rtb As RichTextBox, ByRef L As Long)
    Dim i As Long
    Dim s As String
    Dim temp As Long
    Dim selStart As Long
    Dim SelLength As Long
    
    If L = 0 Then L = 1
    
    temp = L
    
    With rtb
        ' store previous selstart and len
        selStart = .selStart
        SelLength = .SelLength
        
        If InStr(temp, .Text, "ÿc", vbTextCompare) > 0 Then
            .Visible = False
            Do
                i = InStr(temp, .Text, "ÿc", vbTextCompare)
                
                If StrictIsNumeric(Mid$(.Text, i + 2, 1)) Then
                    s = GetColorVal(Mid$(.Text, i + 2, 1))
                    .selStart = i - 1
                    .SelLength = 3
                    .SelText = vbNullString
                    .selStart = i - 1
                    .SelLength = Len(.Text) + 1 - i
                    .SelColor = s
                Else
                    Select Case Mid$(.Text, i + 2, 1)
                        Case "i"
                            .selStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .selStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            If .SelItalic = True Then
                                .SelItalic = False
                            Else
                                .SelItalic = True
                            End If
                            
                        Case "b", "."       'BOLD
                            .selStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .selStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            If .SelBold = True Then
                                .SelBold = False
                            Else
                                .SelBold = True
                            End If
                            
                        Case "u", "."       'underline
                            .selStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .selStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            If .SelUnderline = True Then
                                .SelUnderline = False
                            Else
                                .SelUnderline = True
                            End If
                            
                        Case ";"
                            .selStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .selStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            .SelColor = HTMLToRGBColor("8D00CE")    'Purple
                            
                        Case ":"
                            .selStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .selStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            .SelColor = 186408      '// Lighter green
                            
                        Case "<"
                            .selStart = i - 1
                            .SelLength = 3
                            .SelText = vbNullString
                            .selStart = i - 1
                            .SelLength = Len(.Text) + 1 - 1
                            .SelColor = HTMLToRGBColor("00A200")    'Dark green
                        'Case Else: Debug.Print s
                    End Select
                End If
                temp = temp + 1
                
            Loop While InStr(temp, .Text, "ÿc", vbTextCompare) > 0
            .Visible = True
        End If
        
        '// Check for SC color codes
        temp = L
        
        If InStr(temp, .Text, "Á", vbBinaryCompare) > 0 Then
            .Visible = False
            Do
                i = InStr(temp, .Text, "Á", vbBinaryCompare)
                s = GetScriptColorString(Mid$(.Text, i + 1, 1))
                
                If Len(s) > 0 Then
                    .Visible = False
                    .selStart = i - 1
                    .SelLength = 2
                    .SelText = vbNullString
                    .selStart = i - 1
                    .SelLength = Len(.Text) + 1 - 1
                    .SelColor = s
                    .Visible = True
                End If
                
                temp = temp + 1
                
            Loop While InStr(temp, .Text, "Á", vbBinaryCompare) > 0
            .Visible = True
        End If
        
        ' restore previous selstart and len
        .selStart = selStart
        .SelLength = SelLength
    End With
End Sub

Public Function GetScriptColorString(ByVal scCC As String) As String
    Select Case Asc(scCC)
        Case Asc("Q"): GetScriptColorString = RGB(93, 93, 93)       'Grey
        Case Asc("R"): GetScriptColorString = RGB(30, 224, 54)      'Green
        Case Asc("Z"), Asc("X"), Asc("S"): GetScriptColorString = RGB(160, 169, 116)    'Yellow
        Case Asc("Y"), Asc("["), Asc("Y"): GetScriptColorString = RGB(231, 38, 82)      'Red
        Case Asc("V"), Asc("@"): GetScriptColorString = RGB(98, 77, 232)      'Blue
        Case Asc("W"), Asc("P"): GetScriptColorString = vbWhite               'White
        Case Asc("T"), Asc("U"), Asc("V"): GetScriptColorString = HTMLToRGBColor("00CCCC") 'cyan/teal
        Case Else: GetScriptColorString = vbNullString
    End Select
End Function

Public Function GetColorVal(ByVal d2CC As String) As String
    Select Case CInt(d2CC)
        Case 1: GetColorVal = HTMLToRGBColor("CE3E3E")  'Red
        Case 2: GetColorVal = HTMLToRGBColor("00CE00")  'Green
        Case 3: GetColorVal = HTMLToRGBColor("44409C")  'Blue
        Case 4: GetColorVal = HTMLToRGBColor("A19160")  'Gold
        Case 5: GetColorVal = HTMLToRGBColor("555555")  'Grey
        Case 6: GetColorVal = HTMLToRGBColor("080808")  'Black
        Case 7: GetColorVal = HTMLToRGBColor("A89D65")  'Gold
        Case 8: GetColorVal = HTMLToRGBColor("CE8800")  'Gold-Orange
        Case 9: GetColorVal = HTMLToRGBColor("CECE51")  'Light Yellow
        Case 0: GetColorVal = HTMLToRGBColor("FFFFFF")  'White
    End Select
End Function


'Purppose: Wrap GetTextData CF_UNICODETEXT into single call.
Public Function ClipboardGetUnicodeText() As String
   Dim sText  As String
   ClipboardOpen 0
   GetTextData CF_UNICODETEXT, sText
   ClipboardGetUnicodeText = sText
   ClipboardClose
End Function

'Purppose: Wrap SetTextData CF_UNICODETEXT into single call.
Public Function ClipboardSetUnicodeText(ByVal sText As String) As Boolean
   ClipboardOpen 0
   ClearClipboard
   ClipboardSetUnicodeText = SetTextData(CF_UNICODETEXT, sText)
   ClipboardClose
End Function

Private Function GetTextData(ByVal lFormatId As Long, ByRef sTextOut As String) As Boolean
   ' Returns a string containing text on the clipboard for
   ' format lFormatID:
   Dim lHwndCache       As Long
   Dim bData()          As Byte
   Dim sR               As String

   If (lFormatId = CF_TEXT) Or (lFormatId = CF_UNICODETEXT) Or (lFormatId = 49159) Then
      If (GetBinaryData(lFormatId, bData())) Then
         If (lFormatId = CF_TEXT) Then
            sTextOut = StrConv(bData, vbUnicode)
         Else
            sTextOut = KillNull(CStr(bData))
         End If
         GetTextData = True
      End If
   Else
      If (GetBinaryData(lFormatId, bData())) Then
         sTextOut = StrConv(bData, vbUnicode)
         GetTextData = True
      End If
   End If
End Function

Private Function GetClipboardMemoryHandle(ByVal lFormatId As Long) As Long
   If pbNotReady() Then Exit Function

   ' If the format id is there:
   If (IsClipboardDataAvailableForFormat(lFormatId)) Then
      ' Get the global memory handle to the clipboard data:
      GetClipboardMemoryHandle = GetClipboardData(lFormatId)
   End If
End Function

Property Get IsClipboardDataAvailableForFormat(ByVal lFormatId As Long)
   ' Returns whether data is available for a given format id:
   Dim lR               As Long
   lR = IsClipboardFormatAvailable(lFormatId)
   IsClipboardDataAvailableForFormat = (lR <> 0)
End Property

Private Function GetBinaryData(ByVal lFormatId As Long, ByRef bData() As Byte) As Boolean
   ' Returns a byte array containing binary data on the clipboard for
   ' format lFormatID:
   Dim hMem             As Long, lSize As Long, lPtr As Long

   ' Ensure the return array is clear:
   Erase bData

   hMem = GetClipboardMemoryHandle(lFormatId)
   ' If success:
   If (hMem <> 0) Then
      ' Get the size of this memory block:
      lSize = GlobalSize(hMem)
      ' Get a pointer to the memory:
      lPtr = GlobalLock(hMem)
      If (lSize > 0) Then
         ' Resize the byte array to hold the data:
         ReDim bData(0 To lSize - 1) As Byte
         ' Copy from the pointer into the array:
         CopyMemory bData(0), ByVal lPtr, lSize
      End If
      ' Unlock the memory block:
      GlobalUnlock hMem
      ' Success:
      GetBinaryData = (lSize > 0)
      ' Don't free the memory - it belongs to the clipboard.
   End If
End Function

Public Function ClipboardSetBinaryData(ByVal lFormatId As Long, ByRef bData() As Byte) As Boolean
   ' Puts the binary data contained in bData() onto the clipboard under
   ' format lFormatID:
   Dim lSize            As Long
   Dim lPtr             As Long
   Dim hMem             As Long

   If pbNotReady() Then Exit Function

   ' Determine the size of the binary data to write:
   lSize = UBound(bData) - LBound(bData) + 1
   ' Generate global memory to hold this:
   hMem = GlobalAlloc(GMEM_DDESHARE, lSize)
   If (hMem <> 0) Then
      ' Get pointer to the memory block:
      lPtr = GlobalLock(hMem)
      ' Copy the data into the memory block:
      CopyMemory ByVal lPtr, bData(LBound(bData)), lSize
      ' Unlock the memory block.
      GlobalUnlock hMem

      ' Now set the clipboard data:
      If (SetClipboardData(lFormatId, hMem) <> 0) Then
         ' Success:
         ClipboardSetBinaryData = True
      End If
   End If
   ' We don't free the memory because the clipboard takes
   ' care of that now.

End Function

Private Function pbNotReady() As Boolean
   ' Determines whether a call to Get or Set Data on the
   ' clipboard will work.
   'was If Not (m_bClipboardIsOpen) Or (m_hWnd = 0) Then
   If Not ((m_bClipboardIsOpen) Or (m_hWnd = 0)) Then
      Debug.Assert (1 = 0)
      Err.Raise eccClipboardNotOpen, App.EXEName & ".cCustomClipboard", "Attempt to access the clipboard when clipboard not Open."
      pbNotReady = True
   End If
End Function

Private Function SetTextData(ByVal lFormatId As Long, ByVal sText As String) As Boolean
   Dim bData()          As Byte

   ' Sets the text in sText onto the clipboard under format lFormatID:
   If (Len(sText) > 0) Then
      sText = sText & ChrW$(&H0)
      bData = sText
      SetTextData = ClipboardSetBinaryData(lFormatId, bData())
   End If
End Function

Private Sub ClearClipboard()
   ' Clears all data in the clipboard, and also takes ownership
   ' of the clipboard.  This method will fail
   ' unless OpenClipboard has been called first.
   If (pbNotReady()) Then Exit Sub
   EmptyClipboard
End Sub

Private Sub ClipboardClose()
   ' Closes the clipboard if this class has it open:
   If (m_bClipboardIsOpen) Then
      CloseClipboard
      m_bClipboardIsOpen = False
      m_hWnd = 0
   End If
End Sub

Private Function ClipboardOpen(ByVal hWndOwner As Long) As Boolean
   Dim lR               As Long
   ' Opens the clipboard:
   
   lR = OpenClipboard(hWndOwner)
   If (lR > 0) Then
      m_hWnd = hWndOwner
      m_bClipboardIsOpen = True
      ClipboardOpen = True
   Else
      m_hWnd = 0
      m_bClipboardIsOpen = False
      Err.Raise eccCantOpenClipboard, App.EXEName & ".cCustomClipboard", "Unable to Open Clipboard."
   End If
End Function

'Purpose: Unicode aware MsgBox
'Overrides Vb6 MsgBox. HelpFile/Context not supported
Function MsgBox(Prompt As String, _
   Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
   Optional Title As String) As VbMsgBoxResult

   Dim WshShell As Object
   Set WshShell = CreateObject("WScript.Shell")
   MsgBox = WshShell.Popup(Prompt, 0&, Title, Buttons)
   Set WshShell = Nothing
End Function


