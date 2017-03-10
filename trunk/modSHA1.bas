Attribute VB_Name = "modSHA1"
Option Explicit

'This code is really messy. It was taken from http://vb.wikia.com/wiki/SHA-1.bas by Andy (RealityRipple), and _
    then he added a few functions to it. After cleaning up the other warden code I got over it and I'm just _
    going to leave this the way it is. I fixed up the tabbing a bit. - FrOzeN

Private Type FourBytes
    a As Byte
    b As Byte
    c As Byte
    d As Byte
End Type

Private Type OneLong
    L As Long
End Type

Public Function HashPassword(ByVal Password As String, Optional ByVal Version As enuSHA1Type = shaBrokenROL) As String
    Dim Data()   As Byte
    Dim Result() As Byte

    Data = StringToByteArr(Password)
    Call CalculateSHA1(Data, Result, Version)
    HashPassword = ByteArrToString(Result)
End Function

Public Function DoubleHashPassword(ByVal Password As String, ByVal ClientToken As Long, ByVal ServerToken As Long, Optional ByVal Version As enuSHA1Type = shaBrokenROL) As String
    Dim Data()   As Byte
    Dim Result() As Byte
    Dim Buffer   As New clsDataBuffer

    Data = StringToByteArr(Password)
    Call CalculateSHA1(Data, Result, Version)

    With Buffer
        .InsertDWord ClientToken
        .InsertDWord ServerToken
        .InsertByteArr Result
        Call CalculateSHA1(.GetDataAsByteArr, Result, Version)
    End With
    Set Buffer = Nothing

    DoubleHashPassword = ByteArrToString(Result)
End Function

Public Sub CalculateSHA1(ByRef Message() As Byte, ByRef Result() As Byte, Optional ByVal Version As enuSHA1Type = shaStandard)
    Dim h1 As Long, h2 As Long, h3 As Long, h4 As Long, h5 As Long

    If Version = shaBrokenROL Then
        XSHA1 Message, &H5A827999, &H6ED9EBA1, &H8F1BBCDC, &HCA62C1D6, h1, h2, h3, h4, h5

        ReDim Result(0 To 19)

        CopyMemory Result(0), h1, 4
        CopyMemory Result(4), h2, 4
        CopyMemory Result(8), h3, 4
        CopyMemory Result(12), h4, 4
        CopyMemory Result(16), h5, 4
    Else
        SHA1 Message, &H5A827999, &H6ED9EBA1, &H8F1BBCDC, &HCA62C1D6, h1, h2, h3, h4, h5

        ReDim Result(0 To 19)

        If Version = shaStandard Then
            CopyMemory Result(0), ntohl(h1), 4
            CopyMemory Result(4), ntohl(h2), 4
            CopyMemory Result(8), ntohl(h3), 4
            CopyMemory Result(12), ntohl(h4), 4
            CopyMemory Result(16), ntohl(h5), 4
        ElseIf Version = shaStandardRevEndian Then
            CopyMemory Result(0), h1, 4
            CopyMemory Result(4), h2, 4
            CopyMemory Result(8), h3, 4
            CopyMemory Result(12), h4, 4
            CopyMemory Result(16), h5, 4
        End If
    End If
End Sub

Private Sub SHA1(Message() As Byte, ByVal Key1 As Long, ByVal Key2 As Long, ByVal Key3 As Long, ByVal Key4 As Long, h1 As Long, h2 As Long, h3 As Long, h4 As Long, h5 As Long)
    Dim U As Long, P As Long
    Dim FB As FourBytes, OL As OneLong
    Dim i As Integer
    Dim W(80) As Long
    Dim a As Long, b As Long, c As Long, d As Long, e As Long
    Dim T As Long
    
    h1 = &H67452301: h2 = &HEFCDAB89: h3 = &H98BADCFE: h4 = &H10325476: h5 = &HC3D2E1F0
    
    U = UBound(Message) + 1: OL.L = U32ShiftLeft3(U): a = U \ &H20000000: LSet FB = OL 'U32ShiftRight29(U)
    
    ReDim Preserve Message(0 To (U + 8 And -64) + 63)
    Message(U) = 128
    
    U = UBound(Message)
    Message(U - 4) = a
    Message(U - 3) = FB.d
    Message(U - 2) = FB.c
    Message(U - 1) = FB.b
    Message(U) = FB.a
    
    While P < U
        For i = 0 To 15
            FB.d = Message(P)
            FB.c = Message(P + 1)
            FB.b = Message(P + 2)
            FB.a = Message(P + 3)
            LSet OL = FB
            W(i) = OL.L
            P = P + 4
        Next i
        
        For i = 16 To 79
            W(i) = U32RotateLeft1(W(i - 3) Xor W(i - 8) Xor W(i - 14) Xor W(i - 16))
        Next i
        
        a = h1: b = h2: c = h3: d = h4: e = h5
        
        For i = 0 To 19
            T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), W(i)), Key1), ((b And c) Or ((Not b) And d)))
            e = d: d = c: c = U32RotateLeft30(b): b = a: a = T
        Next i
        
        For i = 20 To 39
            T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), W(i)), Key2), (b Xor c Xor d))
            e = d: d = c: c = U32RotateLeft30(b): b = a: a = T
        Next i
        
        For i = 40 To 59
            T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), W(i)), Key3), ((b And c) Or (b And d) Or (c And d)))
            e = d: d = c: c = U32RotateLeft30(b): b = a: a = T
        Next i
        
        For i = 60 To 79
            T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), W(i)), Key4), (b Xor c Xor d))
            e = d: d = c: c = U32RotateLeft30(b): b = a: a = T
        Next i
        
        h1 = U32Add(h1, a): h2 = U32Add(h2, b): h3 = U32Add(h3, c): h4 = U32Add(h4, d): h5 = U32Add(h5, e)
    Wend
End Sub

Private Sub XSHA1(Message() As Byte, ByVal Key1 As Long, ByVal Key2 As Long, ByVal Key3 As Long, ByVal Key4 As Long, h1 As Long, h2 As Long, h3 As Long, h4 As Long, h5 As Long)
    Dim U As Long, P As Long
    Dim FB As FourBytes, OL As OneLong
    Dim i As Integer
    Dim W(80) As Long
    Dim a As Long, b As Long, c As Long, d As Long, e As Long
    Dim T As Long
    
    h1 = &H67452301: h2 = &HEFCDAB89: h3 = &H98BADCFE: h4 = &H10325476: h5 = &HC3D2E1F0
    
    U = UBound(Message) + 1: OL.L = U32ShiftLeft3(U): a = U \ &H20000000: LSet FB = OL 'U32ShiftRight29(U)
    
    ReDim Preserve Message(0 To (U + 8 And -64) + 63)
    'Message(U) = 128
    
    U = UBound(Message)
    'Message(U - 4) = a
    'Message(U - 3) = FB.d
    'Message(U - 2) = FB.c
    'Message(U - 1) = FB.b
    'Message(U) = FB.a
    
    While P < U
        For i = 0 To 15
            FB.d = Message(P + 3)
            FB.c = Message(P + 2)
            FB.b = Message(P + 1)
            FB.a = Message(P)
            LSet OL = FB
            W(i) = OL.L
            P = P + 4
        Next i
        
        For i = 16 To 79
            W(i) = U32RotateLeftN(1, W(i - 3) Xor W(i - 8) Xor W(i - 14) Xor W(i - 16))
        Next i
        
        a = h1: b = h2: c = h3: d = h4: e = h5
        
        For i = 0 To 19
            T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), W(i)), Key1), ((b And c) Or ((Not b) And d)))
            e = d: d = c: c = U32RotateLeft30(b): b = a: a = T
        Next i
        
        For i = 20 To 39
            T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), W(i)), Key2), (b Xor c Xor d))
            e = d: d = c: c = U32RotateLeft30(b): b = a: a = T
        Next i
        
        For i = 40 To 59
            T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), W(i)), Key3), ((b And c) Or (b And d) Or (c And d)))
            e = d: d = c: c = U32RotateLeft30(b): b = a: a = T
        Next i
        
        For i = 60 To 79
            T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), W(i)), Key4), (b Xor c Xor d))
            e = d: d = c: c = U32RotateLeft30(b): b = a: a = T
        Next i
        
        h1 = U32Add(h1, a): h2 = U32Add(h2, b): h3 = U32Add(h3, c): h4 = U32Add(h4, d): h5 = U32Add(h5, e)
    Wend
End Sub

Private Function U32Add(ByVal a As Long, ByVal b As Long) As Long
    If (a Xor b) < 0 Then
        U32Add = a + b
    Else
        U32Add = (a Xor &H80000000) + b Xor &H80000000
    End If
End Function

Private Function U32ShiftLeft3(ByVal a As Long) As Long
    U32ShiftLeft3 = (a And &HFFFFFFF) * 8
    If a And &H10000000 Then U32ShiftLeft3 = U32ShiftLeft3 Or &H80000000
End Function

Private Function U32ShiftRight29(ByVal a As Long) As Long
    U32ShiftRight29 = (a And &HE0000000) \ &H20000000 And 7
End Function

Private Function U32RotateLeft1(ByVal a As Long) As Long
    U32RotateLeft1 = (a And &H3FFFFFFF) * 2
    If a And &H40000000 Then U32RotateLeft1 = U32RotateLeft1 Or &H80000000
    If a And &H80000000 Then U32RotateLeft1 = U32RotateLeft1 Or 1
End Function

Private Function U32RotateLeft5(ByVal a As Long) As Long
    U32RotateLeft5 = (a And &H3FFFFFF) * 32 Or (a And &HF8000000) \ &H8000000 And 31
    If a And &H4000000 Then U32RotateLeft5 = U32RotateLeft5 Or &H80000000
End Function

Private Function U32RotateLeft30(ByVal a As Long) As Long
    U32RotateLeft30 = (a And 1) * &H40000000 Or (a And &HFFFC) \ 4 And &H3FFFFFFF
    If a And 2 Then U32RotateLeft30 = U32RotateLeft30 Or &H80000000
End Function

Private Function U32RotateLeftN(ByVal a As Long, ByVal Shift As Long) As Long
    Dim RValue As Long, LValue As Long

    Shift = Shift Mod 32
    If Shift < 0 Then
        Shift = Shift + 32
    ElseIf Shift = 0 Then
        U32RotateLeftN = a
        Exit Function
    End If

    ' probably should make this do math rather than recursion, but it can't go deeper than 31 deep...

    'If Shift = 1 Then
    '    RValue = (a And (2 ^ (Shift + 1) - 1)) * &H80000000
    'Else
    '    RValue = (a And (2 ^ (Shift + 1) - 1)) * (2 ^ (32 - Shift))
    'End If

    'If a < 0 Then
    '    LValue = (a And &H7FFFFFFF) \ (2 ^ Shift) Or 2 ^ (31 - Shift)
    'Else
    '    LValue = a \ (2 ^ Shift)
    'End If

    'U32RotateLeftN = RValue Or LValue
    U32RotateLeftN = U32RotateLeftN(U32RotateLeft1(a), Shift - 1)
End Function

