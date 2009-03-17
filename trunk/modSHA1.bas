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

'I added this function as a quick solution and better named method to call. The code it uses is still pretty bad. - FrOzeN
'Public Sub Warden_SHA1(Destination() As Byte, ByRef Source() As Byte)
Public Function Warden_SHA1(ByVal data As String) As String
    Dim arrData() As Byte
    Dim strHash As String, arrHash() As Byte
    Dim arrRet(20) As Byte
    
    arrData = StrConv(data, vbFromUnicode)
    
    strHash = SHA1b(data)
    arrHash = StrConv(strHash, vbFromUnicode)
    
    Call CopyMemory(arrRet(0), arrHash(0), 20)
    Warden_SHA1 = StrConv(arrRet, vbUnicode)
End Function

Private Function SHA1b(ByVal sIn As String) As String
    Dim bIn() As Byte
    StrToByteArray sIn, bIn
    SHA1b = SHAIt(bIn)
End Function

Private Function SHAIt(Message() As Byte) As String
    Dim h1 As Long
    Dim h2 As Long
    Dim h3 As Long
    Dim h4 As Long
    Dim h5 As Long
    
    DefaultSHA1 Message, h1, h2, h3, h4, h5
    SHAIt = LongToStr(h1) & LongToStr(h2) & LongToStr(h3) & LongToStr(h4) & LongToStr(h5)
End Function

Private Sub StrToByteArray(ByVal sStr As String, ByRef ary() As Byte)
    ReDim ary(Len(sStr) - 1) As Byte
    CopyMemory ary(0), ByVal sStr, Len(sStr)
End Sub

Public Function LongToStr(ByVal lVal As Long) As String
    Dim s As String
    s = hex$(lVal)
    
    If Len(s) < 8 Then s = String$(8 - Len(s), "0") & s
    
    LongToStr = Chr$(Val("&H0" & Mid$(s, 1, 2))) & _
                Chr$(Val("&H0" & Mid$(s, 3, 2))) & _
                Chr$(Val("&H0" & Mid$(s, 5, 2))) & _
                Chr$(Val("&H0" & Mid$(s, 7, 2)))
End Function

Public Sub DefaultSHA1(Message() As Byte, h1 As Long, h2 As Long, h3 As Long, h4 As Long, h5 As Long)
    Sha1 Message, &H5A827999, &H6ED9EBA1, &H8F1BBCDC, &HCA62C1D6, h1, h2, h3, h4, h5
End Sub

Public Sub Sha1(Message() As Byte, ByVal Key1 As Long, ByVal Key2 As Long, ByVal Key3 As Long, ByVal Key4 As Long, h1 As Long, h2 As Long, h3 As Long, h4 As Long, h5 As Long)
    Dim U As Long, P As Long
    Dim FB As FourBytes, OL As OneLong
    Dim I As Integer
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
        For I = 0 To 15
            FB.d = Message(P)
            FB.c = Message(P + 1)
            FB.b = Message(P + 2)
            FB.a = Message(P + 3)
            LSet OL = FB
            W(I) = OL.L
            P = P + 4
        Next I
        
        For I = 16 To 79
            W(I) = U32RotateLeft1(W(I - 3) Xor W(I - 8) Xor W(I - 14) Xor W(I - 16))
        Next I
        
        a = h1: b = h2: c = h3: d = h4: e = h5
        
        For I = 0 To 19
            T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), W(I)), Key1), ((b And c) Or ((Not b) And d)))
            e = d: d = c: c = U32RotateLeft30(b): b = a: a = T
        Next I
        
        For I = 20 To 39
            T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), W(I)), Key2), (b Xor c Xor d))
            e = d: d = c: c = U32RotateLeft30(b): b = a: a = T
        Next I
        
        For I = 40 To 59
            T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), W(I)), Key3), ((b And c) Or (b And d) Or (c And d)))
            e = d: d = c: c = U32RotateLeft30(b): b = a: a = T
        Next I
        
        For I = 60 To 79
            T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), W(I)), Key4), (b Xor c Xor d))
            e = d: d = c: c = U32RotateLeft30(b): b = a: a = T
        Next I
        
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
