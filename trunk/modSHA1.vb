Option Strict Off
Option Explicit On
Module modSHA1
	
	'This code is really messy. It was taken from http://vb.wikia.com/wiki/SHA-1.bas by Andy (RealityRipple), and _
	'then he added a few functions to it. After cleaning up the other warden code I got over it and I'm just _
	'going to leave this the way it is. I fixed up the tabbing a bit. - FrOzeN
	
	Private Structure FourBytes
		Dim a As Byte
		Dim b As Byte
		Dim c As Byte
		Dim d As Byte
	End Structure
	
	Private Structure OneLong
		Dim L As Integer
	End Structure
	
	'I added this function as a quick solution and better named method to call. The code it uses is still pretty bad. - FrOzeN
	'Public Sub Warden_SHA1(Destination() As Byte, ByRef Source() As Byte)
	Public Function Warden_SHA1(ByVal data As String) As String
		Dim arrData() As Byte
		Dim strHash As String
		Dim arrHash() As Byte
		Dim arrRet(20) As Byte
		
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
		arrData = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(data, vbFromUnicode))
		
		strHash = SHA1b(data)
		'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
		arrHash = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(strHash, vbFromUnicode))
		
		Call CopyMemory(arrRet(0), arrHash(0), 20)
		'UPGRADE_ISSUE: Constant vbUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		Warden_SHA1 = StrConv(System.Text.UnicodeEncoding.Unicode.GetString(arrRet), vbUnicode)
	End Function
	
	Private Function SHA1b(ByVal sIn As String) As String
		Dim bIn() As Byte
		StrToByteArray(sIn, bIn)
		SHA1b = SHAIt(bIn)
	End Function
	
	Private Function SHAIt(ByRef Message() As Byte) As String
		Dim h1 As Integer
		Dim h2 As Integer
		Dim h3 As Integer
		Dim h4 As Integer
		Dim h5 As Integer
		
		DefaultSHA1(Message, h1, h2, h3, h4, h5)
		SHAIt = LongToStr(h1) & LongToStr(h2) & LongToStr(h3) & LongToStr(h4) & LongToStr(h5)
	End Function
	
	Private Sub StrToByteArray(ByVal sStr As String, ByRef ary() As Byte)
		ReDim ary(Len(sStr) - 1)
		CopyMemory(ary(0), sStr, Len(sStr))
	End Sub
	
	Public Function LongToStr(ByVal lVal As Integer) As String
		Dim s As String
		s = Hex(lVal)
		
		If Len(s) < 8 Then s = New String("0", 8 - Len(s)) & s
		
		LongToStr = Chr(Val("&H0" & Mid(s, 1, 2))) & Chr(Val("&H0" & Mid(s, 3, 2))) & Chr(Val("&H0" & Mid(s, 5, 2))) & Chr(Val("&H0" & Mid(s, 7, 2)))
	End Function
	
	Public Sub DefaultSHA1(ByRef Message() As Byte, ByRef h1 As Integer, ByRef h2 As Integer, ByRef h3 As Integer, ByRef h4 As Integer, ByRef h5 As Integer)
		Sha1(Message, &H5A827999, &H6ED9EBA1, &H8F1BBCDC, &HCA62C1D6, h1, h2, h3, h4, h5)
	End Sub
	
	Public Sub Sha1(ByRef Message() As Byte, ByVal Key1 As Integer, ByVal Key2 As Integer, ByVal Key3 As Integer, ByVal Key4 As Integer, ByRef h1 As Integer, ByRef h2 As Integer, ByRef h3 As Integer, ByRef h4 As Integer, ByRef h5 As Integer)
		Dim U, P As Integer
		Dim FB As FourBytes
		Dim OL As OneLong
		Dim I As Short
		Dim W(80) As Integer
		Dim d, b, a, c, e As Integer
		Dim T As Integer
		
		h1 = &H67452301 : h2 = &HEFCDAB89 : h3 = &H98BADCFE : h4 = &H10325476 : h5 = &HC3D2E1F0
		
		U = UBound(Message) + 1 : OL.L = U32ShiftLeft3(U) : a = U \ &H20000000
		'UPGRADE_ISSUE: LSet cannot assign one type to another. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="899FA812-8F71-4014-BAEE-E5AF348BA5AA"'
		FB = LSet(OL) 'U32ShiftRight29(U)
		
		ReDim Preserve Message(CShort(U + 8 And -64) + 63)
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
				'UPGRADE_ISSUE: LSet cannot assign one type to another. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="899FA812-8F71-4014-BAEE-E5AF348BA5AA"'
				OL = LSet(FB)
				W(I) = OL.L
				P = P + 4
			Next I
			
			For I = 16 To 79
				W(I) = U32RotateLeft1(W(I - 3) Xor W(I - 8) Xor W(I - 14) Xor W(I - 16))
			Next I
			
			a = h1 : b = h2 : c = h3 : d = h4 : e = h5
			
			For I = 0 To 19
				T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), W(I)), Key1), (b And c) Or ((Not b) And d))
				e = d : d = c : c = U32RotateLeft30(b) : b = a : a = T
			Next I
			
			For I = 20 To 39
				T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), W(I)), Key2), b Xor c Xor d)
				e = d : d = c : c = U32RotateLeft30(b) : b = a : a = T
			Next I
			
			For I = 40 To 59
				T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), W(I)), Key3), (b And c) Or (b And d) Or (c And d))
				e = d : d = c : c = U32RotateLeft30(b) : b = a : a = T
			Next I
			
			For I = 60 To 79
				T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), e), W(I)), Key4), b Xor c Xor d)
				e = d : d = c : c = U32RotateLeft30(b) : b = a : a = T
			Next I
			
			h1 = U32Add(h1, a) : h2 = U32Add(h2, b) : h3 = U32Add(h3, c) : h4 = U32Add(h4, d) : h5 = U32Add(h5, e)
		End While
	End Sub
	
	Private Function U32Add(ByVal a As Integer, ByVal b As Integer) As Integer
		If (a Xor b) < 0 Then
			U32Add = a + b
		Else
			U32Add = CShort(a Xor &H80000000) + b Xor &H80000000
		End If
	End Function
	
	Private Function U32ShiftLeft3(ByVal a As Integer) As Integer
		U32ShiftLeft3 = CShort(a And &HFFFFFFF) * 8
		If a And &H10000000 Then U32ShiftLeft3 = U32ShiftLeft3 Or &H80000000
	End Function
	
	Private Function U32ShiftRight29(ByVal a As Integer) As Integer
		U32ShiftRight29 = (a And &HE0000000) \ &H20000000 And 7
	End Function
	
	Private Function U32RotateLeft1(ByVal a As Integer) As Integer
		U32RotateLeft1 = CShort(a And &H3FFFFFFF) * 2
		If a And &H40000000 Then U32RotateLeft1 = U32RotateLeft1 Or &H80000000
		If a And &H80000000 Then U32RotateLeft1 = U32RotateLeft1 Or 1
	End Function
	
	Private Function U32RotateLeft5(ByVal a As Integer) As Integer
		U32RotateLeft5 = CShort(a And &H3FFFFFF) * 32 Or (a And &HF8000000) \ &H8000000 And 31
		If a And &H4000000 Then U32RotateLeft5 = U32RotateLeft5 Or &H80000000
	End Function
	
	Private Function U32RotateLeft30(ByVal a As Integer) As Integer
		U32RotateLeft30 = CShort(a And 1) * &H40000000 Or (a And &HFFFC) \ 4 And &H3FFFFFFF
		If a And 2 Then U32RotateLeft30 = U32RotateLeft30 Or &H80000000
	End Function
End Module