Option Strict Off
Option Explicit On
Friend Class clsCRC32
	
	Private Const CRC32_POLYNOMIAL As Integer = &HEDB88320
	
	Private CRC32Table(255) As Integer
	
	'Public Functions
	Private Sub InitCRC32()
		Static CRC32Initialized As Boolean
		
		Dim i As Integer
		Dim j As Integer
		Dim K As Integer
		Dim XorVal As Integer
		
		If (CRC32Initialized) Then
			Exit Sub
		End If
		
		CRC32Initialized = True
		
		For i = 0 To 255
			K = i
			
			For j = 1 To 8
				If K And 1 Then XorVal = CRC32_POLYNOMIAL Else XorVal = 0
				If K < 0 Then K = ((K And &H7FFFFFFF) \ 2) Or &H40000000 Else K = K \ 2
				K = K Xor XorVal
			Next 
			
			CRC32Table(i) = K
		Next 
	End Sub
	
	Public Function CRC32(ByVal Data As String) As Integer
		Dim i As Integer
		Dim j As Integer
		
		Call InitCRC32()
		
		CRC32 = &HFFFFFFFF
		
		For i = 1 To Len(Data)
			j = CByte(Asc(Mid(Data, i, 1))) Xor (CRC32 And &HFF)
			
			If (CRC32 < 0) Then
				CRC32 = ((CRC32 And &H7FFFFFFF) \ &H100) Or &H800000
			Else
				CRC32 = CRC32 \ &H100
			End If
			
			CRC32 = (CRC32 Xor CRC32Table(j))
		Next 
		
		CRC32 = (Not (CRC32))
	End Function
	
	Public Function GetFileCRC32(ByVal filePath As String) As Integer
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		Dim tmp As String
		Dim f As Short
		
		f = FreeFile
		
		FileOpen(f, filePath, OpenMode.Input)
		Do While (EOF(f) = False)
			tmp = LineInput(f)
			
			str_Renamed = str_Renamed & tmp
		Loop 
		FileClose(f)
		
		GetFileCRC32 = CRC32(str_Renamed)
	End Function
	
	'Modified from code given to me by David Fritts (sneakcharm@yahoo.com)
	Public Function ValidateExecutable() As Boolean
		On Error GoTo ValidateExecutable_Error
		
		'Dim CRC32          As clsCRC32
		Dim strFilePath As String
		Dim intFreeFile As Short
		Dim strBuffer As String
		Dim strFileCRC As New VB6.FixedLengthString(8)
		Dim strComputedCRC As New VB6.FixedLengthString(8)
		Dim lngComputedCRC As Integer
		
		'Set CRC32 = New clsCRC32
		
		'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		strFilePath = My.Application.Info.DirectoryPath & "/" & My.Application.Info.AssemblyName & ".exe"
		
		'Generate a CRC for ourselves
		intFreeFile = FreeFile
		
		'read the sections you want to protect
		FileOpen(intFreeFile, strFilePath, OpenMode.Binary, OpenAccess.Read)
		strBuffer = New String(vbNullChar, LOF(intFreeFile) - 8)
		
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(intFreeFile, strBuffer, 1)
		FileClose(intFreeFile)
		
		'Compute the new CRC
		lngComputedCRC = CRC32(strBuffer)
		strComputedCRC.Value = ZeroOffset(lngComputedCRC, 8)
		
		'Read a CRC from ourselves
		intFreeFile = FreeFile
		
		FileOpen(intFreeFile, strFilePath, OpenMode.Binary)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(intFreeFile, strFileCRC.Value, FileLen(strFilePath) - 7)
		FileClose(intFreeFile)
		
		If (StrComp(strComputedCRC.Value, strFileCRC.Value, CompareMethod.Binary) = 0) Then
			ValidateExecutable = True
		Else
			ValidateExecutable = False
		End If
		
		'Set CRC32 = Nothing
		
ValidateExecutable_Exit: 
		Exit Function
		
ValidateExecutable_Error: 
		ValidateExecutable = True
		
		Debug.Print("Error " & Err.Number & " (" & Err.Description & ") in procedure " & "ValidateExecutable of Module modCRC32Checksum")
		
		Resume ValidateExecutable_Exit
	End Function
End Class