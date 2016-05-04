Option Strict Off
Option Explicit On
Module modRegistry
	' http://www.devx.com/vb2themax/Tip/19134
	
	Private Declare Function RegOpenKeyEx Lib "advapi32.dll"  Alias "RegOpenKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
	Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function RegQueryValueEx Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Any, ByRef lpcbData As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef dest As Any, ByRef source As Any, ByVal numBytes As Integer)
	
	Const KEY_READ As Integer = &H20019 ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
	' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
	' SYNCHRONIZE))
	
	Const REG_SZ As Short = 1
	Const REG_EXPAND_SZ As Short = 2
	Const REG_BINARY As Short = 3
	Const REG_DWORD As Short = 4
	Const REG_MULTI_SZ As Short = 7
	Const ERROR_MORE_DATA As Short = 234
	
	' Read a Registry value
	'
	' Use KeyName = "" for the default value
	' If the value isn't there, it returns the DefaultValue
	' argument, or Empty if the argument has been omitted
	'
	' Supports DWORD, REG_SZ, REG_EXPAND_SZ, REG_BINARY and REG_MULTI_SZ
	' REG_MULTI_SZ values are returned as a null-delimited stream of strings
	' (VB6 users can use SPlit to convert to an array of string)
	
	Function GetRegistryValue(ByVal hKey As Integer, ByVal KeyName As String, ByVal ValueName As String, Optional ByRef DefaultValue As Object = Nothing) As Object
		Dim handle As Integer
		Dim resLong As Integer
		Dim resString As String
		Dim resBinary() As Byte
		Dim length As Integer
		Dim retVal As Integer
		Dim valueType As Integer
		
		' Prepare the default result
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		'UPGRADE_WARNING: Couldn't resolve default property of object GetRegistryValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetRegistryValue = IIf(IsNothing(DefaultValue), Nothing, DefaultValue)
		
		' Open the key, exit if not found.
		If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then
			Exit Function
		End If
		
		' prepare a 1K receiving resBinary
		length = 1024
		ReDim resBinary(length - 1)
		
		' read the registry key
		retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), length)
		' if resBinary was too small, try again
		If retVal = ERROR_MORE_DATA Then
			' enlarge the resBinary, and read the value again
			ReDim resBinary(length - 1)
			retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), length)
		End If
		
		' return a value corresponding to the value type
		Select Case valueType
			Case REG_DWORD
				CopyMemory(resLong, resBinary(0), 4)
				'UPGRADE_WARNING: Couldn't resolve default property of object GetRegistryValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetRegistryValue = resLong
			Case REG_SZ, REG_EXPAND_SZ
				' copy everything but the trailing null char
				resString = Space(length - 1)
				CopyMemory(resString, resBinary(0), length - 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object GetRegistryValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetRegistryValue = resString
			Case REG_BINARY
				' resize the result resBinary
				If length <> UBound(resBinary) + 1 Then
					ReDim Preserve resBinary(length - 1)
				End If
				GetRegistryValue = VB6.CopyArray(resBinary)
			Case REG_MULTI_SZ
				' copy everything but the 2 trailing null chars
				resString = Space(length - 2)
				CopyMemory(resString, resBinary(0), length - 2)
				'UPGRADE_WARNING: Couldn't resolve default property of object GetRegistryValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetRegistryValue = resString
			Case Else
				RegCloseKey(handle)
				Err.Raise(1001,  , "Unsupported value type")
		End Select
		
		' close the registry key
		RegCloseKey(handle)
	End Function
End Module