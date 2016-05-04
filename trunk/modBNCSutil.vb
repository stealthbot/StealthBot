Option Strict Off
Option Explicit On
Module modBNCSutil
	
	'------------------------------------------------------------------------------
	'  BNCSutil
	'  Battle.Net Utility Library
	'
	'  Copyright © 2004-2005 Eric Naeseth
	'------------------------------------------------------------------------------
	'  Visual Basic Declarations
	'  November 20, 2004
	'------------------------------------------------------------------------------
	'  This library is free software; you can redistribute it and/or
	'  modify it under the terms of the GNU Lesser General Public
	'  License as published by the Free Software Foundation; either
	'  version 2.1 of the License, or (at your option) any later version.
	'
	'  This library is distributed in the hope that it will be useful,
	'  but WITHOUT ANY WARRANTY; without even the implied warranty of
	'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
	'  Lesser General Public License for more details.
	'
	'  A copy of the GNU Lesser General Public License is included in the BNCSutil
	'  distribution in the file COPYING.  If you did not receive this copy,
	'  write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330,
	'  Boston, MA  02111-1307  USA
	'------------------------------------------------------------------------------
	
	'  DLL Imports
	'---------------------------
	
	' Library Information
	Private Declare Function BNCSutil_getVersion Lib "BNCSutil.dll" () As Integer
	Private Declare Function BNCSutil_getVersionString_Raw Lib "BNCSutil.dll"  Alias "BNCSutil_getVersionString"(ByVal outbuf As String) As Integer
	
	' CheckRevision
	Private Declare Function extractMPQNumber Lib "BNCSutil.dll" (ByVal MPQName As String) As Integer
	' [!] You should use checkRevision and getExeInfo (see below) instead of their
	'     _Raw counterparts.
	Private Declare Function checkRevision_Raw Lib "BNCSutil.dll"  Alias "checkRevisionFlat"(ByVal ValueString As String, ByVal File1 As String, ByVal File2 As String, ByVal File3 As String, ByVal mpqNumber As Integer, ByRef Checksum As Integer) As Integer
	Private Declare Function getExeInfo_Raw Lib "BNCSutil.dll"  Alias "getExeInfo"(ByVal FileName As String, ByVal exeInfoString As String, ByVal infoBufferSize As Integer, ByRef Version As Integer, ByVal Platform As Integer) As Integer
	
	' Old Logon System
	' [!] You should use doubleHashPassword and hashPassword instead of their
	'     _Raw counterparts.  (See below for those functions.)
	Private Declare Sub doubleHashPassword_Raw Lib "BNCSutil.dll"  Alias "doubleHashPassword"(ByVal Password As String, ByVal ClientToken As Integer, ByVal ServerToken As Integer, ByVal outBuffer As String)
	Private Declare Sub hashPassword_Raw Lib "BNCSutil.dll"  Alias "hashPassword"(ByVal Password As String, ByVal outBuffer As String)
	
	
	' CD-Key Decoding
	
	' Call kd_init() first to set up the decoding system, unless you are only using kd_quick().
	' Then call kd_create() to create a key decoder "handle" each time you want to
	' decode a CD-key.  It will return the handle or -1 on failure.  The handle
	' should then be passed as the "decoder" argument to all the other kd_ functions.
	' Call kd_free() on the handle when finished with the decoder to free the
	' memory it is using.
	
	' Key decoding declarations and logic moved to clsKeyDecoder. (2016-3-26, -Pyro)
	
	
	'New Logon System
	
	' Call nls_init() to get a "handle" to an NLS object (nls_init will return 0
	' if it encounters an error).  This "handle" should be passed as the "NLS"
	' argument to all the other nls_* functions.  You do not need to change the
	' username and password to upper-case as nls_init() will do this for you.
	' Call nls_free() on the handle to free the memory it's using.
	' nls_account_create() and nls_account_logon() generate the bodies of
	' SID_AUTH_ACCOUNTCREATE and SID_AUTH_ACCOUNTLOGIN packets, respectively.
	
	' Logon system declarations and logic moved to clsNLS.
	
	
	
	'  Constants
	'---------------------------
	Public Const BNCSutil_PLATFORM_X86 As Integer = &H1
	Public Const BNCSutil_PLATFORM_WINDOWS As Integer = &H1
	Public Const BNCSutil_PLATFORM_WIN As Integer = &H1
	
	Public Const BNCSutil_PLATFORM_PPC As Integer = &H2
	Public Const BNCSutil_PLATFORM_MAC As Integer = &H2
	
	Public Const BNCSutil_PLATFORM_OSX As Integer = &H3
	
	'BNCSUTIL NLS buffer size constants
	Public Const NLS_ACCOUNTCREATE_ As Integer = 65
	Public Const NLS_ACCOUNTLOGON_ As Integer = 33
	Public Const NLS_GET_S_ As Integer = 32
	Public Const NLS_GET_V_ As Integer = 32
	Public Const NLS_GET_A_ As Integer = 32
	Public Const NLS_GET_K_ As Integer = 40
	Public Const NLS_GET_M1_ As Integer = 20
	
	
	'  VB-Specifc Functions
	'---------------------------
	
	' RequiredVersion must be a version as a.b.c
	' Returns True if the current BNCSutil version is sufficent, False if not.
	' Function will now return the right value - l)ragon
	Public Function bncsutil_checkVersion(ByVal RequiredVersion As String) As Boolean
		Dim i, j As Integer
		Dim Frag() As String
		Dim Req, Check As Integer
		bncsutil_checkVersion = False
		Frag = Split(RequiredVersion, ".")
		j = 0
		For i = UBound(Frag) To 0 Step -1
			Check = Check + (CInt(Val(Frag(i))) * (100 ^ j))
			j = j + 1
		Next i
		'v Somone desided to use Check here instead of Req - l)ragon
		Req = BNCSutil_getVersion()
		If (Check >= Req) Then
			bncsutil_checkVersion = True
		End If
	End Function
	
	Public Function BNCSutil_getVersionString() As String
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		str_Renamed = New String(vbNullChar, 10)
		Call BNCSutil_getVersionString_Raw(str_Renamed)
		BNCSutil_getVersionString = str_Renamed
	End Function
	
	'CheckRevision
	Public Function checkRevision(ByRef ValueString As String, ByRef File1 As String, ByRef File2 As String, ByRef File3 As String, ByRef mpqNumber As Integer, ByRef Checksum As Integer) As Boolean
		checkRevision = (checkRevision_Raw(ValueString, File1, File2, File3, mpqNumber, Checksum) > 0)
	End Function
	
	Public Function checkRevisionA(ByRef ValueString As String, ByRef Files() As String, ByRef mpqNumber As Integer, ByRef Checksum As Integer) As Boolean
		checkRevisionA = (checkRevision_Raw(ValueString, Files(0), Files(1), Files(2), mpqNumber, Checksum) > 0)
	End Function
	
	'EXE Information
	'Information string (file name, date, time, and size) will be placed in InfoString.
	'InfoString does NOT need to be initialized (e.g. InfoString = String$(255, vbNullChar))
	'Returns the file version or 0 on failure.
	Public Function getExeInfo(ByRef EXEFile As String, ByRef InfoString As String, Optional ByVal Platform As Integer = BNCSutil_PLATFORM_WINDOWS) As Integer
		Dim InfoSize, Version, Result As Integer
		Dim i As Integer
		InfoSize = 256
		InfoString = New String(vbNullChar, 256)
		Result = getExeInfo_Raw(EXEFile, InfoString, InfoSize, Version, Platform)
		If Result = 0 Then
			getExeInfo = 0
			Exit Function
		End If
		While Result > InfoSize
			If InfoSize > 1024 Then
				getExeInfo = 0
				Exit Function
			End If
			InfoSize = InfoSize + 256
			InfoString = New String(vbNullChar, InfoSize)
			Result = getExeInfo_Raw(EXEFile, InfoString, InfoSize, Version, Platform)
		End While
		getExeInfo = Version
		i = InStr(InfoString, vbNullChar)
		If i = 0 Then Exit Function
		InfoString = Left(InfoString, i - 1)
	End Function
	
	'OLS Password Hashing
	Public Function doubleHashPassword(ByRef Password As String, ByVal ClientToken As Integer, ByVal ServerToken As Integer) As String
		Dim Hash As New VB6.FixedLengthString(20)
		doubleHashPassword_Raw(Password, ClientToken, ServerToken, Hash.Value)
		doubleHashPassword = Hash.Value
	End Function
	
	Public Function hashPassword(ByRef Password As String) As String
		Dim Hash As New VB6.FixedLengthString(20)
		hashPassword_Raw(Password, Hash.Value)
		hashPassword = Hash.Value
	End Function
End Module