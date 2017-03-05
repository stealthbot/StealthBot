Attribute VB_Name = "modBNCSutil"
Option Explicit

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
Private Declare Function BNCSutil_getVersion Lib "BNCSutil.dll" () As Long
Private Declare Function BNCSutil_getVersionString_Raw Lib "BNCSutil.dll" _
    Alias "BNCSutil_getVersionString" (ByVal outbuf As String) As Long
 
' CheckRevision
Private Declare Function extractMPQNumber Lib "BNCSutil.dll" _
    (ByVal MPQName As String) As Long
' [!] You should use checkRevision and getExeInfo (see below) instead of their
'     _Raw counterparts.
Private Declare Function checkRevision_Raw Lib "BNCSutil.dll" Alias "checkRevisionFlat" _
    (ByVal ValueString As Long, ByVal File1 As String, ByVal File2 As String, _
     ByVal File3 As String, ByVal mpqNumber As Long, ByRef Checksum As Long) As Long
Private Declare Function getExeInfo_Raw Lib "BNCSutil.dll" Alias "getExeInfo" _
    (ByVal FileName As String, ByVal exeInfoString As String, _
    ByVal infoBufferSize As Long, Version As Long, ByVal Platform As Long) As Long

' Old Logon System
' [!] You should use doubleHashPassword and hashPassword instead of their
'     _Raw counterparts.  (See below for those functions.)
Private Declare Sub doubleHashPassword_Raw Lib "BNCSutil.dll" Alias "doubleHashPassword" _
    (ByVal Password As String, ByVal ClientToken As Long, ByVal ServerToken As Long, _
    ByVal outBuffer As Long)
Private Declare Sub hashPassword_Raw Lib "BNCSutil.dll" Alias "hashPassword" _
    (ByVal Password As String, ByVal outBuffer As Long)


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
Public Const BNCSutil_PLATFORM_X86& = &H1
Public Const BNCSutil_PLATFORM_WINDOWS& = &H1
Public Const BNCSutil_PLATFORM_WIN& = &H1

Public Const BNCSutil_PLATFORM_PPC& = &H2
Public Const BNCSutil_PLATFORM_MAC& = &H2

Public Const BNCSutil_PLATFORM_OSX& = &H3

'BNCSUTIL NLS buffer size constants
Public Const NLS_ACCOUNTCREATE_     As Long = 65
Public Const NLS_ACCOUNTLOGON_      As Long = 33
Public Const NLS_GET_S_             As Long = 32
Public Const NLS_GET_V_             As Long = 32
Public Const NLS_GET_A_             As Long = 32
Public Const NLS_GET_K_             As Long = 40
Public Const NLS_GET_M1_            As Long = 20


'  VB-Specifc Functions
'---------------------------

' RequiredVersion must be a version as a.b.c
' Returns True if the current BNCSutil version is sufficent, False if not.
' Function will now return the right value - l)ragon
Public Function bncsutil_checkVersion(ByVal RequiredVersion As String) As Boolean
    Dim i&, j&
    Dim Frag() As String
    Dim Req As Long, Check As Long
    bncsutil_checkVersion = False
    Frag = Split(RequiredVersion, ".")
    j = 0
    For i = UBound(Frag) To 0 Step -1
        Check = Check + (CLng(Val(Frag(i))) * (100 ^ j))
        j = j + 1
    Next i
    'v Somone desided to use Check here instead of Req - l)ragon
    Req = BNCSutil_getVersion()
    If (Check >= Req) Then
        bncsutil_checkVersion = True
    End If
End Function

Public Function BNCSutil_getVersionString() As String
    Dim str As String
    str = String$(10, vbNullChar)
    Call BNCSutil_getVersionString_Raw(str)
    BNCSutil_getVersionString = str
End Function

'CheckRevision
Public Function checkRevision(ValueString As String, File1$, File2$, File3$, mpqNumber As Long, Checksum As Long) As Boolean
    checkRevision = (checkRevision_Raw(ValueString, File1, File2, File3, mpqNumber, Checksum) > 0)
End Function

Public Function checkRevisionA(ValueString As String, Files() As String, mpqNumber As Long, Checksum As Long) As Boolean
    checkRevisionA = (checkRevision_Raw(ValueString, Files(0), Files(1), Files(2), mpqNumber, Checksum) > 0)
End Function

'EXE Information
'Information string (file name, date, time, and size) will be placed in InfoString.
'InfoString does NOT need to be initialized (e.g. InfoString = String$(255, vbNullChar))
'Returns the file version or 0 on failure.
Public Function getExeInfo(EXEFile As String, InfoString As String, Optional ByVal Platform As Long = BNCSutil_PLATFORM_WINDOWS) As Long
    Dim Version As Long, InfoSize As Long, Result As Long
    Dim i&
    InfoSize = 256
    InfoString = String$(256, vbNullChar)
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
        InfoString = String$(InfoSize, vbNullChar)
        Result = getExeInfo_Raw(EXEFile, InfoString, InfoSize, Version, Platform)
    Wend
    getExeInfo = Version
    i = InStr(InfoString, vbNullChar)
    If i = 0 Then Exit Function
    InfoString = Left$(InfoString, i - 1)
End Function

'OLS Password Hashing
Public Function doubleHashPassword(Password As String, ByVal ClientToken As Long, ByVal ServerToken As Long) As String
    Dim Hash(19) As Byte
    Call doubleHashPassword_Raw(Password, ClientToken, ServerToken, VarPtr(Hash(0)))
    doubleHashPassword = ByteArrToString(Hash())
End Function

Public Function hashPassword(Password As String) As String
    Dim Hash(19) As Byte
    Call hashPassword_Raw(Password, VarPtr(Hash(0)))
    hashPassword = ByteArrToString(Hash())
End Function

