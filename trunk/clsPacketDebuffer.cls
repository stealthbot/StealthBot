VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacketDebuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================
'PacketDebuffer Class
'By Bethra, aka. Sorc.Polgara =)
'=================================================
'Modified/Fixed March 4-5, 2005
'  by Andy T, aka Stealth
'  stealth@stealthbot.net
'
' Changes:
'   - Added bounds checking code
'   - Removed unnecessary CopyMemory calls
'   - Fixed existing CopyMemory calls
'   - Added Advance(), DebuffRaw() and HasBytes()
'       functions
'=================================================
Option Explicit

' Uncomment this line if you don't already have CopyMemory declared
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Any, ByVal numBytes As Long)

Public Debuffer As String      '// Debuffering string

'Sets the Debuffer string
' Should be handed a headerless packet to be parsed
Public Function DebuffPacket(PacketData As String)
   Debuffer = PacketData
End Function

'Resets/clears the Debuffer
Public Function Clear()
   Debuffer = vbNullString
End Function

'=======================================================
'Public Functions that debuffer a part from the Debuffer
'=======================================================

'Debuffers a DWORD from the Debuffer
Public Function DebuffDWORD() As Long
    If HasBytes(4) Then
        DebuffDWORD = GetDWORD
        RemoveDWORD
    End If
End Function

'Debuffers a WORD from the Debuffer
Public Function DebuffWORD() As Integer
    If HasBytes(2) Then
        DebuffWORD = GetWORD
        RemoveWORD
    End If
End Function

'Debuffers a BYTE from the Debuffer
Public Function DebuffBYTE() As Byte
    If HasBytes(1) Then
        DebuffBYTE = GetBYTE
        RemoveBYTE
    End If
End Function

'Debuffers a FILETIME from the Debuffer
' Edit by Andy: Returns the 8 bytes of the FILETIME struct
' The end-location will have to typecast this via CopyMemory into a FILETIME
' VB didn't like using a user-defined type as a return type or parameter type
'   in a public class object
Public Function DebuffFILETIME() As String
    If HasBytes(8) Then
        DebuffFILETIME = GetFILETIME
        RemoveFILETIME
    End If
End Function

'Debuffers a null-terminating string from the Debuffer
Public Function DebuffNTString() As String
    If HasBytes(1) Then
        DebuffNTString = GetNTString
        RemoveNTString
    Else
        DebuffNTString = ""
    End If
End Function

'Debuffers x bytes -- for those times when you want it straight-up
'(added by Andy)
Public Function DebuffRaw(ByVal nBytes As Long) As String
    If HasBytes(nBytes) Then
        DebuffRaw = Mid$(Debuffer, 1, nBytes)
        Call Advance(nBytes)
    End If
End Function

'=====================================================
'Public Functions that remove a part from the Debuffer
'=====================================================

'Removes a BYTE from the Debuffer
Public Sub RemoveBYTE()
    Debuffer = Mid$(Debuffer, 2)
End Sub

'Removes a WORD from the Debuffer
Public Sub RemoveWORD()
    Debuffer = Mid$(Debuffer, 3)
End Sub

'Removes a DWORD from the Debuffer
Public Sub RemoveDWORD()
    Debuffer = Mid$(Debuffer, 5)
End Sub

'Removes a FILETIME structure from the Debuffer
Public Sub RemoveFILETIME()
    Debuffer = Mid$(Debuffer, 9)
End Sub

'Removes a null-terminating string from the Debuffer
Public Sub RemoveNTString()
    Dim Pos As Integer
    Pos = InStr(1, Debuffer, Chr(0), vbBinaryCompare)
   
    If Len(Debuffer) > Pos Then
        Debuffer = Mid$(Debuffer, Pos + 1)
    Else
        Call Clear
    End If
End Sub

'Removes nBytes bytes from the buffer
' For those times when you just don't care what's there..
'(added by Andy)
Public Sub Advance(ByVal nBytes As Long)
    If Len(Debuffer) > nBytes Then
        Debuffer = Mid$(Debuffer, nBytes + 1)
    Else
        Call Clear
    End If
End Sub


'=======================================================
'Functions that get parts from the front of the Debuffer
'=======================================================

'Gets a BYTE from the Debuffer
Function GetBYTE() As Byte
    Dim PBYTE As Byte
    PBYTE = Asc(Mid$(Debuffer, 1, 1))
    GetBYTE = PBYTE
End Function

'Gets a WORD from the Debuffer
Function GetWORD() As Integer
    Dim WORD As Integer
    Dim sTemp As String * 2
    
    sTemp = Mid$(Debuffer, 1, 2)
    CopyMemory WORD, ByVal sTemp, 2
    
    GetWORD = WORD
End Function

'Gets a DWORD from the Debuffer
Function GetDWORD() As Long
    Dim DWORD As Long
    Dim sTemp As String * 4
    
    sTemp = Mid$(Debuffer, 1, 4)
    CopyMemory DWORD, ByVal sTemp, 4
    
    GetDWORD = DWORD
End Function

'Gets a FILETIME from the Debuffer
Function GetFILETIME() As String
    GetFILETIME = Mid$(Debuffer, 1, 8)
End Function

'Gets a null-terminating string from the Debuffer
Function GetNTString() As String
    Dim NTString As String
    Dim Pos As Integer
    
    Pos = InStr(1, Debuffer, Chr(0), vbBinaryCompare)
    NTString = Mid$(Debuffer, 1, Pos - 1)
    
    GetNTString = NTString
End Function


'Returns TRUE if the debuffer has >= X bytes in it, else FALSE
'(added by Andy)
Function HasBytes(ByVal X As Integer) As Boolean
    HasBytes = (Len(Debuffer) >= X)
End Function

'                 (Sorc.Polgara's)
'=====================Credits============================
'DarkMinion for using his PacketBuffer class as a guide,
'this is the first class I have ever made =)
'-----------------
'Bot Developement Forum members who helped me understand
'some stuff that I need to know inorder to make this =)
'=======================================================
