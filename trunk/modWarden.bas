Attribute VB_Name = "modWarden"
Option Explicit

'/// Module: modWarden
'/// Requires: modSHA1,
'///           CopyMemory API

'/// This module was written by FrOzeN. Based off code from Andy (RealityRipple), and information from iago's documents.
'/// Topic information was taken from: http://forum.valhallalegends.com/index.php?topic=17356.0
'/// iago's documentation found at: http://www.skullsecurity.org/wiki/index.php/Starcraft_Warden

'/// <summary>
'/// Custom defined data type to easily manage the Random Data (unknown) warden uses
'/// </summary>
Private Type RandomData
    CurrentPosition     As Long
    RandomData(19)      As Byte
    RandomDataSrc1(19)  As Byte
    RandomDataSrc2(19)  As Byte
End Type

'/// <summary>
'/// Different warden address values used for determining which checksum to use.
'/// </summary>
Private Const WARDEN_H1 As Long = &H497FB0
Private Const WARDEN_H2 As Long = &H49C33D
Private Const WARDEN_H3 As Long = &H4A2FF7

'/// <summary>
'/// Variables for storing the warden checksum results.
'/// </summary>
Private Warden_Memory_1 As String
Private Warden_Memory_2 As String
Private Warden_Memory_3 As String

'/// <summary>An instance of the data type RandomData</summary>
Private rd As RandomData

'/// <summary>
'/// Two byte arrays that store the keys used to encrypt and decrypt incoming and outgoing warden packets
'/// </summary>
Private KeyIn(&H101)    As Byte
Private KeyOut(&H101)   As Byte

'/// <summary>
'/// Initializes warden in 3 main steps:
'///     1. Assigns the warden checksum result to variables
'///     2. Uses the first DWORD of KeyHash to seed the random data which is used to encrypt warden packets
'///     3. Generates the warden keys
'/// </summary>
'/// <param name="KeyHashStart">The first DWORD of the Cd-Key hash.</param>
Public Sub InitializeWarden(ByRef KeyHashStart As String)
    Call Initialize_Warden_Memory
    Call Random_Data_Initialize(KeyHashStart)
    Call Generate_Warden_Keys
End Sub

'/// <summary>
'/// Initializes the checksum result values to variables
'/// </summary>
Private Sub Initialize_Warden_Memory()
    Warden_Memory_1 = Chr$(&H84) & Chr$(&H5E) & Chr$(&HC) & Chr$(&H74) & Chr$(&H5) & Chr$(&HE8) & Chr$(&HF6) & Chr$(&H54) & Chr$(&HF9) & Chr$(&HFF) & Chr$(&H8B) & Chr$(&H76) & Chr$(&H4) & Chr$(&H85)
    Warden_Memory_2 = Chr$(&H83) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H8B) & Chr$(&H55) & Chr$(&H8)
    Warden_Memory_3 = Chr$(&HA3) & Chr$(&H68) & Chr$(&HCC) & Chr$(&H59) & Chr$(&H0) & Chr$(&HE8) & Chr$(&HDF) & Chr$(&H23)
End Sub

'/// <summary>
'/// Encrypts a warden packet using warden's random data and the key provided.
'/// </summary>
'/// <param name="PacketIn">Packet to be encrypted.</param>
'/// <param name="Key">Warden key to be used for the encryption.</param>
'/// <param name="PacketOut">The packet returned as encrypted.</param>
Public Sub RunCrypt(ByRef PacketIn() As Byte, ByRef Key() As Byte, ByRef PacketOut() As Byte)
    Dim i As Integer
    Dim z As Long, y As Long
    Dim byteSwap As Byte
    
    ReDim PacketOut(UBound(PacketIn))
    
    CopyMemory PacketOut(0), PacketIn(0), UBound(PacketIn) + 1
    
    y = Key(&H100)
    z = Key(&H101)
    
    For i = 0 To UBound(PacketIn)
        y = (y + 1) And &HFF
        z = (z + Key(y)) And &HFF
        
        'Swap Key(y) with Key(z)
        byteSwap = Key(y)
        Key(y) = Key(z)
        Key(z) = byteSwap
        
        PacketOut(i) = PacketOut(i) Xor Key((CInt(Key(y)) + CInt(Key(z))) And &HFF)
    Next i
    
    Key(&H100) = y
    Key(&H101) = z
End Sub

'/// <summary>
'/// Generates a warden key based on the source given.
'/// </summary>
'/// <param name="Key">The key to be returned.</param>
'/// <param name="Source">The source to be used to generate the key.</param>
Private Sub Generate_Key(ByRef Key() As Byte, ByRef Source() As Byte)
    Dim Value As Long, Position As Long
    Dim i As Integer, j As Integer
    Dim SwapByte As Byte
    
    'Populate keys
    For i = 0 To &HFF
        Key(i) = i
    Next i
    
    For i = 0 To &HFF
        Value = Value + Key(i) + Source(i Mod 16)
        
        'Swap Key(i) with Key(Value And &HFF)
        SwapByte = Key(i)
        Key(i) = Key(Value And &HFF)
        Key(Value And &HFF) = SwapByte
    Next i
End Sub

'/// <summary>
'/// Generates the warden keys using warden's random data.
'/// </summary
Private Sub Generate_Warden_Keys()
    Dim TempData() As Byte
    
    'Generate KeyOut
    Call Random_Data_GetBytes(TempData, &HF)
    Generate_Key KeyOut, TempData
    
    'Generate KeyIn
    Call Random_Data_GetBytes(TempData, &HF)
    Generate_Key KeyIn, TempData
End Sub

'/// <summary>
'/// Returns the warden checksum result based on the 3 addresses provided.
'/// </summary>
'/// <param name="Addresses">The warden addresses to determine the resulting checksum.</param>
Private Function GetWardenChecksum(ByRef Addresses() As Long) As Long
    If CompareAddresses(Addresses, WARDEN_H1, WARDEN_H2, WARDEN_H3) Then
        GetWardenChecksum = &H193E73E8
    ElseIf CompareAddresses(Addresses, WARDEN_H1, WARDEN_H3, WARDEN_H2) Then
        GetWardenChecksum = &H2183172A
    ElseIf CompareAddresses(Addresses, WARDEN_H2, WARDEN_H1, WARDEN_H3) Then
        GetWardenChecksum = &HD6557DEF
    ElseIf CompareAddresses(Addresses, WARDEN_H2, WARDEN_H3, WARDEN_H1) Then
        GetWardenChecksum = &HCA841860
    ElseIf CompareAddresses(Addresses, WARDEN_H3, WARDEN_H2, WARDEN_H1) Then
        GetWardenChecksum = &H9F2AD2C3
    ElseIf CompareAddresses(Addresses, WARDEN_H3, WARDEN_H1, WARDEN_H2) Then
        GetWardenChecksum = &HC04CF757
    Else
        GetWardenChecksum = 0&
    End If
End Function

'/// <summary>
'/// This is helper function for GetWardenChecksum to compare 3 addresses at once.
'/// </summary>
'/// <param name="Addresses">The addresses to be compared against.</param>
'/// <param name="Addr1">The first address.</param>
'/// <param name="Addr2">The second address.</param>
'/// <param name="Addr3">The third address.</param>
Private Function CompareAddresses(ByRef Addresses() As Long, ByVal Addr1 As Long, ByVal Addr2 As Long, ByVal Addr3 As Long) As Boolean
    CompareAddresses = (Addresses(0) = Addr1 And _
                        Addresses(1) = Addr2 And _
                        Addresses(2) = Addr3)
End Function

'/// <summary>
'/// Returns the warden checksum based on the address.
'/// </summary>
'/// <param name="Address">Address to determine which warden memory to use.</param>
Private Function GetWardenMemory(ByVal Address As Long) As String
    Select Case Address
        Case WARDEN_H1
            GetWardenMemory = Warden_Memory_1
        Case WARDEN_H2
            GetWardenMemory = Warden_Memory_2
        Case WARDEN_H3
            GetWardenMemory = Warden_Memory_3
    End Select
End Function

'/// <summary>
'/// This function parses the warden packet and returned the resulting packet which is ready to be sent straight back to Battle.net.
'/// </summary>
'/// <param name="Packet">Warden packet to be processed.</param>
Public Function HandleWarden(ByRef Packet As String) As String
    Dim PacketData() As Byte
    Dim DecryptedPacket() As Byte
    Dim Buffer As String

    'Convert the Packet into a byte array and assign it to PacketData.
    PacketData = StrConv(Packet, vbFromUnicode)
    
    'Decrypt the packet.
    RunCrypt PacketData, KeyIn, DecryptedPacket
    
    Select Case DecryptedPacket(0)
        Case &H0
            'Resize the packet to 1 byte and set it to 0x01
            ReDim PacketData(0)
            PacketData(0) = &H1
        
        Case &H2
            Dim LoopAmount As Integer, Position As Integer, i As Integer
            Dim Values() As String
            Dim Addresses() As Long, Checksum As Long
            
            LoopAmount = (Len(Packet) - 3) / 7

            ReDim Values(LoopAmount - 1) As String
            ReDim Addresses(LoopAmount - 1) As Long
            
            Position = 2
            
            For i = 0 To LoopAmount - 1
                'WORD   (Unknown) - Don't know what this is
                'DWORD  (Addresses)
                'Byte   (ReadLength) - Don't know what this is, called it ReadLength because Andy did
                
                Position = Position + 2     'Skip (Unknown) WORD
                
                'Addresses(i) = GetDWORD(DecryptedPacket)
                CopyMemory Addresses(i), DecryptedPacket(Position), 4
                
                Position = Position + 5     'Move 4 bytes for DWORD, and then 1 extra byte to skip the (ReadLength) byte
                
                'Get the warden memory based on the addresses.
                Values(i) = GetWardenMemory(Addresses(i))
            Next i

            'Get warden's checksum based on the Addresses.
            Checksum = GetWardenChecksum(Addresses)
                        
            If Checksum = 0& Then
                HandleWarden = vbNullString
                Exit Function
            End If
            
            For i = 0 To LoopAmount - 1
                'Packet.InsertByte &H0
                'Packet.InsertString Values(i)
                Buffer = Buffer & Chr$(0) & Values(i)
            Next i
            
            Dim tmpLength As String * 2     'WORD
            Dim tmpChecksum As String * 4   'DWORD
            
            CopyMemory tmpLength, Len(Buffer), 2
            CopyMemory tmpChecksum, Checksum, 4
            
            'Build the final parts of the packet.
            Buffer = Chr$(&H2) & tmpLength & tmpChecksum & Buffer
            
            'Convert the packet back into a string from a byte array.
            PacketData = StrConv(Buffer, vbFromUnicode)
            
        Case Else
            HandleWarden = vbNullString
            Exit Function
    
    End Select
    
    'Encrypt the warden packet with the outgoing key.
    RunCrypt PacketData, KeyOut, DecryptedPacket
    
    'Return the encrypted packet.
    HandleWarden = StrConv(DecryptedPacket, vbUnicode)
End Function

'/// <summary>
'/// Initializes warden's random data.
'/// </summary>
'/// <param name="Seed">The seed to hash warden's random data.</param>
Private Sub Random_Data_Initialize(ByRef Seed As String)
    Dim Seed1() As Byte
    Dim Seed2() As Byte
    
    'Split's the Seed into two WORD's.
    'Each WORD is used as part of the Warden SHA1 seed for hashing random data source 1 and 2 respectively.
    Seed1 = StrConv(Left$(Seed, 2), vbFromUnicode)
    Seed2 = StrConv(Right$(Seed, 2), vbFromUnicode)
    
    'Clear rd.RandomData
    Dim BlankBytes(19) As Byte
    CopyMemory rd.RandomData(0), BlankBytes(0), 20

    'Hash the random data.
    Call Warden_SHA1(rd.RandomDataSrc1, Seed1)
    Call Warden_SHA1(rd.RandomDataSrc2, Seed2)
    
    'Updates warden's random data.
    Call Random_Data_Update
End Sub

'/// <summary>
'/// Pull random data out of the random data variables and update it accordingly.
'/// </summary>
'/// <param name="Destination">Byte array that the data will be assigned to.</param>
'/// <param name="Length">Amount of data (in bytes) that needs to be retrieved.</param>
Private Sub Random_Data_GetBytes(ByRef Destination() As Byte, ByVal Length As Integer)
    Dim i As Integer
    
    'Clears the Destination byte array and resizes it.
    ReDim Destination(Length)
    
    'Loops through the random data and places it into the Destination byte array byte by byte.
    For i = 0 To Length
        Destination(i) = rd.RandomData(rd.CurrentPosition)
        rd.CurrentPosition = rd.CurrentPosition + 1
        
        'Every 20 bytes it updates the random data.
        If rd.CurrentPosition >= &H14 Then
            Call Random_Data_Update
        End If
    Next i
End Sub

'/// <summary>
'/// Updates warden's random data.
'/// </summary>
Private Sub Random_Data_Update()
    Dim TempData(59) As Byte
    
    'Copies warden's random data in a temporary byte array.
    CopyMemory TempData(0), rd.RandomDataSrc1(0), 20
    CopyMemory TempData(20), rd.RandomData(0), 20
    CopyMemory TempData(40), rd.RandomDataSrc2(0), 20

    'Hashes warden's random data using the previous random data as the source.
    Call Warden_SHA1(rd.RandomData, TempData)
    
    rd.CurrentPosition = 0
End Sub
