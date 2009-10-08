Attribute VB_Name = "modOutboundPackets"
'StealthBot project - modOutboundPackets
' Packet creation, checkrevision and NLS code
' March 2005
' Andy T andy@stealthbot.net
Option Explicit

Public g_username As String

' Uses BNCSUtil to decode and hash your cdkey
' Returns the ClientToken that should be used to connect
' 2009-05-17 - Removed Last Argument MPQRevision, Why the hell was it there?
Public Sub DecodeCDKey(ByVal sCDKey As String, ByVal ServerToken As Long, ByVal ClientToken As Long, ByRef KeyHash As String, ByRef Value1 As Long, ByRef ProductID As Long)
    Dim KDh As Long                     ' Key Decoder handler
    Dim HashSize As Long                ' CDKey hash size in bytes
    Dim Result As Long                  ' kd_init() result
    
    sCDKey = Replace(sCDKey, "-", vbNullString)
    sCDKey = Replace(sCDKey, " ", vbNullString)
    sCDKey = KillNull(sCDKey)
    
    Result = kd_init()
    
    KeyHash = vbNullString
    
    If Result = 0 Then
        frmChat.AddChat RTBColors.ErrorMessageText, "BNCSUtil: kd_init() failed! Please use BNLS to connect."
        frmChat.DoDisconnect
        
    Else
        KDh = kd_create(sCDKey, Len(sCDKey))
            
        If (kd_isValid(KDh) = 0) Then
            frmChat.AddChat RTBColors.ErrorMessageText, "Your CD-Key is invalid."
            frmChat.DoDisconnect
            
        Else
            HashSize = kd_calculateHash(KDh, ClientToken, ServerToken)
        
            If HashSize <= 0 Then
                frmChat.AddChat RTBColors.ErrorMessageText, "Your CD-Key is invalid. [kd_calculateHash() <= 0]"
                frmChat.AddChat RTBColors.ErrorMessageText, "Please make sure you typed your CD-Key correctly. This error is often generated when the CD-Key is not the correct length."
                frmChat.DoDisconnect
                
            Else
                KeyHash = String$(HashSize, vbNullChar)
                Call kd_getHash(KDh, KeyHash)
                
                Value1 = kd_val1(KDh)
                ProductID = kd_product(KDh)
                
            End If
            
        End If
        
    End If
    
    If KDh > 0 Then
        Call kd_free(KDh)
    End If
End Sub
