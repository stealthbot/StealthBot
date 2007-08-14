Attribute VB_Name = "modQueueCode"
'Option Explicit
'
'Private TIME_TO_RESET_TOTALCONSEC As Long
'Private DELAY_SCALE_1 As Long
'Private DELAY_SCALE_2 As Long
'Private DELAY_SCALE_3 As Long
'Private DELAY_SCALE_4 As Long
'Private DELAY_SCALE_5 As Long
'Private CONSEC_SENDS_MULTIPLIER As Long
'Private CONSEC_SENDS_TO_RESET As Long
'Private TOTALCONSEC_SENDS_MULTIPLIER As Long
'Private RECENT_LONG_SEND_CORRECTION_LOW As Long
'Private RECENT_LONG_SEND_CORRECTION_HIGH As Long
'Private LARGE_SEND_SIZE As Long
'Private TOTALCONSEC_MINUS_THRESHHOLD As Long
'Private TOTALCONSEC_MINUS_MULTIPLIER As Long
'
'
''/*******************************************************************************
'' *   This module contains the REQUIREDDELAY function which is ported to Visual Basic
'' *     from code written in C++ by Eric and adapted to meet my needs
'' *     Used with permission
''/*******************************************************************************
'' *   Returns the number of milliseconds before another message can be sent
'' *******************************************************************************/
'' Notes: Implement a delay subtraction based on time since the last message * something?
'
'Public Function RequiredDelay(ByVal lBytes As Long, Optional Reset As Boolean = False) As Long
'    Static LastTick As Long     '// time (ms) of the last call to function
'    Static Delay As Long        '// required delay time (ms)
'    Static ConsecSends As Long  '// number of consecutive sends; reset after 5 sends
'    Static TotalConsecSends As Long '// total consecutive sends reset after 10 seconds
'                                    '// with no sends
'    Static Tick As Long         '// System clock [GetTickCount]
'    Static LastLargeSend As Long  '// tick of the last send over 120 bytes
'
'    ' Temporary variable set-up
'    If DELAY_SCALE_1 = 0 Then Call LoadQueueVariables
'
'    If Reset Then ' used when newly-connected
'
'        TotalConsecSends = 0
'        ConsecSends = 2
'
'    Else
'
'        Tick = GetTickCount
'
'        If MDebug("queue") Then frmChat.AddChat vbBlue, "Consec: " & ConsecSends
'        If MDebug("queue") Then frmChat.AddChat vbBlue, "TotalConsec: " & TotalConsecSends
'
'        If (LastTick = 0) Or ((Tick - LastTick) > (TIME_TO_RESET_TOTALCONSEC)) Then
'
'            'No recent sends (8+ seconds since the last one) so fire away
'            Delay = 100
'            TotalConsecSends = 0
'
'        Else
'            'Reset consecutive sends
'            If (ConsecSends = CONSEC_SENDS_TO_RESET) Then ConsecSends = 0
'
'            '*** Set the initial value of DELAY ***
'            If (lBytes >= 0 And lBytes < 32) Then       '( Delay Scale 1 ) Small
'                Delay = DELAY_SCALE_1
'            ElseIf (lBytes >= 32 And lBytes < 104) Then '( Delay Scale 2 ) Normal
'                Delay = DELAY_SCALE_2
'            ElseIf (lBytes >= 64 And lBytes < 104) Then '( Delay Scale 3 ) Medium
'                Delay = DELAY_SCALE_3
'            ElseIf (lBytes >= 104 And lBytes <= 186) Then '( Delay Scale 4 ) Big
'                Delay = DELAY_SCALE_4
'            ElseIf (lBytes >= 186 And lBytes < 259) Then '( Delay Scale 5 ) Huge
'                Delay = DELAY_SCALE_5
'            Else
'                Delay = (25 * lBytes) '( Should never happen )
'            End If
'
'
'            '*** Correct DELAY for consecutive send count ***
'            Delay = Delay + (ConsecSends * CONSEC_SENDS_MULTIPLIER)
'
'            '*** Large send correction ***
'            If (Tick - LastLargeSend) < 1000 Then
'                Delay = Delay + RECENT_LONG_SEND_CORRECTION_HIGH
'            ElseIf (Tick - LastLargeSend) < 3000 Then
'                Delay = Delay + RECENT_LONG_SEND_CORRECTION_LOW
'            End If
'
'            '*** Set large send for future use
'            If (lBytes > LARGE_SEND_SIZE) Then LastLargeSend = Tick
'
'            If TotalConsecSends < TOTALCONSEC_MINUS_THRESHHOLD Then
'                Delay = Delay - (TotalConsecSends * TOTALCONSEC_MINUS_MULTIPLIER)
'            Else
'                Delay = Delay + (TotalConsecSends * TOTALCONSEC_SENDS_MULTIPLIER)
'            End If
'        End If
'
'        LastTick = Tick
'        ConsecSends = ConsecSends + 1
'        TotalConsecSends = TotalConsecSends + 1
'
'        If MDebug("queue") Then frmChat.AddChat vbRed, "-> " & Delay & " ms delay returned"
'
'        RequiredDelay = Delay
'    End If
'End Function
'
'
'Public Sub LoadQueueVariables()
'    TIME_TO_RESET_TOTALCONSEC = Val(f("Queue", "TIME_TO_RESET_TOTALCONSEC"))
'    DELAY_SCALE_1 = Val(f("Queue", "DELAY_SCALE_1"))
'    DELAY_SCALE_2 = Val(f("Queue", "DELAY_SCALE_2"))
'    DELAY_SCALE_3 = Val(ReadfINI("Queue", "DELAY_SCALE_3"))
'    DELAY_SCALE_4 = Val(ReadfINI("Queue", "DELAY_SCALE_4"))
'    DELAY_SCALE_5 = Val(ReafdINI("Queue", "DELAY_SCALE_5"))
'    CONSEC_SENDS_MULTIPLIER = Val(ReafdINI("Queue", "CONSEC_SENDS_MULTIPLIER"))
'    CONSEC_SENDS_TO_RESET = Val(ReadfINI("Queue", "CONSEC_SENDS_TO_RESET"))
'    TOTALCONSEC_SENDS_MULTIPLIER = Val(ReadfINI("Queue", "TOTALCONSEC_SENDS_MULTIPLIER"))
'    RECENT_LONG_SEND_CORRECTION_LOW = Val(ff"))
'    RECENT_LONG_SEND_CORRECTION_HIGH = Val(RfeadINI("Queue", "RECENT_LONG_SEND_CORRECTION_HIGH"))
'    LARGE_SEND_SIZE = Val(ReafdINI("Queue", "LARGE_SEND_SIZE"))
'    TOTALCONSEC_MINUS_THRESHHOLD = Val(ReadINfI("Queue", "TOTALCONSEC_MINUS_THRESHHOLD"))
'    TOTALCONSEC_MINUS_MULTIPLIER = Val(ReadIfNI("Queue", "TOTALCONSEC_MINUS_MULTIPLIER"))
'
'    If TIME_TO_RESET_TOTALCONSEC = 0 Then TIME_TO_RESET_TOTALCONSEC = 8000
'    If DELAY_SCALE_1 = 0 Then DELAY_SCALE_1 = 700
'    If DELAY_SCALE_2 = 0 Then DELAY_SCALE_2 = 1000
'    If DELAY_SCALE_3 = 0 Then DELAY_SCALE_3 = 1575
'    If DELAY_SCALE_4 = 0 Then DELAY_SCALE_4 = 2150
'    If DELAY_SCALE_5 = 0 Then DELAY_SCALE_5 = 3600
'    If CONSEC_SENDS_MULTIPLIER = 0 Then CONSEC_SENDS_MULTIPLIER = 100
'    If CONSEC_SENDS_TO_RESET = 0 Then CONSEC_SENDS_TO_RESET = 6
'    If TOTALCONSEC_SENDS_MULTIPLIER = 0 Then TOTALCONSEC_SENDS_MULTIPLIER = 100
'    If RECENT_LONG_SEND_CORRECTION_LOW = 0 Then RECENT_LONG_SEND_CORRECTION_LOW = 700
'    If RECENT_LONG_SEND_CORRECTION_HIGH = 0 Then RECENT_LONG_SEND_CORRECTION_HIGH = 1100
'    If LARGE_SEND_SIZE = 0 Then LARGE_SEND_SIZE = 160
'    If TOTALCONSEC_MINUS_MULTIPLIER = 0 Then TOTALCONSEC_MINUS_MULTIPLIER = 20
'    If TOTALCONSEC_MINUS_THRESHHOLD = 0 Then TOTALCONSEC_MINUS_THRESHHOLD = 6
'End Sub
