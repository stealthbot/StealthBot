Attribute VB_Name = "modSimilarity"
Option Explicit

' modSimilarity.bas -- project StealthBot
' Code thanks to MrRaza
' Original posting:
' http://forum.valhallalegends.com/phpbbs/index.php?board=17;action=display;threadid=517

Private b1() As Byte
Private b2() As Byte

Public Function Simil(String1 As String, String2 As String) As Double
    Dim l1 As Long
    Dim l2 As Long
    Dim L As Long
    Dim r As Double


    If UCase(String1) = UCase(String2) Then
        r = 1
    Else
        l1 = Len(String1)
        l2 = Len(String2)


        If l1 = 0 Or l2 = 0 Then
            r = 0
        Else
            ReDim b1(1 To l1): ReDim b2(1 To l2)


            For L = 1 To l1
                b1(L) = Asc(UCase(Mid(String1, L, 1)))
            Next


            For L = 1 To l2
                b2(L) = Asc(UCase(Mid(String2, L, 1)))
            Next
            r = SubSim(1, l1, 1, l2) / (l1 + l2) * 2
        End If
    End If
    Simil = r
    Erase b1
    Erase b2
End Function

Private Function SubSim(st1 As Long, end1 As Long, st2 As Long, end2 As Long) As Long
    Dim c1 As Long
    Dim c2 As Long
    Dim ns1 As Long
    Dim ns2 As Long
    Dim I As Long
    Dim max As Long
    If st1 > end1 Or st2 > end2 Or st1 <= 0 Or st2 <= 0 Then Exit Function


    For c1 = st1 To end1


        For c2 = st2 To end2
            I = 0


            Do Until b1(c1 + I) <> b2(c2 + I)
                I = I + 1


                If I > max Then
                    ns1 = c1
                    ns2 = c2
                    max = I
                End If
                If c1 + I > end1 Or c2 + I > end2 Then Exit Do
            Loop
        Next
    Next
    max = max + SubSim(ns1 + max, end1, ns2 + max, end2)
    max = max + SubSim(st1, ns1 - 1, st2, ns2 - 1)


    SubSim = max
End Function
