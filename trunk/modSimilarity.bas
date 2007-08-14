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
    Dim l As Long
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


            For l = 1 To l1
                b1(l) = Asc(UCase(Mid(String1, l, 1)))
            Next


            For l = 1 To l2
                b2(l) = Asc(UCase(Mid(String2, l, 1)))
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
    Dim i As Long
    Dim max As Long
    If st1 > end1 Or st2 > end2 Or st1 <= 0 Or st2 <= 0 Then Exit Function


    For c1 = st1 To end1


        For c2 = st2 To end2
            i = 0


            Do Until b1(c1 + i) <> b2(c2 + i)
                i = i + 1


                If i > max Then
                    ns1 = c1
                    ns2 = c2
                    max = i
                End If
                If c1 + i > end1 Or c2 + i > end2 Then Exit Do
            Loop
        Next
    Next
    max = max + SubSim(ns1 + max, end1, ns2 + max, end2)
    max = max + SubSim(st1, ns1 - 1, st2, ns2 - 1)


    SubSim = max
End Function
