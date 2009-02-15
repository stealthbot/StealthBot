VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsQueue.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Private Const MAX_PRIORITY_LEVEL = 100

Private m_QueueObjs() As clsQueueOBj
Private m_objCount    As Integer
Private m_lastUser    As String
Private m_lastObjID   As Double

' ...
Private Sub Class_Initialize()

    ' ...
    m_lastObjID = 1

    ' ...
    Call Clear
    
End Sub ' end function Class_Initialize

' ...
Private Sub Class_Terminate()

    ' ...
    Call Clear
    
End Sub ' end function Class_Terminate

' ...
Public Function Push(ByRef obj As clsQueueOBj)
    
    On Error GoTo ERROR_HANDLER
    
    Dim Index As Integer ' ...
    Dim I     As Integer ' ...
    
    ' ...
    Index = m_objCount
    
    ' ...
    If (m_objCount) Then
        ' ...
        For I = 0 To m_objCount - 1
            ' ...
            If (obj.PRIORITY < m_QueueObjs(I).PRIORITY) Then
                ' ...
                Index = I
                
                ' ...
                Exit For
            End If
        Next I
    
        ' ...
        ReDim Preserve m_QueueObjs(0 To m_objCount)
    End If

    ' ...
    If (Index < m_objCount) Then
        For I = m_objCount To Index + 1 Step -1
            Set m_QueueObjs(I) = m_QueueObjs(I - 1)
        Next I
    End If
    
    ' ...
    Set m_QueueObjs(Index) = obj
    
    ' ...
    With m_QueueObjs(Index)
        .id = m_lastObjID
    End With
    
    ' ...
    On Error Resume Next
    
    ' ...
    frmChat.SControl.Run "Event_MessageQueued", m_lastObjID, obj.Message, obj.Tag
    
    ' ...
    m_objCount = (m_objCount + 1)
    m_lastObjID = (m_lastObjID + 1)
    
    ' ...
    Exit Function

ERROR_HANDLER:

    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsQueue::Push().")

    Exit Function
    
End Function ' end function Push

' ...
Public Function Pop() As clsQueueOBj

    On Error GoTo ERROR_HANDLER

    Dim I As Integer ' ...

    ' ...
    Set Pop = m_QueueObjs(0)
    
    ' ...
    RemoveItem 0
    
    ' ...
    Exit Function

ERROR_HANDLER:

    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsQueue::Pop().")

    Exit Function
    
End Function ' end function Pop

' ...
Public Function Peek() As clsQueueOBj

    ' ...
    Set Peek = m_QueueObjs(0)
    
End Function ' end function Peek

' ...
Public Function Item(ByVal Index As Integer) As Object

    ' ...
    If ((Index < 0) Or (Index + 1 > m_objCount)) Then
        ' ...
        Set Item = New clsQueueOBj
    
        ' ...
        Exit Function
    End If

    ' ...
    Set Item = m_QueueObjs(Index - 1)
    
End Function ' end function Item

' ...
Public Function ItemByID(ByVal I As Integer) As Object

    ' ...
    Dim j As Integer ' ...
    
    ' ...
    For j = 0 To (m_objCount - 1)
        If (m_QueueObjs(I).id = I) Then
            Set ItemByID = m_QueueObjs(I)
        
            Exit Function
        End If
    Next j
    
    ' ...
    Set ItemByID = New clsQueueOBj

End Function

' ...
Public Property Get Count() As Long

    Count = m_objCount
    
End Property

Public Function RemoveLines(ByVal match As String) As Integer

    Dim curQueueObj As clsQueueOBj ' ...
    Dim I           As Integer     ' ...
    Dim found       As Integer     ' ...
    
    ' ...
    Do
        ' ...
        Set curQueueObj = m_QueueObjs(I)

        ' ...
        If (PrepareCheck(curQueueObj.Message) Like PrepareCheck(match)) Then
            RemoveItem I
            
            found = (found + 1)
            
            I = 0
        Else
            I = (I + 1)
        End If
    Loop While (I < Count())
    
    RemoveLines = found

End Function

' ...
'Public Sub RemoveLines(Optional ByVal strLine As String = vbNullString, Optional ByVal strUser As String = _
'    vbNullString)
'
'    Dim curQueueObj As clsQueueOBj
'    Dim i           As Integer ' ...
'    Dim bln         As Boolean ' ...
'    Dim j           As Integer ' ...
'    Dim tmp         As String  ' ...
'
'    ' ...
'    Do
'        ' reset boolean
'        i = 0
'        bln = False
'
'        ' ...
'        For i = i To (m_objCount - 1)
'            ' ...
'            Set curQueueObj = m_QueueObjs(i)
'
'            ' ...
'            If (strLine <> vbNullString) Then
'                ' ...
'                If (Len(curQueueObj.Message) >= Len(strLine)) Then
'                    ' ...
'                    tmp = Left$(curQueueObj.Message, Len(strLine))
'
'                    ' ...
'                    If (StrComp(tmp, strLine, vbTextCompare) = 0) Then
'                        ' ...
'                        If (strUser <> vbNullString) Then
'                            ' ...
'                            If (curQueueObj.ResponseTo = strUser) Then
'                                ' ...
'                                bln = True
'                            End If
'                        Else
'                            ' ...
'                            bln = True
'                        End If
'                    End If
'                End If
'            Else
'                ' ...
'                If (curQueueObj.ResponseTo = strUser) Then
'                    ' ...
'                    bln = True
'                End If
'            End If
'
'            ' ...
'            If (bln) Then
'                Exit For
'            End If
'        Next i
'
'        ' ...
'        If (bln) Then
'            ' ...
'            For j = i + 1 To m_objCount - 1
'                Set m_QueueObjs(j - 1) = m_QueueObjs(j)
'            Next j
'
'            ' ...
'            ReDim Preserve m_QueueObjs(0 To (m_objCount - 1))
'
'            ' ...
'            m_objCount = (m_objCount - 1)
'        End If
'    Loop While (bln = True)
'
'End Sub ' end function RemoveLines

'Private Function FindLines(Optional ByVal strLine As String = vbNullString, Optional ByVal strUser As String = _
'    vbNullString, Optional PriLower As Integer = MAX_PRIORITY_LEVEL) As Integer
'
'    Dim curQueueObj As clsQueueOBj
'    Dim i           As Integer ' ...
'    Dim bln         As Boolean ' ...
'    Dim j           As Integer ' ...
'    Dim tmp         As String  ' ...
'
'    ' ...
'    For i = 0 To (m_objCount - 1)
'        ' ...
'        Set curQueueObj = m_QueueObjs(i)
'
'        ' ...
'        If (curQueueObj.PRIORITY <= PriLower) Then
'            ' ...
'            If (strLine <> vbNullString) Then
'                ' ...
'                If (Len(curQueueObj.Message) >= Len(strLine)) Then
'                    ' ...
'                    tmp = Left$(curQueueObj.Message, Len(strLine))
'
'                    ' ...
'                    If (StrComp(tmp, strLine, vbTextCompare) = 0) Then
'                        ' ...
'                        If (strUser <> vbNullString) Then
'                            ' ...
'                            If (curQueueObj.ResponseTo = strUser) Then
'                                ' ...
'                                bln = True
'                            End If
'                        Else
'                            ' ...
'                            bln = True
'                        End If
'                    End If
'                End If
'            Else
'                ' ...
'                If (curQueueObj.ResponseTo = strUser) Then
'                    ' ...
'                    bln = True
'                End If
'            End If
'        End If
'
'        ' ...
'        If (bln) Then
'            ' ...
'            FindLines = (FindLines + 1)
'
'            ' ...
'            bln = False
'        End If
'    Next i
'
'End Function

' ...
'Private Function GetNextQueueObj() As Integer
'
'    On Error GoTo ERROR_HANDLER
'
'    Static lastLowestIndex As Integer
'
'    Dim curQueueObj As clsQueueOBj
'    Dim i           As Integer
'    Dim lowest      As Integer
'    Dim lowestIndex As Integer
'
'    ' ...
'    lowest = MAX_PRIORITY_LEVEL
'
'    ' ....
'    For i = 0 To (m_objCount - 1)
'        ' ...
'        Set curQueueObj = m_QueueObjs(i)
'
'        ' ...
'        If (curQueueObj.PRIORITY < lowest) Then
'            ' ...
'            lowest = curQueueObj.PRIORITY
'
'            ' ...
'            lowestIndex = i
'        End If
'    Next i
'
'    ' ...
'    If (m_lastUser <> vbNullString) Then
'        Dim j   As Integer ' ...
'        Dim k   As Integer ' ...
'        Dim bln As Boolean ' ...
'
'        ' ...
'        i = lastLowestIndex
'
'        ' ...
'        For i = i To (m_objCount - 1)
'            ' ...
'            Set curQueueObj = m_QueueObjs(i)
'
'            ' ...
'            If (curQueueObj.PRIORITY = lowest) Then
'                ' ...
'                If (StrComp(curQueueObj.ResponseTo, m_lastUser, vbTextCompare) <> 0) Then
'                    ' ...
'                    For k = 0 To (i - 1)
'                        ' ...
'                        If (m_QueueObjs(k).PRIORITY = lowest) Then
'                            ' ...
'                            If (StrComp(m_QueueObjs(k).ResponseTo, curQueueObj.ResponseTo, _
'                                    vbTextCompare) = 0) Then
'
'                                Exit For
'                            End If
'                        End If
'                    Next k
'
'                    ' ...
'                    If (k >= i) Then
'                        ' ...
'                        lowestIndex = i
'
'                        ' ...
'                        Exit For
'                    End If
'                End If
'            End If
'        Next i
'    End If
'
'    ' ...
'    lastLowestIndex = lowestIndex
'
'    ' ...
'    GetNextQueueObj = lastLowestIndex
'
'    Exit Function
'
'ERROR_HANDLER:
'    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsQueue::GetNextQueueObj().")
'
'    Exit Function
'
'End Function ' end function GetNextQueueObj

' ...
Public Sub RemoveItem(ByVal Index As Integer)

    Dim I As Integer ' ...

    ' ...
    If ((Index < 0) Or (Index + 1 > m_objCount)) Then
        Exit Sub
    End If

    ' ...
    If (m_objCount) Then
        ' ...
        For I = Index To ((m_objCount - 1) - 1)
            Set m_QueueObjs(I) = m_QueueObjs(I + 1)
        Next I
        
        ' ...
        ReDim Preserve m_QueueObjs(0 To m_objCount - 1)
        
        ' ...
        m_objCount = (m_objCount - 1)
    Else
        Call Clear
    End If
    
End Sub ' end function RemoveItem

' ...
Public Sub RemoveItemByID(ByVal I As Integer)

    Dim j As Integer ' ...
    
    ' ...
    For j = 0 To (m_objCount - 1)
        If (m_QueueObjs(I).id = I) Then
            RemoveItem j
            
            Exit Sub
        End If
    Next j

End Sub

' ...
Public Sub Clear()

    Dim I As Integer ' ...

    ' ...
    For I = 0 To m_objCount - 1
        Set m_QueueObjs(I) = Nothing
    Next I

    ' ...
    ReDim m_QueueObjs(0) As clsQueueOBj
    
    ' ...
    Set m_QueueObjs(0) = New clsQueueOBj

    ' ...
    m_objCount = 0
    
End Sub ' end sub Clear
