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

' ...
Private Sub Class_Initialize()
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

    ' ...
    If (m_objCount) Then
        Dim lines As Boolean ' ...
        
        ' ...
        'lines = (FindLines(obj.Message, obj.ResponseTo, obj.PRIORITY) > 0)
    
        ' ...
        'If (lines) Then
        '    ' ...
        '    Exit Function
        'Else
        '    ' ...
        '    Call RemoveLines(obj.Message, obj.ResponseTo)
        'End If
    
        ' ...
        ReDim Preserve m_QueueObjs(0 To m_objCount)
    End If
    
    ' ...
    Set m_QueueObjs(m_objCount) = obj
    
    ' ...
    m_objCount = (m_objCount + 1)
    
    Exit Function

ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsQueue::Push().")

    Exit Function
    
End Function ' end function Push

' ...
Public Function Pop() As clsQueueOBj

    On Error GoTo ERROR_HANDLER

    Dim qindex As Integer ' ...
    Dim I      As Integer ' ...
    
    ' ...
    qindex = GetNextQueueObj()

    ' ...
    Set Pop = m_QueueObjs(qindex)
    
    ' ...
    m_lastUser = Pop.ResponseTo
    
    ' ...
    If (m_objCount) Then
        ' ...
        For I = qindex To (m_objCount - 2)
            ' ...
            Set m_QueueObjs(I) = m_QueueObjs(I + 1)
        Next I
        
        ' ...
        ReDim Preserve m_QueueObjs(0 To (m_objCount - 1))
    Else
        ' ...
        Call Clear
    End If
    
    ' ...
    m_objCount = (m_objCount - 1)
    
    Exit Function

ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsQueue::Pop().")

    Exit Function
End Function ' end function Pop

' ...
Public Function Peek() As clsQueueOBj

    ' ...
    Set Peek = m_QueueObjs(GetNextQueueObj())
    
End Function ' end function Peek

' ...
Public Function Item(lItem As Long) As Object
    ' ...
    If (((lItem - 1) < 0) Or (lItem > m_objCount)) Then
        ' ...
        Set Item = New clsQueueOBj
    
        ' ...
        Exit Function
    End If

    ' ...
    Set Item = m_QueueObjs(lItem - 1)
End Function ' end function Item

' ...
Public Property Get Count() As Long
    Count = m_objCount
End Property

' ...
Public Sub RemoveLines(Optional ByVal strLine As String = vbNullString, Optional ByVal strUser As String = _
    vbNullString)

    Dim curQueueObj As clsQueueOBj
    Dim I           As Integer ' ...
    Dim bln         As Boolean ' ...
    Dim j           As Integer ' ...
    Dim tmp         As String  ' ...
    
    ' ...
    Do
        ' reset boolean
        I = 0
        bln = False
    
        ' ...
        For I = I To (m_objCount - 1)
            ' ...
            Set curQueueObj = m_QueueObjs(I)
        
            ' ...
            If (strLine <> vbNullString) Then
                ' ...
                If (Len(curQueueObj.Message) >= Len(strLine)) Then
                    ' ...
                    tmp = Left$(curQueueObj.Message, Len(strLine))
            
                    ' ...
                    If (StrComp(tmp, strLine, vbTextCompare) = 0) Then
                        ' ...
                        If (strUser <> vbNullString) Then
                            ' ...
                            If (curQueueObj.ResponseTo = strUser) Then
                                ' ...
                                bln = True
                            End If
                        Else
                            ' ...
                            bln = True
                        End If
                    End If
                End If
            Else
                ' ...
                If (curQueueObj.ResponseTo = strUser) Then
                    ' ...
                    bln = True
                End If
            End If
            
            ' ...
            If (bln) Then
                Exit For
            End If
        Next I
        
        ' ...
        If (bln) Then
            ' ...
            'frmChat.AddChat vbRed, "DEBUG: " & m_QueueObjs(I).Message
            
            ' ...
            For j = I + 1 To m_objCount - 1
                Set m_QueueObjs(j - 1) = m_QueueObjs(j)
            Next j
            
            ' ...
            ReDim Preserve m_QueueObjs(0 To (m_objCount - 1))
            
            ' ...
            m_objCount = (m_objCount - 1)
        End If
    Loop While (bln = True)
End Sub ' end function RemoveLines

Private Function FindLines(Optional ByVal strLine As String = vbNullString, Optional ByVal strUser As String = _
    vbNullString, Optional PriLower As Integer = MAX_PRIORITY_LEVEL) As Integer
    
    Dim curQueueObj As clsQueueOBj
    Dim I           As Integer ' ...
    Dim bln         As Boolean ' ...
    Dim j           As Integer ' ...
    Dim tmp         As String  ' ...

    ' ...
    For I = 0 To (m_objCount - 1)
        ' ...
        Set curQueueObj = m_QueueObjs(I)
    
        ' ...
        If (curQueueObj.PRIORITY <= PriLower) Then
            ' ...
            If (strLine <> vbNullString) Then
                ' ...
                If (Len(curQueueObj.Message) >= Len(strLine)) Then
                    ' ...
                    tmp = Left$(curQueueObj.Message, Len(strLine))
            
                    ' ...
                    If (StrComp(tmp, strLine, vbTextCompare) = 0) Then
                        ' ...
                        If (strUser <> vbNullString) Then
                            ' ...
                            If (curQueueObj.ResponseTo = strUser) Then
                                ' ...
                                bln = True
                            End If
                        Else
                            ' ...
                            bln = True
                        End If
                    End If
                End If
            Else
                ' ...
                If (curQueueObj.ResponseTo = strUser) Then
                    ' ...
                    bln = True
                End If
            End If
        End If
        
        ' ...
        If (bln) Then
            ' ...
            FindLines = (FindLines + 1)
            
            ' ...
            bln = False
        End If
    Next I
End Function

' ...
Private Function GetNextQueueObj() As Integer

    On Error GoTo ERROR_HANDLER

    Static lastLowestIndex As Integer

    Dim curQueueObj As clsQueueOBj
    Dim I           As Integer
    Dim lowest      As Integer
    Dim lowestIndex As Integer
    
    ' ...
    lowest = MAX_PRIORITY_LEVEL
    
    ' ....
    For I = 0 To (m_objCount - 1)
        ' ...
        Set curQueueObj = m_QueueObjs(I)
        
        ' ...
        If (curQueueObj.PRIORITY < lowest) Then
            ' ...
            lowest = curQueueObj.PRIORITY
            
            ' ...
            lowestIndex = I
        End If
    Next I
    
    ' ...
    If (m_lastUser <> vbNullString) Then
        Dim j   As Integer ' ...
        Dim k   As Integer ' ...
        Dim bln As Boolean ' ...
    
        ' ...
        I = lastLowestIndex

        ' ...
        For I = I To (m_objCount - 1)
            ' ...
            Set curQueueObj = m_QueueObjs(I)

            ' ...
            If (curQueueObj.PRIORITY = lowest) Then
                ' ...
                If (StrComp(curQueueObj.ResponseTo, m_lastUser, vbTextCompare) <> 0) Then
                    ' ...
                    For k = 0 To (I - 1)
                        ' ...
                        If (m_QueueObjs(k).PRIORITY = lowest) Then
                            ' ...
                            If (StrComp(m_QueueObjs(k).ResponseTo, curQueueObj.ResponseTo, _
                                    vbTextCompare) = 0) Then
                                    
                                Exit For
                            End If
                        End If
                    Next k
                    
                    ' ...
                    If (k >= I) Then
                        ' ...
                        lowestIndex = I

                        ' ...
                        Exit For
                    End If
                End If
            End If
        Next I
    End If

    ' ...
    lastLowestIndex = lowestIndex
    
    ' ...
    GetNextQueueObj = lastLowestIndex
    
    Exit Function

ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsQueue::GetNextQueueObj().")

    Exit Function
    
End Function ' end function GetNextQueueObj

' ...
Private Sub RemoveItem(ByVal index As Integer)

End Sub ' end function RemoveItem

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
