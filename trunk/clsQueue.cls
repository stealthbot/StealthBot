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
Private m_objCount    As Long
Private m_lastUser    As String
Private m_lastObjID   As Double

' ...
Private Sub Class_Initialize()

    ' ...
    m_lastObjID = 1

    ' ...
    Clear
    
End Sub ' end function Class_Initialize

' ...
Private Sub Class_Terminate()

    ' ...
    Clear
    
End Sub ' end function Class_Terminate

' ...
Public Function Push(ByRef obj As clsQueueOBj)

    On Error GoTo ERROR_HANDLER
    
    Dim Index As Long ' ...
    Dim i     As Long ' ...
    
    ' ...
    Index = m_objCount
    
    ' ...
    If (m_objCount >= 1) Then
        ' ...
        For i = 0 To m_objCount - 1
            ' ...
            If (obj.PRIORITY < m_QueueObjs(i).PRIORITY) Then
                ' ...
                Index = i
                
                ' ...
                Exit For
            End If
        Next i
    
        ' ...
        ReDim Preserve m_QueueObjs(0 To m_objCount)
    End If

    ' ...
    If (Index < m_objCount) Then
        For i = m_objCount To Index + 1 Step -1
            Set m_QueueObjs(i) = m_QueueObjs(i - 1)
        Next i
    End If
    
    ' ...
    obj.ID = m_lastObjID
    
    ' ...
    Set m_QueueObjs(Index) = obj
    
    ' ...
    m_objCount = (m_objCount + 1)
    m_lastObjID = (m_lastObjID + 1)

    ' ...
    RunInAll "Event_MessageQueued", obj.ID, obj.Message, obj.Tag
    
    ' ...
    Exit Function

ERROR_HANDLER:

    ' overflow - likely due to message id size
    If (Err.Number = 6) Then
        m_lastObjID = 0
        
        Resume Next
    End If

    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in clsQueue::Push().")

    Exit Function
    
End Function ' end function Push

' ...
Public Function Pop() As clsQueueOBj

    ' ...
    Set Pop = New clsQueueOBj

    ' ...
    Set Pop = m_QueueObjs(0)
    
    ' ...
    RemoveItem 0
    
    RecalculateDelay

End Function ' end function Pop

' ...
Public Function Peek() As clsQueueOBj

    ' ...
    Set Peek = New clsQueueOBj

    ' ...
    Set Peek = m_QueueObjs(0)
    
End Function ' end function Peek

' ...
Public Function Item(ByVal Index As Long) As Object

    ' ...
    If ((Index < 0) Or (Index > m_objCount - 1)) Then
        ' ...
        Set Item = New clsQueueOBj
    
        ' ...
        Exit Function
    End If

    ' ...
    Set Item = m_QueueObjs(Index)
    
End Function ' end function Item

' ...
Public Function ItemByID(ByVal i As Double) As Object

    ' ...
    Dim j As Long ' ...

    ' ...
    For j = 0 To m_objCount - 1
        If (m_QueueObjs(j).ID = i) Then
            Set ItemByID = m_QueueObjs(j)
        
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

' return the number of credits currently stored
Public Property Get Credits() As Long

    Credits = g_BNCSQueue.CurrentCredits

End Property

' returns the number of milliseconds until the message provided "would" be sent
' does not affect the queue
Public Function GetDelayFor(ByVal sText As String) As Long
    
    Dim Delay As Long, ls As Long, cr As Long
    ' call CalculateDelay with the provided text,
    ' and store the lastsent time and credits in the
    ' local variables (and ignore them...)
    Delay = g_BNCSQueue.CalculateDelay(sText, ls, cr)
    
    GetDelayFor = Delay
    
End Function

Public Function RemoveLines(ByVal match As String) As Integer

    Dim curQueueObj As clsQueueOBj ' ...
    Dim i           As Long     ' ...
    Dim found       As Long     ' ...
    
    ' ...
    Do
        ' ...
        Set curQueueObj = m_QueueObjs(i)

        ' ...
        If (PrepareCheck(curQueueObj.Message) Like PrepareCheck(match)) Then
            RemoveItem i
            
            found = (found + 1)
            
            i = 0
        Else
            i = (i + 1)
        End If
    Loop While (i < Count())
    
    If found Then
        RecalculateDelay
    End If
    
    RemoveLines = found

End Function

' ...
Public Sub RemoveItem(ByVal Index As Long)

    Dim i As Long ' ...

    ' ...
    If ((Index < 0) Or (Index > m_objCount - 1)) Then
        Exit Sub
    End If

    ' ...
    If (m_objCount > 1) Then
        ' ...
        For i = Index To ((m_objCount - 1) - 1)
            Set m_QueueObjs(i) = m_QueueObjs(i + 1)
        Next i
        
        ' ...
        ReDim Preserve m_QueueObjs(0 To m_objCount - 1)
        
        ' ...
        m_objCount = (m_objCount - 1)
        
        RecalculateDelay
    Else
        Clear
    End If
    
End Sub ' end function RemoveItem

' ...
Public Sub RemoveItemByID(ByVal i As Double)

    Dim j As Long ' ...
    
    ' ...
    For j = 0 To m_objCount - 1
        If (m_QueueObjs(j).ID = i) Then
            RemoveItem j
            
            RecalculateDelay
            
            Exit Sub
        End If
    Next j

End Sub

' ...
Public Sub Clear()

    Dim i As Long ' ...

    ' ...
    For i = 0 To m_objCount - 1
        Set m_QueueObjs(i) = Nothing
    Next i

    ' ...
    ReDim m_QueueObjs(0)
    
    ' ...
    Set m_QueueObjs(0) = New clsQueueOBj
    
    ' ...
    m_objCount = 0
    
    RecalculateDelay
    
End Sub ' end sub Clear

' after removing an item or clearing, redetermine time to next item
' if items remain
Private Sub RecalculateDelay()

    Dim NewDelay As Long, ExtraDelay As Long

    KillTimer 0&, QueueTimerID
    
    QueueTimerID = 0
    
    g_BNCSQueue.ClearQueue
    
    If (Count) Then
        With Peek()
            NewDelay = g_BNCSQueue.GetDelay(.Message)
            
            If .PRIORITY = PRIORITY.CHANNEL_MODERATION_MESSAGE Then
                ExtraDelay = g_BNCSQueue.BanDelay()
            End If
        End With
        
        QueueTimerID = SetTimer(0&, 0&, NewDelay + ExtraDelay, _
            AddressOf QueueTimerProc)
    End If
    
End Sub
