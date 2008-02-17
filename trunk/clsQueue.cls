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

Private m_QueueObjs() As clsQueueOBj
Private m_objCount    As Integer

' ...
Public Function Push(ByRef obj As clsQueueOBj)
    ' ...
    ReDim Preserve m_QueueObjs(0 To m_objCount + 1)
    
    ' ...
    m_QueueObjs(m_objCount + 1) = obj
    
    ' ...
    m_objCount = (m_objCount + 1)
End Function ' end function Push

' ...
Public Function Pop() As clsQueueOBj
    ' ...
    Set Pop = m_QueueObjs(GetNextQueueObj)
End Function ' end function Pop

' ...
Public Function Peek() As clsQueueOBj
    ' ...
    Set Peek = m_QueueObjs(GetNextQueueObj)
End Function ' end function Peek

' ...
Public Property Get Count() As Long
    Count = (m_objCount + 1)
End Property

' ...
Private Function GetNextQueueObj() As Integer
    Static lastUser        As String
    Static lastLowest      As Integer
    Static lastLowestIndex As Integer

    Dim curQueueObj As clsQueueOBj
    Dim i           As Integer
    Dim lowest      As Integer
    Dim lowestIndex As Integer
    
    ' ....
    For i = 0 To m_objCount
        ' ...
        Set curQueueObj = m_QueueObjs(i)
        
        ' ...
        If (curQueueObj.Priority < lowest) Then
            ' ...
            lowest = curQueueObj.Priority
            
            ' ...
            lowestIndex = i
        End If
    Next i
    
    ' ...
    If (lastUser <> vbNullString) Then
        Dim j As Integer ' ...
    
        ' ...
        If (lastLowest = lowest) Then
            i = (lastLowestIndex + 1)
        Else
            i = 0
        End If
        
        ' ...
        For j = 1 To 2 Step 1
            ' ...
            For i = 0 To m_objCount
                ' ...
                Set curQueueObj = m_QueueObjs(i)
                
                ' ...
                If (curQueueObj.Priority = lowest) Then
                    ' ...
                    If (StrComp(curQueueObj.ResponseTo, lastUser, vbTextCompare) <> 0) Then
                        ' ...
                        lowestIndex = i
                    
                        ' ...
                        Exit For
                    End If
                End If
            Next i
            
            ' ...
            If (lowestIndex <> lastLowestIndex) Then
                Exit For
            End If
            
            ' ...
            i = 0
        Next j
    End If
    
    ' ...
    lastUser = m_QueueObjs(lowestIndex).ResponseTo
    
    ' ...
    lastLowest = lowest
    
    ' ...
    lastLowestIndex = lowestIndex
    
    ' ...
    GetNextQueueObj = lastLowestIndex
End Function ' end function GetNextQueueObj
