Option Strict Off
Option Explicit On
Friend Class clsQueue
	' clsQueue.cls
	' Copyright (C) 2008 Eric Evans
	
	
	Private Const MAX_PRIORITY_LEVEL As Short = 100
	
	Private m_QueueObjs() As clsQueueOBj
	Private m_objCount As Integer
	Private m_lastUser As String
	Private m_lastObjID As Double
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		m_lastObjID = 1
		
		Clear()
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		Clear()
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Function Push(ByRef obj As clsQueueOBj) As Object
		On Error GoTo ERROR_HANDLER
		
		Dim Index As Integer
		Dim I As Integer
		
		Index = m_objCount
		
		If (m_objCount >= 1) Then
			For I = 0 To m_objCount - 1
				If (obj.PRIORITY < m_QueueObjs(I).PRIORITY) Then
					Index = I
					
					Exit For
				End If
			Next I
			
			ReDim Preserve m_QueueObjs(m_objCount)
		End If
		
		If (Index < m_objCount) Then
			For I = m_objCount To Index + 1 Step -1
				m_QueueObjs(I) = m_QueueObjs(I - 1)
			Next I
		End If
		
		obj.ID = m_lastObjID
		
		m_QueueObjs(Index) = obj
		
		m_objCount = (m_objCount + 1)
		m_lastObjID = (m_lastObjID + 1)
		
		RunInAll("Event_MessageQueued", obj.ID, obj.Message, obj.Tag)
		
		Exit Function
		
ERROR_HANDLER: 
		
		' overflow - likely due to message id size
		If (Err.Number = 6) Then
			m_lastObjID = 0
			
			Resume Next
		End If
		
		Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: " & Err.Description & " in clsQueue::Push().")
		
		Exit Function
	End Function
	
	Public Function Pop() As clsQueueOBj
		Pop = New clsQueueOBj
		
		Pop = m_QueueObjs(0)
		
		RemoveItem(0)
		
	End Function ' end function Pop
	
	Public Function Peek() As clsQueueOBj
		Peek = New clsQueueOBj
		
		Peek = m_QueueObjs(0)
	End Function ' end function Peek
	
	Public Function Item(ByVal Index As Integer) As Object
		If ((Index < 0) Or (Index > m_objCount - 1)) Then
			Item = New clsQueueOBj
			
			Exit Function
		End If
		
		Item = m_QueueObjs(Index)
	End Function
	
	Public Function ItemByID(ByVal I As Double) As Object
		Dim j As Integer
		
		For j = 0 To m_objCount - 1
			If (m_QueueObjs(j).ID = I) Then
				ItemByID = m_QueueObjs(j)
				
				Exit Function
			End If
		Next j
		
		ItemByID = New clsQueueOBj
	End Function
	
	Public ReadOnly Property Count() As Integer
		Get
			Count = m_objCount
		End Get
	End Property
	
	Public Function RemoveLines(ByVal match As String) As Short
		Dim curQueueObj As clsQueueOBj
		Dim I As Integer
		Dim found As Integer
		
		Do 
			curQueueObj = m_QueueObjs(I)
			
			If (PrepareCheck(curQueueObj.Message) Like PrepareCheck(match)) Then
				RemoveItem(I)
				
				found = (found + 1)
				
				I = 0
			Else
				I = (I + 1)
			End If
		Loop While (I < Count())
		
		RemoveLines = found
	End Function
	
	Public Sub RemoveItem(ByVal Index As Integer)
		Dim I As Integer
		
		If ((Index < 0) Or (Index > m_objCount - 1)) Then
			Exit Sub
		End If
		
		If (m_objCount > 1) Then
			For I = Index To ((m_objCount - 1) - 1)
				m_QueueObjs(I) = m_QueueObjs(I + 1)
			Next I
			
			ReDim Preserve m_QueueObjs(m_objCount - 1)
			
			m_objCount = (m_objCount - 1)
		Else
			Clear()
		End If
	End Sub
	
	Public Sub RemoveItemByID(ByVal I As Double)
		Dim j As Integer
		
		For j = 0 To m_objCount - 1
			If (m_QueueObjs(j).ID = I) Then
				RemoveItem(j)
				
				Exit Sub
			End If
		Next j
	End Sub
	
	Public Sub Clear()
		Dim I As Integer
		
		For I = 0 To m_objCount - 1
			'UPGRADE_NOTE: Object m_QueueObjs() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			m_QueueObjs(I) = Nothing
		Next I
		
		ReDim m_QueueObjs(0)
		
		m_QueueObjs(0) = New clsQueueOBj
		
		m_objCount = 0
		
		KillTimer(0, QueueTimerID)
		
		QueueTimerID = 0
		
		g_BNCSQueue.ClearQueue()
	End Sub
End Class