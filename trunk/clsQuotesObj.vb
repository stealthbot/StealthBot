Option Strict Off
Option Explicit On
Friend Class clsQuotesObj
	' clsQuotesObj.cls
	' Copyright (C) 2009 Nate Book
	' This object mirrors a collection with the added ability to save quotes as well.
	' -Ribose/2009-08-10
	
	Private Const OBJECT_NAME As String = "clsQuotesObj"
	
	' actual quotes collection
	Private m_Quotes As Collection
	
	' this is cleared so as to not save on each line during load
	Private m_blnIsLoaded As Boolean
	
	' load quotes on object create
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		On Error Resume Next
		
		' load quotes
		LoadQuotes()
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		'UPGRADE_NOTE: Object m_Quotes may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_Quotes = Nothing
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	' call this to add quotes to the collection
	' empty entries are ignored
	' duplicate entries are ignored
	Public Function Add(ByVal QuoteStr As String) As Integer
		
		On Error GoTo ERROR_HANDLER
		
		Add = 0
		
		' trim
		QuoteStr = Trim(QuoteStr)
		
		' is empty?
		If Len(QuoteStr) = 0 Then Exit Function
		
		' already exists?
		If GetIndexOf(QuoteStr) > 0 Then Exit Function
		
		' add
		m_Quotes.Add(QuoteStr)
		
		' is not loading from file?
		If m_blnIsLoaded Then
			
			' save
			AppendQuote(QuoteStr)
			
		End If
		
		Add = m_Quotes.Count()
		
		Exit Function
		
ERROR_HANDLER: 
		
		frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in {2}.Add()", Err.Number, Err.Description, OBJECT_NAME))
	End Function
	
	' call this to reqmove quotes from the collection
	' ignores nonexistant items
	Public Function Remove(ByVal QuoteStr As Object) As String
		
		On Error GoTo ERROR_HANDLER
		
		Remove = vbNullString
		
		Dim Index As Integer
		
		' if its numeric, try seeing if we can remove at that index
		If IsNumeric(QuoteStr) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object QuoteStr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Index = CInt(QuoteStr)
			If Index < 1 Or Index > m_Quotes.Count() Then
				' no? check if its one of the quotes
				'UPGRADE_WARNING: Couldn't resolve default property of object QuoteStr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Index = GetIndexOf(QuoteStr)
				If Index = 0 Then Exit Function
			End If
		Else
			' not numeric- check if its one of the quotes
			'UPGRADE_WARNING: Couldn't resolve default property of object QuoteStr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Index = GetIndexOf(QuoteStr)
			If Index = 0 Then Exit Function
		End If
		
		' store
		'UPGRADE_WARNING: Couldn't resolve default property of object m_Quotes.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Remove = m_Quotes.Item(Index)
		
		m_Quotes.Remove(Index)
		
		' is not loading from file?
		If m_blnIsLoaded Then
			SaveQuotes()
		End If
		
		Exit Function
		
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in {2}.Remove()", Err.Number, Err.Description, OBJECT_NAME))
		
	End Function
	
	' gets the quote by its index
	Public Function Item(ByVal Index As Integer) As String
		
		On Error Resume Next
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_Quotes.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Item = m_Quotes.Item(Index)
		
	End Function
	
	' gets the items in the collection
	Public ReadOnly Property Items() As Collection
		Get
			
			Dim i As Integer
			
			' clone the collection - modification to this collection ignored
			Items = New Collection
			For i = 1 To m_Quotes.Count()
				
				Items.Add(m_Quotes.Item(i))
				
			Next i
			
		End Get
	End Property
	
	
	' gets the number of quotes
	Public ReadOnly Property Count() As Integer
		Get
			
			On Error Resume Next
			
			Count = m_Quotes.Count()
			
		End Get
	End Property
	
	' gets the index of the quote by string
	Public Function GetIndexOf(ByVal QuoteStr As String) As Integer
		
		On Error Resume Next
		
		Dim i As Integer
		
		For i = 1 To m_Quotes.Count()
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_Quotes.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If StrComp(m_Quotes.Item(i), QuoteStr, CompareMethod.Text) = 0 Then
				
				GetIndexOf = i
				
				Exit Function
				
			End If
			
		Next i
		
		GetIndexOf = 0
		
	End Function
	
	' returns a random quote
	Public Function GetRandomQuote() As String
		
		On Error GoTo ERROR_HANDLER
		
		Dim iRand As Short
		Dim sQuote As String
		
		If m_Quotes.Count() = 0 Then
			GetRandomQuote = vbNullString
			Exit Function
		End If
		
		' randomly select quote
		Randomize()
		iRand = (Rnd() * m_Quotes.Count())
		
		' get quote in collection
		If (iRand + 1 <= m_Quotes.Count()) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object m_Quotes.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sQuote = m_Quotes.Item(iRand + 1)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object m_Quotes.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sQuote = m_Quotes.Item(m_Quotes.Count())
		End If
		
		' security check
		'UPGRADE_WARNING: Couldn't resolve default property of object StringFormat(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Left(sQuote, 1) = "/" Then sQuote = StringFormat(" {0}", sQuote)
		GetRandomQuote = sQuote
		
		Exit Function
		
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in {2}.GetRandomQuote()", Err.Number, Err.Description, OBJECT_NAME))
		
	End Function
	
	' this function will load quotes into the collection
	Private Sub LoadQuotes()
		
		On Error GoTo ERROR_HANDLER
		
		m_blnIsLoaded = False
		
		m_Quotes = ListFileLoad(GetFilePath(FILE_QUOTES))
		
		m_blnIsLoaded = True
		
		Exit Sub
		
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in {2}.LoadQuotes()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	' this function will save quotes into quotes.txt when changes are made
	Private Sub SaveQuotes()
		
		On Error GoTo ERROR_HANDLER
		
		Call ListFileSave(GetFilePath(FILE_QUOTES), m_Quotes)
		
		Exit Sub
		
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in {2}.SaveQuotes()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
	
	' this function will save a single quote into quotes.txt on add (instead of saving the whole thing)
	Private Sub AppendQuote(ByVal QuoteStr As String)
		
		On Error GoTo ERROR_HANDLER
		
		Call ListFileAppendItem(GetFilePath(FILE_QUOTES), QuoteStr)
		
		Exit Sub
		
ERROR_HANDLER: 
		frmChat.AddChat(RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in {2}.AppendQuotes()", Err.Number, Err.Description, OBJECT_NAME))
	End Sub
End Class