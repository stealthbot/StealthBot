VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQuotesObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsQuotesObj.cls
' Copyright (C) 2009 Nate Book
' This object mirrors a collection with the added ability to save quotes as well.
' -Ribose/2009-08-10
Option Explicit

' actual quotes collection
Private m_Quotes As Collection

' this is cleared so as to not save on each line during load
Private m_blnIsLoaded As Boolean

' load quotes on object create
Private Sub Class_Initialize()
    
    On Error Resume Next
    
    ' load quotes
    LoadQuotes
    
End Sub

Private Sub Class_Terminate()

    Set m_Quotes = Nothing

End Sub

' call this to add quotes to the collection
' empty entries are ignored
' duplicate entries are ignored
Public Function Add(ByVal QuoteStr As String) As Long

    On Error GoTo ERROR_HANDLER
    
    Add = 0

    ' trim
    QuoteStr = Trim$(QuoteStr)
    
    ' is empty?
    If Len(QuoteStr) = 0 Then Exit Function
    
    ' already exists?
    If GetIndexOf(QuoteStr) > 0 Then Exit Function
        
    ' add
    m_Quotes.Add QuoteStr
    
    ' is not loading from file?
    If m_blnIsLoaded Then
    
        ' save
        AppendQuote QuoteStr
        
    End If
    
    Add = m_Quotes.Count
    
    Exit Function
    
ERROR_HANDLER:
    
    MsgBox "Error (#" & Err.Number & "): " & Err.description & " in Quotes.Add()."
    
End Function

' call this to reqmove quotes from the collection
' ignores nonexistant items
Public Function Remove(ByVal QuoteStr As Variant) As String
    
    On Error GoTo ERROR_HANDLER
    
    Remove = vbNullString

    Dim Index As Long
    
    ' if its numeric, try seeing if we can remove at that index
    If IsNumeric(QuoteStr) Then
        Index = CLng(QuoteStr)
        If Index < 1 Or Index > m_Quotes.Count Then
            ' no? check if its one of the quotes
            Index = GetIndexOf(QuoteStr)
            If Index = 0 Then Exit Function
        End If
    Else
        ' not numeric- check if its one of the quotes
        Index = GetIndexOf(QuoteStr)
        If Index = 0 Then Exit Function
    End If
    
    ' store
    Remove = m_Quotes.Item(Index)

    ' remove
    m_Quotes.Remove Index
    
    ' is not loading from file?
    If m_blnIsLoaded Then
    
        ' save
        SaveQuotes
        
    End If
    
    Exit Function
    
ERROR_HANDLER:
    
    MsgBox "Error (#" & Err.Number & "): " & Err.description & " in Quotes.Remove()."

End Function

' gets the quote by its index
Public Function Item(ByVal Index As Long) As String

    On Error Resume Next
    
    Item = m_Quotes.Item(Index)
    
End Function

' gets the items in the collection
Public Property Get Items() As Collection
    
    Dim I As Long
    
    ' clone the collection - modification to this collection ignored
    Set Items = New Collection
    For I = 1 To m_Quotes.Count
        Items.Add m_Quotes(I)
    Next
    
End Property


' gets the number of quotes
Public Property Get Count() As Long

    On Error Resume Next

    Count = m_Quotes.Count

End Property

' gets the index of the quote by string
Public Function GetIndexOf(ByVal QuoteStr As String) As Long

    On Error Resume Next
    
    Dim I As Long
    
    For I = 1 To m_Quotes.Count
    
        If StrComp(m_Quotes.Item(I), QuoteStr, vbTextCompare) = 0 Then
            
            GetIndexOf = I
            
            Exit Function
            
        End If
    
    Next
    
    GetIndexOf = 0
    
End Function

' returns a random quote
Public Function GetRandomQuote() As String

    On Error GoTo ERROR_HANDLER

    Dim Rand      As Integer
    
	If m_Quotes.Count = 0 Then
        GetRandomQuote = vbNullString
        Exit Function
    End If
    
    ' randomly select quote
    Randomize
    Rand = (Rnd * m_Quotes.Count) + 1
    
    ' get quote in collection
    GetRandomQuote = m_Quotes.Item(Rand)
    
    ' security check
    If Left$(GetRandomQuote, 1) = "/" Then GetRandomQuote = " " & GetRandomQuote
    
    Exit Function

ERROR_HANDLER:

    MsgBox "Error (#" & Err.Number & "): " & Err.description & " in Quotes.GetRandomQuote()."
    
End Function

' this function will load quotes into the collection
Private Sub LoadQuotes()

    On Error GoTo ERROR_HANDLER
    
    Dim f As Integer
    Dim s As String
    Dim strPath As String
    
    m_blnIsLoaded = False
    
    strPath = GetFilePath("quotes.txt")

    Set m_Quotes = New Collection
    
    If (LenB(Dir$(strPath)) > 0) Then
        f = FreeFile
        
        Open (strPath) For Input As #f
            If (LOF(f) > 1) Then
                Do
                    Line Input #f, s
                    
                    Add s
                Loop Until EOF(f)
            End If
        Close #f
    End If
    
    m_blnIsLoaded = True
    
    Exit Sub
    
ERROR_HANDLER:

    MsgBox "Error (#" & Err.Number & "): " & Err.description & " in Quotes.LoadQuotes()."

End Sub

' this function will save quotes into quotes.txt when changes are made
Private Sub SaveQuotes()

    On Error GoTo ERROR_HANDLER

    Dim f As Integer
    Dim I As Long
    Dim strPath As String
    
    strPath = GetFilePath("quotes.txt")
    
    If (LenB(Dir$(strPath)) > 0) Then
        f = FreeFile
        
        Open (strPath) For Output As #f
            ' print each quote
            For I = 1 To m_Quotes.Count
                Print #f, m_Quotes.Item(I)
            Next
        Close #f
    End If
    
    Exit Sub
    
ERROR_HANDLER:

    MsgBox "Error (#" & Err.Number & "): " & Err.description & " in Quotes.SaveQuotes()."

End Sub

' this function will save a single quote into quotes.txt on add (instead of saving the whole thing)
Private Sub AppendQuote(ByVal QuoteStr As String)

    On Error GoTo ERROR_HANDLER

    Dim f As Integer
    Dim I As Long
    Dim strPath As String
    
    strPath = GetFilePath("quotes.txt")
    
    f = FreeFile
    
    If (LenB(Dir$(strPath)) > 0) Then
        Open (strPath) For Append As #f
            ' print new quote
            Print #f, QuoteStr
        Close #f
    Else
        Open (strPath) For Output As #f
            ' print new quote
            Print #f, QuoteStr
        Close #f
    End If
    
    Exit Sub
    
ERROR_HANDLER:

    MsgBox "Error (#" & Err.Number & "): " & Err.description & " in Quotes.AppendQuote()."

End Sub

