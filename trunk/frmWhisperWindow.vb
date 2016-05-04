Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmWhisperWindow
	Inherits System.Windows.Forms.Form
	
	Private m_sWhisperTo As String
	Private m_imyIndex As Short
	Private m_StartDate As Date
	'UPGRADE_NOTE: Shown was upgraded to Shown_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Shown_Renamed As Boolean
	
	'Public MyOldWndProc As Long
	
	
	Public Property StartDate() As Date
		Get
			m_StartDate = m_StartDate
		End Get
		Set(ByVal Value As Date)
			m_StartDate = Value
		End Set
	End Property
	
	
	Public Property sWhisperTo() As String
		Get
			sWhisperTo = m_sWhisperTo
		End Get
		Set(ByVal Value As String)
			If InStr(Value, "*") Then
				Value = Mid(Value, InStr(Value, "*") + 1)
			End If
			
			m_sWhisperTo = Value
			
			Me.Text = "Whisper Window: " & Value
		End Set
	End Property
	
	
	Public Property myIndex() As Short
		Get
			myIndex = m_imyIndex
		End Get
		Set(ByVal Value As Short)
			m_imyIndex = Value
		End Set
	End Property
	
	Private Sub frmWhisperWindow_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Icon = frmChat.Icon
		
		With frmChat.rtbChat
			rtbWhispers.Font = VB6.FontChangeName(rtbWhispers.Font, .Font.Name)
			rtbWhispers.Font = VB6.FontChangeBold(rtbWhispers.Font, .Font.Bold)
			rtbWhispers.Font = VB6.FontChangeSize(rtbWhispers.Font, .Font.SizeInPoints)
			txtSend.Font = VB6.FontChangeName(txtSend.Font, .Font.Name)
			txtSend.Font = VB6.FontChangeBold(txtSend.Font, .Font.Bold)
			txtSend.Font = VB6.FontChangeSize(txtSend.Font, .Font.SizeInPoints)
		End With
		
		frmWhisperWindow_Resize(Me, New System.EventArgs())
		
		'    If Me.MyOldWndProc = 0 Then
		'        Me.MyOldWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf WWNewWndProc)
		'    End If
	End Sub
	
	Private Sub frmWhisperWindow_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call DestroyWW(m_imyIndex)
	End Sub
	
	Public Sub mnuClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuClose.Click
		Call DestroyWW(m_imyIndex)
	End Sub
	
	Public Sub mnuHide_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHide.Click
		Shown_Renamed = False
		Me.Hide()
	End Sub
	
	Public Sub mnuIgnoreAndClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuIgnoreAndClose.Click
		frmChat.AddQ("/ignore " & m_sWhisperTo)
		Call DestroyWW(m_imyIndex)
	End Sub
	
	'UPGRADE_WARNING: Event frmWhisperWindow.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmWhisperWindow_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error Resume Next
		
		Dim SPACER As Integer
		SPACER = VB6.PixelsToTwipsX(rtbWhispers.Left)
		
		With rtbWhispers
			.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - VB6.PixelsToTwipsY(txtSend.Height) - SPACER - 100)
			.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.ClientRectangle.Width) - (SPACER * 2))
			.Font = VB6.FontChangeName(.Font, frmChat.rtbChat.Font.Name)
			.Font = VB6.FontChangeSize(.Font, frmChat.rtbChat.Font.SizeInPoints)
			.BackColor = frmChat.rtbChat.BackColor
		End With
		
		With txtSend
			.Width = rtbWhispers.Width
			.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(rtbWhispers.Top) + VB6.PixelsToTwipsY(rtbWhispers.Height) + 10)
			.Font = VB6.FontChangeName(.Font, frmChat.cboSend.Font.Name)
			.Font = VB6.FontChangeSize(.Font, frmChat.cboSend.Font.SizeInPoints)
			.BackColor = frmChat.cboSend.BackColor
		End With
		
		txtSend.Focus()
	End Sub
	
	Public Sub mnuSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSave.Click
		Dim ToSave() As String
		Dim f, i As Short
		Dim tUsername, tMessage As String
		
		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
		With cdl
			.InitialDirectory = CurDir()
			'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			.Filter = ".htm|HTML Documents"
			.ShowDialog()
			
            If Len(.FileName) > 0 Then
                ToSave = Split(rtbWhispers.Text, vbCrLf)
                f = FreeFile()

                If InStr(1, .FileName, ".") = 0 Then
                    .FileName = .FileName & ".htm"
                End If

                FileOpen(f, .FileName, OpenMode.Output)
                PrintLine(f, "<html><head>")
                PrintLine(f, "<title>StealthBot Conversation Log: " & GetCurrentUsername() & " and " & m_sWhisperTo & "</title></head>")
                PrintLine(f, "<body bgcolor='#000000'>")

                PrintLine(f, "<p><font color='#FFFFFF'><b>")
                PrintLine(f, "StealthBot Conversation Log, between " & GetCurrentUsername() & " and " & m_sWhisperTo & ".<br />")
                PrintLine(f, "Conversation began: " & VB6.Format(m_StartDate, "HH:MM:SS, m/dd/yyyy"))
                PrintLine(f, "</b></font></p>")

                PrintLine(f, "<p>")

                For i = 0 To UBound(ToSave)
                    If Len(ToSave(i)) > 0 Then
                        If InStr(ToSave(i), ":") > 0 Then
                            tMessage = Mid(ToSave(i), InStr(ToSave(i), ":") + 2)
                            tUsername = Split(ToSave(i), " ")(1)
                            tUsername = VB.Left(tUsername, InStr(tUsername, ":") - 1)
                        Else
                            tMessage = ToSave(i)
                        End If

                        If StrComp(tUsername, GetCurrentUsername, CompareMethod.Text) = 0 Then
                            Print(f, "<font size='-1' color='#" & VBHexToHTMLHex(Hex(RTBColors.TalkBotUsername)) & "'><b>")
                        Else
                            Print(f, "<font size='-1' color='#" & VBHexToHTMLHex(Hex(RTBColors.WhisperUsernames)) & "'><b>")
                        End If
                        PrintLine(f, "» " & tUsername & "</b></font>")

                        Print(f, "<font size='-1' color='#" & VBHexToHTMLHex(Hex(RTBColors.WhisperCarats)) & "'><b>")
                        PrintLine(f, ":</b></font> ")

                        Print(f, "<font size='-1' color='#" & VBHexToHTMLHex(Hex(RTBColors.WhisperText)) & "'>")
                        PrintLine(f, tMessage & "</font><br />")

                    End If
                Next i

                PrintLine(f, "</p>")
                PrintLine(f, "</body></html>")
                FileClose(f)

                AddWhisper(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Lime), "» Conversation saved.")
            End If
		End With
	End Sub
	
	Private Sub rtbWhispers_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles rtbWhispers.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		'Disable Ctrl+L, Ctrl+E, and Ctrl+R
		If (Shift = VB6.ShiftConstants.CtrlMask) And ((KeyCode = System.Windows.Forms.Keys.L) Or (KeyCode = System.Windows.Forms.Keys.E) Or (KeyCode = System.Windows.Forms.Keys.R)) Then
			KeyCode = 0
		End If
	End Sub
	
	Private Sub rtbWhispers_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles rtbWhispers.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If (KeyAscii < 32) Then
			GoTo EventExitSub
		End If
		
		txtSend.Focus()
		
		txtSend.SelectedText = Chr(KeyAscii)
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtSend_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSend.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If (KeyAscii = 13) Then
			frmChat.AddQ("/w " & IIf(Dii, "*", "") & m_sWhisperTo & Space(1) & txtSend.Text)
			KeyAscii = 0
			txtSend.Text = ""
		End If
		
		Dim x() As String
		Dim i As Short
		
		If KeyAscii = 22 Then
			On Error Resume Next
			
			If InStr(1, My.Computer.Clipboard.GetText, Chr(13), CompareMethod.Text) <> 0 Then
				
				x = Split(My.Computer.Clipboard.GetText, Chr(10))
				If UBound(x) > 0 Then
					For i = LBound(x) To UBound(x)
						If i = LBound(x) Then x(i) = txtSend.Text & x(i)
						
						x(i) = Replace(x(i), Chr(13), vbNullString)
						
						If x(i) <> vbNullString Then
							frmChat.AddQ("/w " & m_sWhisperTo & Space(1) & x(i))
						End If
					Next i
					txtSend.Text = vbNullString
					KeyAscii = 0
				End If
			End If
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: ParamArray saElements was changed from ByRef to ByVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93C6A0DC-8C99-429A-8696-35FC4DCEFCCC"'
	Sub AddWhisper(ParamArray ByVal saElements() As Object)
		
		On Error Resume Next
		Dim s As String
		Dim L As Integer
		Dim oldSelStart, i, oldSelLength As Short
		
		oldSelStart = txtSend.SelectionStart
		oldSelStart = oldSelStart + txtSend.SelectionLength
		
		If GetForegroundWindow() = Me.Handle.ToInt32 Then
			rtbWhispers.ReadOnly = True
		End If
		
		If Not BotVars.LockChat Then
			With rtbWhispers
				.SelectionStart = Len(.Text)
				.SelectionLength = 0
				.SelectionColor = System.Drawing.ColorTranslator.FromOle(RTBColors.TimeStamps)
				.SelectedText = s
				.SelectionStart = Len(.Text)
			End With
			
			For i = LBound(saElements) To UBound(saElements) Step 2
				'UPGRADE_WARNING: Couldn't resolve default property of object saElements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If InStr(1, saElements(i), Chr(0), CompareMethod.Binary) > 0 Then KillNull(saElements(i))
				
				If Len(saElements(i + 1)) > 0 Then
					With rtbWhispers
						.SelectionStart = Len(.Text)
						L = .SelectionStart
						.SelectionLength = 0
						'UPGRADE_WARNING: Couldn't resolve default property of object saElements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.SelectionColor = System.Drawing.ColorTranslator.FromOle(saElements(i))
						'UPGRADE_WARNING: Couldn't resolve default property of object saElements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.SelectedText = saElements(i + 1) & VB.Left(vbCrLf, -2 * CInt((i + 1) = UBound(saElements)))
						.SelectionStart = Len(.Text)
					End With
				End If
			Next i
			
			Call ColorModify(rtbWhispers, L)
			
			txtSend.SelectionStart = oldSelStart
			txtSend.SelectionLength = oldSelLength
		End If
		
		'    If rtbWhispers.Locked Then
		'        rtbWhispers.Locked = False
		'    End If
	End Sub
End Class