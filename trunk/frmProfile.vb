Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmProfile
	Inherits System.Windows.Forms.Form
	
	Private m_IsWriting As Boolean
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		If m_IsWriting Then SetProfile(rtbLocation.Text, rtbProfile.Text, (rtbSex.Text))
		Me.Close()
	End Sub
	
	Private Sub frmProfile_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Icon = frmChat.Icon
	End Sub
	
	Private Sub frmProfile_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		lblUsername.Text = vbNullString
		rtbAge.Text = vbNullString
		rtbSex.Text = vbNullString
		rtbLocation.Text = vbNullString
		rtbProfile.Text = vbNullString
		
		cboSendHadFocus = True
	End Sub
	
	Public Sub PrepareForProfile(ByVal Username As String, ByVal IsWriting As Boolean)
		' store for later
		m_IsWriting = IsWriting
		
		' set caption
		Text = IIf(IsWriting, "Profile Writer - " & GetCurrentUsername, "Profile Viewer - " & Username)
		
		' set Username
		lblUsername.Text = IIf(IsWriting, GetCurrentUsername, Username)
		
		' set up command buttons
		cmdCancel.Visible = IsWriting
		cmdOK.Text = IIf(IsWriting, "&Write", "&Done")
		
		' set locked based on mode
		rtbAge.ReadOnly = True 'Not IsWriting - always fixed
		rtbSex.ReadOnly = Not IsWriting
		rtbLocation.ReadOnly = Not IsWriting
		rtbProfile.ReadOnly = Not IsWriting
		
		' if we are writing, request our own profile
		If IsWriting Then
			ProfileRequest = True
			RequestProfile(GetCurrentUsername)
		End If
	End Sub
	
	Public Sub SetKey(ByVal KeyName As String, ByVal KeyValue As String)
		Dim rtb As System.Windows.Forms.RichTextBox
		
		' make sure shown
		Show()
		
		'frmChat.AddChat vbWhite, "[Profile] " & KeyName & " == " & KeyValue
		
		Select Case KeyName
			Case "Profile\Age"
				rtb = rtbAge
			Case "Profile\Location"
				rtb = rtbLocation
			Case "Profile\Description"
				rtb = rtbProfile
			Case "Profile\Sex"
				rtb = rtbSex
			Case Else
				Exit Sub
		End Select
		
		rtb.Text = vbNullString
		
		rtb.SelectionStart = 0
		rtb.SelectionLength = 0
		rtb.SelectionColor = System.Drawing.Color.White
		rtb.SelectedText = KeyValue
		
		If m_IsWriting = False Then Call ColorModify(rtb, 0)
		
		'UPGRADE_NOTE: Object rtb may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rtb = Nothing
		
		Activate()
	End Sub
	
	'RTB ADDCHAT SUBROUTINE - originally written by Grok[vL] - modified to support
	'                         logging and timestamps, as well as color decoding.
	'Sub AddText(ByRef rtb As RichTextBox, ParamArray saElements() As Variant)
	'    On Error Resume Next
	'    Dim L As Long
	'    Dim I As Integer
	'
	'    For I = LBound(saElements) To UBound(saElements) Step 2
	'        If InStr(1, saElements(I), Chr(0), vbBinaryCompare) > 0 Then _
	''            KillNull saElements(I)
	'
	'        If Len(saElements(I + 1)) > 0 Then
	'            With rtb
	'                .selStart = Len(.Text)
	'                L = .selStart
	'                .selLength = 0
	'                .SelColor = saElements(I)
	'                .SelText = saElements(I + 1) & Left$(vbCrLf, -2 * CLng((I + 1) = UBound(saElements)))
	'                .selStart = Len(.Text)
	'            End With
	'        End If
	'    Next I
	'
	'    Call ColorModify(rtb, L)
	'End Sub
	
	Private Sub rtbAge_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles rtbAge.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		HandleColorKeyCodes(rtbAge, KeyCode, Shift)
	End Sub
	
	Private Sub rtbSex_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles rtbSex.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		HandleColorKeyCodes(rtbSex, KeyCode, Shift)
	End Sub
	
	Private Sub rtbLocation_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles rtbLocation.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		HandleColorKeyCodes(rtbLocation, KeyCode, Shift)
	End Sub
	
	Private Sub rtbProfile_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles rtbProfile.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		HandleColorKeyCodes(rtbProfile, KeyCode, Shift)
	End Sub
	
	Private Sub HandleColorKeyCodes(ByRef rtb As System.Windows.Forms.RichTextBox, ByRef KeyCode As Short, ByRef Shift As Short)
		Const S_CTRL As Short = 2
		
		If (rtb.ReadOnly) Then
			
			Select Case (KeyCode)
				Case KEY_ENTER
					cmdOK_Click(cmdOK, New System.EventArgs())
					
				Case System.Windows.Forms.Keys.Escape
					cmdCancel_Click(cmdCancel, New System.EventArgs())
					
				Case System.Windows.Forms.Keys.A, System.Windows.Forms.Keys.C, System.Windows.Forms.Keys.X, System.Windows.Forms.Keys.V, KEY_ENTER, System.Windows.Forms.Keys.Up, System.Windows.Forms.Keys.Down, System.Windows.Forms.Keys.Left, System.Windows.Forms.Keys.Right
					' don't disable these
					
				Case Else
					' disable CTRL+L, CTRL+E, CTRL+R, CTRL+I and lots of funny ones
					If (Shift = S_CTRL) Then KeyCode = 0
			End Select
			
			Exit Sub
		End If
		
		If (System.Drawing.ColorTranslator.ToOle(rtb.SelectionColor) <> System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)) Then rtb.SelectionColor = System.Drawing.Color.White
		
		Select Case (KeyCode)
			Case KEY_B
				If (Shift = S_CTRL) Then
					rtb.SelectedText = "ÿcb"
				End If
				
			Case KEY_U
				If (Shift = S_CTRL) Then
					rtb.SelectedText = "ÿcu"
				End If
				
			Case KEY_I
				If (Shift = S_CTRL) Then
					rtb.SelectedText = "ÿci"
				End If
				
			Case KEY_ENTER
				If (Shift = S_CTRL) Then
					cmdOK_Click(cmdOK, New System.EventArgs())
				End If
				
			Case System.Windows.Forms.Keys.Escape
				cmdCancel_Click(cmdCancel, New System.EventArgs())
				
			Case System.Windows.Forms.Keys.A, System.Windows.Forms.Keys.C, System.Windows.Forms.Keys.X, System.Windows.Forms.Keys.V, System.Windows.Forms.Keys.Up, System.Windows.Forms.Keys.Down, System.Windows.Forms.Keys.Left, System.Windows.Forms.Keys.Right
				' don't disable these
				
			Case Else
				' disable CTRL+L, CTRL+E, CTRL+R, CTRL+I and lots of funny ones
				If (Shift = S_CTRL) Then KeyCode = 0
				
		End Select
	End Sub
	
	Private Sub rtbAge_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles rtbAge.KeyPress
		Dim ascii As Short = Asc(eventArgs.KeyChar)
		If (ascii = 13) Then
			ascii = 0
			rtbSex.Focus()
		End If
		eventArgs.KeyChar = Chr(ascii)
		If ascii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub rtbSex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles rtbSex.KeyPress
		Dim ascii As Short = Asc(eventArgs.KeyChar)
		If (ascii = 13) Then
			ascii = 0
			rtbLocation.Focus()
		End If
		eventArgs.KeyChar = Chr(ascii)
		If ascii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub rtbLocation_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles rtbLocation.KeyPress
		Dim ascii As Short = Asc(eventArgs.KeyChar)
		If (ascii = 13) Then
			ascii = 0
			rtbProfile.Focus()
		End If
		eventArgs.KeyChar = Chr(ascii)
		If ascii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class