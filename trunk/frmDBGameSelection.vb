Option Strict Off
Option Explicit On
Friend Class frmDBGameSelection
	Inherits System.Windows.Forms.Form
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		frmDBManager.m_entryname = vbNullString
		
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim Game As String
		
		Select Case (lvGames.FocusedItem.Index)
			Case 1 : Game = PRODUCT_W2BN
			Case 2 : Game = PRODUCT_STAR
			Case 3 : Game = PRODUCT_SEXP
			Case 4 : Game = PRODUCT_D2DV
			Case 5 : Game = PRODUCT_D2XP
			Case 6 : Game = PRODUCT_WAR3
			Case 7 : Game = PRODUCT_W3XP
			Case 8 : Game = PRODUCT_JSTR
			Case 9 : Game = PRODUCT_SSHR
			Case 10 : Game = PRODUCT_DRTL
			Case 11 : Game = PRODUCT_DSHR
			Case 12 : Game = PRODUCT_CHAT
		End Select
		
		frmDBManager.m_entryname = Game
		
		Me.Close()
	End Sub
	
	Private Sub frmDBGameSelection_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'UPGRADE_WARNING: Lower bound of collection lvGames.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
		Call lvGames.Items.Add(GetProductInfo(PRODUCT_W2BN).FullName, 5)
		'UPGRADE_WARNING: Lower bound of collection lvGames.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
		Call lvGames.Items.Add(GetProductInfo(PRODUCT_STAR).FullName, 1)
		'UPGRADE_WARNING: Lower bound of collection lvGames.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
		Call lvGames.Items.Add(GetProductInfo(PRODUCT_SEXP).FullName, 2)
		'UPGRADE_WARNING: Lower bound of collection lvGames.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
		Call lvGames.Items.Add(GetProductInfo(PRODUCT_D2DV).FullName, 3)
		'UPGRADE_WARNING: Lower bound of collection lvGames.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
		Call lvGames.Items.Add(GetProductInfo(PRODUCT_D2XP).FullName, 4)
		'UPGRADE_WARNING: Lower bound of collection lvGames.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
		Call lvGames.Items.Add(GetProductInfo(PRODUCT_WAR3).FullName, 6)
		'UPGRADE_WARNING: Lower bound of collection lvGames.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
		Call lvGames.Items.Add(GetProductInfo(PRODUCT_W3XP).FullName, 11)
		'UPGRADE_WARNING: Lower bound of collection lvGames.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
		Call lvGames.Items.Add(GetProductInfo(PRODUCT_JSTR).FullName, 10)
		'UPGRADE_WARNING: Lower bound of collection lvGames.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
		Call lvGames.Items.Add(GetProductInfo(PRODUCT_SSHR).FullName, 12)
		'UPGRADE_WARNING: Lower bound of collection lvGames.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
		Call lvGames.Items.Add(GetProductInfo(PRODUCT_DRTL).FullName, 8)
		'UPGRADE_WARNING: Lower bound of collection lvGames.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
		Call lvGames.Items.Add(GetProductInfo(PRODUCT_DSHR).FullName, 9)
		'UPGRADE_WARNING: Lower bound of collection lvGames.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
		Call lvGames.Items.Add(GetProductInfo(PRODUCT_CHAT).FullName, 7)
	End Sub
	
	Private Sub lvGames_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lvGames.DoubleClick
		Call cmdOK_Click(cmdOK, New System.EventArgs())
	End Sub
	
	Private Sub lvGames_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles lvGames.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = System.Windows.Forms.Keys.Return Then
			Call cmdOK_Click(cmdOK, New System.EventArgs())
		ElseIf KeyAscii = System.Windows.Forms.Keys.Escape Then 
			Call cmdCancel_Click(cmdCancel, New System.EventArgs())
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class