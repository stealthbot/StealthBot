'UPGRADE_WARNING: The entire project must be compiled once before a form with an ActiveX Control Array can be displayed

Imports System.ComponentModel

<ProvideProperty("Index",GetType(AxMSWinsockLib.AxWinsock))> Public Class AxWinsockArray
	Inherits Microsoft.VisualBasic.Compatibility.VB6.BaseOcxArray
	Implements IExtenderProvider

	Public Sub New()
		MyBase.New()
	End Sub

	Public Sub New(ByVal Container As IContainer)
		MyBase.New(Container)
	End Sub

	Public Shadows Event [Error] (ByVal sender As System.Object, ByVal e As AxMSWinsockLib.DMSWinsockControlEvents_ErrorEvent)
	Public Shadows Event [DataArrival] (ByVal sender As System.Object, ByVal e As AxMSWinsockLib.DMSWinsockControlEvents_DataArrivalEvent)
	Public Shadows Event [ConnectEvent] (ByVal sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event [ConnectionRequest] (ByVal sender As System.Object, ByVal e As AxMSWinsockLib.DMSWinsockControlEvents_ConnectionRequestEvent)
	Public Shadows Event [CloseEvent] (ByVal sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event [SendProgress] (ByVal sender As System.Object, ByVal e As AxMSWinsockLib.DMSWinsockControlEvents_SendProgressEvent)
	Public Shadows Event [SendComplete] (ByVal sender As System.Object, ByVal e As System.EventArgs)

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Function CanExtend(ByVal target As Object) As Boolean Implements IExtenderProvider.CanExtend
		If TypeOf target Is AxMSWinsockLib.AxWinsock Then
			Return BaseCanExtend(target)
		End If
	End Function

	Public Function GetIndex(ByVal o As AxMSWinsockLib.AxWinsock) As Short
		Return BaseGetIndex(o)
	End Function

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Sub SetIndex(ByVal o As AxMSWinsockLib.AxWinsock, ByVal Index As Short)
		BaseSetIndex(o, Index)
	End Sub

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Function ShouldSerializeIndex(ByVal o As AxMSWinsockLib.AxWinsock) As Boolean
		Return BaseShouldSerializeIndex(o)
	End Function

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Sub ResetIndex(ByVal o As AxMSWinsockLib.AxWinsock)
		BaseResetIndex(o)
	End Sub

	Default Public ReadOnly Property Item(ByVal Index As Short) As AxMSWinsockLib.AxWinsock
		Get
			Item = CType(BaseGetItem(Index), AxMSWinsockLib.AxWinsock)
		End Get
	End Property

	Protected Overrides Function GetControlInstanceType() As System.Type
		Return GetType(AxMSWinsockLib.AxWinsock)
	End Function

	Protected Overrides Sub HookUpControlEvents(ByVal o As Object)
		Dim ctl As AxMSWinsockLib.AxWinsock = CType(o, AxMSWinsockLib.AxWinsock)
		MyBase.HookUpControlEvents(o)
		If Not ErrorEvent Is Nothing Then
			AddHandler ctl.Error, New AxMSWinsockLib.DMSWinsockControlEvents_ErrorEventHandler(AddressOf HandleError)
		End If
		If Not DataArrivalEvent Is Nothing Then
			AddHandler ctl.DataArrival, New AxMSWinsockLib.DMSWinsockControlEvents_DataArrivalEventHandler(AddressOf HandleDataArrival)
		End If
		If Not ConnectEventEvent Is Nothing Then
			AddHandler ctl.ConnectEvent, New System.EventHandler(AddressOf HandleConnectEvent)
		End If
		If Not ConnectionRequestEvent Is Nothing Then
			AddHandler ctl.ConnectionRequest, New AxMSWinsockLib.DMSWinsockControlEvents_ConnectionRequestEventHandler(AddressOf HandleConnectionRequest)
		End If
		If Not CloseEventEvent Is Nothing Then
			AddHandler ctl.CloseEvent, New System.EventHandler(AddressOf HandleCloseEvent)
		End If
		If Not SendProgressEvent Is Nothing Then
			AddHandler ctl.SendProgress, New AxMSWinsockLib.DMSWinsockControlEvents_SendProgressEventHandler(AddressOf HandleSendProgress)
		End If
		If Not SendCompleteEvent Is Nothing Then
			AddHandler ctl.SendComplete, New System.EventHandler(AddressOf HandleSendComplete)
		End If
	End Sub

	Private Sub HandleError (ByVal sender As System.Object, ByVal e As AxMSWinsockLib.DMSWinsockControlEvents_ErrorEvent) 
		RaiseEvent [Error] (sender, e)
	End Sub

	Private Sub HandleDataArrival (ByVal sender As System.Object, ByVal e As AxMSWinsockLib.DMSWinsockControlEvents_DataArrivalEvent) 
		RaiseEvent [DataArrival] (sender, e)
	End Sub

	Private Sub HandleConnectEvent (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [ConnectEvent] (sender, e)
	End Sub

	Private Sub HandleConnectionRequest (ByVal sender As System.Object, ByVal e As AxMSWinsockLib.DMSWinsockControlEvents_ConnectionRequestEvent) 
		RaiseEvent [ConnectionRequest] (sender, e)
	End Sub

	Private Sub HandleCloseEvent (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [CloseEvent] (sender, e)
	End Sub

	Private Sub HandleSendProgress (ByVal sender As System.Object, ByVal e As AxMSWinsockLib.DMSWinsockControlEvents_SendProgressEvent) 
		RaiseEvent [SendProgress] (sender, e)
	End Sub

	Private Sub HandleSendComplete (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [SendComplete] (sender, e)
	End Sub

End Class

