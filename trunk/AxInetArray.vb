'UPGRADE_WARNING: The entire project must be compiled once before a form with an ActiveX Control Array can be displayed

Imports System.ComponentModel

<ProvideProperty("Index",GetType(AxInetCtlsObjects.AxInet))> Public Class AxInetArray
	Inherits Microsoft.VisualBasic.Compatibility.VB6.BaseOcxArray
	Implements IExtenderProvider

	Public Sub New()
		MyBase.New()
	End Sub

	Public Sub New(ByVal Container As IContainer)
		MyBase.New(Container)
	End Sub

	Public Shadows Event [StateChanged] (ByVal sender As System.Object, ByVal e As AxInetCtlsObjects.DInetEvents_StateChangedEvent)

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Function CanExtend(ByVal target As Object) As Boolean Implements IExtenderProvider.CanExtend
		If TypeOf target Is AxInetCtlsObjects.AxInet Then
			Return BaseCanExtend(target)
		End If
	End Function

	Public Function GetIndex(ByVal o As AxInetCtlsObjects.AxInet) As Short
		Return BaseGetIndex(o)
	End Function

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Sub SetIndex(ByVal o As AxInetCtlsObjects.AxInet, ByVal Index As Short)
		BaseSetIndex(o, Index)
	End Sub

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Function ShouldSerializeIndex(ByVal o As AxInetCtlsObjects.AxInet) As Boolean
		Return BaseShouldSerializeIndex(o)
	End Function

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Sub ResetIndex(ByVal o As AxInetCtlsObjects.AxInet)
		BaseResetIndex(o)
	End Sub

	Default Public ReadOnly Property Item(ByVal Index As Short) As AxInetCtlsObjects.AxInet
		Get
			Item = CType(BaseGetItem(Index), AxInetCtlsObjects.AxInet)
		End Get
	End Property

	Protected Overrides Function GetControlInstanceType() As System.Type
		Return GetType(AxInetCtlsObjects.AxInet)
	End Function

	Protected Overrides Sub HookUpControlEvents(ByVal o As Object)
		Dim ctl As AxInetCtlsObjects.AxInet = CType(o, AxInetCtlsObjects.AxInet)
		MyBase.HookUpControlEvents(o)
		If Not StateChangedEvent Is Nothing Then
			AddHandler ctl.StateChanged, New AxInetCtlsObjects.DInetEvents_StateChangedEventHandler(AddressOf HandleStateChanged)
		End If
	End Sub

	Private Sub HandleStateChanged (ByVal sender As System.Object, ByVal e As AxInetCtlsObjects.DInetEvents_StateChangedEvent) 
		RaiseEvent [StateChanged] (sender, e)
	End Sub

End Class

