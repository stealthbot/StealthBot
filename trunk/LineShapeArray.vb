Option Strict Off
Option Explicit On
Imports System.Windows.Forms
Imports System.ComponentModel
Imports Microsoft.VisualBasic.Compatibility.VB6
Imports Microsoft.VisualBasic.PowerPacks

<ProvideProperty("Index", GetType(LineShape))> Friend Class LineShapeArray
	Inherits BaseControlArray
	Implements IExtenderProvider
	
	Public Event [Click] As System.EventHandler
	
	Public Sub New()
		MyBase.New()
	End Sub
	
	Public Sub New(ByVal Container As IContainer)
		MyBase.New(Container)
	End Sub
	
	Public Function CanExtend(ByVal Target As Object) As Boolean Implements IExtenderProvider.CanExtend
		If TypeOf Target Is LineShape Then
			Return BaseCanExtend(Target)
		End If
	End Function
	
	Public Function GetIndex(ByVal o As LineShape) As Short
		Return BaseGetIndex(o)
	End Function
	
	Public Sub SetIndex(ByVal o As LineShape, ByVal Index As Short)
		BaseSetIndex(o, Index)
	End Sub
	
	Public Function ShouldSerializeIndex(ByVal o As LineShape) As Boolean
		Return BaseShouldSerializeIndex(o)
	End Function
	
	Public Sub ResetIndex(ByVal o As LineShape)
		BaseResetIndex(o)
	End Sub
	
	Public Shadows Sub Load(ByVal Index As Short)
		MyBase.Load(Index)
		CType(Item(0).Parent, ShapeContainer).Shapes.Add(Item(Index))
	End Sub
	
	Public Shadows Sub Unload(ByVal Index As Short)
		CType(Item(0).Parent, ShapeContainer).Shapes.Remove(Item(Index))
		MyBase.Unload(Index)
	End Sub
	
	Public Default ReadOnly Property Item(ByVal Index As Short) As LineShape
		Get
			Item = CType(BaseGetItem(Index), LineShape)
		End Get
	End Property
	
	Protected Overrides Sub HookUpControlEvents(ByVal o As Object)
		
		Dim ctl As LineShape
		ctl = CType(o, LineShape)
		
		If Not IsNothing(ClickEvent) Then
			addHandler ctl.Click, ClickEvent
		End If
		
	End Sub
	
	Protected Overrides Function GetControlInstanceType() As System.Type
		Return GetType(LineShape)
	End Function
	
End Class