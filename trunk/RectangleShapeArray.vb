Option Strict Off
Option Explicit On
Imports System.Windows.Forms
Imports System.ComponentModel
Imports Microsoft.VisualBasic.Compatibility.VB6
Imports Microsoft.VisualBasic.PowerPacks

<ProvideProperty("Index", GetType(RectangleShape))> Friend Class RectangleShapeArray
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
		If TypeOf Target Is RectangleShape Then
			Return BaseCanExtend(Target)
		End If
	End Function
	
	Public Function GetIndex(ByVal o As RectangleShape) As Short
		Return BaseGetIndex(o)
	End Function
	
	Public Sub SetIndex(ByVal o As RectangleShape, ByVal Index As Short)
		BaseSetIndex(o, Index)
	End Sub
	
	Public Function ShouldSerializeIndex(ByVal o As RectangleShape) As Boolean
		Return BaseShouldSerializeIndex(o)
	End Function
	
	Public Sub ResetIndex(ByVal o As RectangleShape)
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
	
	Public Default ReadOnly Property Item(ByVal Index As Short) As RectangleShape
		Get
			Item = CType(BaseGetItem(Index), RectangleShape)
		End Get
	End Property
	
	Protected Overrides Sub HookUpControlEvents(ByVal o As Object)
		
		Dim ctl As RectangleShape
		ctl = CType(o, RectangleShape)
		
		If Not IsNothing(ClickEvent) Then
			addHandler ctl.Click, ClickEvent
		End If
		
	End Sub
	
	Protected Overrides Function GetControlInstanceType() As System.Type
		Return GetType(RectangleShape)
	End Function
	
End Class