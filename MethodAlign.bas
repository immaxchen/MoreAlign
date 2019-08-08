Attribute VB_Name = "MethodAlign"
Option Explicit

' ================================================== Align CenterX ==================================================

Sub Align_CenterX_To_Left()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterX()
    
    For i = LBound(shapes) + 1 To UBound(shapes)
        shapes(i).left = shapes(1).left + shapes(1).width / 2 - shapes(i).width / 2
    Next
    
End Sub

Sub Align_CenterX_To_Right()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterX()
    
    For i = LBound(shapes) To UBound(shapes) - 1
        shapes(i).left = shapes(UBound(shapes)).left + shapes(UBound(shapes)).width / 2 - shapes(i).width / 2
    Next
    
End Sub

Sub Align_CenterX_To_Top()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterY()
    
    For i = LBound(shapes) + 1 To UBound(shapes)
        shapes(i).left = shapes(1).left + shapes(1).width / 2 - shapes(i).width / 2
    Next
    
End Sub

Sub Align_CenterX_To_Bottom()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterY()
    
    For i = LBound(shapes) To UBound(shapes) - 1
        shapes(i).left = shapes(UBound(shapes)).left + shapes(UBound(shapes)).width / 2 - shapes(i).width / 2
    Next
    
End Sub

' ================================================== Align CenterY ==================================================

Sub Align_CenterY_To_Left()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterX()
    
    For i = LBound(shapes) + 1 To UBound(shapes)
        shapes(i).top = shapes(1).top + shapes(1).height / 2 - shapes(i).height / 2
    Next
    
End Sub

Sub Align_CenterY_To_Right()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterX()
    
    For i = LBound(shapes) To UBound(shapes) - 1
        shapes(i).top = shapes(UBound(shapes)).top + shapes(UBound(shapes)).height / 2 - shapes(i).height / 2
    Next
    
End Sub

Sub Align_CenterY_To_Top()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterY()
    
    For i = LBound(shapes) + 1 To UBound(shapes)
        shapes(i).top = shapes(1).top + shapes(1).height / 2 - shapes(i).height / 2
    Next
    
End Sub

Sub Align_CenterY_To_Bottom()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterY()
    
    For i = LBound(shapes) To UBound(shapes) - 1
        shapes(i).top = shapes(UBound(shapes)).top + shapes(UBound(shapes)).height / 2 - shapes(i).height / 2
    Next
    
End Sub

' ================================================== Align Edge To Center ==================================================

Sub Align_Edge_To_Left_CenterX()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterX()
    
    For i = LBound(shapes) + 1 To UBound(shapes)
        shapes(i).left = shapes(1).left + shapes(1).width / 2
    Next
    
End Sub

Sub Align_Edge_To_Right_CenterX()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterX()
    
    For i = LBound(shapes) To UBound(shapes) - 1
        shapes(i).left = shapes(UBound(shapes)).left + shapes(UBound(shapes)).width / 2 - shapes(i).width
    Next
    
End Sub

Sub Align_Edge_To_Top_CenterY()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterY()
    
    For i = LBound(shapes) + 1 To UBound(shapes)
        shapes(i).top = shapes(1).top + shapes(1).height / 2
    Next
    
End Sub

Sub Align_Edge_To_Bottom_CenterY()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterY()
    
    For i = LBound(shapes) To UBound(shapes) - 1
        shapes(i).top = shapes(UBound(shapes)).top + shapes(UBound(shapes)).height / 2 - shapes(i).height
    Next
    
End Sub
