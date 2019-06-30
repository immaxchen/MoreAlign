Attribute VB_Name = "MethodStack"
Option Explicit

Sub Stack_From_Left()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterX()
    
    For i = LBound(shapes) + 1 To UBound(shapes)
        shapes(i).left = shapes(i - 1).left + shapes(i - 1).width
    Next
    
End Sub

Sub Stack_From_Right()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterX()
    
    For i = UBound(shapes) - 1 To LBound(shapes) Step -1
        shapes(i).left = shapes(i + 1).left - shapes(i).width
    Next
    
End Sub

Sub Stack_From_Top()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterY()
    
    For i = LBound(shapes) + 1 To UBound(shapes)
        shapes(i).top = shapes(i - 1).top + shapes(i - 1).height
    Next
    
End Sub

Sub Stack_From_Bottom()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterY()
    
    For i = UBound(shapes) - 1 To LBound(shapes) Step -1
        shapes(i).top = shapes(i + 1).top - shapes(i).height
    Next
    
End Sub
