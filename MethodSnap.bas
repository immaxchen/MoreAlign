Attribute VB_Name = "MethodSnap"
Option Explicit

Sub Snap_To_Left()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterX()
    
    For i = LBound(shapes) + 1 To UBound(shapes)
        shapes(i).left = shapes(1).left + shapes(1).width
    Next
    
End Sub

Sub Snap_To_Right()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterX()
    
    For i = LBound(shapes) To UBound(shapes) - 1
        shapes(i).left = shapes(UBound(shapes)).left - shapes(i).width
    Next
    
End Sub

Sub Snap_To_Top()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterY()
    
    For i = LBound(shapes) + 1 To UBound(shapes)
        shapes(i).top = shapes(1).top + shapes(1).height
    Next
    
End Sub

Sub Snap_To_Bottom()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterY()
    
    For i = LBound(shapes) To UBound(shapes) - 1
        shapes(i).top = shapes(UBound(shapes)).top - shapes(i).height
    Next
    
End Sub
