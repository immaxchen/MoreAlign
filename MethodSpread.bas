Attribute VB_Name = "MethodSpread"
Option Explicit

Sub Spread_From_Left()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterX()
    
    Dim space As Single
    space = InputBox("Please enter the spacing to spread", "Determine Spacing", 10)
    
    Dim k As Long
    k = 1
    
    For i = LBound(shapes) + 1 To UBound(shapes)
        shapes(i).left = shapes(i).left + k * space
        k = k + 1
    Next
    
End Sub

Sub Spread_From_Right()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterX()
    
    Dim space As Single
    space = InputBox("Please enter the spacing to spread", "Determine Spacing", 10)
    
    Dim k As Long
    k = 1
    
    For i = UBound(shapes) - 1 To LBound(shapes) Step -1
        shapes(i).left = shapes(i).left - k * space
        k = k + 1
    Next
    
End Sub

Sub Spread_From_Top()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterY()
    
    Dim space As Single
    space = InputBox("Please enter the spacing to spread", "Determine Spacing", 10)
    
    Dim k As Long
    k = 1
    
    For i = LBound(shapes) + 1 To UBound(shapes)
        shapes(i).top = shapes(i).top + k * space
        k = k + 1
    Next
    
End Sub

Sub Spread_From_Bottom()
    
    Dim i As Long
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterY()
    
    Dim space As Single
    space = InputBox("Please enter the spacing to spread", "Determine Spacing", 10)
    
    Dim k As Long
    k = 1
    
    For i = UBound(shapes) - 1 To LBound(shapes) Step -1
        shapes(i).top = shapes(i).top - k * space
        k = k + 1
    Next
    
End Sub
