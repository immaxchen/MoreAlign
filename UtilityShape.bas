Attribute VB_Name = "UtilityShape"
Option Explicit

' ==================================================  Get shape objects collection  ==================================================

Private Function GetShapes() As Variant
    
    Dim i As Long
    
    Dim n As Long
    n = ActiveWindow.Selection.ShapeRange.Count
    
    Dim output() As Variant
    ReDim output(1 To n)
    
    For i = 1 To n
        Set output(i) = ActiveWindow.Selection.ShapeRange(i)
    Next
    
    GetShapes = output
    
End Function

Public Function GetShapesOrderByCenterX() As Variant
    
    Dim shapes() As Variant
    shapes = GetShapes()
    
    Dim centerX() As Variant
    centerX = GetShapeCenterX(shapes)
    
    QuickSortObjectByValue shapes, centerX, LBound(shapes), UBound(shapes)
    GetShapesOrderByCenterX = shapes
    
End Function

Public Function GetShapesOrderByCenterY() As Variant
    
    Dim shapes() As Variant
    shapes = GetShapes()
    
    Dim centerY() As Variant
    centerY = GetShapeCenterY(shapes)
    
    QuickSortObjectByValue shapes, centerY, LBound(shapes), UBound(shapes)
    GetShapesOrderByCenterY = shapes
    
End Function

' ==================================================  Get shape center X, Y values  ==================================================

Public Function GetShapeCenterX(shapes As Variant) As Variant
    
    Dim left() As Single
    left = GetShapeLeft(shapes)
    
    Dim width() As Single
    width = GetShapeWidth(shapes)
    
    GetShapeCenterX = AddArr(left, MulVal(width, 0.5))
    
End Function

Public Function GetShapeCenterY(shapes As Variant) As Variant
    
    Dim top() As Single
    top = GetShapeTop(shapes)
    
    Dim height() As Single
    height = GetShapeHeight(shapes)
    
    GetShapeCenterY = AddArr(top, MulVal(height, 0.5))
    
End Function

' ==================================================  Get shape basic values  ==================================================

Public Function GetShapeLeft(shapes As Variant) As Variant
    
    Dim i As Long
    
    Dim n As Long
    n = UBound(shapes)
    
    Dim output() As Single
    ReDim output(1 To n)
    
    For i = 1 To n
        output(i) = shapes(i).left
    Next
    
    GetShapeLeft = output
    
End Function

Public Function GetShapeTop(shapes As Variant) As Variant
    
    Dim i As Long
    
    Dim n As Long
    n = UBound(shapes)
    
    Dim output() As Single
    ReDim output(1 To n)
    
    For i = 1 To n
        output(i) = shapes(i).top
    Next
    
    GetShapeTop = output
    
End Function

Public Function GetShapeWidth(shapes As Variant) As Variant
    
    Dim i As Long
    
    Dim n As Long
    n = UBound(shapes)
    
    Dim output() As Single
    ReDim output(1 To n)
    
    For i = 1 To n
        output(i) = shapes(i).width
    Next
    
    GetShapeWidth = output
    
End Function

Public Function GetShapeHeight(shapes As Variant) As Variant
    
    Dim i As Long
    
    Dim n As Long
    n = UBound(shapes)
    
    Dim output() As Single
    ReDim output(1 To n)
    
    For i = 1 To n
        output(i) = shapes(i).height
    Next
    
    GetShapeHeight = output
    
End Function
