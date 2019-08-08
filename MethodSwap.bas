Attribute VB_Name = "MethodSwap"
Option Explicit

Sub Swap_By_Corner()
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterX()
    
    If Not LBound(shapes) = 1 Or Not UBound(shapes) = 2 Then
        MsgBox "Please Select Exactly 2 Items !"
        Exit Sub
    End If
    
    Dim tmpVal As Single
    
    tmpVal = shapes(1).left
    shapes(1).left = shapes(2).left
    shapes(2).left = tmpVal
    
    tmpVal = shapes(1).top
    shapes(1).top = shapes(2).top
    shapes(2).top = tmpVal
    
End Sub

Sub Swap_By_Center()
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterX()
    
    If Not LBound(shapes) = 1 Or Not UBound(shapes) = 2 Then
        MsgBox "Please Select Exactly 2 Items !"
        Exit Sub
    End If
    
    Dim tmpVal As Single
    
    tmpVal = shapes(1).left + shapes(1).width / 2
    shapes(1).left = shapes(2).left + shapes(2).width / 2 - shapes(1).width / 2
    shapes(2).left = tmpVal - shapes(2).width / 2
    
    tmpVal = shapes(1).top + shapes(1).height / 2
    shapes(1).top = shapes(2).top + shapes(2).height / 2 - shapes(1).height / 2
    shapes(2).top = tmpVal - shapes(2).height / 2
    
End Sub
