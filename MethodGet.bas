Attribute VB_Name = "MethodGet"
Option Explicit

Sub Get_Distance_X()
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterX()
    
    If Not LBound(shapes) = 1 Or Not UBound(shapes) = 2 Then
        MsgBox "Please Select Exactly 2 Items !"
        Exit Sub
    End If
    
    Dim tmpVal As Single
    tmpVal = shapes(2).left - (shapes(1).left + shapes(1).width)
    
    MsgBox tmpVal
    
End Sub

Sub Get_Distance_Y()
    
    Dim shapes() As Variant
    shapes = GetShapesOrderByCenterX()
    
    If Not LBound(shapes) = 1 Or Not UBound(shapes) = 2 Then
        MsgBox "Please Select Exactly 2 Items !"
        Exit Sub
    End If
    
    Dim tmpVal As Single
    tmpVal = shapes(2).top - (shapes(1).top + shapes(1).height)
    
    MsgBox tmpVal
    
End Sub
