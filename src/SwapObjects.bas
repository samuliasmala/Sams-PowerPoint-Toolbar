Attribute VB_Name = "SwapObjects"
Option Explicit

'If exactly two objects are selected, then their locations are swapped
Sub SwapObjects(ByVal control As IRibbonControl)
    Dim shp1 As Shape, shp2 As Shape
    Dim topTmp As Single, leftTmp As Single

    With ActiveWindow.Selection
        If .Type = ppSelectionShapes And .ShapeRange.Count = 2 Then
            Set shp1 = .ShapeRange.Item(1)
            Set shp2 = .ShapeRange.Item(2)
            
            topTmp = shp1.Top
            leftTmp = shp1.Left
            
            shp1.Top = shp2.Top
            shp1.Left = shp2.Left
            
            shp2.Top = topTmp
            shp2.Left = leftTmp
        Else
            MsgBox "Select two shapes to swap their places!", vbExclamation, "No Two Shapes Selected"
        End If
    End With
End Sub
