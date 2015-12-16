Attribute VB_Name = "ResizeObjects"
Option Explicit

Sub ObjectsToSameSize(ByVal control As IRibbonControl)
    Call ResizeObjectsToSameSize
End Sub

Sub ObjectsToSameWidth(ByVal control As IRibbonControl)
    Call ResizeObjectsToSameSize(True)
End Sub

Sub ObjectsToSameHeight(ByVal control As IRibbonControl)
    Call ResizeObjectsToSameSize(False, True)
End Sub

'Resize all selected objects to same size, width or height as the firstly selected object
Sub ResizeObjectsToSameSize(Optional ByVal skipHeight As Boolean = False, Optional ByVal skipWidth As Boolean = False)
    Dim shp As Shape
    Dim objHeight As Single, objWidth As Single, newHeight As Single, newWidth As Single
        
    With ActiveWindow.Selection
        If .Type = ppSelectionShapes And .ShapeRange.Count > 0 Then
            'Set the height and width from the first object
            objHeight = .ShapeRange.Item(1).height
            objWidth = .ShapeRange.Item(1).width
            
            'Loop in case of multiples shapes are selected
            For Each shp In .ShapeRange
                If Not skipHeight Then newHeight = objHeight Else newHeight = shp.height
                If Not skipWidth Then newWidth = objWidth Else newWidth = shp.width
                Call ChangeShapeSize(shp, newHeight, newWidth)
            Next shp
        Else
            MsgBox "Select one or more shapes!", vbExclamation, "No Shape Selected"
        End If
    End With
End Sub

'Change the height of all selected objects to match their widths, efectively making them square
Sub MakeObjectsSquare(ByVal control As IRibbonControl)
    Dim shp As Shape
        
    With ActiveWindow.Selection
        If .Type = ppSelectionShapes And .ShapeRange.Count > 0 Then
            'Loop in case multiples shapes selected
            For Each shp In .ShapeRange
                Call ChangeShapeSize(shp, shp.width, shp.width)
            Next shp
        Else
            MsgBox "Select one or more shapes!", vbExclamation, "No Shape Selected"
        End If
    End With
End Sub

'Disable "LockAspectRation" during the resize
Sub ChangeShapeSize(shp As Shape, ByVal newHeight As Single, ByVal newWidth As Single)
    Dim orgAspectRatio As MsoTriState
    
    orgAspectRatio = shp.LockAspectRatio
    shp.LockAspectRatio = msoFalse
    shp.height = newHeight
    shp.width = newWidth
    shp.LockAspectRatio = orgAspectRatio
End Sub
