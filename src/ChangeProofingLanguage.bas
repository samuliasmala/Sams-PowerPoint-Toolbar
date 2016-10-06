Attribute VB_Name = "ChangeProofingLanguage"
Option Explicit

Sub ChangeProofingLanguageToEnglishUS(ByVal control As IRibbonControl)
    Call ChangeProofingLanguage(msoLanguageIDEnglishUS)
End Sub

Sub ChangeProofingLanguageToFinnish(ByVal control As IRibbonControl)
    Call ChangeProofingLanguage(msoLanguageIDFinnish)
End Sub

'Change the proofing language of all shapes and tables and set the default language
Sub ChangeProofingLanguage(languageID As MsoLanguageID)
    Dim j As Integer, k As Integer
  
    'Set document language for this and future documents
    ActivePresentation.DefaultLanguageID = languageID
    
    'Change language for active slide
    For k = 1 To ActivePresentation.Slides(ActiveWindow.View.Slide.SlideNumber).Shapes.Count
       Call ChangeAllSubShapes(ActivePresentation.Slides(ActiveWindow.View.Slide.SlideNumber).Shapes(k), languageID)
    Next k
    
    'Change language for all slides
    'For j = 1 To ActivePresentation.Slides.Count
    '    For k = 1 To ActivePresentation.Slides(j).Shapes.Count
    '        Call ChangeAllSubShapes(ActivePresentation.Slides(j).Shapes(k), languageID)
    '    Next k
    'Next j
End Sub

'The language of the shape and all sub-shapes if its a group
Sub ChangeAllSubShapes(targetShape As Shape, languageID As MsoLanguageID)
    Dim i As Integer, r As Integer, c As Integer
    Dim origHeight As Single, origTop As Single

    'Set languageID for text boxes while maintaining the shape size even with autofit
    If targetShape.HasTextFrame Then
        origHeight = targetShape.height
        origTop = targetShape.Top
        
        targetShape.TextFrame.TextRange.languageID = languageID
        
        targetShape.height = origHeight
        targetShape.Top = origTop
    End If
    
    'Set languageID for embedded tables
    If targetShape.HasTable Then
        For r = 1 To targetShape.Table.Rows.Count
            For c = 1 To targetShape.Table.Columns.Count
                targetShape.Table.Cell(r, c).Shape.TextFrame.TextRange.languageID = languageID
            Next
        Next
    End If

    'Set languageID recursively for groups and smartart
    Select Case targetShape.Type
        Case msoGroup, msoSmartArt
            For i = 1 To targetShape.GroupItems.Count
                Call ChangeAllSubShapes(targetShape.GroupItems.Item(i), languageID)
            Next i
    End Select
End Sub

