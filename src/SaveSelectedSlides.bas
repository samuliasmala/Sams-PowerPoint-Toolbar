Attribute VB_Name = "SaveSelectedSlides"
Option Explicit

'Save selected files in a new file. Prompt the filename of the new file
Sub SaveSelectedSlides()
    Dim i As Integer
    Dim startTime As Single
    Dim dlgSaveAs As FileDialog
    Dim fdfs As FileDialogFilters
    Dim fdf As FileDialogFilter
    Dim ap As Presentation, vv As Presentation
    
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)
    
    'Check that slides are selected
    If ActivePresentation.Windows.Item(1).Selection.Type <> ppSelectionSlides Then
        MsgBox "No slides selected!"
        Exit Sub
    End If

    With dlgSaveAs
        Set fdfs = .Filters
        .InitialFileName = ActivePresentation.Path & "\" & ActivePresentation.Name
        
        For i = 1 To fdfs.Count
            Set fdf = fdfs.Item(i)
            If ActivePresentation.Name Like fdf.Extensions Then
                .FilterIndex = i
                Exit For
            End If
        Next
    
        If .Show = -1 Then
            'Check that new file name is different from the original
            If .InitialFileName <> .SelectedItems.Item(1) Then
                .Execute
                
                'Check that the active presentation is correct one and delete slides
                Set ap = ActivePresentation
                If ap.FullName = .SelectedItems.Item(1) Then
                    Call DeleteSlidesNotSelected(ap)
                    ap.Save
                End If
                
                'Open source file and activate the original window
                Set vv = Presentations.Open(.InitialFileName)
                ap.Windows.Item(1).Activate
                
                'Timer to give some time to open the original presentation
                startTime = Timer
                While Timer < startTime + 3 And Not ActivePresentation Is vv
                    DoEvents
                Wend
                
                ap.Windows.Item(1).Activate
            End If
        End If
    End With
End Sub

'Delete slides that are not selected from the presentation
Sub DeleteSlidesNotSelected(p As Presentation)
    Dim i As Integer
    Dim selectedSlides As SlideRange
    Dim deleteSlides() As Variant
    
    'Check that selection contains slides
    If p.Windows.Item(1).Selection.Type <> ppSelectionSlides Then
        MsgBox "No slides selected!"
        Exit Sub
    End If
    
    'Move selected slides to the beginning of the presentation
    Set selectedSlides = p.Windows.Item(1).Selection.SlideRange
    selectedSlides.MoveTo (1)

    'Select slides to be deleted
    If selectedSlides.Count < p.Slides.Count Then
        ReDim deleteSlides(selectedSlides.Count + 1 To p.Slides.Count)
        For i = selectedSlides.Count + 1 To p.Slides.Count
            deleteSlides(i) = i
        Next
        
        'Delete slides
        p.Slides.Range(deleteSlides).Delete
    End If

End Sub

