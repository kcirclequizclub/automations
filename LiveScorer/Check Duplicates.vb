Sub CheckDuplicateScoreboards()
    Dim slide As slide
    Dim shp As shape
    Dim scoreboardCount As Integer
    Dim duplicateSlides As String
    Dim slideNumber As String
    
    ' Initialize the variable to store slides with duplicates
    duplicateSlides = ""

    ' Loop through all slides in the presentation
    For Each slide In ActivePresentation.Slides
        scoreboardCount = 0

        ' Check each shape on the slide
        For Each shp In slide.Shapes
            If shp.Type = msoGroup Then
                If shp.Name = "Scoreboard" Then
                    scoreboardCount = scoreboardCount + 1
                End If
            End If
        Next shp

        ' If more than one scoreboard group is found, record the slide number
        If scoreboardCount > 1 Then
            slideNumber = slide.SlideIndex
            If duplicateSlides = "" Then
                duplicateSlides = "Slide " & slideNumber
            Else
                duplicateSlides = duplicateSlides & ", Slide " & slideNumber
            End If
        End If
    Next slide

    ' Display message based on the result
    If duplicateSlides = "" Then
        MsgBox "All good!", vbInformation, "Check Complete"
    Else
        MsgBox duplicateSlides & " has/have duplicate Scoreboards. Please Resolve.", vbExclamation, "Duplicates Found"
    End If
End Sub
