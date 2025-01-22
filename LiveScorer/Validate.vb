Sub AddScoreboardToSlides()
    Dim slide As slide
    Dim sourceSlide As slide
    Dim ScoreboardShape As shape
    Dim shp As shape
    Dim groupExists As Boolean

    ' Ensure Slide 2 exists and has the "Scoreboard" group
    On Error Resume Next
    Set sourceSlide = ActivePresentation.Slides(2)
    Set ScoreboardShape = sourceSlide.Shapes("Scoreboard")
    On Error GoTo 0

    If ScoreboardShape Is Nothing Then
        MsgBox "Slide 2 of the template is required for the slideshow to run.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Loop through all slides in the presentation
    For Each slide In ActivePresentation.Slides
        ' Skip Slide 2 as it already has the "Scoreboard" group
        If slide.SlideIndex = 2 Then GoTo NextSlide

        groupExists = False

        ' Check if the group named "Scoreboard" exists on the slide
        For Each shp In slide.Shapes
            If shp.Name = "Scoreboard" Then
                groupExists = True
                Exit For
            End If
        Next shp

        ' If "Scoreboard" does not exist, copy it from Slide 2 and set visibility to false
        If Not groupExists Then
            ScoreboardShape.Copy
            slide.Shapes.Paste.Name = "Scoreboard"
            slide.Shapes("Scoreboard").Visible = msoFalse
        End If

NextSlide:
    Next slide

    MsgBox "Scoreboard validated.", vbInformation, "Process Complete"
End Sub
