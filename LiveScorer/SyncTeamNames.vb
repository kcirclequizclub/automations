Sub SyncTeamNames()
    Dim slide As slide
    Dim scoreboardGroup As shape
    Dim shp As shape
    Dim sourceSlide As slide
    Dim sourceText As String
    Dim groupExists As Boolean
    Dim fieldName As String
    Dim i As Integer

    ' Loop through field names from AA to JJ
    For i = 1 To 10
        ' Generate the field name (AA, BB, CC, ..., JJ)
        fieldName = Chr(64 + i) & Chr(64 + i) ' AA, BB, CC, ..., JJ

        ' Ensure Slide 2 exists and has the "Scoreboard" group with the text field
        On Error Resume Next
        Set sourceSlide = ActivePresentation.Slides(2)
        Set scoreboardGroup = sourceSlide.Shapes("Scoreboard")
        On Error GoTo 0

        If scoreboardGroup Is Nothing Then
            MsgBox "The 'Scoreboard' group was not found on Slide 2. Please ensure it exists.", vbExclamation, "Error"
            Exit Sub
        End If

        ' Find the text field in the "Scoreboard" group on Slide 2
        On Error Resume Next
        Set shp = scoreboardGroup.GroupItems(fieldName)
        On Error GoTo 0

        ' Check if the text field exists in the "Scoreboard" group on Slide 2
        If shp Is Nothing Then
            MsgBox "Text field '" & fieldName & "' not found in the 'Scoreboard' group on Slide 2. Please ensure it exists.", vbExclamation, "Error"
            Exit Sub
        End If

        ' Get the value of the text field on Slide 2
        sourceText = shp.TextFrame.TextRange.Text

        ' Loop through all slides and update the text field values in the "Scoreboard" group
        For Each slide In ActivePresentation.Slides
            ' Check if the "Scoreboard" group exists on the current slide
            On Error Resume Next
            Set scoreboardGroup = slide.Shapes("Scoreboard")
            On Error GoTo 0

            If Not scoreboardGroup Is Nothing Then
                ' Loop through each shape inside the "Scoreboard" group
                For Each shp In scoreboardGroup.GroupItems
                    ' Check if the shape is a text box and is named the current field (AA, BB, CC, ..., JJ)
                    If shp.HasTextFrame Then
                        If shp.Name = fieldName Then
                            ' Set the text value to be the same as the field on Slide 2
                            shp.TextFrame.TextRange.Text = sourceText
                        End If
                    End If
                Next shp
            End If
        Next slide
    Next i
End Sub



