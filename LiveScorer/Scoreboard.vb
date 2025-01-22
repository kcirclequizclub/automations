Dim counter As TextRange
Function GetTotalSlides() As Integer
    ' Returns the total number of slides in the active presentation
    GetTotalSlides = ActivePresentation.Slides.Count
End Function
Sub counterReset()
Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
        Set counter = ActivePresentation.Slides(i).Shapes("C1").TextFrame.TextRange
        counter = 0
        Set counter = ActivePresentation.Slides(i).Shapes("C2").TextFrame.TextRange
        counter = 0
        Set counter = ActivePresentation.Slides(i).Shapes("C3").TextFrame.TextRange
        counter = 0
        Set counter = ActivePresentation.Slides(i).Shapes("C4").TextFrame.TextRange
        counter = 0
        Set counter = ActivePresentation.Slides(i).Shapes("C5").TextFrame.TextRange
        counter = 0
        Set counter = ActivePresentation.Slides(i).Shapes("C6").TextFrame.TextRange
        counter = 0
        Set counter = ActivePresentation.Slides(i).Shapes("C7").TextFrame.TextRange
        counter = 0
        Set counter = ActivePresentation.Slides(i).Shapes("C8").TextFrame.TextRange
        counter = 0
        Set counter = ActivePresentation.Slides(i).Shapes("C9").TextFrame.TextRange
        counter = 0
        Set counter = ActivePresentation.Slides(i).Shapes("C10").TextFrame.TextRange
        counter = 0
    Next i
End Sub
Sub counter1NEG()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C1").TextFrame.TextRange
        counter = Int(counter) - 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter1PART()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C1").TextFrame.TextRange
        counter = Int(counter) + 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter1CORRECT()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C1").TextFrame.TextRange
        counter = Int(counter) + 10
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub

Sub counter2NEG()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C2").TextFrame.TextRange
        counter = Int(counter) - 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter2PART()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C2").TextFrame.TextRange
        counter = Int(counter) + 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter2CORRECT()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C2").TextFrame.TextRange
        counter = Int(counter) + 10
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub

Sub counter3NEG()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C3").TextFrame.TextRange
        counter = Int(counter) - 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter3PART()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C3").TextFrame.TextRange
        counter = Int(counter) + 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter3CORRECT()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C3").TextFrame.TextRange
        counter = Int(counter) + 10
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub

Sub counter4NEG()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C4").TextFrame.TextRange
        counter = Int(counter) - 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter4PART()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C4").TextFrame.TextRange
        counter = Int(counter) + 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter4CORRECT()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C4").TextFrame.TextRange
        counter = Int(counter) + 10
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter5NEG()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C5").TextFrame.TextRange
        counter = Int(counter) - 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter5PART()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C5").TextFrame.TextRange
        counter = Int(counter) + 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter5CORRECT()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C5").TextFrame.TextRange
        counter = Int(counter) + 10
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub

Sub counter6NEG()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C6").TextFrame.TextRange
        counter = Int(counter) - 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter6PART()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C6").TextFrame.TextRange
        counter = Int(counter) + 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter6CORRECT()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C6").TextFrame.TextRange
        counter = Int(counter) + 10
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub

Sub counter7NEG()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C7").TextFrame.TextRange
        counter = Int(counter) - 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter7PART()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C7").TextFrame.TextRange
        counter = Int(counter) + 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter7CORRECT()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C7").TextFrame.TextRange
        counter = Int(counter) + 10
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub

Sub counter8NEG()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C8").TextFrame.TextRange
        counter = Int(counter) - 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter8PART()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C8").TextFrame.TextRange
        counter = Int(counter) + 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter8CORRECT()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C8").TextFrame.TextRange
        counter = Int(counter) + 10
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub

Sub counter9NEG()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C9").TextFrame.TextRange
        counter = Int(counter) - 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter9PART()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C9").TextFrame.TextRange
        counter = Int(counter) + 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter9CORRECT()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C9").TextFrame.TextRange
        counter = Int(counter) + 10
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub

Sub counter10NEG()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C10").TextFrame.TextRange
        counter = Int(counter) - 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter10PART()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C10").TextFrame.TextRange
        counter = Int(counter) + 5
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
Sub counter10CORRECT()
    Dim i As Integer
    Dim P As Integer
    
    P = GetTotalSlides() ' Get the total number of slides
    For i = 1 To P 'Slide Range
    Set counter = ActivePresentation.Slides(i).Shapes("C10").TextFrame.TextRange
        counter = Int(counter) + 10
        Debug.Print "Function1: Processing Slide " & i
    Next i
End Sub
