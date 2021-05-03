Attribute VB_Name = "storyline"
Option Explicit

Public Sub unfold_storyline()
    
    Dim story As Presentation
    Dim storyline As Variant
    Dim headline As Variant
    Dim story_slide As Slide

    Set story = ActivePresentation
    If InStr(story.Slides(1).Shapes(2).TextFrame.TextRange.Text, vbLf) Then
        storyline = Split(story.Slides(1).Shapes(2).TextFrame.TextRange.Text, vbLf)
    Else
        storyline = Split(story.Slides(1).Shapes(2).TextFrame.TextRange.Text, vbCr)
    End If
    For Each headline In storyline
        Set story_slide = story.Slides.AddSlide(story.Slides.Count + 1, story.SlideMaster.CustomLayouts(2))
        story_slide.Shapes.Title.TextFrame.TextRange.Text = CStr(headline)
    Next
End Sub
