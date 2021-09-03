Attribute VB_Name = "TSupport"
'this module is for sharing helper functions between your test cases

Option Explicit

Public Sub close_testfiles()

    Dim test_presentation As Presentation
    
    For Each test_presentation In Application.Presentations
        If Not test_presentation.Name = "storyline.pptm" Then
            test_presentation.Saved = msoTrue
            test_presentation.Close
        End If
    Next
End Sub
