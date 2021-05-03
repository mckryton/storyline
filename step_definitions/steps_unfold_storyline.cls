VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "steps_unfold_storyline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_a_new_presentation_with_one_slide_6B6FEF40B1BD()
    'Given a new presentation with one slide
    Set TStepVars.story = Application.Presentations.Add(msoTrue)
    Set TStepVars.story_setup = TStepVars.story.Slides.AddSlide(1, TStepVars.story.SlideMaster.CustomLayouts(2))
End Sub

Public Sub Given_the_slide_contains_this_text_STR_127D297E656F(step_expressions As Collection)
    'And the slide contains this text {string}
    TStepVars.storyline = step_expressions(1)
    TStepVars.story.Slides(1).Shapes(2).TextFrame.TextRange.Text = TStepVars.storyline
End Sub

Public Sub When_the_storyline_is_unfolded_418C7D14F6DC()
    'When the storyline is unfolded
    storyline.unfold_storyline
End Sub

Public Sub Then_three_new_slides_are_added_to_the_presentatio_C4A0155FBF51()
    'Then three new slides are added to the presentation after the first slide
    TSpec.expect(TStepVars.story.Slides.Count).to_be 4
    TSpec.expect(TStepVars.story.Slides(1)).to_be TStepVars.story_setup
End Sub

Public Sub Then_the_headline_of_each_slide_matches_the_corres_4253515F62AB()
    'And the headline of each slide matches the corresponding line in the storyline
    Dim headlines As Variant
    
    headlines = Split(TStepVars.storyline, vbLf)
    TSpec.expect(TStepVars.story.Slides(2).Shapes.Title.TextFrame.TextRange.Text).to_be headlines(0)
    TSpec.expect(TStepVars.story.Slides(3).Shapes.Title.TextFrame.TextRange.Text).to_be headlines(1)
    TSpec.expect(TStepVars.story.Slides(4).Shapes.Title.TextFrame.TextRange.Text).to_be headlines(2)
End Sub

