Feature: unfold storyline
  A good way to define the story line for a new presentation is to just to list
  all the headlines on one single slide. The storyline macro can then "unfold
  the storyline" by distributing each headline on a single slide.


  Rule: every line of text from the first slide should appear as headline in a separate slide

    @happy_path
    Example: storyline with 3 lines
      Given a new presentation with one slide
        And the slide contains this text
          """
            story topic one
            story topic two
            story topic three
          """
       When the storyline is unfolded
       Then three new slides are added to the presentation after the first slide
        And the headline of each slide matches the corresponding line in the storyline
