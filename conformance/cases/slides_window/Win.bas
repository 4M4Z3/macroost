Option Explicit

' Run with run_ppt.ps1 -WithWindow: ActivePresentation needs a presentation with a window.

Function W1() As String
    W1 = ActivePresentation.Name & "|" & ActivePresentation.Slides.Count & "|" & (ActivePresentation Is Presentations(1)) & "|" & TypeName(ActivePresentation)
End Function

Function W2() As String
    Dim s As Slide
    Set s = ActivePresentation.Slides.Add(1, ppLayoutText)
    W2 = s.SlideIndex & "|" & s.Name & "|" & ActivePresentation.Slides.Count & "|" & Application.ActivePresentation.Slides(1).Shapes.Count
End Function

Function W3() As String
    W3 = Application.ActivePresentation.PageSetup.SlideWidth & "|" & Application.Presentations.Count & "|" & TypeName(Application)
End Function
