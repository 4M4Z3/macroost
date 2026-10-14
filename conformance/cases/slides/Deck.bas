Option Explicit

' Each case runs in a fresh copy of the generated presentation: 2 blank slides, 960 x 540 points,
' layouts Title Slide / Title and Content / Title and Text / Title Only / Blank.
' Every case returns a string describing what it observed.

Private Function P() As Presentation
    Set P = Presentations(1)
End Function

Private Function S1() As Slide
    Set S1 = Presentations(1).Slides(1)
End Function

Private Function Geo(s As Shape) As String
    Geo = s.Left & "," & s.Top & "," & s.Width & "," & s.Height
End Function

Private Function NamesOf(sl As Slide) As String
    Dim sh As Shape, r As String
    For Each sh In sl.Shapes
        r = r & sh.Name & ";"
    Next
    NamesOf = r
End Function

Function T01() As String
    T01 = Presentations.Count & "|" & P.Slides.Count & "|" & P.Name & "|" & TypeName(P)
End Function

Function T02() As String
    Dim s As Slide
    Set s = P.Slides(2)
    T02 = s.SlideIndex & "|" & s.SlideNumber & "|" & s.Name & "|" & s.SlideID & "|" & P.Slides(1).SlideID
End Function

Function T03() As String
    Dim s As Slide
    Set s = P.Slides.Add(3, ppLayoutBlank)
    T03 = P.Slides.Count & "|" & s.SlideIndex & "|" & s.Name & "|" & s.Layout & "|" & s.Shapes.Count & "|" & s.SlideID
End Function

Function T04() As String
    Dim s As Slide
    Set s = P.Slides.Add(1, ppLayoutText)
    T04 = s.Shapes.Count & "|" & s.Layout & "|" & P.Slides(2).SlideIndex & "|" & P.Slides(2).Name & "|" & NamesOf(s)
End Function

Function T05() As String
    Dim s As Slide
    Set s = P.Slides.Add(1, ppLayoutText)
    T05 = Geo(s.Shapes(1)) & "|" & Geo(s.Shapes(2))
End Function

Function T06() As String
    Dim s As Slide
    Set s = P.Slides.Add(1, ppLayoutText)
    T06 = s.Shapes(1).Type & "|" & s.Shapes(1).HasTextFrame & "|" & s.Shapes(1).PlaceholderFormat.Type & "|" & s.Shapes(2).PlaceholderFormat.Type & "|" & s.Shapes.Placeholders.Count
End Function

Function T07() As String
    Dim s As Slide
    Set s = P.Slides.Add(1, ppLayoutTitle)
    T07 = s.Shapes.Count & "|" & NamesOf(s) & "|" & Geo(s.Shapes(1)) & "|" & Geo(s.Shapes(2)) & "|" & s.Shapes(1).PlaceholderFormat.Type & "|" & s.Shapes(2).PlaceholderFormat.Type
End Function

Function T08() As String
    Dim s As Slide
    Set s = P.Slides.Add(2, ppLayoutTitleOnly)
    T08 = s.Shapes.Count & "|" & NamesOf(s) & "|" & s.Layout & "|" & s.SlideIndex
End Function

Function T09() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 20, 200, 50)
    T09 = t.Name & "|" & t.Type & "|" & Geo(t) & "|" & t.HasTextFrame & "|" & t.Id & "|" & TypeName(t)
End Function

Function T10() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 20, 200, 50)
    t.TextFrame.TextRange.Text = "Hello"
    T10 = t.TextFrame.TextRange.Text & "|" & t.Height & "|" & t.Width
End Function

Function T11() As String
    Dim r As Shape
    Set r = S1.Shapes.AddShape(msoShapeRectangle, 100, 100, 50, 40)
    T11 = r.Name & "|" & r.Type & "|" & r.AutoShapeType & "|" & r.HasTextFrame & "|" & Geo(r) & "|" & r.TextFrame.HasText
End Function

Function T12() As String
    Dim o As Shape
    Set o = S1.Shapes.AddShape(msoShapeOval, 0, 0, 30, 30)
    T12 = o.Name & "|" & o.AutoShapeType & "|" & o.Type
End Function

Function T13() As String
    With S1.Shapes
        .AddTextbox msoTextOrientationHorizontal, 0, 0, 100, 20
        .AddShape msoShapeRectangle, 0, 0, 10, 10
        .AddShape msoShapeOval, 0, 0, 10, 10
        .AddTable 2, 2, 0, 0, 100, 40
        .AddTextbox msoTextOrientationHorizontal, 0, 0, 100, 20
    End With
    T13 = NamesOf(S1) & "|" & S1.Shapes.Count
End Function

Function T14() As String
    S1.Shapes.AddShape msoShapeRectangle, 5, 6, 10, 10
    S1.Shapes.AddShape msoShapeRectangle, 7, 8, 10, 10
    T14 = S1.Shapes(1).Left & "|" & S1.Shapes("Rectangle 2").Top & "|" & S1.Shapes(2).Name & "|" & S1.Shapes.Item(1).Name
End Function

Function T15() As String
    Dim r As Shape
    Set r = S1.Shapes.AddShape(msoShapeRectangle, 5, 6, 10, 10)
    r.Name = "Box"
    T15 = S1.Shapes("Box").Left & "|" & S1.Shapes("box").Name & "|" & r.Name
End Function

Function T16() As String
    S1.Shapes.AddShape msoShapeRectangle, 1, 1, 10, 10
    S1.Shapes.AddShape msoShapeRectangle, 2, 2, 10, 10
    S1.Shapes.AddShape msoShapeRectangle, 3, 3, 10, 10
    S1.Shapes(2).Delete
    T16 = S1.Shapes.Count & "|" & NamesOf(S1) & "|" & S1.Shapes(2).Left
End Function

Function T17() As String
    Dim r As Shape
    Set r = S1.Shapes.AddShape(msoShapeRectangle, 1, 1, 10, 10)
    r.Left = 10.25: r.Top = 20.5: r.Width = 33.3: r.Height = 0.1
    T17 = Geo(r) & "|" & TypeName(r.Left)
End Function

Function T18() As String
    P.Slides.Add 3, ppLayoutBlank
    P.Slides(1).Delete
    T18 = P.Slides.Count & "|" & P.Slides(1).Name & "|" & P.Slides(1).SlideIndex & "|" & P.Slides(2).Name & "|" & P.Slides(2).SlideNumber
End Function

Function T19() As String
    P.Slides.Add 3, ppLayoutBlank
    T19 = P.Slides.Range(Array(1, 3)).Count & "|" & P.Slides.Range.Count & "|" & P.Slides.Range(2).Count & "|" & P.Slides.Range(Array(1, 3))(2).Name
End Function

Function T20() As String
    Dim r As SlideRange
    S1.Shapes.AddShape msoShapeRectangle, 1, 1, 10, 10
    Set r = S1.Duplicate
    T20 = P.Slides.Count & "|" & r.SlideIndex & "|" & r.Name & "|" & r.Shapes.Count & "|" & P.Slides(3).Name & "|" & P.Slides(3).SlideIndex & "|" & TypeName(r)
End Function

Function T21() As String
    Dim r As Shape, d As ShapeRange
    Set r = S1.Shapes.AddShape(msoShapeRectangle, 100, 100, 50, 40)
    Set d = r.Duplicate
    T21 = d.Name & "|" & d.Left & "," & d.Top & "|" & S1.Shapes.Count & "|" & TypeName(d) & "|" & d.Count
End Function

Function T22() As String
    S1.Shapes.AddShape msoShapeRectangle, 1, 1, 10, 10
    S1.Shapes.AddShape msoShapeOval, 1, 1, 10, 10
    S1.Shapes.AddShape msoShapeRectangle, 1, 1, 10, 10
    T22 = S1.Shapes(1).ZOrderPosition & S1.Shapes(2).ZOrderPosition & S1.Shapes(3).ZOrderPosition
    S1.Shapes(3).ZOrder msoSendToBack
    T22 = T22 & "|" & NamesOf(S1) & "|" & S1.Shapes("Rectangle 3").ZOrderPosition
End Function

Function T23() As String
    S1.Shapes.AddShape msoShapeRectangle, 1, 1, 10, 10
    S1.Shapes.AddShape msoShapeOval, 1, 1, 10, 10
    S1.Shapes.AddShape msoShapeRectangle, 1, 1, 10, 10
    S1.Shapes(1).ZOrder msoBringToFront
    T23 = NamesOf(S1) & "|" & S1.Shapes("Rectangle 1").ZOrderPosition
End Function

Function T24() As String
    S1.Shapes.AddShape msoShapeRectangle, 1, 1, 10, 10
    S1.Shapes.AddShape msoShapeOval, 1, 1, 10, 10
    S1.Shapes.AddShape msoShapeRectangle, 1, 1, 10, 10
    S1.Shapes(1).ZOrder msoBringForward
    T24 = NamesOf(S1)
    S1.Shapes(3).ZOrder msoSendBackward
    T24 = T24 & "|" & NamesOf(S1)
End Function

Function T25() As String
    With P.SlideShowSettings
        T25 = .StartingSlide & "|" & .EndingSlide & "|" & .ShowType & "|" & .LoopUntilStopped & "|" & .AdvanceMode & "|" & .RangeType & "|" & .ShowWithNarration & "|" & .ShowWithAnimation
    End With
End Function

Function T26() As String
    With P.PageSetup
        T26 = .SlideWidth & "|" & .SlideHeight & "|" & .FirstSlideNumber & "|" & .SlideOrientation & "|" & TypeName(.SlideWidth)
    End With
End Function

Function T27() As String
    P.PageSetup.FirstSlideNumber = 5
    T27 = P.Slides(1).SlideNumber & "|" & P.Slides(2).SlideNumber & "|" & P.Slides(2).SlideIndex
End Function

Function T28() As String
    Dim tr As TextRange
    Set tr = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50).TextFrame.TextRange
    tr.Text = "One" & vbCr & "Two" & vbCr & "Three"
    T28 = tr.Paragraphs.Count & "|" & tr.Paragraphs(2).Text & "|" & Len(tr.Paragraphs(2).Text) & "|" & Len(tr.Paragraphs(3).Text) & "|" & tr.Paragraphs(2, 2).Count
End Function

Function T29() As String
    Dim tr As TextRange
    Set tr = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50).TextFrame.TextRange
    tr.Text = "Hello World"
    T29 = tr.Runs.Count
    tr.Characters(1, 3).Font.Bold = msoTrue
    T29 = T29 & "|" & tr.Runs.Count & "|" & tr.Runs(1).Text & "|" & tr.Runs(2).Text & "|" & tr.Runs(2).Font.Bold
End Function

Function T30() As String
    Dim tr As TextRange
    Set tr = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50).TextFrame.TextRange
    tr.Text = "abcdef"
    T30 = tr.Characters.Count & "|" & tr.Characters(2, 3).Text & "|" & tr.Characters(5).Text & "|" & tr.Length & "|" & tr.Characters(2, 3).Start
End Function

Function T31() As String
    Dim f As Font
    Dim t As Shape
    Set t = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50)
    t.TextFrame.TextRange.Text = "x"
    Set f = t.TextFrame.TextRange.Font
    T31 = f.Bold & "|" & f.Size & "|" & f.Color.RGB & "|" & f.Name & "|" & f.Italic & "|" & f.Underline & "|" & TypeName(f.Size)
End Function

Function T32() As String
    Dim tr As TextRange
    Set tr = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50).TextFrame.TextRange
    tr.Text = "x"
    tr.Font.Size = 24
    tr.Font.Bold = True
    tr.Font.Italic = msoTrue
    T32 = tr.Font.Size & "|" & tr.Font.Bold & "|" & tr.Font.Italic
    tr.Font.Size = 10.5
    T32 = T32 & "|" & tr.Font.Size
End Function

Function T33() As String
    Dim tr As TextRange
    Set tr = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50).TextFrame.TextRange
    tr.Text = "x"
    tr.Font.Color.RGB = RGB(255, 0, 0)
    T33 = tr.Font.Color.RGB & "|" & tr.Font.Color.Type
    tr.Font.Color.RGB = RGB(1, 2, 3)
    T33 = T33 & "|" & tr.Font.Color.RGB
End Function

Function T34() As String
    Dim s As Slide
    Set s = P.Slides.Add(1, ppLayoutText)
    s.Shapes(1).TextFrame.TextRange.Text = "Title"
    s.Shapes(2).TextFrame.TextRange.Text = "Body"
    T34 = s.Shapes(1).TextFrame.TextRange.Font.Size & "|" & s.Shapes(2).TextFrame.TextRange.Font.Size & "|" & s.Shapes(1).TextFrame.TextRange.Font.Name & "|" & s.Shapes(2).TextFrame.TextRange.Font.Name
End Function

Function T35() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTable(2, 3, 10, 10, 300, 100)
    T35 = t.Name & "|" & t.Type & "|" & t.HasTable & "|" & t.Table.Rows.Count & "|" & t.Table.Columns.Count & "|" & t.HasTextFrame & "|" & TypeName(t.Table)
End Function

Function T36() As String
    Dim tb As Table
    Set tb = S1.Shapes.AddTable(2, 3, 10, 10, 300, 100).Table
    tb.Cell(1, 2).Shape.TextFrame.TextRange.Text = "B1"
    tb.Cell(2, 3).Shape.TextFrame.TextRange.Text = "C2"
    T36 = tb.Cell(1, 2).Shape.TextFrame.TextRange.Text & "|" & tb.Cell(2, 3).Shape.TextFrame.TextRange.Text & "|[" & tb.Cell(1, 1).Shape.TextFrame.TextRange.Text & "]"
End Function

Function T37() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTable(2, 3, 10, 10, 300, 100)
    T37 = Geo(t) & "|" & t.Table.Columns(1).Width & "|" & t.Table.Rows(1).Height & "|" & t.Table.Cell(1, 1).Shape.TextFrame.TextRange.Font.Size
End Function

Function T38() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTable(3, 4, 0, 0, 400, 120)
    t.Table.Columns(2).Width = 150
    T38 = t.Width & "|" & t.Table.Columns(2).Width & "|" & t.Table.Columns(1).Width
End Function

Function T39() As String
    Dim tr As TextRange, r As TextRange
    Set tr = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50).TextFrame.TextRange
    tr.Text = "Hello World World"
    Set r = tr.Replace("World", "There")
    T39 = tr.Text & "|" & r.Text & "|" & r.Start & "|" & r.Length
End Function

Function T40() As String
    Dim tr As TextRange, r As TextRange
    Set tr = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50).TextFrame.TextRange
    tr.Text = "Hello"
    Set r = tr.Replace("zzz", "y")
    T40 = TypeName(r) & "|" & (r Is Nothing) & "|" & tr.Text
End Function

Function T41() As String
    Dim tr As TextRange
    Set tr = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50).TextFrame.TextRange
    tr.Text = "One" & vbCr & "Two"
    T41 = tr.Length & "|" & tr.Start & "|" & tr.Paragraphs(2).Start & "|" & tr.Paragraphs(1).Length & "|" & tr.Paragraphs(2).Length
End Function

Function T42() As String
    Dim tr As TextRange
    Set tr = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50).TextFrame.TextRange
    tr.Text = "One" & vbCr & "Two"
    T42 = Asc(Right(tr.Paragraphs(1).Text, 1)) & "|" & Len(tr.Text) & "|" & InStr(tr.Text, vbCr)
End Function

Function T43() As String
    Dim tr As TextRange
    Set tr = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50).TextFrame.TextRange
    tr.Text = "mid"
    tr.InsertAfter "-end"
    tr.InsertBefore "start-"
    T43 = tr.Text & "|" & tr.InsertAfter("!").Text & "|" & tr.Text
End Function

Function T44() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50)
    T44 = t.TextFrame.HasText & "|" & Len(t.TextFrame.TextRange.Text)
    t.TextFrame.TextRange.Text = "a"
    T44 = T44 & "|" & t.TextFrame.HasText
End Function

Function T45() As String
    Dim sh As Shape, n As Long
    S1.Shapes.AddTextbox msoTextOrientationHorizontal, 0, 0, 100, 20
    S1.Shapes.AddShape msoShapeRectangle, 0, 0, 10, 10
    For Each sh In S1.Shapes
        n = n + sh.Type
    Next
    T45 = n & "|" & S1.Shapes.Count & "|" & P.Slides(2).Shapes.Count
End Function

Function T46() As String
    Dim s As Slide, r As String
    P.Slides.Add 2, ppLayoutBlank
    For Each s In P.Slides
        r = r & s.Name & "=" & s.SlideIndex & ";"
    Next
    T46 = r
End Function

Function T47() As String
    S1.Shapes.AddShape msoShapeRectangle, 0, 0, 10, 10
    S1.Shapes.AddShape msoShapeOval, 0, 0, 10, 10
    S1.Shapes.AddShape msoShapeRectangle, 0, 0, 10, 10
    T47 = S1.Shapes.Range(Array(1, 3)).Count & "|" & S1.Shapes.Range(Array("Oval 2")).Count & "|" & S1.Shapes.Range.Count & "|" & S1.Shapes.Range(Array(1, 3))(2).Name
End Function

Function T48() As String
    Dim s As Slide
    Set s = P.Slides.Add(1, ppLayoutText)
    T48 = s.Shapes.HasTitle & "|" & s.Shapes.Title.Name & "|" & P.Slides(2).Shapes.HasTitle
End Function

Function T49() As String
    Dim s As Slide
    On Error Resume Next
    Set s = P.Slides(5)
    T49 = Err.Number & "|" & (s Is Nothing)
End Function

Function T50() As String
    Dim sh As Shape
    S1.Shapes.AddShape msoShapeRectangle, 0, 0, 10, 10
    On Error Resume Next
    Set sh = S1.Shapes("NoSuch")
    T50 = Err.Number & "|" & (sh Is Nothing)
    Err.Clear
    Set sh = S1.Shapes(2)
    T50 = T50 & "|" & Err.Number
End Function

Function T51() As String
    T51 = S1.Layout & "|" & S1.CustomLayout.Name & "|" & P.SlideMaster.CustomLayouts.Count & "|" & P.SlideMaster.CustomLayouts(2).Name
End Function

Function T52() As String
    Dim i As Long
    For i = 1 To 3
        P.Slides.Add P.Slides.Count + 1, ppLayoutBlank
    Next
    T52 = P.Slides.Count & "|" & P.Slides(5).Name & "|" & P.Slides(5).SlideID
End Function

Function T53() As String
    P.Slides.Add 3, ppLayoutBlank
    P.Slides(3).MoveTo 1
    T53 = P.Slides(1).Name & "|" & P.Slides(2).Name & "|" & P.Slides(3).Name & "|" & P.Slides("Slide2").SlideIndex
End Function

Function T54() As String
    Dim r As Shape
    Set r = S1.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
    T54 = r.Rotation & "|" & r.Visible & "|" & r.HorizontalFlip
    r.Rotation = 45
    r.Visible = msoFalse
    T54 = T54 & "|" & r.Rotation & "|" & r.Visible
End Function

Function T55() As String
    Dim a As Shape, b As Shape, c As Shape
    Set a = S1.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
    Set b = S1.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
    b.Delete
    Set c = S1.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
    T55 = a.Id & "|" & c.Id & "|" & c.Name
End Function

Function T56() As String
    Dim s As Slide
    P.Slides(2).Delete
    Set s = P.Slides.Add(2, ppLayoutBlank)
    T56 = s.SlideID & "|" & s.Name & "|" & P.Slides.FindBySlideID(256).Name
End Function

Function T57() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTable(2, 2, 0, 0, 100, 50)
    T57 = TypeName(P.Slides) & "|" & TypeName(S1) & "|" & TypeName(S1.Shapes) & "|" & TypeName(t.TextFrame) & "|" & TypeName(t.Table.Cell(1, 1)) & "|" & TypeName(t.Table.Cell(1, 1).Shape.TextFrame.TextRange) & "|" & TypeName(P.PageSetup) & "|" & TypeName(P.SlideShowSettings)
End Function

Function T58() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 100, 20)
    With t.TextFrame
        T58 = .MarginLeft & "|" & .MarginRight & "|" & .MarginTop & "|" & .MarginBottom & "|" & .WordWrap & "|" & .AutoSize & "|" & .Orientation & "|" & .VerticalAnchor
    End With
End Function

Function T59() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 100, 20)
    t.TextFrame.TextRange.Text = "x"
    T59 = t.TextFrame.TextRange.ParagraphFormat.Alignment
    t.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    T59 = T59 & "|" & t.TextFrame.TextRange.ParagraphFormat.Alignment
End Function

Function T60() As String
    Dim r As Shape
    Set r = S1.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 50)
    r.TextFrame.TextRange.Text = "In box"
    With r.TextFrame
        T60 = .TextRange.ParagraphFormat.Alignment & "|" & .TextRange.Font.Color.RGB & "|" & .TextRange.Font.Size & "|" & .WordWrap & "|" & .AutoSize & "|" & .VerticalAnchor & "|" & r.Height
    End With
End Function

Function T61() As String
    Dim r As Shape
    Set r = S1.Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 50)
    T61 = r.Fill.ForeColor.RGB & "|" & r.Line.ForeColor.RGB & "|" & r.Fill.Visible & "|" & r.Line.Visible & "|" & r.Line.Weight
    r.Fill.ForeColor.RGB = RGB(0, 255, 0)
    T61 = T61 & "|" & r.Fill.ForeColor.RGB
End Function

Function T62() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 100, 20)
    T62 = t.Fill.Visible & "|" & t.Line.Visible
End Function

Function T63() As String
    Dim tb As Table
    Set tb = S1.Shapes.AddTable(2, 2, 0, 0, 200, 80).Table
    tb.Rows.Add
    tb.Columns.Add
    T63 = tb.Rows.Count & "|" & tb.Columns.Count & "|" & S1.Shapes(1).Width & "|" & tb.Columns(3).Width & "|" & tb.Rows(3).Height
End Function

Function T64() As String
    Dim s As Slide
    On Error Resume Next
    Set s = P.Slides.Add(5, ppLayoutBlank)
    T64 = Err.Number & "|" & P.Slides.Count
End Function

Function T65() As String
    Dim r As Shape, sh As Shape
    Set r = S1.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
    r.Name = "Gone"
    r.Delete
    On Error Resume Next
    Set sh = S1.Shapes("Gone")
    T65 = Err.Number & "|" & S1.Shapes.Count
End Function

Function T66() As String
    Dim tr As TextRange
    Set tr = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50).TextFrame.TextRange
    tr.Text = "abcdef"
    tr.Characters(3, 2).Font.Bold = msoTrue
    T66 = tr.Runs.Count & "|" & tr.Runs(1).Text & "|" & tr.Runs(2).Text & "|" & tr.Runs(3).Text & "|" & tr.Font.Bold & "|" & tr.Runs(1).Font.Bold
End Function

Function T67() As String
    Dim tr As TextRange
    Set tr = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50).TextFrame.TextRange
    tr.Text = "The quick brown fox"
    T67 = tr.Words.Count & "|" & tr.Words(2).Text & "|" & tr.Words(4).Text & "|" & Len(tr.Words(2).Text)
End Function

Function T68() As String
    T68 = P.Slides("Slide2").SlideIndex & "|" & P.Slides(1).Shapes.Count & "|" & P.Slides.Item(2).Name
End Function

Function T69() As String
    Dim r As SlideRange
    Set r = P.Slides(2).Duplicate
    T69 = r.SlideID & "|" & r.Name & "|" & r.SlideIndex & "|" & P.Slides.Count
End Function

Function T70() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 20, 200, 50)
    t.TextFrame.TextRange.Text = "Hello" & vbCr & "World"
    T70 = t.Height & "|" & t.Top
End Function

Function T71() As String
    Dim tr As TextRange
    Set tr = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50).TextFrame.TextRange
    tr.Text = "a,b,c"
    tr.Replace ",", ";"
    T71 = tr.Text
    tr.Text = "aaa"
    Set tr = tr.Replace("a", "bb")
    T71 = T71 & "|" & tr.Text & "|" & tr.Parent.Parent.TextRange.Text
End Function

Function T72() As String
    Dim tr As TextRange
    Set tr = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50).TextFrame.TextRange
    tr.Text = "abc"
    tr.Characters(2, 1).Text = "XYZ"
    T72 = tr.Text & "|" & tr.Length
    tr.Paragraphs(1).Text = "new"
    T72 = T72 & "|" & tr.Text
End Function

Function T73() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50)
    t.TextFrame.TextRange.Text = "Big"
    t.TextFrame.TextRange.Font.Size = 36
    T73 = t.TextFrame.TextRange.Font.Size & "|" & t.Height
End Function

Function T74() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50)
    t.TextFrame.AutoSize = ppAutoSizeNone
    t.TextFrame.TextRange.Text = "x"
    T74 = t.Height & "|" & t.TextFrame.AutoSize
End Function

Function T75() As String
    Dim s As Slide
    Set s = P.Slides.Add(1, ppLayoutText)
    s.Shapes(1).Delete
    T75 = s.Shapes.Count & "|" & s.Shapes.HasTitle & "|" & s.Shapes(1).Name
End Function

Function T76() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50)
    T76 = t.TextFrame.TextRange.Paragraphs.Count & "|" & t.TextFrame.TextRange.Runs.Count & "|" & t.TextFrame.TextRange.Characters.Count
End Function

Function T77() As String
    Dim t As Variant, r As String, sh As Shape
    For Each t In Array(msoShapeRoundedRectangle, msoShapeIsoscelesTriangle, msoShape5pointStar, msoShapeRightArrow, msoShapeDiamond, msoShapeHexagon)
        Set sh = S1.Shapes.AddShape(t, 0, 0, 20, 20)
        r = r & sh.Name & "=" & sh.AutoShapeType & ";"
    Next
    T77 = r
End Function

Function T78() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTable(2, 2, 0, 0, 100, 40)
    T78 = t.Height & "|" & t.Table.Rows(1).Height & "|" & t.Width
End Function

Function T79() As String
    Dim tb As Table
    Set tb = S1.Shapes.AddTable(2, 2, 0, 0, 200, 100).Table
    tb.Cell(1, 1).Shape.TextFrame.TextRange.Text = "h"
    tb.Cell(2, 1).Shape.TextFrame.TextRange.Text = "b"
    With tb.Cell(1, 1).Shape.TextFrame.TextRange.Font
        T79 = .Bold & "|" & .Color.RGB & "|" & .Size
    End With
    With tb.Cell(2, 1).Shape.TextFrame.TextRange.Font
        T79 = T79 & "|" & .Bold & "|" & .Color.RGB & "|" & .Size
    End With
End Function

Function T80() As String
    Dim np As Presentation
    Set np = Presentations.Add(msoFalse)
    T80 = Presentations.Count & "|" & np.Slides.Count & "|" & np.PageSetup.SlideWidth & "|" & np.PageSetup.SlideHeight & "|" & Presentations(1).Name
    np.Slides.Add 1, ppLayoutText
    T80 = T80 & "|" & np.Slides(1).Shapes.Count & "|" & np.Slides(1).CustomLayout.Name & "|" & np.SlideMaster.CustomLayouts.Count
End Function

Function T81() As String
    Dim tr As TextRange
    Set tr = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50).TextFrame.TextRange
    tr.Text = "a" & vbLf & "b" & vbCrLf & "c"
    T81 = Len(tr.Text) & "|" & tr.Paragraphs.Count & "|" & Asc(Mid(tr.Text, 2, 1)) & "|" & Asc(Mid(tr.Text, 4, 1))
End Function

Function T82() As String
    Dim s As Slide
    Set s = P.Slides.Add(1, ppLayoutText)
    With s.Shapes(1).TextFrame
        T82 = .AutoSize & "|" & .VerticalAnchor & "|" & .WordWrap & "|" & .TextRange.ParagraphFormat.Alignment & "|" & s.Shapes(1).Fill.Visible & "|" & s.Shapes(1).Line.Visible
    End With
End Function

Function T83() As String
    Dim s As Slide, r As String, lay As Variant
    For Each lay In Array(ppLayoutObject, ppLayoutTwoColumnText, ppLayoutTitleOnly)
        Set s = P.Slides.Add(1, lay)
        r = r & s.Layout & ":" & s.Shapes.Count & ":" & NamesOf(s) & ":" & s.CustomLayout.Name & "/"
    Next
    T83 = r & P.SlideMaster.CustomLayouts.Count
End Function

Function T84() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50)
    t.TextFrame.TextRange.Text = "Hello"
    t.TextFrame.TextRange.Text = ""
    T84 = t.Height & "|" & t.TextFrame.TextRange.Paragraphs.Count
End Function

Function T85() As String
    Dim t As Shape
    Set t = S1.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 50)
    t.TextFrame.TextRange.Text = "One" & vbCr & "Two" & vbCr & "Three"
    T85 = t.Height
    t.TextFrame.TextRange.Paragraphs(2).Font.Size = 30
    T85 = T85 & "|" & t.Height
End Function
