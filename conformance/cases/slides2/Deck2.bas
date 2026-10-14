Option Explicit

' Held-out cases, written after the model was built from conformance/cases/slides: they test whether the
' model generalizes. Fresh copy of the generated presentation per case (2 blank slides).

Private Function P() As Presentation
    Set P = Presentations(1)
End Function

Function U01() As String
    Dim s As Slide, sh As Shape, r As String
    For Each s In P.Slides
        Set sh = s.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 400, 40)
        sh.TextFrame.TextRange.Text = "Slide " & s.SlideIndex
    Next
    For Each s In P.Slides
        r = r & s.Shapes(1).TextFrame.TextRange.Text & "/" & s.Shapes(1).Name & ";"
    Next
    U01 = r
End Function

Function U02() As String
    Dim s As Slide, i As Long
    For i = 1 To 3
        Set s = P.Slides.Add(P.Slides.Count + 1, ppLayoutText)
        s.Shapes.Title.TextFrame.TextRange.Text = "Title " & i
        s.Shapes(2).TextFrame.TextRange.Text = "Point A" & vbCr & "Point B"
    Next
    U02 = P.Slides.Count & "|" & P.Slides(4).Shapes.Title.TextFrame.TextRange.Text & "|" & P.Slides(5).Shapes(2).TextFrame.TextRange.Paragraphs.Count & "|" & P.Slides(5).Name
End Function

Function U03() As String
    Dim s As Slide, sh As Shape, n As Long
    P.Slides.Add(1, ppLayoutText).Shapes(2).TextFrame.TextRange.Text = "foo bar foo"
    P.Slides(2).Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 100, 20).TextFrame.TextRange.Text = "Foo"
    For Each s In P.Slides
        For Each sh In s.Shapes
            If sh.HasTextFrame Then
                If Not sh.TextFrame.TextRange.Replace("foo", "baz") Is Nothing Then n = n + 1
            End If
        Next
    Next
    U03 = n & "|" & P.Slides(1).Shapes(2).TextFrame.TextRange.Text & "|" & P.Slides(2).Shapes(1).TextFrame.TextRange.Text
End Function

Function U04() As String
    Dim tb As Table, r As Long, c As Long
    Set tb = P.Slides(1).Shapes.AddTable(3, 3, 20, 20, 600, 200).Table
    For r = 1 To 3
        For c = 1 To 3
            tb.Cell(r, c).Shape.TextFrame.TextRange.Text = r * c
        Next
    Next
    U04 = tb.Cell(3, 3).Shape.TextFrame.TextRange.Text & "|" & tb.Cell(2, 3).Shape.TextFrame.TextRange.Text & "|" & tb.Rows(2).Height & "|" & tb.Columns(3).Width & "|" & P.Slides(1).Shapes(1).Height
End Function

Function U05() As String
    Dim i As Long
    For i = 1 To 5
        P.Slides(1).Shapes.AddShape msoShapeRectangle, i * 10, i * 10, 20, 20
    Next
    For i = P.Slides(1).Shapes.Count To 1 Step -1
        If i Mod 2 = 0 Then P.Slides(1).Shapes(i).Delete
    Next
    U05 = P.Slides(1).Shapes.Count & "|" & P.Slides(1).Shapes(2).Name & "|" & P.Slides(1).Shapes(3).Left
End Function

Function U06() As String
    Dim sh As Shape
    Set sh = P.Slides(1).Shapes.AddShape(msoShapeOval, 100, 50, 80, 60)
    sh.Left = sh.Left + sh.Width / 2
    sh.Top = P.PageSetup.SlideHeight - sh.Height
    sh.Width = sh.Width * 1.5
    U06 = sh.Left & "|" & sh.Top & "|" & sh.Width & "|" & sh.Height
End Function

Function U07() As String
    Dim tr As TextRange
    Set tr = P.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 400, 40).TextFrame.TextRange
    tr.Text = "Alpha" & vbCr & "Beta" & vbCr & "Gamma"
    tr.Paragraphs(2).Font.Bold = msoTrue
    tr.Paragraphs(3).ParagraphFormat.Alignment = ppAlignRight
    U07 = tr.Paragraphs(1).Font.Bold & "|" & tr.Paragraphs(2).Font.Bold & "|" & tr.Font.Bold & "|" & tr.Paragraphs(3).ParagraphFormat.Alignment & "|" & tr.ParagraphFormat.Alignment & "|" & tr.Runs.Count
End Function

Function U08() As String
    Dim tr As TextRange, f As TextRange
    Set tr = P.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 400, 40).TextFrame.TextRange
    tr.Text = "one two three two"
    Set f = tr.Find("two")
    U08 = f.Start & "|" & f.Length & "|" & f.Text
    Set f = tr.Find("two", 5)
    U08 = U08 & "|" & f.Start
    Set f = tr.Find("TWO", 0, msoTrue)
    U08 = U08 & "|" & (f Is Nothing)
End Function

Function U09() As String
    Dim r As SlideRange, s As Slide, out As String
    P.Slides.Add 3, ppLayoutBlank
    Set r = P.Slides.Range(Array(3, 1))
    For Each s In r
        out = out & s.Name & ";"
    Next
    U09 = out & r.Count & "|" & r(1).SlideIndex
End Function

Function U10() As String
    Dim sr As ShapeRange
    With P.Slides(1).Shapes
        .AddShape msoShapeRectangle, 0, 0, 10, 10
        .AddShape msoShapeRectangle, 0, 0, 10, 10
        .AddShape msoShapeRectangle, 0, 0, 10, 10
    End With
    Set sr = P.Slides(1).Shapes.Range(Array(1, 3))
    sr.Left = 77
    sr.Delete
    U10 = P.Slides(1).Shapes.Count & "|" & P.Slides(1).Shapes(1).Name & "|" & P.Slides(1).Shapes(1).Left
End Function

Function U11() As String
    Dim s As Slide
    On Error Resume Next
    Set s = P.Slides.Add(1, ppLayoutBlank)
    s.Shapes.AddShape msoShapeRectangle, 0, 0, 10, 10
    Dim t As Long
    t = s.Shapes(1).PlaceholderFormat.Type
    U11 = Err.Number
    Err.Clear
    Dim tb As Object
    Set tb = s.Shapes(1).Table
    U11 = U11 & "|" & Err.Number
End Function

Function U12() As String
    Dim t As Shape, x As String
    Set t = P.Slides(1).Shapes.AddTable(2, 2, 0, 0, 100, 60)
    On Error Resume Next
    x = t.TextFrame.TextRange.Text
    U12 = Err.Number & "|" & t.TextFrame.HasText
End Function

Function U13() As String
    Dim tb As Table
    Set tb = P.Slides(1).Shapes.AddTable(3, 3, 0, 0, 300, 120).Table
    tb.Rows(2).Delete
    tb.Columns(1).Delete
    U13 = tb.Rows.Count & "|" & tb.Columns.Count & "|" & P.Slides(1).Shapes(1).Width & "|" & P.Slides(1).Shapes(1).Height
End Function

Function U14() As String
    Dim tb As Table
    Set tb = P.Slides(1).Shapes.AddTable(2, 2, 0, 0, 200, 100).Table
    tb.Rows(1).Height = 10
    tb.Rows(2).Height = 70
    U14 = tb.Rows(1).Height & "|" & tb.Rows(2).Height & "|" & P.Slides(1).Shapes(1).Height
End Function

Function U15() As String
    P.PageSetup.SlideHeight = 600
    U15 = P.PageSetup.SlideHeight & "|" & P.PageSetup.SlideWidth & "|" & P.PageSetup.SlideOrientation
    P.PageSetup.SlideWidth = 500
    U15 = U15 & "|" & P.PageSetup.SlideOrientation
End Function

Function U16() As String
    Dim r As SlideRange
    P.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 100, 20).TextFrame.TextRange.Text = "orig"
    Set r = P.Slides(1).Duplicate
    r.Shapes(1).TextFrame.TextRange.Text = "copy"
    U16 = P.Slides(1).Shapes(1).TextFrame.TextRange.Text & "|" & P.Slides(2).Shapes(1).TextFrame.TextRange.Text & "|" & P.Slides(2).Shapes(1).Name
End Function

Function U17() As String
    P.Slides.Add 3, ppLayoutBlank
    P.Slides(1).MoveTo P.Slides.Count
    U17 = P.Slides(1).Name & "|" & P.Slides(2).Name & "|" & P.Slides(3).Name & "|" & P.Slides(3).SlideIndex
End Function

Function U18() As String
    P.PageSetup.FirstSlideNumber = 0
    P.Slides.Add 1, ppLayoutBlank
    U18 = P.Slides(1).SlideNumber & "|" & P.Slides(3).SlideNumber
    P.Slides(1).Delete
    U18 = U18 & "|" & P.Slides(1).SlideNumber
End Function

Function U19() As String
    Dim tr As TextRange
    Set tr = P.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 400, 40).TextFrame.TextRange
    tr.Text = "Hello"
    tr.Characters(1, 1).Font.Italic = msoTrue
    tr.Characters(5, 1).Font.Italic = msoTrue
    tr.Characters(3, 1).Font.Size = 30
    U19 = tr.Runs.Count & "|" & tr.Runs(2).Text & "|" & tr.Runs(3).Font.Size & "|" & tr.Font.Italic
End Function

Function U20() As String
    Dim tr As TextRange, ins As TextRange
    Set tr = P.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 400, 40).TextFrame.TextRange
    tr.Text = "world"
    Set ins = tr.InsertBefore("hello ")
    U20 = ins.Text & "|" & ins.Start & "|" & tr.Text & "|" & tr.Length
End Function

Function U21() As String
    Dim tr As TextRange
    Set tr = P.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 400, 40).TextFrame.TextRange
    tr.Text = "a  b" & vbCr & "c"
    U21 = tr.Words.Count & "|" & Len(tr.Words(1).Text) & "|" & tr.Characters.Count & "|" & tr.Paragraphs(1).Words.Count
End Function

Function U22() As String
    Dim s As Slide, sh As Shape, out As String
    Set s = P.Slides.Add(1, ppLayoutTitle)
    s.Shapes(1).TextFrame.TextRange.Text = "Deck"
    s.Shapes(2).TextFrame.TextRange.Text = "Sub"
    For Each sh In s.Shapes.Placeholders
        out = out & sh.PlaceholderFormat.Type & "=" & sh.TextFrame.TextRange.Font.Size & ";"
    Next
    U22 = out & s.Shapes.Title.TextFrame.TextRange.Text
End Function

Function U23() As String
    Dim t As Shape
    Set t = P.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 40, 250, 100)
    t.TextFrame.TextRange.Text = "Line 1" & vbCr & "Line 2" & vbCr & "Line 3" & vbCr & "Line 4"
    U23 = t.Top & "|" & t.Height
    t.TextFrame.TextRange.Font.Size = 12
    U23 = U23 & "|" & t.Height
End Function

Function U24() As String
    Dim t As Shape
    Set t = P.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 250, 100)
    t.TextFrame.TextRange.Text = "x"
    t.TextFrame.MarginTop = 10
    t.TextFrame.MarginBottom = 0
    t.TextFrame.TextRange.Text = "y"
    U24 = t.Height & "|" & t.TextFrame.MarginTop
End Function

Function U25() As String
    Dim sh As Shape
    Set sh = P.Slides(1).Shapes.AddShape(msoShapeRectangle, 0, 0, 100, 100)
    sh.Fill.ForeColor.RGB = RGB(10, 20, 30)
    sh.Line.Visible = msoFalse
    sh.Line.Weight = 2.5
    sh.TextFrame.TextRange.Text = "Box"
    sh.TextFrame.TextRange.Font.Color.RGB = vbBlack
    U25 = sh.Fill.ForeColor.RGB & "|" & sh.Line.Visible & "|" & sh.Line.Weight & "|" & sh.TextFrame.TextRange.Font.Color.RGB & "|" & sh.Height
End Function

Function U26() As String
    Dim sh As Shape, d As ShapeRange
    Set sh = P.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 100, 20)
    sh.Name = "Label"
    sh.TextFrame.TextRange.Text = "dup me"
    Set d = sh.Duplicate
    U26 = d.Name & "|" & d.TextFrame.TextRange.Text & "|" & d.Top & "|" & P.Slides(1).Shapes.Count
End Function

Function U27() As String
    Dim s As Slide
    Set s = P.Slides.Add(2, ppLayoutTitleOnly)
    U27 = s.CustomLayout.Name & "|" & s.Shapes(1).Name & "|" & s.Shapes(1).Top & "|" & s.Shapes(1).TextFrame.TextRange.Font.Name
End Function

Function U28() As String
    Dim tr As TextRange
    Set tr = P.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 400, 40).TextFrame.TextRange
    tr.Text = "The cat sat on the mat"
    Do While Not tr.Replace("at", "og") Is Nothing
    Loop
    U28 = tr.Text
End Function

Function U29() As String
    Dim tr As TextRange
    Set tr = P.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 400, 40).TextFrame.TextRange
    tr.Text = "abc"
    tr.Characters(2, 1).Delete
    U29 = tr.Text & "|" & tr.Length
    tr.Text = ""
    U29 = U29 & "|" & tr.Length & "|" & tr.Font.Size
End Function

Function U30() As String
    Dim s As Slide
    Set s = P.Slides.FindBySlideID(257)
    U30 = s.SlideIndex & "|" & s.Name
    s.Name = "Intro"
    U30 = U30 & "|" & P.Slides("Intro").SlideID & "|" & P.Slides(2).Name
End Function

Function U31() As String
    Dim sh As Shape
    Set sh = P.Slides(1).Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
    sh.Rotation = 400
    U31 = sh.Rotation
    sh.Rotation = -30
    U31 = U31 & "|" & sh.Rotation
End Function

Function U32() As String
    Dim tr As TextRange
    Set tr = P.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 400, 40).TextFrame.TextRange
    tr.Text = "abc"
    tr.Font.Name = "Arial"
    tr.InsertAfter "def"
    U32 = tr.Font.Name & "|" & tr.Characters(5, 1).Font.Name & "|" & tr.Runs.Count
End Function

Function U33() As String
    Dim n As Long
    On Error Resume Next
    P.Slides(1).Shapes.AddTable 0, 2
    n = Err.Number
    Err.Clear
    P.Slides(1).Shapes.AddTable 2, 2
    U33 = n & "|" & Err.Number & "|" & P.Slides(1).Shapes.Count & "|" & P.Slides(1).Shapes(1).Width & "|" & P.Slides(1).Shapes(1).Height
End Function

Function U34() As String
    Dim c As Long
    c = P.Slides.Count
    P.Slides.Range.Delete
    U34 = c & "|" & P.Slides.Count
    P.Slides.Add 1, ppLayoutText
    U34 = U34 & "|" & P.Slides(1).Name & "|" & P.Slides(1).SlideID
End Function

Function U35() As String
    Dim tr As TextRange
    Set tr = P.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 400, 40).TextFrame.TextRange
    tr.Text = "x"
    tr.Font.Size = 20
    tr.Text = "new text"
    U35 = tr.Font.Size & "|" & P.Slides(1).Shapes(1).Height
End Function
