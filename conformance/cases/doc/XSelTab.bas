Option Explicit

Private Function V(s As String) As String
    V = Replace(Replace(Replace(s, Chr(13), "P"), Chr(7), "B"), vbTab, "T")
End Function

Function XS01() As String
    ActiveDocument.Content.Text = "One" & vbCr & "Two" & vbCr & "Three"
    Selection.HomeKey wdStory
    Selection.MoveDown wdParagraph, 1
    XS01 = Selection.Start
End Function

Function XS02() As String
    ActiveDocument.Content.Text = "One" & vbCr & "Two" & vbCr & "Three"
    Selection.EndKey wdStory
    Selection.MoveUp wdParagraph, 1
    XS02 = Selection.Start
End Function

Function XS03() As String
    ActiveDocument.Content.Text = "Hello world"
    Selection.SetRange 3, 3
    Selection.Expand wdWord
    XS03 = "[" & Selection.Text & "]|" & Selection.Start & "|" & Selection.End
End Function

Function XS04() As String
    ActiveDocument.Content.Text = "abc"
    Selection.SetRange 1, 1
    Selection.TypeText "X" & vbCr & "Y"
    XS04 = V(ActiveDocument.Content.Text) & "|" & Selection.Start & "|" & ActiveDocument.Paragraphs.Count
End Function

Function XS05() As String
    ActiveDocument.Content.Text = "abc"
    Selection.WholeStory
    Selection.TypeText "new"
    XS05 = V(ActiveDocument.Content.Text) & "|" & Selection.Start
End Function

Function XS06() As String
    ActiveDocument.Content.Text = "abc"
    Selection.WholeStory
    Selection.Delete
    XS06 = Len(ActiveDocument.Content.Text) & "|" & Selection.Start & "|" & Selection.End
End Function

Function XS07() As String
    Dim n As Long
    ActiveDocument.Content.Text = "Hello big world"
    Selection.EndKey wdStory
    n = Selection.MoveLeft(wdWord, 1)
    XS07 = n & "|" & Selection.Start
End Function

Function XS08() As String
    ActiveDocument.Content.Text = "abcdef"
    Selection.SetRange 2, 2
    Selection.MoveLeft wdCharacter, 2, wdExtend
    XS08 = Selection.Start & "|" & Selection.End & "|" & Selection.Text
End Function

Function XS09() As String
    ActiveDocument.Content.Text = "abc"
    Selection.SetRange 0, 3
    Selection.Collapse wdCollapseEnd
    Selection.TypeText "!"
    XS09 = V(ActiveDocument.Content.Text)
End Function

Function XS10() As String
    ActiveDocument.Content.Text = "Hello"
    Selection.HomeKey wdStory
    Selection.EndKey wdStory, wdExtend
    XS10 = Selection.Start & "|" & Selection.End & "|" & V(Selection.Text)
End Function

Function XS11() As String
    ActiveDocument.Content.Text = "Hello"
    Selection.EndKey wdLine
    XS11 = Selection.Start
End Function

Function XS12() As String
    ActiveDocument.Content.Text = "abc"
    ActiveDocument.Range(1, 2).Select
    XS12 = Selection.Text & "|" & Selection.Type & "|" & TypeName(Selection.Range)
End Function

Function XS13() As String
    ActiveDocument.Content.Text = "abc"
    Selection.SetRange 1, 2
    Selection.Range.Text = "XYZ"
    XS13 = ActiveDocument.Content.Text & "|" & Selection.Start & "|" & Selection.End
End Function

Function XS14() As String
    ActiveDocument.Content.Text = "Hello world"
    Selection.SetRange 0, 5
    Selection.Font.Italic = True
    XS14 = ActiveDocument.Words(1).Italic & "|" & ActiveDocument.Range(0, 5).Italic
End Function

Function XS15() As String
    Selection.TypeText "abc"
    Selection.TypeBackspace
    Selection.TypeBackspace
    XS15 = V(ActiveDocument.Content.Text) & "|" & Selection.Start
End Function

Function XT01() As String
    Dim t As Table
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Range(0, 0), 2, 2)
    t.Rows.Add t.Rows(1)
    t.Cell(1, 1).Range.Text = "new"
    XT01 = t.Rows.Count & "|" & V(t.Range.Text)
End Function

Function XT02() As String
    Dim t As Table
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Range(0, 0), 2, 3)
    t.Cell(1, 1).Range.Text = "a"
    t.Cell(1, 2).Range.Text = "b"
    t.Columns(1).Delete
    XT02 = t.Columns.Count & "|" & V(t.Range.Text)
End Function

Function XT03() As String
    Dim t As Table
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Range(0, 0), 2, 2)
    t.Cell(1, 1).Range.Text = "a" & vbCr & "b"
    XT03 = ActiveDocument.Paragraphs.Count & "|" & V(t.Cell(1, 1).Range.Text) & "|" & t.Cell(1, 2).Range.Start
End Function

Function XT04() As String
    Dim t As Table
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Range(0, 0), 2, 2)
    t.Cell(2, 1).Range.Text = "x"
    t.Cell(2, 1).Range.Delete
    XT04 = V(t.Range.Text) & "|" & t.Range.End
End Function

Function XT05() As String
    Dim t As Table
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Range(0, 0), 2, 2)
    t.Cell(1, 2).Range.Text = "abc"
    XT05 = t.Cell(1, 2).Range.Characters.Count & "|" & t.Cell(1, 2).Range.Paragraphs.Count & "|" & V(t.Cell(1, 2).Range.Paragraphs(1).Range.Text)
End Function

Function XT06() As String
    Dim t As Table
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Range(0, 0), 2, 2)
    XT06 = ActiveDocument.Range(1, 1).Information(wdWithInTable) & "|" & ActiveDocument.Range(6, 6).Information(wdWithInTable)
End Function

Function XT07() As String
    Dim t As Table
    ActiveDocument.Content.Text = "abc"
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Range(1, 1), 1, 1)
    XT07 = V(ActiveDocument.Content.Text) & "|" & ActiveDocument.Paragraphs.Count
End Function

Function XT08() As String
    Dim t As Table, c As Cell, s As String
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Range(0, 0), 2, 2)
    For Each c In t.Range.Cells
        s = s & c.RowIndex & c.ColumnIndex & ";"
    Next
    XT08 = s
End Function

Function XT09() As String
    Dim t As Table
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Range(0, 0), 2, 2)
    t.Cell(1, 1).Range.Text = "abc"
    t.Cell(1, 1).Range.Font.Bold = True
    XT09 = t.Cell(1, 1).Range.Font.Bold & "|" & t.Cell(1, 2).Range.Font.Bold & "|" & t.Range.Font.Bold
End Function

Function XT10() As String
    Dim t As Table
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Range(0, 0), 2, 2)
    t.Cell(1, 1).Range.Text = "x"
    t.Cell(2, 2).Range.Text = "y"
    ActiveDocument.Content.Find.Execute FindText:="y", ReplaceWith:="zz", Replace:=wdReplaceAll
    XT10 = V(ActiveDocument.Content.Text)
End Function

Function XT11() As String
    Dim t As Table, t2 As Table
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Range(0, 0), 1, 1)
    ActiveDocument.Content.InsertAfter "between"
    ActiveDocument.Content.InsertParagraphAfter
    Set t2 = ActiveDocument.Tables.Add(ActiveDocument.Paragraphs.Last.Range, 1, 2)
    XT11 = ActiveDocument.Tables.Count & "|" & ActiveDocument.Tables(2).Columns.Count & "|" & V(ActiveDocument.Content.Text)
End Function

Function XT12() As String
    Dim t As Table
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Range(0, 0), 2, 2)
    t.Cell(1, 1).Range.Text = "abc"
    XT12 = t.Cell(1, 1).Range.Words.Count & "|" & t.Cell(1, 1).Range.Sentences.Count & "|" & ActiveDocument.Sentences.Count
End Function

Function XT13() As String
    Dim t As Table
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Range(0, 0), 3, 2)
    t.Rows(2).Delete
    t.Rows(1).Delete
    XT13 = t.Rows.Count & "|" & ActiveDocument.Tables.Count & "|" & ActiveDocument.Content.End
End Function

Function XT14() As String
    Dim t As Table
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Range(0, 0), 2, 2)
    XT14 = t.Rows(1).Range.Start & "-" & t.Rows(1).Range.End & "|" & t.Rows(2).Range.Start & "-" & t.Rows(2).Range.End & "|" & t.Columns(2).Cells(2).Range.Start
End Function

Function XK01() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "bm", ActiveDocument.Range(6, 11)
    ActiveDocument.Range(6, 6).InsertBefore "big "
    XK01 = ActiveDocument.Bookmarks("bm").Start & "|" & ActiveDocument.Bookmarks("bm").Range.Text
End Function

Function XK02() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "bm", ActiveDocument.Range(6, 11)
    ActiveDocument.Bookmarks("bm").Range.InsertBefore "big "
    XK02 = ActiveDocument.Bookmarks("bm").Start & "|" & ActiveDocument.Bookmarks("bm").Range.Text
End Function

Function XK03() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "bm", ActiveDocument.Range(3, 3)
    ActiveDocument.Range(3, 3).InsertAfter "XX"
    XK03 = ActiveDocument.Bookmarks("bm").Start & "|" & ActiveDocument.Bookmarks("bm").End
End Function

Function XK04() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "bm", ActiveDocument.Range(0, 5)
    ActiveDocument.Range(0, 11).Delete
    XK04 = ActiveDocument.Bookmarks.Exists("bm") & "|" & ActiveDocument.Bookmarks.Count
End Function

Function XK05() As String
    ActiveDocument.Content.Text = "Hello world"
    Selection.SetRange 2, 4
    ActiveDocument.Bookmarks.Add "sel"
    XK05 = ActiveDocument.Bookmarks("sel").Start & "|" & ActiveDocument.Bookmarks("sel").End
End Function

Function XK06() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "a_1", ActiveDocument.Range(0, 1)
    ActiveDocument.Bookmarks.Add "B2", ActiveDocument.Range(1, 2)
    XK06 = ActiveDocument.Bookmarks(1).Name & "|" & ActiveDocument.Bookmarks(2).Name
End Function

Function XK07() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "bm", ActiveDocument.Range(0, 5)
    ActiveDocument.Bookmarks("bm").Select
    XK07 = Selection.Text & "|" & Selection.Start
End Function

Function XK08() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "bm", ActiveDocument.Range(2, 8)
    ActiveDocument.Range(0, 4).Delete
    XK08 = ActiveDocument.Bookmarks("bm").Start & "|" & ActiveDocument.Bookmarks("bm").End & "|" & ActiveDocument.Bookmarks("bm").Range.Text
End Function

Function XD01() As String
    ActiveDocument.Content.Text = "Title" & vbCr & "Body"
    ActiveDocument.Paragraphs(1).Style = "Heading 2"
    XD01 = ActiveDocument.Paragraphs(1).Style & "|" & ActiveDocument.Paragraphs(1).Range.ParagraphFormat.Style & "|" & ActiveDocument.Range(0, 2).Style
End Function

Function XD02() As String
    ActiveDocument.Content.Text = "Title"
    ActiveDocument.Paragraphs(1).Style = wdStyleTitle
    ActiveDocument.Content.InsertParagraphAfter
    ActiveDocument.Content.InsertAfter "Next"
    XD02 = ActiveDocument.Paragraphs(1).Style & "|" & ActiveDocument.Paragraphs(2).Style
End Function

Function XD03() As String
    ActiveDocument.Content.Text = "abc"
    ActiveDocument.Paragraphs(1).Style = "No Such Style"
    XD03 = "set"
End Function

Function XD04() As String
    ActiveDocument.Content.Text = "abc"
    ActiveDocument.Content.Font.Name = "Arial"
    ActiveDocument.Content.Font.Size = 20
    ActiveDocument.Content.Font.Reset
    XD04 = ActiveDocument.Content.Font.Name & "|" & ActiveDocument.Content.Font.Size
End Function

Function XD05() As String
    ActiveDocument.Content.Text = "abc"
    ActiveDocument.Content.ParagraphFormat.Alignment = wdAlignParagraphCenter
    ActiveDocument.Content.ParagraphFormat.LeftIndent = 20
    ActiveDocument.Content.ParagraphFormat.Reset
    XD05 = ActiveDocument.Content.ParagraphFormat.Alignment & "|" & ActiveDocument.Content.ParagraphFormat.LeftIndent
End Function

Function XD06() As String
    ActiveDocument.Content.Text = "abc"
    ActiveDocument.Content.Font.Size = 0.5
    XD06 = "set"
End Function

Function XD07() As String
    ActiveDocument.Content.Text = "abc"
    ActiveDocument.Content.Font.Size = 11.3
    XD07 = ActiveDocument.Content.Font.Size
End Function

Function XD08() As String
    ActiveDocument.Content.Text = "Hello"
    ActiveDocument.Range(0, 1).Font.Superscript = True
    ActiveDocument.Range(1, 2).Font.StrikeThrough = True
    XD08 = ActiveDocument.Range(0, 1).Font.Superscript & "|" & ActiveDocument.Range(0, 1).Font.Subscript & "|" & ActiveDocument.Range(1, 2).Font.StrikeThrough & "|" & ActiveDocument.Content.Font.Superscript
End Function

Function XD09() As String
    ActiveDocument.Content.Text = "Hello"
    ActiveDocument.Range(0, 5).HighlightColorIndex = wdYellow
    XD09 = ActiveDocument.Range(0, 5).HighlightColorIndex & "|" & ActiveDocument.Content.HighlightColorIndex
End Function

Function XD10() As String
    ActiveDocument.Content.Text = "abc"
    XD10 = ActiveDocument.Range(1, 1).Font.Bold & "|" & ActiveDocument.Range(1, 1).Font.Name & "|" & ActiveDocument.Range(1, 1).Bold
End Function
