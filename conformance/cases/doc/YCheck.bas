Option Explicit

' Third batch: confirms rules inferred from the first two batches.

Function YC01() As String
    ActiveDocument.Content.Text = "ab cd"
    ActiveDocument.Range(0, 2).Font.Bold = True
    ActiveDocument.Range(3, 5).Font.Bold = True
    YC01 = ActiveDocument.Range(0, 5).Font.Bold & "|" & ActiveDocument.Range(2, 3).Font.Bold & "|" & ActiveDocument.Content.Font.Bold
End Function

Function YC02() As String
    ActiveDocument.Content.Text = "Heading text"
    ActiveDocument.Paragraphs(1).Style = wdStyleHeading1
    ActiveDocument.Range(7, 7).InsertParagraphAfter
    YC02 = ActiveDocument.Paragraphs.Count & "|" & ActiveDocument.Paragraphs(1).Style & "|" & ActiveDocument.Paragraphs(2).Style
End Function

Function YC03() As String
    Dim n As Long
    ActiveDocument.Content.Text = "Hello big world"
    n = ActiveDocument.Range(0, 6).Delete
    YC03 = n & "|" & ActiveDocument.Content.Text
End Function

Function YC04() As String
    ActiveDocument.Content.Text = "a, b; c: d"
    YC04 = ActiveDocument.Words.Count & "|[" & ActiveDocument.Words(2).Text & "]|[" & ActiveDocument.Words(4).Text & "]"
End Function

Function YC05() As String
    ActiveDocument.Content.Text = "Hello"
    Selection.EndKey wdStory
    Selection.TypeParagraph
    Selection.TypeText "World"
    YC05 = ActiveDocument.Paragraphs.Count & "|" & Selection.Start & "|" & ActiveDocument.Paragraphs(2).Range.Text
End Function

Function YC06() As String
    Dim r As Range
    ActiveDocument.Content.Text = "x y x"
    Set r = ActiveDocument.Range(2, 3)
    r.Find.Execute FindText:="x", Wrap:=wdFindContinue
    YC06 = r.Start & "|" & r.End
End Function

Function YC07() As String
    Dim t As Table
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Range(0, 0), 3, 1)
    YC07 = ActiveDocument.Sentences.Count & "|" & ActiveDocument.Paragraphs.Count & "|" & ActiveDocument.Words.Count
End Function

Function YC08() As String
    ActiveDocument.Content.Text = "abc"
    ActiveDocument.Content.Font.Size = 13.75
    YC08 = ActiveDocument.Content.Font.Size
End Function

Function YC09() As String
    ActiveDocument.Content.Text = "a   b"
    ActiveDocument.Range(1, 4).Font.Italic = True
    YC09 = ActiveDocument.Range(1, 4).Font.Italic & "|" & ActiveDocument.Range(0, 5).Font.Italic
End Function

Function YC10() As String
    Dim r As Range
    ActiveDocument.Content.Text = "Hello world"
    Set r = ActiveDocument.Range(0, 5)
    r.Collapse wdCollapseEnd
    r.Text = "!"
    YC10 = ActiveDocument.Content.Text & "|" & r.Start & "|" & r.End
End Function
