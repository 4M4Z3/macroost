Option Explicit

' Second batch: edge behaviors of ranges, units and editing.

Private Function V(s As String) As String
    V = Replace(Replace(Replace(Replace(s, Chr(13), "P"), Chr(7), "B"), Chr(11), "L"), vbTab, "T")
End Function

Function XA01() As String
    Dim r As Range
    ActiveDocument.Content.Text = "One" & vbCr & "Two"
    Set r = ActiveDocument.Paragraphs(1).Range
    r.Collapse wdCollapseEnd
    XA01 = r.Start & "|" & r.End
End Function

Function XA02() As String
    Dim r As Range
    ActiveDocument.Content.Text = "Hello"
    Set r = ActiveDocument.Content
    r.Collapse wdCollapseEnd
    XA02 = r.Start & "|" & r.End
End Function

Function XA03() As String
    Dim r As Range, n As Long
    ActiveDocument.Content.Text = "Hello big world"
    Set r = ActiveDocument.Range(0, 0)
    n = r.Move(wdWord, 2)
    XA03 = n & "|" & r.Start & "|" & r.End
End Function

Function XA04() As String
    Dim r As Range, n As Long
    ActiveDocument.Content.Text = "Hello big world"
    Set r = ActiveDocument.Range(3, 3)
    n = r.MoveEnd(wdWord, 1)
    XA04 = n & "|" & r.Start & "|" & r.End & "|[" & r.Text & "]"
End Function

Function XA05() As String
    Dim r As Range, n As Long
    ActiveDocument.Content.Text = "Hello big world"
    Set r = ActiveDocument.Range(12, 12)
    n = r.MoveStart(wdWord, -1)
    XA05 = n & "|" & r.Start & "|" & r.End
End Function

Function XA06() As String
    ActiveDocument.Content.Text = "line1" & vbLf & "line2"
    XA06 = ActiveDocument.Paragraphs.Count & "|" & V(ActiveDocument.Content.Text)
End Function

Function XA07() As String
    ActiveDocument.Content.Text = "a" & vbCrLf & "b"
    XA07 = ActiveDocument.Paragraphs.Count & "|" & V(ActiveDocument.Content.Text) & "|" & ActiveDocument.Content.End
End Function

Function XA08() As String
    ActiveDocument.Content.Text = "a" & Chr(11) & "b"
    XA08 = ActiveDocument.Paragraphs.Count & "|" & ActiveDocument.Words.Count & "|" & ActiveDocument.Sentences.Count & "|" & V(ActiveDocument.Content.Text)
End Function

Function XA09() As String
    Dim r As Range
    ActiveDocument.Content.Text = "Hello big world"
    Set r = ActiveDocument.Range(2, 8)
    XA09 = r.Words.Count & "|[" & r.Words(1).Text & "]|[" & r.Words(r.Words.Count).Text & "]"
End Function

Function XA10() As String
    Dim r As Range
    ActiveDocument.Content.Text = "First para" & vbCr & "Second para"
    Set r = ActiveDocument.Range(3, 14)
    XA10 = r.Paragraphs.Count & "|[" & V(r.Paragraphs(1).Range.Text) & "]|[" & V(r.Paragraphs(2).Range.Text) & "]"
End Function

Function XA11() As String
    Dim r As Range
    ActiveDocument.Content.Text = "Hello world"
    Set r = ActiveDocument.Range(6, 11)
    ActiveDocument.Range(6, 6).InsertBefore "big "
    XA11 = r.Start & "|" & r.End & "|" & r.Text
End Function

Function XA12() As String
    Dim r As Range
    ActiveDocument.Content.Text = "Hello world"
    Set r = ActiveDocument.Range(0, 5)
    ActiveDocument.Range(5, 5).InsertAfter "!!"
    XA12 = r.Start & "|" & r.End & "|" & r.Text
End Function

Function XA13() As String
    Dim r As Range
    ActiveDocument.Content.Text = "Hello world"
    Set r = ActiveDocument.Range(3, 3)
    ActiveDocument.Range(3, 3).InsertAfter "XY"
    XA13 = r.Start & "|" & r.End
End Function

Function XA14() As String
    Dim r As Range
    ActiveDocument.Content.Text = "Hello world"
    Set r = ActiveDocument.Range(2, 9)
    ActiveDocument.Range(0, 4).Delete
    XA14 = r.Start & "|" & r.End & "|" & r.Text
End Function

Function XA15() As String
    Dim r As Range, n As Long
    ActiveDocument.Content.Text = "abcdef"
    Set r = ActiveDocument.Range(2, 2)
    n = r.Delete(wdCharacter, 2)
    XA15 = n & "|" & ActiveDocument.Content.Text
End Function

Function XA16() As String
    Dim r As Range, n As Long
    ActiveDocument.Content.Text = "Hello big world"
    Set r = ActiveDocument.Range(6, 6)
    n = r.Delete(wdWord, 1)
    XA16 = n & "|" & ActiveDocument.Content.Text
End Function

Function XA17() As String
    ActiveDocument.Content.Text = "abc"
    ActiveDocument.Content.Delete
    XA17 = Len(ActiveDocument.Content.Text) & "|" & ActiveDocument.Paragraphs.Count
End Function

Function XA18() As String
    Dim r As Range
    ActiveDocument.Content.Text = "One" & vbCr & "Two"
    Set r = ActiveDocument.Paragraphs(2).Range
    r.InsertParagraphBefore
    XA18 = ActiveDocument.Paragraphs.Count & "|" & V(ActiveDocument.Content.Text) & "|" & V(r.Text)
End Function

Function XA19() As String
    Dim r As Range
    ActiveDocument.Content.Text = "One" & vbCr & "Two"
    Set r = ActiveDocument.Paragraphs(1).Range
    r.InsertAfter "X"
    XA19 = V(ActiveDocument.Content.Text) & "|" & V(r.Text)
End Function

Function XA20() As String
    Dim r As Range
    ActiveDocument.Content.Text = "One" & vbCr & "Two"
    Set r = ActiveDocument.Paragraphs(2).Range
    r.InsertAfter "X"
    XA20 = V(ActiveDocument.Content.Text) & "|" & V(r.Text)
End Function

Function XA21() As String
    Dim r As Range
    ActiveDocument.Content.Text = "One" & vbCr & "Two"
    Set r = ActiveDocument.Paragraphs(1).Range
    r.Text = "Uno"
    XA21 = V(ActiveDocument.Content.Text) & "|" & ActiveDocument.Paragraphs.Count & "|" & V(r.Text)
End Function

Function XA22() As String
    Dim p As Paragraph
    ActiveDocument.Content.Text = "a" & vbCr & "b" & vbCr & "c"
    Set p = ActiveDocument.Paragraphs(2)
    XA22 = V(p.Next.Range.Text) & "|" & V(p.Previous.Range.Text) & "|" & (ActiveDocument.Paragraphs(3).Next Is Nothing)
End Function

Function XA23() As String
    Dim r As Range
    ActiveDocument.Content.Text = "One. Two. Three."
    Set r = ActiveDocument.Range(6, 6)
    r.Expand wdSentence
    XA23 = "[" & V(r.Text) & "]|" & r.Start & "|" & r.End
End Function

Function XA24() As String
    Dim r As Range, n As Long
    ActiveDocument.Content.Text = "abc"
    Set r = ActiveDocument.Range(1, 1)
    n = r.Expand(wdStory)
    XA24 = n & "|" & r.Start & "|" & r.End
End Function

Function XA25() As String
    ActiveDocument.Content.Text = "  leading"
    XA25 = ActiveDocument.Words.Count & "|[" & ActiveDocument.Words(1).Text & "]|" & ActiveDocument.ComputeStatistics(wdStatisticWords)
End Function

Function XA26() As String
    ActiveDocument.Content.Text = "Wait... what?! (yes) ""quoted"" end"
    XA26 = ActiveDocument.Words.Count & "|" & ActiveDocument.Sentences.Count & "|[" & ActiveDocument.Words(2).Text & "]|[" & ActiveDocument.Sentences(1).Text & "]"
End Function

Function XA27() As String
    ActiveDocument.Content.Text = "e-mail co-op 50% $5 a/b"
    XA27 = ActiveDocument.Words.Count & "|" & ActiveDocument.ComputeStatistics(wdStatisticWords) & "|[" & ActiveDocument.Words(2).Text & "]"
End Function

Function XA28() As String
    ActiveDocument.Content.Text = "Hi. " & vbCr & "There"
    XA28 = ActiveDocument.Sentences.Count & "|[" & V(ActiveDocument.Sentences(1).Text) & "]"
End Function

Function XA29() As String
    ActiveDocument.Content.Text = "abc"
    XA29 = ActiveDocument.Content.StoryLength & "|" & ActiveDocument.Range(0, 2).StoryLength & "|" & ActiveDocument.Content.Characters.Last.Start
End Function

Function XA30() As String
    Dim r As Range, r2 As Range
    ActiveDocument.Content.Text = "abcdef"
    Set r = ActiveDocument.Range(1, 4)
    Set r2 = ActiveDocument.Range(2, 3)
    XA30 = r2.InRange(r) & "|" & r.InRange(r2) & "|" & r.IsEqual(ActiveDocument.Range(1, 4))
End Function

Function XA31() As String
    Dim r As Range
    ActiveDocument.Content.Text = "abc"
    Set r = ActiveDocument.Range(0, 2)
    r.End = 1
    r.Start = 2
    XA31 = r.Start & "|" & r.End
End Function

Function XA32() As String
    Dim r As Range
    ActiveDocument.Content.Text = "abc def"
    Set r = ActiveDocument.Words(1)
    r.Text = "xyz"
    XA32 = ActiveDocument.Content.Text & "|" & r.Text
End Function

Function XA33() As String
    ActiveDocument.Content.Text = "x"
    ActiveDocument.Content.InsertAfter "y"
    ActiveDocument.Content.InsertBefore "w"
    ActiveDocument.Content.InsertParagraphBefore
    XA33 = V(ActiveDocument.Content.Text) & "|" & ActiveDocument.Paragraphs.Count
End Function

Function XA34() As String
    Dim r As Range
    ActiveDocument.Content.Text = "abc"
    Set r = ActiveDocument.Range(1, 2)
    r.InsertParagraph
    XA34 = V(ActiveDocument.Content.Text) & "|" & V(r.Text) & "|" & r.Start & "|" & r.End
End Function

Function XA35() As String
    Dim r As Range
    ActiveDocument.Content.Text = "Hello world"
    Set r = ActiveDocument.Content
    r.Collapse wdCollapseStart
    r.InsertAfter "A "
    XA35 = ActiveDocument.Content.Text & "|" & r.Text
End Function

Function XA36() As String
    ActiveDocument.Content.Text = "x"
    XA36 = ActiveDocument.Range(1, 1).Text = "" & "|" & ActiveDocument.Range(1, 2).Text = vbCr
End Function

Function XA37() As String
    ActiveDocument.Content.Text = "one two"
    ActiveDocument.Words(2).InsertBefore "big "
    XA37 = ActiveDocument.Content.Text & "|" & ActiveDocument.Words.Count
End Function

Function XA38() As String
    Dim i As Long, s As String
    ActiveDocument.Content.Text = "a b c"
    For i = ActiveDocument.Words.Count To 1 Step -1
        s = s & "[" & V(ActiveDocument.Words(i).Text) & "]"
    Next
    XA38 = s
End Function

Function XA39() As String
    Dim r As Range
    ActiveDocument.Content.Text = "abc"
    Set r = ActiveDocument.Range(1, 2)
    r.Text = ""
    XA39 = ActiveDocument.Content.Text & "|" & r.Start & "|" & r.End
End Function

Function XA40() As String
    ActiveDocument.Content.Text = "1234 5678"
    XA40 = ActiveDocument.Words.Count & "|" & ActiveDocument.ComputeStatistics(wdStatisticWords) & "|" & ActiveDocument.ComputeStatistics(wdStatisticCharactersWithSpaces)
End Function
