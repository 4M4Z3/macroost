Option Explicit

' Each case starts from a freshly opened document holding one empty paragraph and returns a
' string describing what happened. Chr(13) is the paragraph mark the host counts in every range.

Function B01() As String
    ActiveDocument.Content.Text = "Hello"
    B01 = ActiveDocument.Content.Text & "|" & ActiveDocument.Content.End & "|" & ActiveDocument.Paragraphs.Count
End Function

Function B02() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.InsertAfter "abc"
    d.Content.InsertParagraphAfter
    d.Content.InsertAfter "def"
    B02 = d.Content.Text & "|" & d.Paragraphs.Count & "|" & d.Content.End
End Function

Function B03() As String
    ActiveDocument.Content.Text = "Hello world"
    B03 = ActiveDocument.Range(0, 5).Text & "|" & ActiveDocument.Range(6, 11).Text
End Function

Function B04() As String
    Dim r As Range
    ActiveDocument.Content.Text = "Hello world"
    Set r = ActiveDocument.Range(6, 11)
    B04 = r.Start & "|" & r.End & "|" & Len(r.Text)
End Function

Function B05() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "Hello world. How are you?"
    B05 = d.Characters.Count & "|" & d.Words.Count & "|" & d.Sentences.Count & "|" & d.Paragraphs.Count
End Function

Function B06() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "Hello world. How are you?"
    B06 = "[" & d.Words(1).Text & "][" & d.Words(2).Text & "][" & d.Words(3).Text & "][" & d.Words(4).Text & "]"
End Function

Function B07() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "abc"
    B07 = d.Characters(1).Text & "|" & Asc(d.Characters(d.Characters.Count).Text) & "|" & d.Characters.Count
End Function

Function B08() As String
    ActiveDocument.Content.Text = "world"
    ActiveDocument.Content.InsertBefore "Hello "
    B08 = ActiveDocument.Content.Text
End Function

Function B09() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "One" & vbCr & "Two" & vbCr & "Three"
    B09 = d.Paragraphs.Count & "|" & d.Paragraphs(2).Range.Text & "|" & d.Paragraphs(2).Range.Start & "|" & d.Paragraphs(2).Range.End
End Function

Function B10() As String
    Dim r As Range, s As String
    ActiveDocument.Content.Text = "Hello world"
    Set r = ActiveDocument.Range(2, 7)
    r.Collapse wdCollapseStart
    s = r.Start & "," & r.End
    Set r = ActiveDocument.Range(2, 7)
    r.Collapse wdCollapseEnd
    B10 = s & "|" & r.Start & "," & r.End
End Function

Function B11() As String
    Dim r As Range, n As Long
    ActiveDocument.Content.Text = "Hello world"
    Set r = ActiveDocument.Range(0, 2)
    n = r.MoveEnd(wdCharacter, 3)
    B11 = n & "|" & r.End & "|" & r.Text
End Function

Function B12() As String
    ActiveDocument.Content.Text = "Hello cruel world"
    ActiveDocument.Range(5, 11).Delete
    B12 = ActiveDocument.Content.Text & "|" & ActiveDocument.Content.End
End Function

Function B13() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "Hello world, again." & vbCr & "Second line here"
    B13 = d.ComputeStatistics(wdStatisticWords) & "|" & d.ComputeStatistics(wdStatisticCharacters) & "|" & d.ComputeStatistics(wdStatisticParagraphs) & "|" & d.Content.ComputeStatistics(wdStatisticWords)
End Function

Function B14() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Range(0, 5).Text = "Goodbye"
    B14 = ActiveDocument.Content.Text & "|" & ActiveDocument.Content.End
End Function

Function B15() As String
    Dim r As Range
    ActiveDocument.Content.Text = "abc"
    Set r = ActiveDocument.Range(0, 3)
    r.InsertAfter "def"
    B15 = r.Text & "|" & r.Start & "|" & r.End
End Function

Function B16() As String
    Dim r As Range
    ActiveDocument.Content.Text = "abc"
    Set r = ActiveDocument.Range(1, 2)
    r.InsertBefore "X"
    B16 = r.Text & "|" & r.Start & "|" & r.End & "|" & ActiveDocument.Content.Text
End Function

Function B17() As String
    Dim r As Range
    ActiveDocument.Content.Text = "abc"
    Set r = ActiveDocument.Range(0, 3)
    r.InsertParagraphAfter
    B17 = Len(r.Text) & "|" & ActiveDocument.Paragraphs.Count & "|" & ActiveDocument.Content.Text
End Function

Function B18() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "First" & vbCr & "Middle" & vbCr & "Last"
    B18 = d.Paragraphs.First.Range.Text & "|" & d.Paragraphs.Last.Range.Text
End Function

Function B19() As String
    Dim p As Paragraph, s As String
    ActiveDocument.Content.Text = "a" & vbCr & "bb" & vbCr & "ccc"
    For Each p In ActiveDocument.Paragraphs
        s = s & Len(p.Range.Text) & ";"
    Next
    B19 = s
End Function

Function B20() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "Hello, world! Bye."
    B20 = d.Words.Count & "|" & d.ComputeStatistics(wdStatisticWords) & "|" & d.Range(0, 6).Words.Count
End Function

Function B21() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "One. Two! Three? Four"
    B21 = d.Sentences.Count & "|[" & d.Sentences(2).Text & "]|[" & d.Sentences(4).Text & "]"
End Function

Function B22() As String
    ActiveDocument.Content.InsertAfter "a" & vbCr & "b" & vbCr & "c"
    B22 = ActiveDocument.Paragraphs.Count & "|" & ActiveDocument.Content.End
End Function

Function B23() As String
    Dim r As Range, s As String
    ActiveDocument.Content.Text = "abcdef"
    Set r = ActiveDocument.Range(1, 3)
    r.SetRange 2, 5
    s = r.Text
    r.Start = 0
    B23 = s & "|" & r.Text & "|" & r.Start & "|" & r.End
End Function

Function B24() As String
    Dim r As Range
    ActiveDocument.Content.Text = "abc"
    Set r = ActiveDocument.Range(0, 1000)
    B24 = r.Start & "|" & r.End
End Function

Function B25() As String
    Dim r As Range
    ActiveDocument.Content.Text = "abcdef"
    Set r = ActiveDocument.Range(4, 2)
    B25 = r.Start & "|" & r.End & "|" & r.Text
End Function

Function B26() As String
    Dim r As Range, r2 As Range
    ActiveDocument.Content.Text = "abcdef"
    Set r = ActiveDocument.Range(0, 3)
    Set r2 = r.Duplicate
    r2.End = 5
    B26 = r.Text & "|" & r2.Text
End Function

Function B27() As String
    Dim d As Document
    Set d = ActiveDocument
    B27 = d.Characters.Count & "|" & d.Words.Count & "|" & d.Sentences.Count & "|" & d.Paragraphs.Count & "|" & d.ComputeStatistics(wdStatisticWords)
End Function

Function B28() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "a  b   c"
    B28 = d.Words.Count & "|[" & d.Words(1).Text & "]|[" & d.Words(2).Text & "]"
End Function

Function B29() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "Para one" & vbCr & "Para two"
    B29 = d.Paragraphs(1).Range.Characters.Count & "|" & d.Paragraphs(2).Range.Words.Count & "|" & d.Paragraphs(1).Range.Characters.First.Text & "|" & (d.Paragraphs(2).Range.Characters.Last.Text = vbCr)
End Function

Function B30() As String
    Dim r As Range
    ActiveDocument.Content.Text = "abc"
    Set r = ActiveDocument.Range(1, 1)
    B30 = "[" & r.Text & "]|" & r.Start & "|" & r.End & "|" & r.Characters.Count
End Function

Function B31() As String
    Dim r As Range, n As Long
    ActiveDocument.Content.Text = "abcdef"
    Set r = ActiveDocument.Range(1, 3)
    n = r.MoveStart(wdCharacter, -1)
    B31 = n & "|" & r.Start & "|" & r.End & "|" & r.Text
End Function

Function B32() As String
    Dim r As Range, n As Long
    ActiveDocument.Content.Text = "abc"
    Set r = ActiveDocument.Range(0, 2)
    n = r.MoveEnd(wdCharacter, 10)
    B32 = n & "|" & r.End
End Function

Function B33() As String
    Dim r As Range
    ActiveDocument.Content.Text = "Hello big world"
    Set r = ActiveDocument.Range(7, 7)
    r.Expand wdWord
    B33 = "[" & r.Text & "]|" & r.Start & "|" & r.End
End Function

Function B34() As String
    Dim r As Range
    ActiveDocument.Content.Text = "First para" & vbCr & "Second para"
    Set r = ActiveDocument.Range(13, 13)
    r.Expand wdParagraph
    B34 = "[" & r.Text & "]|" & r.Start & "|" & r.End
End Function

Function B35() As String
    ActiveDocument.Content.Text = ""
    B35 = Len(ActiveDocument.Content.Text) & "|" & ActiveDocument.Paragraphs.Count & "|" & Asc(ActiveDocument.Content.Text)
End Function

Function B36() As String
    Dim r As Range
    ActiveDocument.Content.Text = "Hello world"
    Set r = ActiveDocument.Range(6, 11)
    ActiveDocument.Range(0, 0).InsertBefore "Big "
    B36 = r.Text & "|" & r.Start & "|" & r.End
End Function

Function B37() As String
    ActiveDocument.Content.Text = "abc" & vbCr
    B37 = ActiveDocument.Paragraphs.Count & "|" & Len(ActiveDocument.Content.Text)
End Function

Function B38() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "x" & vbTab & "y"
    B38 = d.Words.Count & "|" & d.Characters.Count & "|" & d.ComputeStatistics(wdStatisticWords)
End Function

Function B39() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "It's 3.5 dollars"
    B39 = d.Words.Count & "|[" & d.Words(1).Text & "]|[" & d.Words(2).Text & "]"
End Function

Function B40() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "One" & vbCr & "Two"
    d.Paragraphs(1).Range.Delete
    B40 = d.Content.Text & "|" & d.Paragraphs.Count
End Function
