Option Explicit

Function S01() As String
    Selection.TypeText "Hello"
    S01 = Selection.Start & "|" & Selection.End & "|" & ActiveDocument.Content.Text
End Function

Function S02() As String
    Selection.TypeText "Hello"
    Selection.HomeKey wdStory
    Selection.TypeText "X"
    S02 = ActiveDocument.Content.Text & "|" & Selection.Start
End Function

Function S03() As String
    ActiveDocument.Content.Text = "abc"
    Selection.EndKey wdStory
    Selection.TypeText "Z"
    S03 = ActiveDocument.Content.Text & "|" & Selection.Start
End Function

Function S04() As String
    Dim n As Long
    ActiveDocument.Content.Text = "abcdef"
    Selection.HomeKey wdStory
    n = Selection.MoveRight(wdCharacter, 2)
    S04 = n & "|" & Selection.Start & "|" & Selection.End
End Function

Function S05() As String
    ActiveDocument.Content.Text = "abcdef"
    Selection.HomeKey wdStory
    Selection.MoveRight wdCharacter, 1
    Selection.MoveRight wdCharacter, 3, wdExtend
    S05 = Selection.Text & "|" & Selection.Start & "|" & Selection.End
End Function

Function S06() As String
    Dim n As Long
    ActiveDocument.Content.Text = "abcdef"
    Selection.EndKey wdStory
    n = Selection.MoveLeft(wdCharacter, 2)
    S06 = n & "|" & Selection.Start
End Function

Function S07() As String
    Selection.TypeText "One"
    Selection.TypeParagraph
    Selection.TypeText "Two"
    S07 = ActiveDocument.Paragraphs.Count & "|" & ActiveDocument.Content.Text & "|" & Selection.Start
End Function

Function S08() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Range(0, 5).Select
    Selection.Text = "Bye"
    S08 = ActiveDocument.Content.Text & "|" & Selection.Start & "|" & Selection.End & "|" & Selection.Text
End Function

Function S09() As String
    ActiveDocument.Content.Text = "Hello world"
    Selection.SetRange 6, 11
    S09 = Selection.Text & "|" & Selection.Range.Text & "|" & Selection.Range.Start
End Function

Function S10() As String
    Dim s As String
    ActiveDocument.Content.Text = "Hello world"
    Selection.SetRange 2, 8
    Selection.Collapse wdCollapseEnd
    s = Selection.Start & "," & Selection.End
    Selection.SetRange 2, 8
    Selection.Collapse
    S10 = s & "|" & Selection.Start & "," & Selection.End
End Function

Function S11() As String
    Dim n As Long
    ActiveDocument.Content.Text = "Hello big world"
    Selection.HomeKey wdStory
    n = Selection.MoveRight(wdWord, 1)
    S11 = n & "|" & Selection.Start & "|" & Selection.MoveRight(wdWord, 1, wdExtend) & "|[" & Selection.Text & "]"
End Function

Function S12() As String
    ActiveDocument.Content.Text = "abc"
    Selection.EndKey wdStory
    S12 = Selection.MoveRight(wdCharacter, 5) & "|" & Selection.Start
End Function

Function S13() As String
    ActiveDocument.Content.Text = "abcdef"
    Selection.SetRange 1, 3
    Selection.Delete
    S13 = ActiveDocument.Content.Text & "|" & Selection.Start & "|" & Selection.End
End Function

Function S14() As String
    ActiveDocument.Content.Text = "abcdef"
    Selection.SetRange 3, 3
    Selection.TypeBackspace
    S14 = ActiveDocument.Content.Text & "|" & Selection.Start
End Function

Function S15() As String
    ActiveDocument.Content.Text = "abcdef"
    Selection.SetRange 2, 2
    Selection.Delete
    S15 = ActiveDocument.Content.Text & "|" & Selection.Start
End Function

Function S16() As String
    Dim s As String
    ActiveDocument.Content.Text = "abc"
    Selection.SetRange 1, 1
    s = Selection.Type
    Selection.SetRange 0, 2
    S16 = s & "|" & Selection.Type
End Function

Function S17() As String
    ActiveDocument.Content.Text = "Hello big world. Two."
    Selection.SetRange 0, 15
    S17 = Selection.Words.Count & "|" & Selection.Characters.Count & "|" & Selection.Sentences.Count & "|" & Selection.Paragraphs.Count
End Function

Function S18() As String
    ActiveDocument.Content.Text = "abc"
    Selection.SetRange 0, 3
    Selection.InsertAfter "def"
    S18 = Selection.Text & "|" & Selection.Start & "|" & Selection.End
End Function

Function S19() As String
    ActiveDocument.Content.Text = "One" & vbCr & "Two"
    Selection.SetRange 5, 5
    S19 = Selection.Paragraphs(1).Range.Text & "|" & Selection.Paragraphs.Count
End Function

Function S20() As String
    Selection.TypeText "abc"
    Selection.HomeKey wdStory, wdExtend
    S20 = Selection.Text & "|" & Selection.Start & "|" & Selection.End
End Function

Function S21() As String
    ActiveDocument.Content.Text = "abcdef"
    Selection.SetRange 2, 4
    Selection.MoveRight wdCharacter, 1
    S21 = Selection.Start & "|" & Selection.End
End Function

Function S22() As String
    ActiveDocument.Content.Text = "abcdef"
    Selection.SetRange 2, 4
    Selection.MoveLeft wdCharacter, 1
    S22 = Selection.Start & "|" & Selection.End
End Function

Function S23() As String
    ActiveDocument.Content.Text = "abcdef"
    Selection.SetRange 1, 2
    Selection.TypeText "XYZ"
    S23 = ActiveDocument.Content.Text & "|" & Selection.Start & "|" & Selection.End
End Function

Function S24() As String
    ActiveDocument.Content.Text = "Hello"
    Selection.WholeStory
    S24 = Selection.Start & "|" & Selection.End & "|" & Selection.Text
End Function

Function S25() As String
    ActiveDocument.Content.Text = "Hello world"
    Selection.HomeKey wdStory
    Selection.Find.Execute FindText:="world"
    S25 = Selection.Start & "|" & Selection.End & "|" & Selection.Text
End Function

Function S26() As String
    Selection.TypeText "abc"
    Selection.SetRange 0, 3
    Selection.Font.Bold = True
    Selection.Collapse wdCollapseEnd
    Selection.TypeText "d"
    S26 = ActiveDocument.Range(3, 4).Font.Bold & "|" & ActiveDocument.Content.Text
End Function
