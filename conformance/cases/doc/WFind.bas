Option Explicit

Function R01() As String
    Dim r As Range, ok As Boolean
    ActiveDocument.Content.Text = "Hello world"
    Set r = ActiveDocument.Content
    ok = r.Find.Execute(FindText:="world")
    R01 = ok & "|" & r.Start & "|" & r.End & "|" & r.Text
End Function

Function R02() As String
    Dim r As Range, ok As Boolean
    ActiveDocument.Content.Text = "Hello world"
    Set r = ActiveDocument.Range(2, 8)
    ok = r.Find.Execute(FindText:="zebra")
    R02 = ok & "|" & r.Start & "|" & r.End
End Function

Function R03() As String
    ActiveDocument.Content.Text = "banana bandana"
    ActiveDocument.Content.Find.Execute FindText:="an", ReplaceWith:="AN", Replace:=wdReplaceAll
    R03 = ActiveDocument.Content.Text
End Function

Function R04() As String
    Dim s As String
    ActiveDocument.Content.Text = "Cat cat CAT"
    ActiveDocument.Content.Find.Execute FindText:="cat", MatchCase:=True, ReplaceWith:="dog", Replace:=wdReplaceAll
    s = ActiveDocument.Content.Text
    ActiveDocument.Content.Text = "Cat cat CAT"
    ActiveDocument.Content.Find.Execute FindText:="cat", MatchCase:=False, ReplaceWith:="dog", Replace:=wdReplaceAll
    R04 = s & "|" & ActiveDocument.Content.Text
End Function

Function R05() As String
    ActiveDocument.Content.Text = "the theme other the"
    ActiveDocument.Content.Find.Execute FindText:="the", MatchWholeWord:=True, ReplaceWith:="a", Replace:=wdReplaceAll
    R05 = ActiveDocument.Content.Text
End Function

Function R06() As String
    ActiveDocument.Content.Text = "x1 x2 x3"
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Text = "x"
        .Replacement.ClearFormatting
        .Replacement.Text = "yy"
        .Forward = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
    R06 = ActiveDocument.Content.Text
End Function

Function R07() As String
    Dim r As Range, n As Long, s As String
    ActiveDocument.Content.Text = "a-b-c-d"
    Set r = ActiveDocument.Content
    With r.Find
        .Text = "-"
        .Forward = True
        .Wrap = wdFindStop
    End With
    Do While r.Find.Execute
        n = n + 1
        s = s & r.Start & ","
        r.Collapse wdCollapseEnd
    Loop
    R07 = n & "|" & s
End Function

Function R08() As String
    ActiveDocument.Content.Text = "one one one"
    ActiveDocument.Content.Find.Execute FindText:="one", ReplaceWith:="two", Replace:=wdReplaceOne
    R08 = ActiveDocument.Content.Text
End Function

Function R09() As String
    ActiveDocument.Content.Text = "a b a"
    ActiveDocument.Content.Find.Execute FindText:="a", ReplaceWith:="alpha", Replace:=wdReplaceAll
    R09 = ActiveDocument.Content.Text & "|" & ActiveDocument.Content.End
End Function

Function R10() As String
    ActiveDocument.Content.Text = "aaa aaa aaa"
    ActiveDocument.Range(4, 7).Find.Execute FindText:="a", ReplaceWith:="b", Replace:=wdReplaceAll
    R10 = ActiveDocument.Content.Text
End Function

Function R11() As String
    Dim r As Range, s As String
    ActiveDocument.Content.Text = "find me"
    Set r = ActiveDocument.Content
    r.Find.Execute FindText:="me"
    s = r.Find.Found
    Set r = ActiveDocument.Content
    r.Find.Execute FindText:="you"
    R11 = s & "|" & r.Find.Found
End Function

Function R12() As String
    ActiveDocument.Content.Text = "remove the the duplicate"
    ActiveDocument.Content.Find.Execute FindText:="the ", ReplaceWith:="", Replace:=wdReplaceAll
    R12 = ActiveDocument.Content.Text
End Function

Function R13() As String
    ActiveDocument.Content.Text = "one" & vbCr & "two" & vbCr & "three"
    ActiveDocument.Content.Find.Execute FindText:="^p", ReplaceWith:=" ", Replace:=wdReplaceAll
    R13 = ActiveDocument.Content.Text & "|" & ActiveDocument.Paragraphs.Count
End Function

Function R14() As String
    ActiveDocument.Content.Text = "a" & vbTab & "b" & vbTab & "c"
    ActiveDocument.Content.Find.Execute FindText:="^t", ReplaceWith:=";", Replace:=wdReplaceAll
    R14 = ActiveDocument.Content.Text
End Function

Function R15() As String
    ActiveDocument.Content.Text = "Hello hello HELLO"
    ActiveDocument.Content.Find.Execute FindText:="hello", ReplaceWith:="bye", Replace:=wdReplaceAll
    R15 = ActiveDocument.Content.Text
End Function

Function R16() As String
    Dim ok As Boolean
    ActiveDocument.Content.Text = "abc"
    ok = ActiveDocument.Content.Find.Execute(FindText:="b", ReplaceWith:="B", Replace:=wdReplaceAll)
    R16 = ok & "|" & ActiveDocument.Content.Find.Execute(FindText:="zz", ReplaceWith:="B", Replace:=wdReplaceAll) & "|" & ActiveDocument.Content.Text
End Function

Function R17() As String
    Dim r As Range
    ActiveDocument.Content.Text = "Word one word two"
    Set r = ActiveDocument.Content
    r.Find.MatchCase = True
    r.Find.Execute FindText:="word"
    R17 = r.Start & "|" & r.End
End Function

Function R18() As String
    Dim r As Range, ok As Boolean
    ActiveDocument.Content.Text = "Hello world"
    Set r = ActiveDocument.Range(0, 4)
    ok = r.Find.Execute(FindText:="world", Wrap:=wdFindStop)
    R18 = ok & "|" & r.Start & "|" & r.End
End Function

Function R19() As String
    ActiveDocument.Content.Text = "cat category concat cat."
    ActiveDocument.Content.Find.Execute FindText:="cat", MatchWholeWord:=True, MatchCase:=True, ReplaceWith:="dog", Replace:=wdReplaceAll
    R19 = ActiveDocument.Content.Text
End Function

Function R20() As String
    ActiveDocument.Content.Text = "a.b.c"
    ActiveDocument.Content.Find.Execute FindText:=".", ReplaceWith:="^p", Replace:=wdReplaceAll
    R20 = ActiveDocument.Paragraphs.Count & "|" & ActiveDocument.Content.Text
End Function

Function R21() As String
    Dim r As Range, ok As Boolean
    ActiveDocument.Content.Text = "Hello World"
    Set r = ActiveDocument.Content
    ok = r.Find.Execute("WORLD", True)
    R21 = ok & "|" & r.Find.Execute("WORLD", False) & "|" & r.Text
End Function
