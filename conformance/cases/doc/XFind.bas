Option Explicit

Private Function V(s As String) As String
    V = Replace(Replace(Replace(s, Chr(13), "P"), Chr(7), "B"), vbTab, "T")
End Function

Function XF01() As String
    Dim r As Range, ok As Boolean
    ActiveDocument.Content.Text = "ab ab ab"
    Set r = ActiveDocument.Range(8, 8)
    ok = r.Find.Execute(FindText:="ab", Forward:=False)
    XF01 = ok & "|" & r.Start & "|" & r.End
End Function

Function XF02() As String
    Dim r As Range, ok As Boolean
    ActiveDocument.Content.Text = "ab cd ab"
    Set r = ActiveDocument.Range(4, 4)
    ok = r.Find.Execute(FindText:="ab", Wrap:=wdFindStop)
    XF02 = ok & "|" & r.Start & "|" & r.Find.Execute(FindText:="cd", Wrap:=wdFindContinue) & "|" & r.Start
End Function

Function XF03() As String
    ActiveDocument.Content.Text = "x1 y22 z333"
    ActiveDocument.Content.Find.Execute FindText:="^#", ReplaceWith:="#", Replace:=wdReplaceAll
    XF03 = ActiveDocument.Content.Text
End Function

Function XF04() As String
    ActiveDocument.Content.Text = "cat hat"
    ActiveDocument.Content.Find.Execute FindText:="^?at", ReplaceWith:="[^&]", Replace:=wdReplaceAll
    XF04 = ActiveDocument.Content.Text
End Function

Function XF05() As String
    ActiveDocument.Content.Text = "a1b2"
    ActiveDocument.Content.Find.Execute FindText:="^$", ReplaceWith:="L", Replace:=wdReplaceAll
    XF05 = ActiveDocument.Content.Text
End Function

Function XF06() As String
    ActiveDocument.Content.Text = "Hello, hello. HELLO! hello-world"
    ActiveDocument.Content.Find.Execute FindText:="hello", MatchWholeWord:=True, ReplaceWith:="X", Replace:=wdReplaceAll
    XF06 = ActiveDocument.Content.Text
End Function

Function XF07() As String
    Dim n As Long, r As Range
    ActiveDocument.Content.Text = "aaaa"
    Set r = ActiveDocument.Content
    r.Find.Text = "aa"
    Do While r.Find.Execute
        n = n + 1
        r.Collapse wdCollapseEnd
    Loop
    XF07 = n
End Function

Function XF08() As String
    ActiveDocument.Content.Text = "aaaa"
    ActiveDocument.Content.Find.Execute FindText:="aa", ReplaceWith:="a", Replace:=wdReplaceAll
    XF08 = ActiveDocument.Content.Text
End Function

Function XF09() As String
    ActiveDocument.Content.Text = "a b"
    ActiveDocument.Content.Find.Execute FindText:=" ", ReplaceWith:="^t", Replace:=wdReplaceAll
    XF09 = V(ActiveDocument.Content.Text)
End Function

Function XF10() As String
    ActiveDocument.Content.Text = "one" & vbCr & vbCr & "two"
    ActiveDocument.Content.Find.Execute FindText:="^p^p", ReplaceWith:="^p", Replace:=wdReplaceAll
    XF10 = V(ActiveDocument.Content.Text) & "|" & ActiveDocument.Paragraphs.Count
End Function

Function XF11() As String
    Dim r As Range
    ActiveDocument.Content.Text = "red green red"
    Set r = ActiveDocument.Content
    With r.Find
        .Text = "red"
        .Replacement.Text = "blue"
        .Execute Replace:=wdReplaceOne
    End With
    XF11 = ActiveDocument.Content.Text & "|" & r.Start & "|" & r.End & "|" & r.Text
End Function

Function XF12() As String
    Dim r As Range
    ActiveDocument.Content.Text = "red green red"
    Set r = ActiveDocument.Content
    r.Find.Execute FindText:="red", ReplaceWith:="blue", Replace:=wdReplaceAll
    XF12 = r.Start & "|" & r.End & "|" & r.Text
End Function

Function XF13() As String
    ActiveDocument.Content.Text = "abc"
    XF13 = ActiveDocument.Content.Find.Execute(FindText:="") & "|" & ActiveDocument.Content.Find.Text
End Function

Function XF14() As String
    Dim f As Find
    ActiveDocument.Content.Text = "abc"
    Set f = ActiveDocument.Content.Find
    XF14 = f.Forward & "|" & f.Wrap & "|" & f.MatchCase & "|" & f.MatchWholeWord & "|" & f.Format & "|[" & f.Text & "]|[" & f.Replacement.Text & "]"
End Function

Function XF15() As String
    Dim r As Range
    ActiveDocument.Content.Text = "abc abc"
    Set r = ActiveDocument.Content
    r.Find.Execute FindText:="b", MatchCase:=True
    XF15 = r.Find.MatchCase & "|" & r.Find.Text & "|" & r.Start
End Function

Function XF16() As String
    ActiveDocument.Content.Text = "The THE the"
    ActiveDocument.Content.Find.Execute FindText:="THE", ReplaceWith:="a", Replace:=wdReplaceAll
    XF16 = ActiveDocument.Content.Text
End Function

Function XF17() As String
    ActiveDocument.Content.Text = "Cat"
    ActiveDocument.Content.Find.Execute FindText:="cat", ReplaceWith:="dOG", Replace:=wdReplaceAll
    XF17 = ActiveDocument.Content.Text
End Function

Function XF18() As String
    ActiveDocument.Content.Text = "Abc abc"
    ActiveDocument.Range(0, 3).Font.Bold = True
    ActiveDocument.Content.Find.Execute FindText:="abc", ReplaceWith:="xyzw", Replace:=wdReplaceAll
    XF18 = ActiveDocument.Content.Text & "|" & ActiveDocument.Range(0, 4).Font.Bold & "|" & ActiveDocument.Range(5, 9).Font.Bold
End Function

Function XF19() As String
    Dim r As Range, ok As Boolean
    ActiveDocument.Content.Text = "start middle end"
    Set r = ActiveDocument.Range(6, 12)
    ok = r.Find.Execute(FindText:="end")
    XF19 = ok & "|" & r.Start & "|" & r.End
End Function

Function XF20() As String
    ActiveDocument.Content.Text = "x^y"
    ActiveDocument.Content.Find.Execute FindText:="^^", ReplaceWith:="-", Replace:=wdReplaceAll
    XF20 = ActiveDocument.Content.Text
End Function

Function XF21() As String
    ActiveDocument.Content.Text = "Tab" & vbTab & "here"
    XF21 = ActiveDocument.Content.Find.Execute(FindText:="^t") & "|" & ActiveDocument.Content.Find.Execute(FindText:="^w")
End Function

Function XF22() As String
    Dim r As Range
    ActiveDocument.Content.Text = "aXbXc"
    Set r = ActiveDocument.Range(1, 4)
    r.Find.Execute FindText:="X", ReplaceWith:="--", Replace:=wdReplaceAll
    XF22 = ActiveDocument.Content.Text & "|" & r.Start & "|" & r.End
End Function

Function XF23() As String
    ActiveDocument.Content.Text = "abc"
    Selection.HomeKey wdStory
    With Selection.Find
        .Text = "c"
        .Execute
    End With
    XF23 = Selection.Find.Found & "|" & Selection.Start & "|" & Selection.Text
End Function

Function XF24() As String
    Dim r As Range
    ActiveDocument.Content.Text = "one two"
    Set r = ActiveDocument.Content
    XF24 = r.Find.Execute("TWO", False, True) & "|" & r.Text
End Function
