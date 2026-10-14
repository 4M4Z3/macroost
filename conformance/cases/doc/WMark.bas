Option Explicit

Function K01() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "bm1", ActiveDocument.Range(0, 5)
    K01 = ActiveDocument.Bookmarks.Exists("bm1") & "|" & ActiveDocument.Bookmarks.Exists("nope") & "|" & ActiveDocument.Bookmarks.Count
End Function

Function K02() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "Greeting", ActiveDocument.Range(0, 5)
    K02 = ActiveDocument.Bookmarks("Greeting").Range.Text & "|" & ActiveDocument.Bookmarks("greeting").Name & "|" & ActiveDocument.Bookmarks(1).Start & "|" & ActiveDocument.Bookmarks(1).End
End Function

Function K03() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "w", ActiveDocument.Range(6, 11)
    ActiveDocument.Range(0, 0).InsertBefore "Big "
    K03 = ActiveDocument.Bookmarks("w").Start & "|" & ActiveDocument.Bookmarks("w").End & "|" & ActiveDocument.Bookmarks("w").Range.Text
End Function

Function K04() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "bm", ActiveDocument.Range(0, 5)
    ActiveDocument.Bookmarks("bm").Delete
    K04 = ActiveDocument.Bookmarks.Exists("bm") & "|" & ActiveDocument.Bookmarks.Count & "|" & ActiveDocument.Content.Text
End Function

Function K05() As String
    ActiveDocument.Content.Text = "aaa bbb ccc"
    ActiveDocument.Bookmarks.Add "zeta", ActiveDocument.Range(0, 3)
    ActiveDocument.Bookmarks.Add "alpha", ActiveDocument.Range(8, 11)
    ActiveDocument.Bookmarks.Add "mid", ActiveDocument.Range(4, 7)
    K05 = ActiveDocument.Bookmarks(1).Name & "," & ActiveDocument.Bookmarks(2).Name & "," & ActiveDocument.Bookmarks(3).Name
End Function

Function K06() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "bm", ActiveDocument.Range(0, 5)
    ActiveDocument.Bookmarks("bm").Range.Text = "Howdy"
    K06 = ActiveDocument.Bookmarks.Exists("bm") & "|" & ActiveDocument.Content.Text
End Function

Function K07() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "1bad", ActiveDocument.Range(0, 5)
    K07 = "added"
End Function

Function K08() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "bm", ActiveDocument.Range(0, 5)
    ActiveDocument.Bookmarks.Add "bm", ActiveDocument.Range(6, 11)
    K08 = ActiveDocument.Bookmarks.Count & "|" & ActiveDocument.Bookmarks("bm").Range.Text
End Function

Function K09() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "here", ActiveDocument.Range(3, 3)
    K09 = ActiveDocument.Bookmarks("here").Empty & "|" & ActiveDocument.Bookmarks("here").Start & "|" & ActiveDocument.Bookmarks.Exists("HERE")
End Function

Function K10() As String
    Dim b As Bookmark, s As String
    ActiveDocument.Content.Text = "aaa bbb ccc"
    ActiveDocument.Bookmarks.Add "c2", ActiveDocument.Range(8, 11)
    ActiveDocument.Bookmarks.Add "c1", ActiveDocument.Range(0, 3)
    For Each b In ActiveDocument.Bookmarks
        s = s & b.Name & ";"
    Next
    K10 = s
End Function

Function K11() As String
    ActiveDocument.Content.Text = "Hello world"
    K11 = ActiveDocument.Bookmarks("missing").Name
End Function

Function K12() As String
    Dim r As Range
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "bm", ActiveDocument.Range(6, 11)
    Set r = ActiveDocument.Bookmarks("bm").Range
    r.InsertAfter "!"
    K12 = ActiveDocument.Bookmarks("bm").Range.Text & "|" & r.Text
End Function

Function K13() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Bookmarks.Add "bm", ActiveDocument.Range(0, 11)
    ActiveDocument.Range(2, 4).Delete
    K13 = ActiveDocument.Bookmarks("bm").End & "|" & ActiveDocument.Bookmarks("bm").Range.Text
End Function

Function K14() As String
    ActiveDocument.Content.Text = "Hello world"
    ActiveDocument.Range(0, 5).Bookmarks.Add "x"
    K14 = ActiveDocument.Bookmarks.Count & "|" & ActiveDocument.Range(0, 3).Bookmarks.Count & "|" & ActiveDocument.Range(6, 11).Bookmarks.Count
End Function
