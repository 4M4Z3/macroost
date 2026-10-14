Option Explicit

Function D01() As String
    D01 = Documents.Count & "|" & ActiveDocument.Name & "|" & (ActiveDocument Is Documents(1)) & "|" & Documents(ActiveDocument.Name).Name
End Function

Function D02() As String
    Dim n As Long, d As Document
    n = Documents.Count
    Set d = Documents.Add
    D02 = n & "|" & Documents.Count & "|" & (ActiveDocument Is d) & "|" & Len(d.Content.Text)
End Function

Function D03() As String
    ActiveDocument.Content.Text = "abc"
    D03 = ActiveDocument.Range.Text & "|" & ActiveDocument.Range.End & "|" & ActiveDocument.Range(1).Text
End Function

Function D04() As String
    D04 = TypeName(ActiveDocument) & "|" & TypeName(Documents) & "|" & TypeName(Selection) & "|" & TypeName(Application) & "|" & TypeName(ActiveDocument.Paragraphs) & "|" & TypeName(ActiveDocument.Words)
End Function

Function D05() As String
    D05 = ThisDocument.Name & "|" & (ThisDocument Is ActiveDocument)
End Function

Function D06() As String
    Dim d As Document
    Set d = Documents.Add
    d.Content.Text = "new"
    D06 = d.Content.Text & "|" & Documents(2).Content.Text & "|" & ActiveDocument.Content.Text
End Function

Function D07() As String
    ActiveDocument.Content.Text = "abc"
    D07 = Application.ActiveDocument.Content.Text & "|" & Application.Documents.Count & "|" & Application.Selection.Start
End Function

Function D08() As String
    Dim w As Range, s As String
    ActiveDocument.Content.Text = "red green blue"
    For Each w In ActiveDocument.Words
        s = s & "[" & w.Text & "]"
    Next
    D08 = s
End Function

Function D09() As String
    Dim c As Range, s As String
    ActiveDocument.Content.Text = "ab"
    For Each c In ActiveDocument.Characters
        s = s & Asc(c.Text) & ","
    Next
    D09 = s
End Function

Function D10() As String
    Dim s As Range, out As String
    ActiveDocument.Content.Text = "One. Two."
    For Each s In ActiveDocument.Sentences
        out = out & "[" & s.Text & "]"
    Next
    D10 = out
End Function

Function D11() As String
    ActiveDocument.Content.Text = "abc"
    D11 = ActiveDocument.Paragraphs(5).Range.Text
End Function

Function D12() As String
    ActiveDocument.Content.Text = "abc"
    D12 = ActiveDocument.Words(10).Text
End Function

Function D13() As String
    ActiveDocument.Content.Text = "Hello world"
    D13 = ActiveDocument.Content.Characters(1).Start & "|" & ActiveDocument.Content.Words(2).Start & "|" & ActiveDocument.Content.Words(2).End
End Function

Function D14() As String
    Dim r As Range
    ActiveDocument.Content.Text = "abc"
    Set r = ActiveDocument.Content
    r.Text = "xyz"
    D14 = r.Text & "|" & r.Start & "|" & r.End & "|" & ActiveDocument.Content.Text
End Function

Function D15() As String
    Dim r As Range
    ActiveDocument.Content.Text = "Hello world"
    Set r = ActiveDocument.Range(0, 5)
    r.Text = "Hi"
    D15 = r.Text & "|" & r.Start & "|" & r.End
End Function

Function D16() As String
    Dim r As Range
    ActiveDocument.Content.Text = "Hello world"
    Set r = ActiveDocument.Range(6, 6)
    r.Text = "big "
    D16 = r.Text & "|" & r.Start & "|" & r.End & "|" & ActiveDocument.Content.Text
End Function
