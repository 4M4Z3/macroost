Option Explicit

Private Function NewTable(rows As Long, cols As Long) As Table
    Set NewTable = ActiveDocument.Tables.Add(ActiveDocument.Range(0, 0), rows, cols)
End Function

Private Function CellText(t As Table, r As Long, c As Long) As String
    Dim s As String
    s = t.Cell(r, c).Range.Text
    CellText = Left(s, Len(s) - 2)
End Function

Function T01() As String
    Dim t As Table
    Set t = NewTable(2, 3)
    T01 = ActiveDocument.Tables.Count & "|" & t.Rows.Count & "|" & t.Columns.Count
End Function

Function T02() As String
    Dim t As Table, s As String
    Set t = NewTable(2, 2)
    t.Cell(1, 1).Range.Text = "A"
    s = t.Cell(1, 1).Range.Text
    T02 = Len(s) & "|" & Left(s, 1) & "|" & Asc(Mid(s, 2, 1)) & "|" & Asc(Mid(s, 3, 1))
End Function

Function T03() As String
    Dim t As Table, s As String, i As Long, out As String
    Set t = NewTable(2, 2)
    t.Cell(1, 1).Range.Text = "a"
    t.Cell(1, 2).Range.Text = "b"
    t.Cell(2, 1).Range.Text = "c"
    t.Cell(2, 2).Range.Text = "d"
    s = ActiveDocument.Content.Text
    For i = 1 To Len(s)
        out = out & Asc(Mid(s, i, 1)) & " "
    Next
    T03 = out
End Function

Function T04() As String
    Dim t As Table
    Set t = NewTable(2, 2)
    t.Rows.Add
    T04 = t.Rows.Count & "|" & t.Range.Cells.Count
End Function

Function T05() As String
    Dim t As Table
    Set t = NewTable(2, 2)
    T05 = t.Range.Start & "|" & t.Range.End & "|" & ActiveDocument.Content.End
End Function

Function T06() As String
    Dim t As Table
    Set t = NewTable(2, 3)
    T06 = ActiveDocument.Paragraphs.Count & "|" & ActiveDocument.Characters.Count & "|" & ActiveDocument.Words.Count
End Function

Function T07() As String
    Dim t As Table
    Set t = NewTable(2, 2)
    t.Cell(2, 2).Range.Text = "Hello"
    T07 = CellText(t, 2, 2) & "|" & Len(t.Cell(2, 2).Range.Text) & "|" & t.Cell(2, 2).Range.Start & "|" & t.Cell(2, 2).Range.End
End Function

Function T08() As String
    Dim t As Table, s As String, r As Long, c As Long
    Set t = NewTable(2, 3)
    For r = 1 To 2
        For c = 1 To 3
            s = s & t.Cell(r, c).Range.Start & "-" & t.Cell(r, c).Range.End & " "
        Next
    Next
    T08 = s
End Function

Function T09() As String
    Dim t As Table
    Set t = NewTable(1, 2)
    t.Cell(1, 1).Range.Text = "x"
    t.Cell(1, 1).Range.InsertAfter "y"
    T09 = CellText(t, 1, 1)
End Function

Function T10() As String
    Dim t As Table
    Set t = NewTable(3, 4)
    T10 = t.Rows(1).Cells.Count & "|" & t.Columns(2).Cells.Count & "|" & t.Range.Cells.Count
End Function

Function T11() As String
    Dim t As Table
    Set t = NewTable(2, 2)
    t.Columns.Add
    T11 = t.Columns.Count & "|" & t.Rows(1).Cells.Count
End Function

Function T12() As String
    Dim t As Table
    Set t = NewTable(2, 2)
    t.Delete
    T12 = ActiveDocument.Tables.Count & "|" & Len(ActiveDocument.Content.Text) & "|" & ActiveDocument.Paragraphs.Count
End Function

Function T13() As String
    Dim t As Table
    Set t = NewTable(2, 2)
    t.Cell(1, 2).Range.Text = "q"
    t.Cell(2, 1).Range.Text = "rs"
    T13 = Replace(Replace(t.Range.Text, Chr(13), "P"), Chr(7), "B") & "|" & t.Range.Characters.Count & "|" & t.Cell(2, 2).Range.Start
End Function

Function T14() As String
    Dim t As Table, s As String
    Set t = NewTable(1, 2)
    t.Cell(1, 1).Range.Text = "a"
    s = t.Rows(1).Range.Text
    T14 = Len(s) & "|" & Replace(Replace(s, Chr(13), "P"), Chr(7), "B")
End Function

Function T15() As String
    Dim t As Table
    ActiveDocument.Content.Text = "Before"
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Paragraphs(1).Range, 1, 1)
    T15 = ActiveDocument.Tables.Count & "|" & Len(ActiveDocument.Content.Text) & "|" & Replace(Replace(ActiveDocument.Content.Text, Chr(13), "P"), Chr(7), "B")
End Function

Function T16() As String
    Dim t As Table
    ActiveDocument.Content.Text = "Before" & vbCr & "After"
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Range(7, 7), 1, 2)
    t.Cell(1, 2).Range.Text = "z"
    T16 = Replace(Replace(ActiveDocument.Content.Text, Chr(13), "P"), Chr(7), "B") & "|" & ActiveDocument.Paragraphs.Count
End Function

Function T17() As String
    Dim t As Table
    Set t = NewTable(2, 2)
    T17 = TypeName(t) & "|" & TypeName(t.Cell(1, 1)) & "|" & TypeName(t.Rows) & "|" & TypeName(t.Rows(1)) & "|" & t.Cell(2, 1).RowIndex & "|" & t.Cell(2, 1).ColumnIndex
End Function

Function T18() As String
    Dim t As Table, r As Row, s As String
    Set t = NewTable(3, 1)
    t.Cell(1, 1).Range.Text = "p"
    t.Cell(2, 1).Range.Text = "q"
    t.Cell(3, 1).Range.Text = "r"
    For Each r In t.Rows
        s = s & Left(r.Cells(1).Range.Text, 1)
    Next
    T18 = s
End Function

Function T19() As String
    Dim t As Table
    Set t = NewTable(2, 2)
    t.Rows(1).Delete
    T19 = t.Rows.Count & "|" & ActiveDocument.Tables(1).Rows.Count
End Function

Function T20() As String
    Dim t As Table
    Set t = NewTable(2, 2)
    t.Cell(1, 1).Range.Text = "Hello world"
    T20 = ActiveDocument.ComputeStatistics(wdStatisticWords) & "|" & t.Range.Words.Count & "|" & t.Cell(1, 1).Range.Words.Count
End Function
