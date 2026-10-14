Option Explicit

' Comments, hyperlinks and defined names.

Function M001() As String
    Reset
    Range("B2").AddComment "Check this"
    M001 = Range("B2").Comment.Text & "|" & ActiveSheet.Comments.Count & "|" & (Range("B3").Comment Is Nothing) & "|" & TypeName(Range("B2").Comment)
End Function

Function M002() As String
    Reset
    Range("B2").AddComment "one"
    Range("B2").Comment.Text "two"
    M002 = Range("B2").Comment.Text
    Range("B2").Comment.Delete
    M002 = M002 & "|" & (Range("B2").Comment Is Nothing) & "|" & ActiveSheet.Comments.Count
End Function

Function M003() As String
    Reset
    Range("A1").AddComment "a"
    On Error Resume Next
    Range("A1").AddComment "b"
    M003 = Err.Number & "|" & Range("A1").Comment.Text
    Err.Clear
    Range("A1:B2").ClearComments
    M003 = M003 & "|" & (Range("A1").Comment Is Nothing)
End Function

Function M004() As String
    Dim c As Comment, s As String
    Reset
    Range("C3").AddComment "c3"
    Range("A1").AddComment "a1"
    Range("B5").AddComment "b5"
    For Each c In ActiveSheet.Comments
        s = s & c.Parent.Address(0, 0) & "=" & c.Text & ","
    Next
    M004 = s & "|" & ActiveSheet.Comments(1).Text
End Function

Function M005() As String
    Reset
    Range("A1").AddComment "note"
    Range("A1").Value = 5
    Range("A1").ClearContents
    M005 = (Range("A1").Comment Is Nothing)
    Range("A1").Clear
    M005 = M005 & "|" & (Range("A1").Comment Is Nothing)
    Range("A2").AddComment "x"
    Rows(1).Delete
    M005 = M005 & "|" & Range("A1").Comment.Text
End Function

Function M006() As String
    Reset
    Range("A1").AddComment "copied"
    Range("A1").Value = 1
    Range("A1").Copy Range("C1")
    M006 = Range("C1").Comment.Text & "|" & ActiveSheet.Comments.Count
End Function

Function M007() As String
    Reset
    ActiveSheet.Hyperlinks.Add Anchor:=Range("A1"), Address:="https://example.com", TextToDisplay:="Example"
    M007 = Range("A1").Value & "|" & ActiveSheet.Hyperlinks.Count & "|" & Range("A1").Hyperlinks(1).Address & "|" & Range("A1").Hyperlinks.Count & "|" & Range("A2").Hyperlinks.Count
End Function

Function M008() As String
    Reset
    ActiveSheet.Hyperlinks.Add Range("B2"), "", "Sheet1!D4", "tip", "Go"
    With Range("B2").Hyperlinks(1)
        M008 = .SubAddress & "|" & .TextToDisplay & "|" & .ScreenTip & "|" & Range("B2").Value & "|" & .Address
    End With
End Function

Function M009() As String
    Reset
    ActiveSheet.Hyperlinks.Add Range("A1"), "https://example.com"
    M009 = Range("A1").Value
    Range("A1").Hyperlinks.Delete
    M009 = M009 & "|" & ActiveSheet.Hyperlinks.Count & "|" & Range("A1").Value
    ActiveSheet.Hyperlinks.Add Range("A2"), "mailto:x@example.com", , , "mail"
    Range("A2").ClearHyperlinks
    M009 = M009 & "|" & ActiveSheet.Hyperlinks.Count & "|" & Range("A2").Value
End Function

Function M010() As String
    Reset
    Range("A1:A3").Value = 2
    ThisWorkbook.Names.Add Name:="Data", RefersTo:="=Sheet1!$A$1:$A$3"
    M010 = ThisWorkbook.Names.Count & "|" & ThisWorkbook.Names("Data").RefersTo & "|" & Range("Data").Address & "|" & Application.WorksheetFunction.Sum(Range("Data"))
End Function

Function M011() As String
    Reset
    ThisWorkbook.Names.Add "Rate", RefersTo:=Range("B2")
    Range("B2").Value = 0.5
    Range("C1").Formula = "=Rate*10"
    M011 = ThisWorkbook.Names("Rate").RefersTo & "|" & ThisWorkbook.Names("Rate").RefersToRange.Address & "|" & Range("C1").Value & "|" & ThisWorkbook.Names("rate").Name
End Function

Function M012() As String
    Reset
    Range("A1").Name = "Header"
    M012 = ThisWorkbook.Names("Header").RefersTo & "|" & Range("A1").Name.Name
    ThisWorkbook.Names("Header").Delete
    M012 = M012 & "|" & ThisWorkbook.Names.Count
End Function

Function M013() As String
    Reset
    ActiveSheet.Names.Add Name:="Local", RefersTo:="=$A$1"
    M013 = ThisWorkbook.Names.Count & "|" & ActiveSheet.Names.Count & "|" & ThisWorkbook.Names(1).Name & "|" & ActiveSheet.Names("Local").RefersTo
End Function

Function M014() As String
    Reset
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="A1", RefersTo:="=Sheet1!$B$1"
    M014 = Err.Number
    Err.Clear
    ThisWorkbook.Names.Add Name:="bad name", RefersTo:="=Sheet1!$B$1"
    M014 = M014 & "|" & Err.Number
    Err.Clear
    Dim n As Name
    Set n = ThisWorkbook.Names("Missing")
    M014 = M014 & "|" & Err.Number
End Function

Function M015() As String
    Reset
    ThisWorkbook.Names.Add Name:="Tbl", RefersTo:="=Sheet1!$A$1:$B$2"
    Rows(1).Insert
    M015 = ThisWorkbook.Names("Tbl").RefersTo
    Columns("A").Delete
    M015 = M015 & "|" & ThisWorkbook.Names("Tbl").RefersTo
End Function

Function M016() As String
    Reset
    ThisWorkbook.Names.Add Name:="Konst", RefersTo:="=42"
    M016 = ThisWorkbook.Names("Konst").RefersTo & "|" & ThisWorkbook.Names("Konst").Value & "|" & Application.Evaluate("Konst*2") & "|" & [Konst]
End Function

Function M017() As String
    Reset
    ThisWorkbook.Names.Add Name:="Spot", RefersTo:="=Sheet1!$C$3"
    Range("Spot").Value = "here"
    M017 = Range("C3").Value & "|" & [Spot].Address & "|" & Worksheets("Sheet1").Range("Spot").Row
    ThisWorkbook.Names("Spot").RefersTo = "=Sheet1!$D$4"
    M017 = M017 & "|" & Range("Spot").Address
End Function

Function M018() As String
    Dim n As Name, s As String
    Reset
    ThisWorkbook.Names.Add "Zed", "=Sheet1!$A$1"
    ThisWorkbook.Names.Add "Alpha", "=Sheet1!$B$1"
    ThisWorkbook.Names.Add "mid", "=Sheet1!$C$1"
    For Each n In ThisWorkbook.Names
        s = s & n.Name & ","
    Next
    M018 = s & "|" & ThisWorkbook.Names(1).Name
End Function

Function M019() As String
    Reset
    Range("A1").AddComment
    M019 = "[" & Range("A1").Comment.Text & "]"
    Range("A1").Comment.Visible = True
    M019 = M019 & "|" & Range("A1").Comment.Visible
End Function

Function M020() As String
    Reset
    ThisWorkbook.Names.Add Name:="Grid", RefersToR1C1:="=Sheet1!R1C1:R2C2"
    M020 = ThisWorkbook.Names("Grid").RefersTo & "|" & ThisWorkbook.Names("Grid").RefersToR1C1
End Function
