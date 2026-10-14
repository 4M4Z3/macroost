Option Explicit

' Addresses, offsets, resizing, unions, intersections, areas, rows and columns.

Function A001() As String
    Reset
    With Range("B3:D7")
        A001 = .Address & "|" & .Address(False, False) & "|" & .Address(True, False) & "|" & .Address(False, True) & "|" & .Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End With
End Function

Function A002() As String
    Reset
    With Range("C5")
        A002 = .Address(ReferenceStyle:=xlR1C1) & "|" & .Address(False, False, xlR1C1) & "|" & .Address(False, False, xlR1C1, , Range("A1")) & "|" & .Address(False, False, xlR1C1, RelativeTo:=Range("E2")) & "|" & .Address(True, False, xlR1C1, , Range("B2"))
    End With
End Function

Function A003() As String
    Reset
    A003 = Range("A1:C3").Address(External:=True) & "|" & Range("A1").Address(0, 0, xlA1, True) & "|" & Range("B2:C3").Address(ReferenceStyle:=xlR1C1, External:=True)
End Function

Function A004() As String
    Reset
    A004 = Range("A:A").Address & "|" & Range("2:4").Address & "|" & Range("B:D").Address(False, False) & "|" & Cells.Address & "|" & Range("A:A").Address(ReferenceStyle:=xlR1C1) & "|" & Range("3:3").Address(ReferenceStyle:=xlR1C1)
End Function

Function A005() As String
    Reset
    A005 = Range("A1,C3:D4,F6").Address & "|" & Range("A1,C3:D4,F6").Areas.Count & "|" & Range("A1,C3:D4,F6").Count & "|" & Range("A1,C3:D4,F6").Address(False, False, xlR1C1)
End Function

Function A006() As String
    Reset
    A006 = Range("C3").Offset(-2, -2).Address & "|" & Range("C3").Offset(0, 5).Address & "|" & Range("C3:D4").Offset(1).Address & "|" & Range("C3:D4").Offset(, 1).Address & "|" & Range("A1,C3").Offset(1, 1).Address
End Function

Function A007() As String
    Reset
    On Error Resume Next
    Dim r As Range
    Set r = Range("A1").Offset(-1, 0)
    A007 = Err.Number & "|" & Err.Description & "|" & (r Is Nothing)
    Err.Clear
    Set r = Range("XFD1").Offset(0, 1)
    A007 = A007 & "|" & Err.Number
End Function

Function A008() As String
    Reset
    A008 = Range("B2").Resize(3).Address & "|" & Range("B2").Resize(, 4).Address & "|" & Range("B2:D9").Resize(2, 2).Address & "|" & Range("B2").Resize(1, 1).Address & "|" & Range("B2:C3").Resize(3).Address
End Function

Function A009() As String
    Reset
    On Error Resume Next
    Dim r As Range
    Set r = Range("B2").Resize(0, 1)
    A009 = Err.Number & "|" & Err.Description
End Function

Function A010() As String
    Reset
    A010 = Union(Range("A1:B2"), Range("D4")).Address & "|" & Union(Range("A1:B2"), Range("B2:C3")).Address & "|" & Union(Range("A1"), Range("A2")).Address & "|" & Union(Range("A1:A3"), Range("A1")).Count & "|" & Union(Range("A1:A3"), Range("A1")).Areas.Count
End Function

Function A011() As String
    Dim r As Range
    Reset
    Set r = Intersect(Range("A1:C5"), Range("B3:E9"))
    A011 = r.Address
    Set r = Intersect(Range("A1:A2"), Range("C1:C2"))
    A011 = A011 & "|" & (r Is Nothing)
    Set r = Intersect(Range("A:A"), Range("3:3"))
    A011 = A011 & "|" & r.Address
    Set r = Intersect(Range("A1:D4"), Range("B2:C3"), Range("C3:F9"))
    A011 = A011 & "|" & r.Address
End Function

Function A012() As String
    Dim r As Range
    Reset
    Set r = Intersect(Range("A1:B10,D1:E10"), Range("2:3"))
    A012 = r.Address & "|" & r.Areas.Count & "|" & r.Count
End Function

Function A013() As String
    Dim r As Range, a As Range, s As String
    Reset
    Set r = Range("A1:B2,D4:E5,G7")
    For Each a In r.Areas
        s = s & a.Address(False, False) & "=" & a.Count & ","
    Next
    A013 = s & "|" & r.Areas(2).Address & "|" & r.Areas(3).Cells.Count
End Function

Function A014() As String
    Reset
    A014 = Range("C5").EntireRow.Address & "|" & Range("C5:D7").EntireColumn.Address & "|" & Range("C5").EntireRow.Count & "|" & Range("C5").EntireRow.Cells.Count & "|" & Range("B2:C3").EntireRow.Rows.Count
End Function

Function A015() As String
    Reset
    A015 = Rows(3).Address & "|" & Columns(2).Address & "|" & Columns("C:E").Address & "|" & Rows("2:5").Address & "|" & Range("B2:E9").Rows(2).Address & "|" & Range("B2:E9").Columns(3).Address
End Function

Function A016() As String
    Reset
    A016 = Rows.Count & "|" & Columns.Count & "|" & Cells.Count & "|" & Cells.CountLarge & "|" & Range("A1:C3").Rows.Count & "|" & Range("A1,B1:B5").Rows.Count
End Function

Function A017() As String
    Reset
    A017 = Cells(2, 3).Address & "|" & Cells(2, "D").Address & "|" & Cells(5).Address & "|" & Cells(16385).Address & "|" & Range("B2:D4").Cells(3).Address & "|" & Range("B2:D4").Cells(10).Address
End Function

Function A018() As String
    Reset
    A018 = Range(Cells(2, 2), Cells(4, 5)).Address & "|" & Range("A1", "C3").Address & "|" & Range(Range("D4"), Range("B2")).Address & "|" & Range("B2:C3", "E5").Address
End Function

Function A019() As String
    Reset
    A019 = Range("B2:D5").Range("A1").Address & "|" & Range("B2:D5").Range("B2:C3").Address & "|" & Range("C3").Range("A2").Address & "|" & Range("B2").Cells(0, 0).Address & "|" & Range("B2").Cells(-1, 1).Address
End Function

Function A020() As String
    Reset
    A020 = Range("B2:D5").Item(2).Address & "|" & Range("B2:D5").Item(2, 3).Address & "|" & Range("B2:D5")(1, 1).Address & "|" & Range("B2:D5")(5).Address
End Function

Function A021() As String
    Dim r As Range, s As String
    Reset
    For Each r In Range("A1:C2").Rows
        s = s & r.Address(False, False) & ","
    Next
    s = s & "|"
    For Each r In Range("A1:C2").Columns
        s = s & r.Address(False, False) & ","
    Next
    A021 = s
End Function

Function A022() As String
    Reset
    A022 = Range("B2:D5").Row & "|" & Range("B2:D5").Column & "|" & Range("D5:B2").Address & "|" & Range("B2:D5").Rows(Range("B2:D5").Rows.Count).Row & "|" & Range("XFD1048576").Row & "|" & Range("XFD1048576").Column
End Function

Function A023() As String
    Reset
    On Error Resume Next
    Dim r As Range
    Set r = Range("A0")
    A023 = Err.Number & "|" & Err.Description
    Err.Clear
    Set r = Range("ZZZZ1")
    A023 = A023 & "|" & Err.Number
    Err.Clear
    Set r = Range("")
    A023 = A023 & "|" & Err.Number
End Function

Function A024() As String
    Reset
    A024 = Range("$B$2").Address & "|" & Range("b2:c3").Address & "|" & Range(" B2 ").Address & "|" & Range("A1:A1").Address & "|" & Range("C3:A1").Address
End Function

Function A025() As String
    Reset
    A025 = Range("B2").Offset(1, 1).Offset(-1, -1).Address & "|" & Range("A1").Resize(2, 3).Offset(1, 1).Resize(1).Address & "|" & Range("B2:D4").Offset(1, 1).Cells(1, 1).Address
End Function

Function A026() As String
    Reset
    A026 = Range("A1").Next.Address & "|" & Range("B1").Previous.Address & "|" & Range("C3:D4").Next.Address
End Function

Function A027() As String
    Dim r As Range
    Reset
    Set r = Range("B2:C3")
    A027 = r.Cells(1).Address & "|" & r.Cells(r.Cells.Count).Address & "|" & r.Rows(1).Cells.Count & "|" & r.Columns(2).Address & "|" & r.Columns("B").Address
End Function

Function A028() As String
    Reset
    A028 = Range("A1").Address(ReferenceStyle:=xlR1C1, RowAbsolute:=False, ColumnAbsolute:=True, RelativeTo:=Range("C3")) & "|" & Range("E7").Address(False, False, xlR1C1, False, Range("E7")) & "|" & Range("A1:B2").Address(False, False, xlR1C1, , Range("B2"))
End Function

Function A029() As String
    Reset
    A029 = Union(Range("A1"), Range("C1"), Range("E1")).Address(False, False) & "|" & Union(Rows(1), Rows(3)).Address & "|" & Union(Columns("A"), Columns("C")).Count
End Function

Function A030() As String
    Reset
    A030 = Range("Sheet1!B2").Address(External:=True) & "|" & Worksheets("Sheet1").Range("C4").Address(0, 0) & "|" & Application.Range("D5").Address
End Function

Function A031() As String
    Reset
    A031 = Range("B2:D4").Rows(2).Cells(3).Address & "|" & Range("B2:D4").Columns(1).Cells(3).Address & "|" & Range("B2:D4").Rows.Item(3).Address
End Function

Function A032() As String
    Reset
    A032 = TypeName(Range("A1")) & "|" & TypeName(Range("A1").Rows) & "|" & TypeName(Range("A1").Areas) & "|" & TypeName(Cells) & "|" & TypeName(Range("A1").EntireRow)
End Function

Function A033() As String
    Reset
    A033 = Range("A1:C3").Columns.Count & "|" & Range("A1:C3").EntireColumn.Columns.Count & "|" & Range("A1:C3").EntireColumn.Count & "|" & Range("A1:C3").EntireRow.Count & "|" & Columns("B:C").Count & "|" & Rows("2:9").Count
End Function

Function A034() As String
    Reset
    A034 = Range("A1:B2").Offset(1, 1).Address(0, 0) & "|" & Range("A1:B2").Resize(3, 3).Address(0, 0) & "|" & Range("A1:B2").Cells(3, 3).Address(0, 0)
End Function

Function A035() As String
    Reset
    A035 = Range("B2").Address(True, True, xlR1C1) & "|" & Range("B2").Address(False, True, xlR1C1) & "|" & Range("B2").Address(True, False, xlR1C1)
End Function
