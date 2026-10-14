Option Explicit

' Default members: Range(...)(i), Selection(i), Cells(r, c) chains, objects used as values.

Function Q001() As String
    Reset
    Range("A1:C3").Value = 1
    Range("A1:C3")(2, 2).Value = "mid"
    Range("A1:C3")(4).Value = "four"
    Q001 = Range("B2").Value & "|" & Range("A2").Value & "|" & Range("B2:C3")(1).Address & "|" & Range("B2:C3")(3, 3).Address
End Function

Function Q002() As String
    Reset
    Range("B2:D4").Select
    Selection(1).Value = "first"
    Selection(2, 1).Value = "below"
    Q002 = Range("B2").Value & "|" & Range("B3").Value & "|" & Selection(5).Address & "|" & Selection.Cells(9).Address
End Function

Function Q003() As String
    Reset
    Worksheets("Sheet1").Range("A1:B2")(2, 2) = 7
    Worksheets("Sheet1").Cells(3, 1) = "c"
    ActiveSheet.Cells(4, 1) = 4
    Q003 = Range("B2").Value & "|" & Range("A3").Value & "|" & Worksheets("Sheet1").Range("A4") & "|" & Worksheets(1).Cells(4, 1)
End Function

Function Q004() As String
    Dim v As Variant, r As Range
    Reset
    Range("A1").Value = 42
    v = Range("A1")
    Set r = Range("A1")
    Q004 = TypeName(v) & "|" & v & "|" & TypeName(r) & "|" & r & "|" & (r = 42) & "|" & (Range("A1") + 1)
End Function

Function Q005() As String
    Reset
    Range("A1:A3").Value = 3
    Range("C1:C3") = Range("A1:A3").Value
    Range("D1:D3") = Range("A1:A3")
    Q005 = Range("C3").Value & "|" & Range("D2").Value & "|" & IsEmpty(Range("E1"))
End Function

Function Q006() As String
    Reset
    Cells(1, 1) = 1
    Cells(1, 2) = Cells(1, 1) + 1
    Cells(2, 1).Value = Cells(1, 2) * 10
    Q006 = Cells(1, 2) & "|" & Cells(2, 1) & "|" & Cells(2, 1).Value & "|" & Cells(1)
End Function

Function Q007() As String
    Dim ws As Worksheet
    Reset
    Set ws = Worksheets("Sheet1")
    ws.Range("A1") = "x"
    ws.Cells(2, 1) = "y"
    With ws
        .Range("A3") = .Range("A1") & .Cells(2, 1)
        Q007 = .Range("A3") & "|" & .Cells(3, 1).Value
    End With
End Function

Function Q008() As String
    Reset
    Range("A1").Value = 5
    Range("A2").Value = Range("A1").Value
    Range("A3").Value = Range("A1")
    Range("A4") = Range("A1")
    Q008 = Range("A2") + Range("A3") + Range("A4")
End Function

Function Q009() As String
    Dim c As Range
    Reset
    Range("A1:A3").Value = 2
    For Each c In Range("A1:A3")
        c.Value = c.Value * c.Row
    Next
    For Each c In Range("A1:A3").Cells
        c = c + 1
    Next
    Q009 = Range("A1") & "," & Range("A2") & "," & Range("A3")
End Function

Function Q010() As String
    Reset
    Range("A1:B2").Value = 9
    Q010 = Range("A1:B2").Cells(2, 2) & "|" & Range("A1:B2").Rows(2).Cells(1) & "|" & Range("A1:B2").Columns(2)(1) & "|" & Range("A1").Offset(1, 1)
End Function

Function Q011() As String
    Dim a As Variant
    Reset
    Range("A1:B2").Value = 4
    a = Range("A1:B2")
    Q011 = TypeName(a) & "|" & a(2, 2)
    a = Range("A1:B2").Cells
    Q011 = Q011 & "|" & TypeName(a)
End Function

Function Q012() As String
    Reset
    Range("A1").Value = "t"
    Q012 = Len(Range("A1")) & "|" & UCase(Range("A1")) & "|" & IsEmpty(Range("Z9")) & "|" & IsNumeric(Range("A1")) & "|" & CStr(Range("A1"))
End Function

Function Q013() As String
    Reset
    Range("A1:C1").Value = Array(1, 2, 3)
    Q013 = Application.WorksheetFunction.Sum(Range("A1:C1")) & "|" & Application.Sum(Range("A1"), Range("C1")) & "|" & Application.WorksheetFunction.Max(Range("A1"), Range("B1"))
End Function

Function Q014() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Src"
    Worksheets("Src").Range("A1:A3").Value = 6
    Worksheets("Sheet1").Range("B1:B3").Value = Worksheets("Src").Range("A1:A3").Value
    Worksheets("Sheet1").Range("C1:C3") = Worksheets("Src").Range("A1:A3").Value
    Q014 = Worksheets("Sheet1").Range("B3") & "|" & Worksheets("Sheet1").Range("C2")
End Function

Function Q015() As String
    Reset
    Range("A1:A3").Select
    Selection = "all"
    Q015 = Range("A3").Value
    ActiveCell = "active"
    Q015 = Q015 & "|" & Range("A1").Value & "|" & Selection(2)
End Function

Function Q016() As String
    Reset
    Range("A1:B3").Value = 1
    Range("A1:B3").Rows(2) = 5
    Range("A1:B3").Columns(1).Value = 7
    Q016 = Dump(Range("A1:B3"))
End Function

Function Q017() As String
    Reset
    Range("A1").Value = 3
    Q017 = Range("A1").Value2 * Range("A1") & "|" & Range("A1") & Range("A1") & "|" & (Range("A1") = Range("A1").Value)
End Function

Function Q018() As String
    Reset
    Range("A1").Value = 1
    Q018 = Worksheets("Sheet1").Range("A1")(1).Address & "|" & Range("A1")(2).Address & "|" & Range("A1")(1, 2).Address & "|" & Range("C3")(0, 0).Address
End Function

Function Q019() As String
    Reset
    Range("A1:A3").Value = 3
    Range("C1:C3").Value = Range("A1:A3")
    Range("D1").Value = Range("A1")
    Q019 = "[" & Range("C1").Value & "][" & Range("C2").Value & "]|" & Range("D1").Value
End Function
