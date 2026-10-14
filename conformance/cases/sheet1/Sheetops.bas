Option Explicit

' Each case starts from a cleared Sheet1 and returns a string describing what happened.

Private Sub Reset()
    Worksheets("Sheet1").Cells.Clear
    Worksheets("Sheet1").Activate
    Worksheets("Sheet1").Range("A1").Select
End Sub

Private Sub FillA(n As Long)
    Dim i As Long
    For i = 1 To n
        Worksheets("Sheet1").Cells(i, 1).Value = i
    Next
End Sub

Function X001() As String
    Reset: FillA 10
    X001 = Worksheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Address
End Function

Function X002() As String
    Reset: FillA 10
    X002 = Range("A1").End(xlDown).Row & "|" & Range("A11").End(xlUp).Row & "|" & Range("B1").End(xlToRight).Address
End Function

Function X003() As String
    Reset
    Range("A1").Value = 1: Range("A3").Value = 3: Range("A4").Value = 4: Range("A7").Value = 7
    X003 = Range("A1").End(xlDown).Address & "|" & Range("A3").End(xlDown).Address & "|" & Range("A4").End(xlDown).Address & "|" & Range("A7").End(xlDown).Address & "|" & Range("A1048576").End(xlUp).Address
End Function

Function X004() As String
    Reset
    X004 = Range("B2").Offset(1, 2).Address & "|" & Range("B2").Resize(3, 2).Address & "|" & Range("B2:D5").Offset(-1, -1).Address(False, False) & "|" & Range("A1").Offset(2).Address
End Function

Function X005() As String
    Reset
    X005 = Range("B2:D5").Rows.Count & "|" & Range("B2:D5").Columns.Count & "|" & Range("B2:D5").Count & "|" & Range("B2:D5").Cells(2, 3).Address & "|" & Range("B2:D5").Cells(5).Address & "|" & Range("A:A").Rows.Count
End Function

Function X006() As String
    Reset
    Range("A1:C2").Value = 7
    Range("A3").Value = "x"
    X006 = Application.WorksheetFunction.Sum(Range("A1:C2")) & "|" & Range("A3").Value & "|" & TypeName(Range("A1:C2").Value) & "|" & UBound(Range("A1:C2").Value, 1) & "x" & UBound(Range("A1:C2").Value, 2)
End Function

Function X007() As String
    Dim a(1 To 2, 1 To 3) As Variant, v As Variant
    Reset
    a(1, 1) = 1: a(1, 2) = "two": a(1, 3) = 3.5
    a(2, 1) = True: a(2, 2) = Empty: a(2, 3) = #1/2/2026#
    Range("B2:D3").Value = a
    v = Range("B2:D3").Value
    X007 = TypeName(v(1, 1)) & "|" & v(1, 2) & "|" & TypeName(v(1, 3)) & "|" & TypeName(v(2, 1)) & "|" & TypeName(v(2, 2)) & "|" & TypeName(v(2, 3)) & "|" & Range("D3").Text
End Function

Function X008() As String
    Reset: FillA 5
    Range("A2").EntireRow.Insert
    X008 = Range("A2").Value & "|" & IsEmpty(Range("A2").Value) & "|" & Range("A3").Value & "|" & Cells(Rows.Count, 1).End(xlUp).Row
End Function

Function X009() As String
    Reset: FillA 5
    Range("A2:A3").Delete Shift:=xlUp
    X009 = Range("A2").Value & "|" & Cells(Rows.Count, 1).End(xlUp).Row
End Function

Function X010() As String
    Reset: FillA 3
    Range("A1:A3").Copy
    Range("C1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Range("A1:A3").Copy Destination:=Range("E2")
    X010 = Range("C3").Value & "|" & Range("E4").Value & "|" & Range("E1").Value & "|" & IsEmpty(Range("E1").Value)
End Function

Function X011() As String
    Dim f As Range
    Reset
    Range("B5").Value = "Total"
    Range("C9").Value = "total due"
    Set f = Range("A1:D10").Find(What:="total", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    X011 = f.Address
    Set f = Range("A1:D10").Find(What:="due", LookIn:=xlValues, LookAt:=xlPart)
    X011 = X011 & "|" & f.Address
    Set f = Range("A1:D10").Find(What:="nothing here", LookIn:=xlValues)
    X011 = X011 & "|" & (f Is Nothing)
End Function

Function X012() As String
    Reset
    Range("A1").Formula = "=1+2"
    Range("A2").Formula = "=A1*10"
    X012 = Range("A1").Value & "|" & Range("A2").Value & "|" & Range("A2").Formula & "|" & Range("A2").HasFormula
End Function

Function X013() As String
    Reset
    Range("A1").Value = 1234.5
    Range("A1").NumberFormat = "#,##0.00"
    Range("A2").Value = "10/14/2026"
    X013 = Range("A1").Text & "|" & TypeName(Range("A2").Value) & "|" & Range("A1").NumberFormat
End Function

Function X014() As String
    Reset
    Range("A1:C3").Interior.Color = RGB(255, 0, 0)
    Range("A1").Font.Bold = True
    X014 = Range("B2").Interior.Color & "|" & Range("A1").Font.Bold & "|" & Range("B2").Font.Bold & "|" & RGB(255, 0, 0)
End Function

Function X015() As String
    Reset
    Range("B2").Value = 1: Range("D6").Value = 2
    X015 = ActiveSheet.UsedRange.Address & "|" & Range("C3").CurrentRegion.Address & "|" & Range("B2").CurrentRegion.Address
End Function

Function X016() As String
    Reset
    Range("C5").Select
    X016 = Selection.Address & "|" & ActiveCell.Address & "|" & ActiveSheet.Name
    Range("B2:C3").Select
    X016 = X016 & "|" & Selection.Address & "|" & ActiveCell.Address
End Function

Function X017() As String
    Dim c As Range, s As String
    Reset
    Range("A1:B2").Value = 5
    For Each c In Range("A1:B2")
        s = s & c.Address(False, False) & ","
    Next
    X017 = s
End Function

Function X018() As String
    Reset
    X018 = Worksheets.Count & "|" & Worksheets(1).Name & "|" & ThisWorkbook.Worksheets("Sheet1").Index
End Function

Function X019() As String
    Reset
    Range("A1").Value = "a"
    Range("A1").ClearContents
    X019 = IsEmpty(Range("A1").Value) & "|" & Len(Range("A1").Value) & "|" & (Range("A1").Value = "") & "|" & (Range("A1").Value = 0)
End Function

Function X020() As String
    Reset
    Cells(3, 2) = 42
    Range("B4") = "s"
    [C1] = 7
    X020 = Range("B3").Value & "|" & Cells(4, 2) & "|" & Range("C1").Value & "|" & Cells(3, "B").Value
End Function

Function X021() As String
    Reset
    X021 = Range("A1").Address(RowAbsolute:=False) & "|" & Range("A1:B2").Address(False, True) & "|" & Range("AA100").Column & "|" & Columns("AB").Column & "|" & Range("XFD1").Column
End Function

Function X022() As String
    Reset
    Range("A1").Value = 3: Range("A2").Value = 1: Range("A3").Value = 2
    X022 = Application.WorksheetFunction.Max(Range("A1:A3")) & "|" & Application.WorksheetFunction.Match(2, Range("A1:A3"), 0) & "|" & Application.WorksheetFunction.CountA(Range("A1:A10")) & "|" & Application.WorksheetFunction.Average(Range("A1:A3"))
End Function

Function X023() As String
    Reset
    ActiveSheet.Protect
    X023 = ActiveSheet.ProtectContents
    ActiveSheet.Unprotect
    X023 = X023 & "|" & ActiveSheet.ProtectContents
End Function

Function X024() As String
    Reset
    Range("A1").Value = "x"
    Range("A1").EntireRow.Hidden = True
    X024 = Rows(1).Hidden & "|" & Range("A1").EntireRow.Address & "|" & Range("C3").EntireColumn.Address
    Range("A1").EntireRow.Hidden = False
End Function

Function X025() As String
    Reset
    Range("A1").Value = 1: Range("A2").Value = 2
    Range("A1:A2").Copy
    Range("B1").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False
    X025 = Range("B2").Value & "|" & Application.CutCopyMode
End Function
