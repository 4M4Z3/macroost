Option Explicit

' End in every direction, CurrentRegion, UsedRange and SpecialCells.

Function N001() As String
    Reset
    N001 = Range("C5").End(xlUp).Address & "|" & Range("C5").End(xlDown).Address & "|" & Range("C5").End(xlToLeft).Address & "|" & Range("C5").End(xlToRight).Address
End Function

Function N002() As String
    Reset
    Range("B2:F2").Value = "x"
    N002 = Range("B2").End(xlToRight).Address & "|" & Range("F2").End(xlToLeft).Address & "|" & Range("D2").End(xlToRight).Address & "|" & Range("A2").End(xlToRight).Address & "|" & Range("G2").End(xlToLeft).Address
End Function

Function N003() As String
    Reset
    Range("C1").Value = 1: Range("C2").Value = 2: Range("C5").Value = 5: Range("C6").Value = 6: Range("C9").Value = 9
    N003 = Range("C1").End(xlDown).Address(0, 0) & "," & Range("C2").End(xlDown).Address(0, 0) & "," & Range("C5").End(xlDown).Address(0, 0) & "," & Range("C6").End(xlDown).Address(0, 0) & "," & Range("C9").End(xlDown).Address(0, 0)
    N003 = N003 & "|" & Range("C9").End(xlUp).Address(0, 0) & "," & Range("C6").End(xlUp).Address(0, 0) & "," & Range("C5").End(xlUp).Address(0, 0) & "," & Range("C3").End(xlUp).Address(0, 0) & "," & Range("C20").End(xlUp).Address(0, 0)
End Function

Function N004() As String
    Reset
    Range("A1").Value = 1
    N004 = Range("A1").End(xlUp).Address & "|" & Range("A1").End(xlToLeft).Address & "|" & Range("A1").End(xlDown).Address & "|" & Range("A1").End(xlToRight).Address
End Function

Function N005() As String
    Reset
    Range("A1048576").Value = "last"
    Range("XFD1").Value = "right"
    N005 = Range("A1").End(xlDown).Address & "|" & Range("A1048576").End(xlUp).Address & "|" & Range("A1").End(xlToRight).Address & "|" & Range("XFD1").End(xlToLeft).Address
End Function

Function N006() As String
    Reset
    Range("A1:A5").Value = 1
    Range("A3").ClearContents
    N006 = Cells(Rows.Count, "A").End(xlUp).Row & "|" & Range("A1").End(xlDown).Row & "|" & Range("A4").End(xlUp).Row
End Function

Function N007() As String
    Reset
    Range("B2").Formula = "="""""
    Range("B3").Value = 1
    N007 = Range("B1").End(xlDown).Address & "|" & Range("B5").End(xlUp).Address & "|" & Len(Range("B2").Value)
End Function

Function N008() As String
    Reset
    Range("A1:C1").Value = "h"
    Range("A2:C4").Value = 1
    N008 = Range("A1").End(xlDown).End(xlToRight).Address & "|" & Range("A1").End(xlToRight).End(xlDown).Address & "|" & Range("C4").End(xlUp).End(xlToLeft).Address
End Function

Function N009() As String
    Reset
    Range("B2:D4").Value = 1
    Range("F6").Value = 1
    N009 = Range("C3").CurrentRegion.Address & "|" & Range("F6").CurrentRegion.Address & "|" & Range("A1").CurrentRegion.Address & "|" & Range("H9").CurrentRegion.Address
End Function

Function N010() As String
    Reset
    Range("B2:C3").Value = 1
    Range("D4").Value = 1
    N010 = Range("B2").CurrentRegion.Address & "|" & Range("E5").CurrentRegion.Address & "|" & Range("E6").CurrentRegion.Address
End Function

Function N011() As String
    Reset
    Range("A1:D1").Value = "head"
    Range("A2:D10").Value = 5
    Range("A6").EntireRow.ClearContents
    N011 = Range("A1").CurrentRegion.Address & "|" & Range("B8").CurrentRegion.Address & "|" & Range("A1").CurrentRegion.Rows.Count
End Function

Function N012() As String
    Reset
    N012 = ActiveSheet.UsedRange.Address
    Range("C3").Value = 1
    N012 = N012 & "|" & ActiveSheet.UsedRange.Address
    Range("E8").Value = 2
    N012 = N012 & "|" & ActiveSheet.UsedRange.Address
    Range("E8").ClearContents
    N012 = N012 & "|" & ActiveSheet.UsedRange.Address
End Function

Function N013() As String
    Reset
    Range("B2").Value = 1
    Range("D6").Value = 2
    Rows(6).Delete
    N013 = ActiveSheet.UsedRange.Address
    Columns("B").Delete
    N013 = N013 & "|" & ActiveSheet.UsedRange.Address
End Function

Function N014() As String
    Reset
    Range("C3").Interior.Color = vbYellow
    N014 = ActiveSheet.UsedRange.Address
    Range("F2").Value = 1
    N014 = N014 & "|" & ActiveSheet.UsedRange.Address & "|" & ActiveSheet.UsedRange.Count
End Function

Function N015() As String
    Reset
    Range("B2").Value = 1
    Range("C3").Formula = "=B2+1"
    Range("D4").Value = "t"
    N015 = Cells.SpecialCells(xlCellTypeConstants).Address & "|" & Cells.SpecialCells(xlCellTypeFormulas).Address & "|" & Cells.SpecialCells(xlCellTypeLastCell).Address
End Function

Function N016() As String
    Reset
    Range("A1").Value = 1
    Range("A2").Value = "text"
    Range("A3").Value = True
    Range("A4").Formula = "=1/0"
    Range("A5").Formula = "=""s"""
    Range("A6").Formula = "=2"
    N016 = Range("A1:A6").SpecialCells(xlCellTypeConstants, xlNumbers).Address & "|" & Range("A1:A6").SpecialCells(xlCellTypeConstants, xlTextValues).Address & "|" & Range("A1:A6").SpecialCells(xlCellTypeConstants, xlLogical).Address & "|" & Range("A1:A6").SpecialCells(xlCellTypeFormulas, xlErrors).Address & "|" & Range("A1:A6").SpecialCells(xlCellTypeFormulas, xlTextValues).Address & "|" & Range("A1:A6").SpecialCells(xlCellTypeFormulas, xlNumbers).Address
End Function

Function N017() As String
    Reset
    Range("A1:C3").Value = 1
    Range("B2").ClearContents
    Range("A3").ClearContents
    N017 = Range("A1:C3").SpecialCells(xlCellTypeBlanks).Address & "|" & Range("A1:C3").SpecialCells(xlCellTypeBlanks).Count
End Function

Function N018() As String
    Reset
    Range("A1:A3").Value = 1
    On Error Resume Next
    Dim r As Range
    Set r = Range("A1:A3").SpecialCells(xlCellTypeFormulas)
    N018 = Err.Number & "|" & Err.Description & "|" & (r Is Nothing)
    Err.Clear
    Set r = Range("A1:A3").SpecialCells(xlCellTypeBlanks)
    N018 = N018 & "|" & Err.Number
End Function

Function N019() As String
    Reset
    Range("A1:A6").Value = 1
    Rows(2).Hidden = True
    Rows(4).Hidden = True
    N019 = Range("A1:A6").SpecialCells(xlCellTypeVisible).Address & "|" & Range("A1:A6").SpecialCells(xlCellTypeVisible).Count & "|" & Range("A1:A6").SpecialCells(xlCellTypeVisible).Areas.Count
End Function

Function N020() As String
    Reset
    Range("A1:C4").Value = 1
    Columns("B").Hidden = True
    N020 = Range("A1:C4").SpecialCells(xlCellTypeVisible).Address
End Function

Function N021() As String
    Reset
    Range("B2").Value = 1
    Range("F9").Value = 1
    Range("F9").ClearContents
    N021 = Cells.SpecialCells(xlCellTypeLastCell).Address & "|" & ActiveSheet.UsedRange.Address
End Function

Function N022() As String
    Reset
    Range("A1").Value = 1
    Range("C1").Value = 1
    Range("E1").Value = 1
    N022 = Range("A1:E1").SpecialCells(xlCellTypeConstants).Address & "|" & Range("A1:E1").SpecialCells(xlCellTypeConstants).Areas.Count & "|" & Range("A1:E1").SpecialCells(xlCellTypeBlanks).Address
End Function

Function N023() As String
    Reset
    Range("A1").Value = 1
    Range("B5").Value = 2
    N023 = Range("A1").SpecialCells(xlCellTypeConstants).Address & "|" & Range("A1").SpecialCells(xlCellTypeBlanks).Address
End Function

Function N024() As String
    Reset
    Range("A1:A10").Value = 1
    Range("A11").Value = "x"
    N024 = Range("A1").End(xlDown).Address & "|" & Range("A:A").SpecialCells(xlCellTypeConstants, xlNumbers).Address & "|" & Range("A:A").SpecialCells(xlCellTypeConstants).Count
End Function

Function N025() As String
    Dim lastRow As Long, lastCol As Long
    Reset
    Range("A1:D1").Value = "h"
    Range("A2:A20").Value = 1
    Range("D25").Value = 9
    lastRow = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    N025 = lastRow & "|" & lastCol & "|" & Cells(Rows.Count, 1).End(xlUp).Row & "|" & Cells(1, Columns.Count).End(xlToLeft).Column
End Function

Function N026() As String
    Reset
    Range("B2:C3").Value = 1
    N026 = Range("B2").CurrentRegion.Offset(1).Resize(Range("B2").CurrentRegion.Rows.Count - 1).Address
End Function

Function N027() As String
    Reset
    Range("A1:B2").Value = 1
    Rows(1).Insert
    N027 = ActiveSheet.UsedRange.Address & "|" & Range("A2").CurrentRegion.Address
    Range("A1").EntireColumn.Insert
    N027 = N027 & "|" & ActiveSheet.UsedRange.Address
End Function

Function N028() As String
    Reset
    Range("A1:C3").Value = 1
    Range("A1:C3").ClearContents
    N028 = ActiveSheet.UsedRange.Address & "|" & Application.WorksheetFunction.CountA(Cells)
End Function

Function N029() As String
    Reset
    Range("A1:A3").Value = 1
    Range("A2").Formula = "=A1+1"
    N029 = Range("A1:A3").SpecialCells(xlCellTypeConstants).Address & "|" & Range("A1:A3").SpecialCells(xlCellTypeFormulas).Address & "|" & Range("A1:A3").SpecialCells(xlCellTypeFormulas).Count
End Function

Function N030() As String
    Reset
    Range("B2:D2").Value = 1
    Range("C2").ClearContents
    N030 = Range("B2").End(xlToRight).Address & "|" & Range("D2").End(xlToLeft).Address & "|" & Range("C2").End(xlToRight).Address & "|" & Range("C2").End(xlToLeft).Address
End Function
