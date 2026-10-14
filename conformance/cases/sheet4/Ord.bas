Option Explicit

' Sort, AutoFilter with visible cells, RemoveDuplicates.

Private Sub Table()
    With Worksheets("Sheet1")
        .Range("A1:C1").Value = Array("Name", "Region", "Amount")
        .Range("A2:C2").Value = Array("Cara", "East", 30)
        .Range("A3:C3").Value = Array("abe", "West", 10)
        .Range("A4:C4").Value = Array("Bob", "East", 20)
        .Range("A5:C5").Value = Array("Dee", "West", 40)
        .Range("A6:C6").Value = Array("Bob", "North", 5)
    End With
End Sub

Private Function Col(ByVal c As Long, ByVal n As Long) As String
    Dim i As Long, s As String
    For i = 1 To n
        If i > 1 Then s = s & ","
        s = s & Worksheets("Sheet1").Cells(i, c).Value
    Next
    Col = s
End Function

Function O001() As String
    Reset: Table
    Range("A1:C6").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
    O001 = Col(1, 6) & "|" & Col(3, 6)
End Function

Function O002() As String
    Reset: Table
    Range("A1:C6").Sort Key1:=Range("C2"), Order1:=xlDescending, Header:=xlYes
    O002 = Col(1, 6) & "|" & Col(3, 6)
End Function

Function O003() As String
    Reset: Table
    Range("A1:C6").Sort Key1:=Range("B1"), Order1:=xlAscending, Key2:=Range("C1"), Order2:=xlDescending, Header:=xlYes
    O003 = Col(1, 6) & "|" & Col(2, 6)
End Function

Function O004() As String
    Reset: Table
    Range("A2:C6").Sort Key1:=Range("A2"), Header:=xlNo
    O004 = Col(1, 6)
End Function

Function O005() As String
    Reset: Table
    Range("A1:C6").Sort Key1:=Range("A1"), Header:=xlGuess
    O005 = Col(1, 6)
End Function

Function O006() As String
    Reset
    Range("A1:A7").Value = Application.WorksheetFunction.Transpose(Array(3, "b", True, 1, "A", "", 2))
    Range("A1:A7").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlNo
    O006 = Col(1, 7)
    Range("A1:A7").Sort Key1:=Range("A1"), Order1:=xlDescending, Header:=xlNo
    O006 = O006 & "|" & Col(1, 7)
End Function

Function O007() As String
    Reset
    Range("A1:A4").Value = Application.WorksheetFunction.Transpose(Array(4, 2, 3, 1))
    Range("B1:B4").Formula = "=A1*10"
    Range("A1:B4").Sort Key1:=Range("A1"), Header:=xlNo
    O007 = Col(1, 4) & "|" & Col(2, 4) & "|" & Range("B1").Formula
End Function

Function O008() As String
    Reset: Table
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("C2:C6"), Order:=xlAscending
        .SetRange Range("A1:C6")
        .Header = xlYes
        .Apply
    End With
    O008 = Col(1, 6)
End Function

Function O009() As String
    Reset: Table
    Range("A1:C6").AutoFilter Field:=2, Criteria1:="East"
    O009 = ActiveSheet.AutoFilterMode & "|" & Rows(3).Hidden & "|" & Rows(4).Hidden & "|" & Range("A1:C6").SpecialCells(xlCellTypeVisible).Address & "|" & ActiveSheet.FilterMode
    ActiveSheet.AutoFilterMode = False
    O009 = O009 & "|" & Rows(3).Hidden & "|" & ActiveSheet.AutoFilterMode
End Function

Function O010() As String
    Reset: Table
    Range("A1").AutoFilter Field:=3, Criteria1:=">15"
    O010 = Range("A1:C6").SpecialCells(xlCellTypeVisible).Address & "|" & Application.WorksheetFunction.Subtotal(9, Range("C2:C6")) & "|" & Application.WorksheetFunction.Sum(Range("C2:C6"))
End Function

Function O011() As String
    Reset: Table
    Range("A1:C6").AutoFilter Field:=1, Criteria1:="Bob", Operator:=xlOr, Criteria2:="Dee"
    O011 = Range("A2:A6").SpecialCells(xlCellTypeVisible).Address & "|" & Range("A2:A6").SpecialCells(xlCellTypeVisible).Count
    Range("A1:C6").AutoFilter Field:=3, Criteria1:=">=10", Operator:=xlAnd, Criteria2:="<40"
    O011 = O011 & "|" & Range("A2:A6").SpecialCells(xlCellTypeVisible).Address
End Function

Function O012() As String
    Reset: Table
    Range("A1:C6").AutoFilter Field:=2, Criteria1:="West"
    Range("A1:C6").AutoFilter
    O012 = ActiveSheet.AutoFilterMode & "|" & Rows(2).Hidden & "|" & Rows(4).Hidden
End Function

Function O013() As String
    Reset: Table
    Range("A1:C6").AutoFilter Field:=1, Criteria1:="B*"
    O013 = Range("A2:A6").SpecialCells(xlCellTypeVisible).Address
    Range("A1:C6").AutoFilter Field:=1, Criteria1:="<>Bob"
    O013 = O013 & "|" & Range("A2:A6").SpecialCells(xlCellTypeVisible).Address
End Function

Function O014() As String
    Dim c As Range, s As String
    Reset: Table
    Range("A1:C6").AutoFilter Field:=2, Criteria1:="East"
    For Each c In Range("A2:A6").SpecialCells(xlCellTypeVisible)
        s = s & c.Value & ","
    Next
    Range("A2:C6").SpecialCells(xlCellTypeVisible).Copy Range("E1")
    O014 = s & "|" & Dump(Range("E1:G2"))
End Function

Function O015() As String
    Reset: Table
    Range("A1:C6").AutoFilter Field:=3, Criteria1:="20"
    Range("A2:C6").SpecialCells(xlCellTypeVisible).EntireRow.Delete
    ActiveSheet.AutoFilterMode = False
    O015 = Col(1, 5)
End Function

Function O016() As String
    Reset: Table
    Range("A1:C6").RemoveDuplicates Columns:=1, Header:=xlYes
    O016 = Col(1, 6) & "|" & Col(2, 6)
End Function

Function O017() As String
    Reset
    Range("A1:B6").Value = 1
    Range("A3").Value = 2
    Range("B5").Value = 2
    Range("A1:B6").RemoveDuplicates Columns:=Array(1, 2), Header:=xlNo
    O017 = Dump(Range("A1:B6"))
End Function

Function O018() As String
    Reset
    Range("A1:A5").Value = Application.WorksheetFunction.Transpose(Array("x", "X", "y", "x", "z"))
    Range("A1:A5").RemoveDuplicates Columns:=1, Header:=xlNo
    O018 = Col(1, 5)
End Function

Function O020() As String
    Reset: Table
    Range("A1:C6").Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlYes, MatchCase:=False, Orientation:=xlTopToBottom
    O020 = Col(1, 6)
    Range("A1:C6").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes, MatchCase:=True
    O020 = O020 & "|" & Col(1, 6)
End Function

Function O021() As String
    Reset
    Range("A1:E1").Value = Array(5, 3, 4, 1, 2)
    Range("A1:E1").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlNo, Orientation:=xlLeftToRight
    O021 = Range("A1").Value & Range("B1").Value & Range("C1").Value & Range("D1").Value & Range("E1").Value
End Function

Function O022() As String
    Reset
    Range("A1:A5").Value = Application.WorksheetFunction.Transpose(Array("10", 9, "8", 100, "x"))
    Range("A1:A5").Sort Key1:=Range("A1"), Header:=xlNo
    O022 = Col(1, 5) & "|" & TypeName(Range("A1").Value)
End Function

' Runs last: after an xlFilterValues filter the host's next Range.Sort with Orientation and
' MatchCase arguments in the same session leaves the data unsorted (observed), so no sort case
' follows it.
Function O023() As String
    Reset: Table
    Range("A1:C6").AutoFilter Field:=2, Criteria1:=Array("East", "North"), Operator:=xlFilterValues
    O023 = Range("A2:A6").SpecialCells(xlCellTypeVisible).Address & "|" & ActiveSheet.AutoFilter.Range.Address
End Function
