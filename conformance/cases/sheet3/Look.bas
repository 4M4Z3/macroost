Option Explicit

' Interior, Font, Borders, alignment, wrapping, merging, sizes, hidden rows/columns, locking.

Function L001() As String
    Reset
    Range("A1").Interior.Color = RGB(255, 255, 0)
    Range("A2").Interior.ColorIndex = 3
    L001 = Range("A1").Interior.Color & "|" & Range("A1").Interior.ColorIndex & "|" & Range("A2").Interior.Color & "|" & Range("A2").Interior.ColorIndex & "|" & Range("A3").Interior.Color & "|" & Range("A3").Interior.ColorIndex
End Function

Function L002() As String
    Reset
    L002 = Range("A1").Interior.Pattern & "|" & Range("A1").Interior.PatternColorIndex
    Range("A1").Interior.Color = vbGreen
    L002 = L002 & "|" & Range("A1").Interior.Pattern
    Range("A1").Interior.Pattern = xlNone
    L002 = L002 & "|" & Range("A1").Interior.ColorIndex & "|" & Range("A1").Interior.Color
    Range("A1").Interior.ColorIndex = xlNone
    L002 = L002 & "|" & Range("A1").Interior.ColorIndex
End Function

Function L003() As String
    Reset
    Range("A1:B2").Interior.Color = vbRed
    Range("B2").Interior.Color = vbBlue
    L003 = IsNull(Range("A1:B2").Interior.Color) & "|" & IsNull(Range("A1:B2").Interior.ColorIndex) & "|" & Range("A1:A2").Interior.Color & "|" & Range("B2").Interior.ColorIndex
End Function

Function L004() As String
    Reset
    Range("A1").Interior.ColorIndex = 6
    Range("A2").Interior.ColorIndex = 15
    Range("A3").Interior.ColorIndex = 37
    Range("A4").Interior.Color = RGB(250, 10, 10)
    L004 = Range("A1").Interior.Color & "|" & Range("A2").Interior.Color & "|" & Range("A3").Interior.Color & "|" & Range("A4").Interior.ColorIndex
End Function

Function L005() As String
    Reset
    With Range("A1").Font
        .Bold = True
        .Italic = True
        .Underline = xlUnderlineStyleSingle
        .Strikethrough = True
        .Size = 14
        .Name = "Arial"
        .Color = RGB(0, 0, 255)
    End With
    With Range("A1").Font
        L005 = .Bold & "|" & .Italic & "|" & .Underline & "|" & .Strikethrough & "|" & .Size & "|" & .Name & "|" & .Color & "|" & .ColorIndex & "|" & .FontStyle
    End With
End Function

Function L006() As String
    Reset
    With Range("B2").Font
        L006 = .Bold & "|" & .Italic & "|" & .Underline & "|" & .Strikethrough & "|" & .Size & "|" & .Name & "|" & .Color & "|" & .ColorIndex & "|" & .FontStyle
    End With
End Function

Function L007() As String
    Reset
    Range("A1:A3").Font.Bold = True
    Range("A2").Font.Bold = False
    L007 = IsNull(Range("A1:A3").Font.Bold) & "|" & Range("A1:A3").Font.Size & "|" & Range("A1,A3").Font.Bold
    Range("A1").Font.Underline = xlUnderlineStyleDouble
    L007 = L007 & "|" & Range("A1").Font.Underline
    Range("A1").Font.Underline = False
    L007 = L007 & "|" & Range("A1").Font.Underline
    Range("A1").Font.Underline = True
    L007 = L007 & "|" & Range("A1").Font.Underline
End Function

Function L008() As String
    Reset
    Range("A1").Font.ColorIndex = 3
    Range("A2").Font.Color = vbRed
    Range("A3").Font.ColorIndex = xlAutomatic
    L008 = Range("A1").Font.Color & "|" & Range("A2").Font.ColorIndex & "|" & Range("A3").Font.ColorIndex & "|" & Range("A3").Font.Color
End Function

Function L009() As String
    Reset
    On Error Resume Next
    Range("A1").Font.Size = 0
    L009 = Err.Number & "|" & Range("A1").Font.Size
    Err.Clear
    Range("A1").Font.Size = 10.5
    L009 = L009 & "|" & Err.Number & "|" & Range("A1").Font.Size
    Err.Clear
    Range("A1").Font.FontStyle = "Bold Italic"
    L009 = L009 & "|" & Range("A1").Font.Bold & "|" & Range("A1").Font.Italic
End Function

Function L010() As String
    Reset
    With Range("B2").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(255, 0, 0)
    End With
    With Range("B2").Borders(xlEdgeBottom)
        L010 = .LineStyle & "|" & .Weight & "|" & .Color & "|" & .ColorIndex
    End With
    L010 = L010 & "|" & Range("B3").Borders(xlEdgeTop).LineStyle & "|" & Range("B2").Borders(xlEdgeTop).LineStyle & "|" & Range("B2").Borders(xlEdgeTop).Weight
End Function

Function L011() As String
    Reset
    Range("B2:D4").BorderAround xlContinuous, xlMedium
    L011 = Range("B2").Borders(xlEdgeTop).LineStyle & "|" & Range("B2").Borders(xlEdgeLeft).Weight & "|" & Range("C3").Borders(xlEdgeTop).LineStyle & "|" & Range("D4").Borders(xlEdgeRight).Weight & "|" & Range("B2:D4").Borders(xlEdgeBottom).LineStyle & "|" & Range("B2:D4").Borders(xlInsideHorizontal).LineStyle
End Function

Function L012() As String
    Reset
    Range("B2:C3").Borders.LineStyle = xlContinuous
    L012 = Range("B2:C3").Borders(xlInsideVertical).LineStyle & "|" & Range("B2:C3").Borders(xlInsideHorizontal).LineStyle & "|" & Range("C3").Borders(xlEdgeLeft).LineStyle & "|" & Range("B2:C3").Borders.LineStyle
    Range("B2:C3").Borders.LineStyle = xlNone
    L012 = L012 & "|" & Range("B2").Borders(xlEdgeBottom).LineStyle & "|" & Range("B2:C3").Borders.LineStyle
End Function

Function L013() As String
    Reset
    Range("A1").Borders(xlEdgeLeft).LineStyle = xlDash
    Range("A2").Borders(xlEdgeLeft).LineStyle = xlDot
    Range("A3").Borders(xlEdgeLeft).LineStyle = xlDouble
    Range("A4").Borders(xlEdgeLeft).Weight = xlHairline
    L013 = Range("A1").Borders(xlEdgeLeft).LineStyle & "|" & Range("A2").Borders(xlEdgeLeft).LineStyle & "|" & Range("A3").Borders(xlEdgeLeft).LineStyle & "|" & Range("A3").Borders(xlEdgeLeft).Weight & "|" & Range("A4").Borders(xlEdgeLeft).LineStyle & "|" & Range("A4").Borders(xlEdgeLeft).Weight
End Function

Function L014() As String
    Reset
    Range("A1").HorizontalAlignment = xlCenter
    Range("A2").HorizontalAlignment = xlRight
    Range("A3").VerticalAlignment = xlTop
    L014 = Range("A1").HorizontalAlignment & "|" & Range("A2").HorizontalAlignment & "|" & Range("A3").HorizontalAlignment & "|" & Range("A3").VerticalAlignment & "|" & Range("A4").VerticalAlignment & "|" & IsNull(Range("A1:A2").HorizontalAlignment)
End Function

Function L015() As String
    Reset
    Range("A1").WrapText = True
    L015 = Range("A1").WrapText & "|" & Range("A2").WrapText & "|" & IsNull(Range("A1:A2").WrapText)
    Range("A1").HorizontalAlignment = xlCenterAcrossSelection
    L015 = L015 & "|" & Range("A1").HorizontalAlignment & "|" & Range("A1").IndentLevel & "|" & Range("A1").Orientation
End Function

Function L016() As String
    Reset
    Range("B2:D3").Merge
    L016 = Range("B2").MergeCells & "|" & Range("C3").MergeCells & "|" & Range("C3").MergeArea.Address & "|" & Range("B2:D3").MergeCells & "|" & Range("A1:B2").MergeCells & "|" & IsNull(Range("A1:B2").MergeCells) & "|" & Range("E5").MergeArea.Address
End Function

Function L017() As String
    Reset
    Range("B2").Value = "keep"
    Range("C2").Value = "lost"
    Application.DisplayAlerts = False
    Range("B2:C2").Merge
    Application.DisplayAlerts = True
    L017 = Range("B2").Value & "|" & IsEmpty(Range("C2").Value) & "|" & Range("B2:C2").MergeCells
    Range("B2").UnMerge
    L017 = L017 & "|" & Range("B2:C2").MergeCells & "|" & Range("B2").MergeArea.Address
End Function

Function L018() As String
    Reset
    Range("A1:C2").Merge Across:=True
    L018 = Range("A1").MergeArea.Address & "|" & Range("B2").MergeArea.Address & "|" & Range("A1:C2").MergeCells
    Range("A1:C2").MergeCells = False
    L018 = L018 & "|" & Range("A1:C2").MergeCells
    Range("E1:F2").MergeCells = True
    L018 = L018 & "|" & Range("F2").MergeArea.Address
End Function

Function L019() As String
    Reset
    Range("B2:C3").Merge
    Range("B2").Value = 5
    L019 = Range("C3").Value & "|" & Range("B2").Value & "|" & Range("B2:C3").Cells.Count & "|" & Range("B2").MergeArea.Cells(1, 1).Address & "|" & Range("B2").MergeArea.Rows.Count
End Function

Function L020() As String
    Reset
    L020 = Rows(1).RowHeight & "|" & Columns(1).ColumnWidth & "|" & ActiveSheet.StandardWidth & "|" & ActiveSheet.StandardHeight
    Rows(2).RowHeight = 30
    Columns("B").ColumnWidth = 20
    L020 = L020 & "|" & Rows(2).RowHeight & "|" & Range("B5").ColumnWidth & "|" & Range("B2").RowHeight & "|" & IsNull(Range("A1:A2").RowHeight) & "|" & IsNull(Range("A1:B1").ColumnWidth)
End Function

Function L021() As String
    Reset
    Range("A1:A3").RowHeight = 25.5
    Range("C1:E1").ColumnWidth = 3
    L021 = Range("A2").RowHeight & "|" & Range("A1:A3").RowHeight & "|" & Columns("D").ColumnWidth & "|" & Range("A1:A3").Height & "|" & Rows(4).Height
End Function

Function L022() As String
    Reset
    On Error Resume Next
    Rows(1).RowHeight = 500
    L022 = Err.Number & "|" & Rows(1).RowHeight
    Err.Clear
    Columns(1).ColumnWidth = 300
    L022 = L022 & "|" & Err.Number & "|" & Columns(1).ColumnWidth
    Err.Clear
    Rows(1).RowHeight = 0
    L022 = L022 & "|" & Err.Number & "|" & Rows(1).RowHeight & "|" & Rows(1).Hidden
End Function

Function L023() As String
    Reset
    Rows("2:3").Hidden = True
    Columns("C").Hidden = True
    L023 = Rows(2).Hidden & "|" & Rows(4).Hidden & "|" & Columns("C").Hidden & "|" & Range("C5").EntireColumn.Hidden & "|" & Rows(2).RowHeight & "|" & Columns("C").ColumnWidth & "|" & IsNull(Rows("1:2").Hidden)
    Rows("2:3").Hidden = False
    L023 = L023 & "|" & Rows(3).Hidden & "|" & Rows(3).RowHeight
End Function

Function L024() As String
    Reset
    On Error Resume Next
    Range("A1").Hidden = True
    L024 = Err.Number & "|" & Err.Description
    Err.Clear
    L024 = L024 & "|" & Range("A1").EntireRow.Hidden & "|" & Err.Number
End Function

Function L025() As String
    Reset
    Columns(2).ColumnWidth = 0
    L025 = Columns(2).Hidden & "|" & Columns(2).ColumnWidth
    Columns(2).Hidden = False
    L025 = L025 & "|" & Columns(2).ColumnWidth
End Function

Function L026() As String
    Reset
    L026 = Range("A1").Locked & "|" & Range("A1").FormulaHidden
    Range("A1").Locked = False
    Range("A1").FormulaHidden = True
    L026 = L026 & "|" & Range("A1").Locked & "|" & Range("A1").FormulaHidden & "|" & IsNull(Range("A1:A2").Locked)
End Function

Function L027() As String
    Reset
    Range("A1").Value = 1
    Range("A1").Font.Bold = True
    Range("A1").Interior.Color = vbYellow
    Range("A1").NumberFormat = "0.00"
    Range("A1").ClearFormats
    L027 = Range("A1").Value & "|" & Range("A1").Font.Bold & "|" & Range("A1").Interior.ColorIndex & "|" & Range("A1").NumberFormat & "|" & Range("A1").Text
End Function

Function L028() As String
    Reset
    Range("A1").Value = 1
    Range("A1").Font.Italic = True
    Range("A1").Clear
    L028 = IsEmpty(Range("A1").Value) & "|" & Range("A1").Font.Italic
    Range("A2").Value = 2
    Range("A2").Font.Italic = True
    Range("A2").ClearContents
    L028 = L028 & "|" & IsEmpty(Range("A2").Value) & "|" & Range("A2").Font.Italic
End Function

Function L029() As String
    Reset
    Rows(3).Font.Bold = True
    Columns("C").Interior.ColorIndex = 4
    L029 = Range("Z3").Font.Bold & "|" & Range("C100").Interior.ColorIndex & "|" & Range("C3").Font.Bold & "|" & Range("B2").Font.Bold & "|" & Range("C3").Interior.ColorIndex
End Function

Function L030() As String
    Reset
    Cells.Font.Size = 9
    L030 = Range("XFD1048576").Font.Size & "|" & Range("A1").Font.Size
    Cells.Font.Size = 11
    L030 = L030 & "|" & Range("A1").Font.Size
End Function

Function L031() As String
    Reset
    Range("A1").Font.Bold = True
    Range("A1").Interior.Color = vbYellow
    Range("A1").Copy
    Range("C1:C2").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    L031 = Range("C2").Font.Bold & "|" & Range("C2").Interior.Color & "|" & IsEmpty(Range("C2").Value)
End Function

Function L032() As String
    Reset
    Range("A1").Value = "x"
    Range("A1").IndentLevel = 2
    Range("A1").Orientation = 90
    Range("A1").ShrinkToFit = True
    L032 = Range("A1").IndentLevel & "|" & Range("A1").Orientation & "|" & Range("A1").ShrinkToFit & "|" & Range("A1").HorizontalAlignment
End Function

Function L033() As String
    Reset
    Range("A1").Interior.Color = &HFF8000
    Range("A1").Font.Color = &H808080
    L033 = Range("A1").Interior.Color & "|" & Range("A1").Font.Color & "|" & Range("A1").Font.ColorIndex & "|" & Range("A1").Interior.ColorIndex
End Function

Function L034() As String
    Reset
    Range("A1").Borders(xlEdgeTop).ColorIndex = 5
    L034 = Range("A1").Borders(xlEdgeTop).LineStyle & "|" & Range("A1").Borders(xlEdgeTop).Color & "|" & Range("A1").Borders(xlEdgeTop).ColorIndex & "|" & Range("A1").Borders.Count
End Function

Function L035() As String
    Reset
    Range("A1:B2").Borders(xlDiagonalDown).LineStyle = xlContinuous
    L035 = Range("A1").Borders(xlDiagonalDown).LineStyle & "|" & Range("A1").Borders(xlDiagonalUp).LineStyle & "|" & Range("B2").Borders(xlDiagonalDown).LineStyle
End Function
