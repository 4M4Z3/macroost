Option Explicit

' Range.AutoFill, FillDown/FillRight, DataSeries.

Private Function Col(ByVal c As Long, ByVal r1 As Long, ByVal r2 As Long) As String
    Dim i As Long, s As String
    For i = r1 To r2
        If i > r1 Then s = s & ","
        s = s & CStr(Worksheets("Sheet1").Cells(i, c).Value)
    Next
    Col = s
End Function

Function G001() As String
    Reset
    Range("A1").Value = 1
    Range("A2").Value = 2
    Range("A1:A2").AutoFill Destination:=Range("A1:A6")
    G001 = Col(1, 1, 6)
End Function

Function G002() As String
    Reset
    Range("A1").Value = 5
    Range("A1").AutoFill Destination:=Range("A1:A4"), Type:=xlFillDefault
    G002 = Col(1, 1, 4)
    Range("B1").Value = 5
    Range("B1").AutoFill Destination:=Range("B1:B4"), Type:=xlFillSeries
    G002 = G002 & "|" & Col(2, 1, 4)
End Function

Function G003() As String
    Reset
    Range("A1").Value = 1
    Range("A2").Value = 3
    Range("A1:A2").AutoFill Destination:=Range("A1:A5"), Type:=xlFillCopy
    G003 = Col(1, 1, 5)
    Range("B1:B2").Value = Application.WorksheetFunction.Transpose(Array(10, 7))
    Range("B1:B2").AutoFill Range("B1:B5"), xlFillSeries
    G003 = G003 & "|" & Col(2, 1, 5)
End Function

Function G004() As String
    Reset
    Range("A1:A3").Value = Application.WorksheetFunction.Transpose(Array(1, 2, 3))
    Range("B1").Formula = "=A1*10"
    Range("B1").AutoFill Destination:=Range("B1:B3")
    G004 = Range("B3").Formula & "|" & Range("B3").Value & "|" & Range("B2").Formula
End Function

Function G005() As String
    Reset
    Range("A1").Value = DateSerial(2026, 1, 30)
    Range("A1").AutoFill Destination:=Range("A1:A4"), Type:=xlFillDefault
    G005 = Col(1, 1, 4)
    Range("B1").Value = DateSerial(2026, 1, 31)
    Range("B1").AutoFill Destination:=Range("B1:B3"), Type:=xlFillMonths
    G005 = G005 & "|" & Col(2, 1, 3)
End Function

Function G006() As String
    Reset
    Range("A1").Value = "Item 1"
    Range("A1").AutoFill Destination:=Range("A1:A3")
    G006 = Col(1, 1, 3)
    Range("B1").Value = "Mon"
    Range("B1").AutoFill Destination:=Range("B1:B3")
    G006 = G006 & "|" & Col(2, 1, 3)
    Range("C1").Value = "text"
    Range("C1").AutoFill Destination:=Range("C1:C3")
    G006 = G006 & "|" & Col(3, 1, 3)
End Function

Function G007() As String
    Reset
    Range("A1").Value = 1
    Range("A1").Font.Bold = True
    Range("A1").AutoFill Destination:=Range("A1:A3"), Type:=xlFillFormats
    G007 = Range("A3").Font.Bold & "|" & IsEmpty(Range("A3").Value)
    Range("B1").Value = 2
    Range("B1").Font.Italic = True
    Range("B1").AutoFill Destination:=Range("B1:B3"), Type:=xlFillValues
    G007 = G007 & "|" & Col(2, 1, 3) & "|" & Range("B3").Font.Italic
End Function

Function G008() As String
    Reset
    Range("A1").Value = 1
    Range("B1").Value = 2
    Range("A1:B1").AutoFill Destination:=Range("A1:E1")
    G008 = Range("C1").Value & Range("D1").Value & Range("E1").Value
    Range("A2").Formula = "=A1+1"
    Range("A2").AutoFill Destination:=Range("A2:C2")
    G008 = G008 & "|" & Range("C2").Formula & "|" & Range("C2").Value
End Function

Function G009() As String
    Reset
    Range("A1:A5").Value = 1
    Range("B1").Formula = "=A1*2"
    Range("B1:B5").FillDown
    G009 = Range("B5").Formula & "|" & Range("B5").Value
    Range("C1").Value = "r"
    Range("C1:E1").FillRight
    G009 = G009 & "|" & Range("E1").Value
End Function

Function G010() As String
    Reset
    On Error Resume Next
    Range("A1").Value = 1
    Range("A1").AutoFill Destination:=Range("B5:B9")
    G010 = Err.Number & "|" & Err.Description
End Function

Function G011() As String
    Reset
    Range("A5").Value = 10
    Range("A6").Value = 8
    Range("A5:A6").AutoFill Destination:=Range("A2:A6")
    G011 = Col(1, 2, 6)
End Function

Function G012() As String
    Reset
    Range("A1").Value = 1
    Range("A1:A5").DataSeries Rowcol:=xlColumns, Type:=xlLinear, Step:=2
    G012 = Col(1, 1, 5)
    Range("B1").Value = 1
    Range("B1:B4").DataSeries Rowcol:=xlColumns, Type:=xlGrowth, Step:=3
    G012 = G012 & "|" & Col(2, 1, 4)
End Function

Function G013() As String
    Reset
    Range("A1:A3").Value = Application.WorksheetFunction.Transpose(Array(1, 2, 3))
    Range("B1").Formula = "=SUM($A$1:A1)"
    Range("B1").AutoFill Range("B1:B3")
    G013 = Range("B3").Formula & "|" & Range("B3").Value
End Function

Function G014() As String
    Reset
    Range("A1").Value = 1.5
    Range("A2").Value = 2
    Range("A1:A2").AutoFill Range("A1:A5")
    G014 = Col(1, 1, 5)
    Range("B1").Value = "Q1"
    Range("B1").AutoFill Range("B1:B6")
    G014 = G014 & "|" & Col(2, 1, 6)
End Function
