Option Explicit

' Formulas the workbook calculates, read back through Value; Evaluate and [ ] syntax.

Private Function F(ByVal formula As String) As String
    With Worksheets("Sheet1").Range("Z1")
        .Clear
        .Formula = formula
        If IsError(.Value) Then
            F = .Text
        Else
            F = CStr(.Value)
        End If
    End With
End Function

Private Sub Nums()
    Worksheets("Sheet1").Range("A1:A5").Value = Application.WorksheetFunction.Transpose(Array(10, 20, 30, 40, 50))
    Worksheets("Sheet1").Range("B1:B5").Value = Application.WorksheetFunction.Transpose(Array("x", "y", "x", "z", "x"))
End Sub

Function C001() As String
    Reset: Nums
    C001 = F("=SUM(A1:A5)") & "|" & F("=AVERAGE(A1:A5)") & "|" & F("=MAX(A1:A5)-MIN(A1:A5)") & "|" & F("=COUNT(A1:B5)") & "|" & F("=COUNTA(A1:B5)") & "|" & F("=SUM(A1,A3,100)")
End Function

Function C002() As String
    Reset: Nums
    C002 = F("=SUMIF(B1:B5,""x"",A1:A5)") & "|" & F("=COUNTIF(B1:B5,""x"")") & "|" & F("=SUMIFS(A1:A5,B1:B5,""x"",A1:A5,"">10"")") & "|" & F("=COUNTIFS(A1:A5,"">=20"",A1:A5,""<50"")") & "|" & F("=AVERAGEIF(B1:B5,""x"",A1:A5)")
End Function

Function C003() As String
    Reset: Nums
    C003 = F("=IF(A1>5,""big"",""small"")") & "|" & F("=IF(A1>50,1)") & "|" & F("=IFERROR(1/0,""err"")") & "|" & F("=IF(AND(A1>5,A2<30),""both"",""no"")") & "|" & F("=IF(OR(A1>100,A2=20),1,0)") & "|" & F("=NOT(TRUE)")
End Function

Function C004() As String
    Reset: Nums
    C004 = F("=VLOOKUP(30,A1:B5,2,FALSE)") & "|" & F("=INDEX(B1:B5,4)") & "|" & F("=MATCH(40,A1:A5,0)") & "|" & F("=INDEX(A1:A5,MATCH(""z"",B1:B5,0))") & "|" & F("=VLOOKUP(35,A1:B5,2,TRUE)") & "|" & F("=VLOOKUP(99,A1:B5,2,FALSE)")
End Function

Function C005() As String
    Reset
    C005 = F("=ROUND(2.567,1)") & "|" & F("=ROUNDUP(2.01,0)") & "|" & F("=ROUNDDOWN(2.99,0)") & "|" & F("=INT(-3.5)") & "|" & F("=MOD(17,5)") & "|" & F("=ABS(-4)") & "|" & F("=SQRT(81)") & "|" & F("=POWER(2,8)") & "|" & F("=2^10")
End Function

Function C006() As String
    Reset
    C006 = F("=LEFT(""Hello"",2)") & "|" & F("=RIGHT(""Hello"",3)") & "|" & F("=MID(""Hello"",2,3)") & "|" & F("=LEN(""Hello"")") & "|" & F("=UPPER(""abc"")") & "|" & F("=LOWER(""ABC"")") & "|" & F("=PROPER(""john smith"")") & "|" & F("=TRIM(""  a  b "")")
End Function

Function C007() As String
    Reset
    C007 = F("=CONCATENATE(""a"",""b"",1)") & "|" & F("=""a""&""b""&2") & "|" & F("=SUBSTITUTE(""a.b.c"",""."",""-"")") & "|" & F("=FIND(""l"",""Hello"")") & "|" & F("=SEARCH(""L"",""Hello"")") & "|" & F("=REPT(""-"",3)") & "|" & F("=TEXT(1234.5,""#,##0.0"")") & "|" & F("=VALUE(""12.5"")")
End Function

Function C008() As String
    Reset
    C008 = F("=DATE(2026,10,14)") & "|" & F("=YEAR(DATE(2026,10,14))") & "|" & F("=MONTH(DATE(2026,10,14))") & "|" & F("=DAY(DATE(2026,10,14))") & "|" & F("=WEEKDAY(DATE(2026,10,14))") & "|" & F("=EDATE(DATE(2026,1,31),1)") & "|" & F("=EOMONTH(DATE(2026,2,1),0)") & "|" & F("=DATEDIF(DATE(2026,1,1),DATE(2026,10,14),""d"")")
End Function

Function C009() As String
    Reset
    Range("A1").Value = 5
    Range("B1").Formula = "=A1*2"
    Range("C1").Formula = "=B1+A1"
    C009 = Range("C1").Value
    Range("A1").Value = 10
    C009 = C009 & "|" & Range("C1").Value & "|" & Range("B1").Value
End Function

Function C010() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Rates"
    Worksheets("Rates").Range("A1").Value = 0.25
    Worksheets("Rates").Range("B1:B3").Value = 4
    Worksheets("Sheet1").Range("A1").Formula = "=Rates!A1*100"
    Worksheets("Sheet1").Range("A2").Formula = "=SUM(Rates!B1:B3)"
    Worksheets("Sheet1").Range("A3").Formula = "='Rates'!$A$1+1"
    C010 = Worksheets("Sheet1").Range("A1").Value & "|" & Worksheets("Sheet1").Range("A2").Value & "|" & Worksheets("Sheet1").Range("A3").Value
End Function

Function C011() As String
    Reset
    Application.Calculation = xlCalculationManual
    Range("A1").Value = 1
    Range("B1").Formula = "=A1+1"
    Range("A1").Value = 5
    C011 = Range("B1").Value
    Calculate
    C011 = C011 & "|" & Range("B1").Value & "|" & Application.Calculation
    Application.Calculation = xlCalculationAutomatic
    C011 = C011 & "|" & Application.Calculation
End Function

Function C012() As String
    Reset: Nums
    C012 = Application.Evaluate("SUM(A1:A5)") & "|" & Application.Evaluate("=A2*2") & "|" & Evaluate("1+2*3") & "|" & [A3] & "|" & [SUM(A1:A2)] & "|" & Application.Evaluate("A1:A2").Address & "|" & [B1:B2].Address
End Function

Function C013() As String
    Reset
    C013 = Evaluate("""a""&""b""") & "|" & Evaluate("LEN(""abc"")") & "|" & ErrCode(Evaluate("1/0")) & "|" & Evaluate("TRUE") & "|" & Evaluate("MAX(3,7)") & "|" & ErrCode(Evaluate("NOPE()"))
End Function

Function C014() As String
    Reset: Nums
    C014 = Worksheets("Sheet1").Evaluate("A1+A2") & "|" & ActiveSheet.Evaluate("COUNTIF(B1:B5,""x"")") & "|" & TypeName(Evaluate("A1")) & "|" & TypeName(Evaluate("{1,2,3}")) & "|" & UBound(Evaluate("{1,2,3}"))
End Function

Function C015() As String
    Reset
    Range("A1:A3").Value = 2
    Range("A4").Formula = "=SUM(A1:A3)"
    Range("A1").Value = 10
    [A2] = 20
    C015 = Range("A4").Value & "|" & [A4].Value & "|" & [A4].Formula
End Function

Function C016() As String
    Reset
    Range("A1").Formula = "=1/0"
    Range("A2").Formula = "=A1+1"
    Range("A3").Formula = "=ISERROR(A1)"
    Range("A4").Formula = "=IFERROR(A2,-1)"
    Range("A5").Formula = "=ISNUMBER(A4)"
    C016 = Range("A2").Text & "|" & Range("A3").Value & "|" & Range("A4").Value & "|" & Range("A5").Value & "|" & F("=ISBLANK(A9)") & "|" & F("=NA()")
End Function

Function C017() As String
    Reset
    Range("A1").Value = "10"
    Range("A2").Value = 5
    Range("A1").NumberFormat = "@"
    Range("A1").Value = "10"
    C017 = F("=A1+A2") & "|" & F("=SUM(A1:A2)") & "|" & F("=A1&A2") & "|" & F("=A1=10") & "|" & F("=""10""=A1")
End Function

Function C018() As String
    Reset
    C018 = F("=TRUE+1") & "|" & F("=""3""*2") & "|" & F("=-2^2") & "|" & F("=10/4") & "|" & F("=7%") & "|" & F("=1=1") & "|" & F("=""a""<""B""") & "|" & F("=""abc""=""ABC""")
End Function

Function C019() As String
    Reset
    Range("A1:A4").Value = Application.WorksheetFunction.Transpose(Array(3, 1, 4, 1))
    C019 = F("=LARGE(A1:A4,2)") & "|" & F("=SMALL(A1:A4,1)") & "|" & F("=RANK(4,A1:A4)") & "|" & F("=SUMPRODUCT(A1:A4,A1:A4)") & "|" & F("=PRODUCT(A1:A4)") & "|" & F("=MEDIAN(A1:A4)") & "|" & F("=COUNTBLANK(A1:A6)")
End Function

Function C020() As String
    Reset
    Range("A1").Value = 1
    Range("A2").Formula = "=A1+1"
    Range("A2").Copy Range("A3:A6")
    C020 = Range("A6").Value & "|" & Range("A6").Formula & "|" & F("=SUM(A1:A6)") & "|" & F("=ROWS(A1:A6)") & "|" & F("=COLUMNS(A1:C1)") & "|" & F("=ROW(A5)") & "|" & F("=COLUMN(C1)")
End Function

Function C021() As String
    Reset
    Range("A1").Value = DateSerial(2026, 10, 14)
    Columns("B").ColumnWidth = 20
    Range("B1").Formula = "=A1+30"
    Range("C1").Formula = "=TEXT(A1,""mmm d"")"
    C021 = Range("B1").Text & "|" & TypeName(Range("B1").Value) & "|" & Range("C1").Value & "|" & Range("B1").NumberFormat
End Function

Function C022() As String
    Reset
    Range("A1:C1").Value = Array(1, 2, 3)
    Range("A2").FormulaR1C1 = "=SUM(R1C1:R1C3)"
    Range("B2").FormulaR1C1 = "=R[-1]C*2"
    Range("A3:C3").FormulaR1C1 = "=R[-2]C+1"
    C022 = Range("A2").Value & "|" & Range("B2").Value & "|" & Range("C3").Formula & "|" & Range("C3").Value
End Function

Function C023() As String
    Reset
    Range("A1").Formula = "=CHOOSE(2,""a"",""b"",""c"")"
    Range("A2").Formula = "=HLOOKUP(""b"",{""a"",""b"";1,2},2,FALSE)"
    Range("A3").Formula = "=SUM({1,2,3})"
    Range("A4").Formula = "=MAX(-1,MIN(5,9))"
    C023 = Range("A1").Value & "|" & Range("A2").Value & "|" & Range("A3").Value & "|" & Range("A4").Value
End Function

Function C024() As String
    Reset
    Range("A1").Formula = "=OFFSET(B1,1,1)"
    Range("C2").Value = "off"
    Range("A2").Formula = "=INDIRECT(""C2"")"
    Range("A3").Formula = "=ADDRESS(2,3)"
    Range("A4").Formula = "=COUNTA(B:C)"
    C024 = Range("A1").Value & "|" & Range("A2").Value & "|" & Range("A3").Value & "|" & Range("A4").Value
End Function

Function C025() As String
    Reset
    Range("A1").Formula = "=SUM(A2:A3)"
    Range("A2").Value = 1
    Range("A3").Formula = "=A2*3"
    Range("A2").Value = 2
    C025 = Range("A1").Value & "|" & Application.WorksheetFunction.Sum(Range("A1:A3"))
End Function
