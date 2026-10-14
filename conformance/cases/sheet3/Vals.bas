Option Explicit

' Value, Value2, Formula, FormulaR1C1, Text, HasFormula and typed-text coercion.

' Wide columns, so Text shows the number format rather than the column-fit (####) rule.
Private Sub Prep()
    Reset
    Worksheets("Sheet1").Columns("A:E").ColumnWidth = 60
End Sub


Private Function Desc(ByVal r As Range) As String
    Desc = TypeName(r.Value) & ":" & CStr(r.Value2) & ":" & r.Text & ":" & r.NumberFormat
End Function

Function V001() As String
    Prep
    Range("A1").Value = "123"
    Range("A2").Value = "1e3"
    Range("A3").Value = " 12 "
    Range("A4").Value = "1,234"
    Range("A5").Value = "-5.5"
    V001 = Desc(Range("A1")) & "|" & Desc(Range("A2")) & "|" & Desc(Range("A3")) & "|" & Desc(Range("A4")) & "|" & Desc(Range("A5"))
End Function

Function V002() As String
    Prep
    Range("A1").Value = "50%"
    Range("A2").Value = "12.5%"
    Range("A3").Value = "$1,234.50"
    Range("A4").Value = "$7"
    Range("A5").Value = "(42)"
    V002 = Desc(Range("A1")) & "|" & Desc(Range("A2")) & "|" & Desc(Range("A3")) & "|" & Desc(Range("A4")) & "|" & Desc(Range("A5"))
End Function

Function V003() As String
    Prep
    Range("A1").Value = "12/31/2026"
    Range("A2").Value = "2026-03-04"
    Range("A3").Value = "3:30 PM"
    Range("A4").Value = "14:05"
    Range("A5").Value = "1/2/2026 10:30"
    V003 = Desc(Range("A1")) & "|" & Desc(Range("A2")) & "|" & Desc(Range("A3")) & "|" & Desc(Range("A4")) & "|" & Desc(Range("A5"))
End Function

Function V004() As String
    Prep
    Range("A1").Value = "TRUE"
    Range("A2").Value = "false"
    Range("A3").Value = "#N/A"
    Range("A4").Value = "'0012"
    Range("A5").Value = "0012"
    V004 = Desc(Range("A1")) & "|" & Desc(Range("A2")) & "|" & TypeName(Range("A3").Value) & ":" & Range("A3").Text & "|" & Desc(Range("A4")) & "|" & Desc(Range("A5"))
End Function

Function V005() As String
    Prep
    Range("A1").Value = 5
    Range("A2").Value = CLng(5)
    Range("A3").Value = CCur(5.25)
    Range("A4").Value = CDate("2026-10-14")
    Range("A5").Value = True
    V005 = Desc(Range("A1")) & "|" & Desc(Range("A2")) & "|" & Desc(Range("A3")) & "|" & Desc(Range("A4")) & "|" & Desc(Range("A5"))
End Function

Function V006() As String
    Prep
    Range("A1").Value = #10/14/2026 6:30:00 PM#
    Range("A2").Value = #6:00:00 AM#
    Range("A3").Value2 = 46309
    V006 = Desc(Range("A1")) & "|" & Desc(Range("A2")) & "|" & Desc(Range("A3"))
End Function

Function V007() As String
    Prep
    Range("A1").NumberFormat = "@"
    Range("A1").Value = "00123"
    Range("A2").NumberFormat = "@"
    Range("A2").Value = 45
    Range("A3").NumberFormat = "@"
    Range("A3").Value = "=1+1"
    V007 = Desc(Range("A1")) & "|" & Desc(Range("A2")) & "|" & Desc(Range("A3")) & "|" & Range("A3").HasFormula
End Function

Function V008() As String
    Prep
    Range("A1").Value = "=2*3"
    Range("A2").Formula = "=A1+1"
    Range("A3").FormulaR1C1 = "=R[-1]C*2"
    Range("A4").Value = "=SUM(A1:A3)"
    V008 = Range("A1").Value & "|" & Range("A2").Value & "|" & Range("A3").Formula & "|" & Range("A3").Value & "|" & Range("A4").FormulaR1C1 & "|" & Range("A4").Value
End Function

Function V009() As String
    Prep
    Range("B2").Formula = "=$A$1+A1+$A1+A$1"
    V009 = Range("B2").FormulaR1C1 & "|" & Range("B2").Formula
    Range("B2").FormulaR1C1 = "=RC[-1]+R1C1+R[1]C[1]+R2C"
    V009 = V009 & "|" & Range("B2").Formula
End Function

Function V010() As String
    Prep
    Range("A1:A3").Formula = "=ROW()*10"
    Range("B1:B3").Formula = "=A1+1"
    V010 = Range("A3").Value & "|" & Range("B3").Formula & "|" & Range("B2").Value & "|" & Range("B1:B3").HasFormula
End Function

Function V011() As String
    Prep
    Range("A1").Value = 1
    Range("A2").Formula = "=A1"
    V011 = Range("A1:A2").HasFormula & "|" & IsNull(Range("A1:A2").HasFormula) & "|" & Range("A1").HasFormula & "|" & Range("A2").HasFormula & "|" & Range("A3").HasFormula
End Function

Function V012() As String
    Prep
    Range("A1").Value = 1234.5678
    V012 = Range("A1").Formula & "|" & Range("A1").FormulaR1C1 & "|" & Range("A1").FormulaLocal & "|" & Range("A1").Text & "|" & Range("A1").Value2
End Function

Function V013() As String
    Prep
    Range("A1").Value = "hello"
    Range("A2").Value = True
    Range("A3").Formula = "=1/0"
    V013 = Range("A1").Formula & "|" & Range("A2").Formula & "|" & Range("A3").Text & "|" & Range("A3").Formula & "|" & CStr(Range("A3").Value = CVErr(xlErrDiv0))
End Function

Function V014() As String
    Dim v As Variant
    Prep
    Range("A1").Formula = "=NA()"
    Range("A2").Formula = "=1/0"
    Range("A3").Formula = "=""a""+1"
    Range("A4").Formula = "=NOSUCHFN(1)"
    Range("A5").Formula = "=A1048577"
    v = Range("A1").Value
    V014 = IsError(v) & "|" & ErrCode(Range("A1").Value) & "|" & ErrCode(Range("A2").Value) & "|" & ErrCode(Range("A3").Value) & "|" & ErrCode(Range("A5").Value) & "|" & Range("A4").Text & "|" & Range("A3").Text
End Function

Function V015() As String
    Dim a As Variant
    Prep
    Range("A1:B2").Formula = "=1"
    a = Range("A1:B2").Formula
    V015 = TypeName(a) & "|" & LBound(a, 1) & "|" & UBound(a, 2) & "|" & a(2, 2)
End Function

Function V016() As String
    Dim a(1 To 3) As Variant
    Prep
    a(1) = 1: a(2) = "b": a(3) = 3
    Range("A1:C1").Value = a
    Range("A2:E2").Value = a
    V016 = Range("C1").Value & "|" & Range("D2").Text & "|" & IsError(Range("E2").Value)
End Function

Function V017() As String
    Dim a(0 To 1, 0 To 1) As Variant
    Prep
    a(0, 0) = 1: a(0, 1) = 2: a(1, 0) = 3: a(1, 1) = 4
    Range("B2:C3").Value = a
    V017 = Range("B2").Value & Range("C2").Value & Range("B3").Value & Range("C3").Value
    Range("B5:D7").Value = a
    V017 = V017 & "|" & Range("D5").Text & "|" & Range("B7").Text
End Function

Function V018() As String
    Prep
    Range("A1:A3").Value = Application.WorksheetFunction.Transpose(Array(1, 2, 3))
    V018 = Range("A1").Value & Range("A2").Value & Range("A3").Value & "|" & TypeName(Application.WorksheetFunction.Transpose(Array(1, 2, 3)))
End Function

Function V019() As String
    Dim v As Variant
    Prep
    Range("A1").Value = 1
    v = Range("A1:A1").Value
    V019 = TypeName(v)
    v = Range("A1:A2").Value
    V019 = V019 & "|" & TypeName(v) & "|" & IsEmpty(v(2, 1))
    v = Range("A1:A2").Text
    V019 = V019 & "|" & TypeName(v) & "|" & IsNull(v)
End Function

Function V020() As String
    Prep
    Range("A1").Value = 0.1 + 0.2
    Range("A2").Value = 1E+20
    Range("A3").Value = 1.23456789012345E-05
    Range("A4").Value = 123456789012
    Range("A5").Value = -0.000001
    V020 = Range("A1").Text & "|" & Range("A2").Text & "|" & Range("A3").Text & "|" & Range("A4").Text & "|" & Range("A5").Text
End Function

Function V021() As String
    Prep
    Range("A1").Value = 12345678901234#
    Range("A2").Value = 1234567.891
    Range("A3").Value = 3.14159265358979
    Range("A4").Value = 100
    Range("A5").Value = 2.5E-10
    V021 = Range("A1").Text & "|" & Range("A2").Text & "|" & Range("A3").Text & "|" & Range("A4").Text & "|" & Range("A5").Text & "|" & Range("A3").Value
End Function

Function V022() As String
    Prep
    Range("A1").Value = "line1" & vbLf & "line2"
    V022 = Len(Range("A1").Value) & "|" & Range("A1").WrapText & "|" & InStr(Range("A1").Text, vbLf)
End Function

Function V023() As String
    Prep
    Range("A1").Value = ""
    Range("A2").Value = Empty
    Range("A3").Formula = ""
    V023 = IsEmpty(Range("A1").Value) & "|" & IsEmpty(Range("A2").Value) & "|" & IsEmpty(Range("A3").Value) & "|" & Application.WorksheetFunction.CountA(Range("A1:A3"))
End Function

Function V024() As String
    Prep
    Range("A1").Value = "  padded  "
    Range("A2").Value = "=  1 + 2"
    Range("A3").Value = "+5"
    Range("A4").Value = "-"
    Range("A5").Value = "="
    V024 = "[" & Range("A1").Value & "]|" & Range("A2").Formula & "|" & Range("A2").Value & "|" & Desc(Range("A3")) & "|" & Desc(Range("A4")) & "|" & Desc(Range("A5"))
End Function

Function V025() As String
    Prep
    Range("A1").Value = "Jan-2026"
    Range("A2").Value = "3/4/5"
    Range("A3").Value = "Jan 5, 2026"
    Range("A4").Value = "5-Mar-2026"
    Range("A5").Value = "0 1/2"
    V025 = Desc(Range("A1")) & "|" & Desc(Range("A2")) & "|" & Desc(Range("A3")) & "|" & Desc(Range("A4")) & "|" & Desc(Range("A5"))
End Function

Function V026() As String
    Prep
    Range("A1").Value = 45000
    Range("A1").NumberFormat = "m/d/yyyy"
    V026 = TypeName(Range("A1").Value) & "|" & TypeName(Range("A1").Value2) & "|" & Range("A1").Value & "|" & Range("A1").Value2 & "|" & Range("A1").Text
End Function

Function V027() As String
    Prep
    Range("A1").Value = 12.345
    Range("A1").NumberFormat = "$#,##0.00"
    V027 = TypeName(Range("A1").Value) & "|" & TypeName(Range("A1").Value2) & "|" & Range("A1").Value & "|" & Range("A1").Text
    Range("A1").NumberFormat = "0.00"
    V027 = V027 & "|" & TypeName(Range("A1").Value)
End Function

Function V028() As String
    Prep
    Range("A1").Formula = "=TODAY()-TODAY()"
    Range("A2").Formula = "=DATE(2026,10,14)"
    Range("A3").Formula = "=TIME(18,30,0)"
    V028 = Range("A1").Value & "|" & TypeName(Range("A2").Value) & "|" & Range("A2").Text & "|" & Range("A2").NumberFormat & "|" & Range("A3").Text & "|" & Range("A3").NumberFormat
End Function

Function V029() As String
    Prep
    Range("A1").Value = 5
    Range("A1").Value = Range("A1").Value + 1
    Range("A2").Value = Range("A1")
    Range("A3") = Range("A2") * 2
    V029 = Range("A1").Value & "|" & Range("A2").Value & "|" & Range("A3").Value & "|" & TypeName(Range("A3").Value)
End Function

Function V030() As String
    Prep
    Cells(1, 1).Value = "x"
    Cells(1, 1).Value = Cells(1, 1).Value & "y"
    Range("A2").Value = Len(Range("A1"))
    V030 = Range("A1") & "|" & Range("A2") & "|" & Range("A1").Value2
End Function

Function V031() As String
    Prep
    Range("A1").Value = "=A2"
    Range("A2").Value = 7
    V031 = Range("A1").Value & "|" & Range("A1").Text
    Range("A2").Value = "s"
    V031 = V031 & "|" & Range("A1").Value
    Range("A2").ClearContents
    V031 = V031 & "|" & Range("A1").Value & "|" & TypeName(Range("A1").Value)
End Function

Function V032() As String
    Prep
    Range("A1").Value = 1
    Range("A1").Formula = 2
    Range("A2").Formula = "text"
    Range("A3").FormulaR1C1 = 3.5
    V032 = TypeName(Range("A1").Value) & "|" & Range("A1").HasFormula & "|" & Range("A2").Value & "|" & Range("A3").Value & "|" & Range("A3").HasFormula
End Function

Function V033() As String
    Prep
    Range("A1").Value = 100
    Range("A1").Value = Range("A1").Value * 1.1
    V033 = Range("A1").Value & "|" & Range("A1").Text & "|" & (Range("A1").Value = 110)
End Function

Function V034() As String
    Prep
    Range("A1").Value = "abc"
    Range("A2").Value = 1
    Range("A3").Formula = "=A1&A2"
    Range("A4").Formula = "=A2&""x"""
    Range("A5").Formula = "=LEN(A3)"
    V034 = Range("A3").Value & "|" & Range("A4").Value & "|" & Range("A5").Value & "|" & TypeName(Range("A5").Value)
End Function

Function V035() As String
    Prep
    Range("A1:C1").Value = Array("a", 2, True)
    V035 = Range("A1").Value & "|" & TypeName(Range("B1").Value) & "|" & TypeName(Range("C1").Value) & "|" & Range("C1").Text
End Function
