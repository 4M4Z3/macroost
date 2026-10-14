Option Explicit

' Range.Text under number formats: numbers, percents, currency, accounting, dates, times,
' scientific, fractions, text, custom sections and colors.

' Wide columns, so Text shows the number format rather than the column-fit (####) rule.
Private Sub Prep()
    Reset
    Worksheets("Sheet1").Columns("A:E").ColumnWidth = 60
End Sub


Private Function TX(ByVal v As Variant, ByVal fmt As String) As String
    TX = Application.WorksheetFunction.Text(v, fmt)
End Function

Private Function T(ByVal v As Variant, ByVal fmt As String) As String
    With Worksheets("Sheet1").Range("A1")
        .ClearFormats
        .Value = v
        .NumberFormat = fmt
        T = .Text
    End With
End Function

Function F001() As String
    Prep
    F001 = T(1234.567, "0") & "|" & T(1234.567, "0.00") & "|" & T(1234.567, "#,##0") & "|" & T(1234.567, "#,##0.00") & "|" & T(0.5, "0") & "|" & T(-0.4, "0") & "|" & T(2.5, "0") & "|" & T(0.123, "#.##")
End Function

Function F002() As String
    Prep
    F002 = T(0.256, "0%") & "|" & T(0.256, "0.0%") & "|" & T(1.5, "0.00%") & "|" & T(-0.05, "0%") & "|" & T(0, "0%")
End Function

Function F003() As String
    Prep
    F003 = T(1234.5, "$#,##0.00") & "|" & T(-1234.5, "$#,##0.00") & "|" & T(-1234.5, "$#,##0.00_);($#,##0.00)") & "|" & T(-1234.5, "$#,##0.00_);[Red]($#,##0.00)") & "|" & T(0, "$#,##0")
End Function

Function F004() As String
    Prep
    F004 = "[" & TX(1234.5, "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)") & "]|[" & TX(-1234.5, "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)") & "]|[" & TX(0, "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)") & "]|[" & TX("txt", "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)") & "]"
End Function

Function F005() As String
    Prep
    F005 = "[" & TX(1234.5, "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)") & "]|[" & TX(-7, "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)") & "]|[" & TX(0, "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)") & "]"
End Function

Function F006() As String
    Dim d As Double
    Prep
    d = DateSerial(2026, 3, 7) + TimeSerial(14, 5, 9)
    F006 = T(d, "m/d/yyyy") & "|" & T(d, "mm/dd/yyyy") & "|" & T(d, "d-mmm-yy") & "|" & T(d, "dd-mmm-yyyy") & "|" & T(d, "mmm-yy") & "|" & T(d, "mmmm d, yyyy") & "|" & T(d, "dddd") & "|" & T(d, "ddd, mmm d")
End Function

Function F007() As String
    Dim d As Double
    Prep
    d = DateSerial(2026, 3, 7) + TimeSerial(14, 5, 9)
    F007 = T(d, "h:mm") & "|" & T(d, "h:mm AM/PM") & "|" & T(d, "hh:mm:ss") & "|" & T(d, "h:mm:ss AM/PM") & "|" & T(d, "m/d/yyyy h:mm") & "|" & T(d, "yyyy-mm-dd hh:mm:ss") & "|" & T(d, "h AM/PM") & "|" & T(d, "mm:ss")
End Function

Function F008() As String
    Prep
    F008 = T(1.5, "[h]:mm") & "|" & T(2.75, "[h]:mm:ss") & "|" & T(0.0625, "[mm]:ss") & "|" & T(0.5, "h:mm A/P") & "|" & T(0.999, "hh:mm")
End Function

Function F009() As String
    Prep
    F009 = T(12345.678, "0.00E+00") & "|" & T(0.000123, "0.00E+00") & "|" & T(12345.678, "##0.0E+0") & "|" & T(-5, "0E+00") & "|" & T(1, "0.0E-0")
End Function

Function F010() As String
    Prep
    F010 = T(1.5, "# ?/?") & "|" & T(0.333, "# ??/??") & "|" & T(2.25, "?/?") & "|" & T(5, "# ?/?") & "|" & T(0.125, "# ?/8") & "|" & T(-1.75, "# ?/4")
End Function

Function F011() As String
    Prep
    F011 = T(123, "@") & "|" & T("abc", "@") & "|" & T("abc", """pre ""@") & "|" & T(5, """Qty: ""0") & "|" & T(5, "0"" units""") & "|" & T(3, "0\x")
End Function

Function F012() As String
    Prep
    F012 = T(5, "0;(0);""zero""") & "|" & T(-5, "0;(0);""zero""") & "|" & T(0, "0;(0);""zero""") & "|" & T("t", "0;(0);""zero"";""text:""@") & "|" & T(-5, "0;0") & "|" & T(-5, "0")
End Function

Function F013() As String
    Prep
    F013 = T(5, "[Red]0") & "|" & T(-5, "0;[Red]-0") & "|" & T(5, "[Blue]0.0;[Red]-0.0") & "|" & T(150, "[>100]""big"";[<=100]""small""") & "|" & T(50, "[>100]""big"";[<=100]""small""")
End Function

Function F014() As String
    Prep
    F014 = T(123456789, "000-00-0000") & "|" & T(5551234, "###-####") & "|" & T(42, "00000") & "|" & T(3.14159, "0.000") & "|" & T(7, "000.0")
End Function

Function F015() As String
    Prep
    F015 = T(1234567.89, "#,##0,") & "|" & T(1234567.89, "#,##0.0,,") & "|" & T(1234567.89, "#,##0,,""M""") & "|" & T(999, "#,##0,")
End Function

Function F016() As String
    Prep
    F016 = T(12345678901#, "General") & "|" & T(123456789012#, "General") & "|" & T(0.00001234, "General") & "|" & T(1234.5678901234, "General") & "|" & T(-1E-20, "General") & "|" & T(1E+100, "General")
End Function

Function F017() As String
    Prep
    F017 = T(True, "0.00") & "|" & T("text", "0.00") & "|" & T(CVErr(xlErrNA), "0.00") & "|" & T(False, "General")
End Function

Function F018() As String
    Dim d As Double
    Prep
    d = DateSerial(2026, 12, 25)
    F018 = T(d, "yyyy") & "|" & T(d, "yy") & "|" & T(d, "mmmmm") & "|" & T(d, "m") & "|" & T(d, "d") & "|" & T(d, "dd") & "|" & T(d, "ddd") & "|" & T(d, "[$-409]mmmm d, yyyy")
End Function

Function F019() As String
    Prep
    F019 = T(0, "m/d/yyyy") & "|" & T(1, "m/d/yyyy") & "|" & T(60, "m/d/yyyy") & "|" & T(61, "m/d/yyyy") & "|" & T(0.5, "m/d/yyyy h:mm")
End Function

Function F020() As String
    Prep
    F020 = T(-1, "m/d/yyyy") & "|" & T(3000000, "m/d/yyyy") & "|" & T(2958465, "m/d/yyyy")
End Function

Function F021() As String
    Prep
    F021 = T(1234.5, "#,##0.00 ""USD""") & "|" & T(0.75, "0%"" done""") & "|" & T(12, "0 ""items""") & "|" & T(-3.2, "0.0 ;(0.0)")
End Function

Function F022() As String
    Prep
    F022 = T(1234.5, "[$$-409]#,##0.00") & "|" & T(1234.5, "#,##0.00 [$USD]")
End Function

Function F023() As String
    Prep
    F023 = T(1.005, "0.00") & "|" & T(2.675, "0.00") & "|" & T(-2.5, "0") & "|" & T(0.045, "0.0%") & "|" & T(1234.5, "#,##0")
End Function

Function F024() As String
    Prep
    F024 = T(0.000001, "0.00") & "|" & T(-0.001, "0.00") & "|" & T(-0.001, "0.00;-0.00;""z""") & "|" & T(1000000, "#,###") & "|" & T(0, "#,###") & "|" & T(0, "#,##0")
End Function

Function F025() As String
    Prep
    F025 = T(45, "0_)") & "|" & TX(45, "0* ") & "|" & T(7, "?0") & "|" & T(7.5, "0.0?") & "|" & T(7.25, "0.0?") & "|" & T(7, "#.??")
End Function

Function F026() As String
    Prep
    With Range("B1")
        .Value = 0.5
        .NumberFormat = "0.0%"
        F026 = .Text & "|" & .Value & "|" & .NumberFormat
        .NumberFormat = "General"
        F026 = F026 & "|" & .Text & "|" & .NumberFormat
    End With
End Function

Function F027() As String
    Prep
    Range("A1:A3").NumberFormat = "0.00"
    Range("A2").NumberFormat = "0%"
    F027 = Range("A1").NumberFormat & "|" & IsNull(Range("A1:A3").NumberFormat) & "|" & Range("A1,A3").NumberFormat & "|" & Range("C1:C3").NumberFormat
End Function

Function F028() As String
    Prep
    Range("A1").Value = 1234.5
    Range("A1").Style = "Currency"
    F028 = Range("A1").NumberFormat & "|" & Range("A1").Style
    Range("A1").Style = "Percent"
    F028 = F028 & "|" & Range("A1").NumberFormat & "|" & Range("A1").Text
    Range("A1").Style = "Comma"
    F028 = F028 & "|" & Range("A1").NumberFormat & "|" & Range("A1").Style
End Function

Function F029() As String
    Prep
    Range("A1").Value = 3
    Range("A1").NumberFormat = "0"
    Range("A1").Value = 3.7
    F029 = Range("A1").Text & "|" & Range("A1").Value
    Range("A2").Value = "4.5"
    Range("A2").NumberFormat = "0"
    F029 = F029 & "|" & Range("A2").Text
End Function

Function F030() As String
    Prep
    F030 = Format(1234.5, "#,##0.00") & "|" & Application.WorksheetFunction.Text(1234.5, "#,##0.00") & "|" & Application.WorksheetFunction.Text(0.25, "0%") & "|" & Application.WorksheetFunction.Text(DateSerial(2026, 1, 9), "mmm d, yyyy") & "|" & Application.WorksheetFunction.Text(-3, "0;(0)")
End Function

Function F031() As String
    Prep
    F031 = T(-1234.5, "#,##0.00;(#,##0.00)") & "|" & T(0.1, "0.00;-0.00;0.00") & "|" & T(-0.004, "0.00") & "|" & T(12, "0.0#") & "|" & T(12.345, "0.0#")
End Function

Function F032() As String
    Prep
    F032 = T(0.5, "h:mm:ss.0") & "|" & T(1.000011574, "s.00") & "|" & T(36526.75, "dd/mm/yy hh:mm") & "|" & T(36526, "dddd, mmmm dd, yyyy")
End Function
