Option Explicit

Function Show(ByVal v As Variant) As String
    If IsNull(v) Then
        Show = "Null"
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function

Function D01() As String: D01 = Show(CDec(1)): End Function
Function D02() As String: D02 = Show(CDec(0.1)): End Function
Function D03() As String: D03 = Show(CDec("1.50")): End Function
Function D04() As String: D04 = Show(CDec("0.1") * 3): End Function
Function D05() As String: D05 = Show(CDec(0.1) * 3 = 0.3): End Function
Function D06() As String: D06 = Show(CDec("79228162514264337593543950335")): End Function
Function D07() As String
    On Error Resume Next
    D07 = Show(CDec("79228162514264337593543950336"))
    If Err.Number <> 0 Then D07 = "ERR " & Err.Number
End Function
Function D08() As String: D08 = Show(CDec("-79228162514264337593543950335")): End Function
Function D09() As String: D09 = Show(CDec("0.0000000000000000000000000001")): End Function
Function D10() As String: D10 = Show(CDec("1.23456789012345678901234567890")): End Function
Function D11() As String: D11 = Show(CDec(1) / 3): End Function
Function D12() As String: D12 = Show(CDec(2) / 3): End Function
Function D13() As String: D13 = Show(CDec(1) / 8): End Function
Function D14() As String: D14 = Show(CDec(10) / 4): End Function
Function D15() As String: D15 = Show(CDec("1.5") + 1): End Function
Function D16() As String: D16 = Show(CDec("1.5") + CInt(2)): End Function
Function D17() As String: D17 = Show(CDec("1.5") + CLng(2)): End Function
Function D18() As String: D18 = Show(CDec("1.5") + CDbl(2.25)): End Function
Function D19() As String: D19 = Show(CDec("1.5") + CSng(2.5)): End Function
Function D20() As String: D20 = Show(CDec("1.5") + CCur(2.5)): End Function
Function D21() As String: D21 = Show(CDec("1.5") + True): End Function
Function D22() As String: D22 = Show(CDec("1.5") + Empty): End Function
Function D23() As String: D23 = Show(CDec("1.5") + "2.5"): End Function
Function D24() As String: D24 = Show(CDec("1.5") & "x"): End Function
Function D25() As String: D25 = Show(CDec("1.5") - 3): End Function
Function D26() As String: D26 = Show(CDec("1.5") * CDec("1.5")): End Function
Function D27() As String: D27 = Show(CDec(7) \ 2): End Function
Function D28() As String: D28 = Show(CDec("7.5") Mod 2): End Function
Function D29() As String: D29 = Show(CDec(2) ^ 10): End Function
Function D30() As String: D30 = Show(-CDec("1.5")): End Function
Function D31() As String
    On Error Resume Next
    D31 = Show(CDec(1) / 0)
    If Err.Number <> 0 Then D31 = "ERR " & Err.Number
End Function
Function D32() As String
    On Error Resume Next
    D32 = Show(CDec("79228162514264337593543950335") + 1)
    If Err.Number <> 0 Then D32 = "ERR " & Err.Number
End Function
Function D33() As String
    On Error Resume Next
    D33 = Show(CDec("79228162514264337593543950335") * 2)
    If Err.Number <> 0 Then D33 = "ERR " & Err.Number
End Function
Function D34() As String: D34 = Show(CDec("1.10") = CDec("1.1")): End Function
Function D35() As String: D35 = Show(CDec("2") > 1.5): End Function
Function D36() As String: D36 = Show(CDec("0.1") + CDec("0.2") = CDec("0.3")): End Function
Function D37() As String: D37 = Show(VarType(CDec(1))): End Function
Function D38() As String: D38 = Show(IsNumeric(CDec(1))): End Function
Function D39() As String: D39 = Show(CDbl(CDec("1.25"))): End Function
Function D40() As String: D40 = Show(CLng(CDec("2.5"))): End Function
Function D41() As String: D41 = Show(CInt(CDec("3.5"))): End Function
Function D42() As String: D42 = Show(CStr(CDec("1000000"))): End Function
Function D43() As String: D43 = Show(Int(CDec("-2.5"))): End Function
Function D44() As String: D44 = Show(Fix(CDec("-2.5"))): End Function
Function D45() As String: D45 = Show(Abs(CDec("-2.5"))): End Function
Function D46() As String: D46 = Show(Sgn(CDec("-2.5"))): End Function
Function D47() As String: D47 = Show(Round(CDec("2.5"))): End Function
Function D48() As String: D48 = Show(Round(CDec("2.345"), 2)): End Function
Function D49() As String: D49 = Show(Round(CDec("2.355"), 2)): End Function
Function D50() As String: D50 = Show(Format(CDec("1234.5"), "#,##0.00")): End Function
Function D51() As String: D51 = Show(CDec("1E3")): End Function
Function D52() As String: D52 = Show(CDec("1.5E-3")): End Function
Function D53() As String: D53 = Show(CDec(1E+20)): End Function
Function D54() As String: D54 = Show(CDec(123456789.123456)): End Function
Function D55() As String: D55 = Show(CDec(1 / 3)): End Function
Function D56() As String: D56 = Show(CDec(CCur(1.2345))): End Function
Function D57() As String: D57 = Show(CDec(#1/2/2026#)): End Function
Function D58() As String: D58 = Show(CDec(True)): End Function
Function D59() As String
    On Error Resume Next
    D59 = Show(CDec("abc"))
    If Err.Number <> 0 Then D59 = "ERR " & Err.Number
End Function
Function D60() As String: D60 = Show(CDec("  12.5  ")): End Function
Function D61() As String
    Dim v As Variant
    v = CDec("0.5")
    v = v + v
    D61 = Show(v)
End Function
Function D62() As String
    Dim x As Double
    x = CDec("1.25")
    D62 = Show(x)
End Function
Function D63() As String
    Dim s As String
    s = CDec("1.25")
    D63 = Show(s)
End Function
Function D64() As String: D64 = Show(CDec("1.5") * 0): End Function
Function D65() As String: D65 = Show(CDec("-0")): End Function
Function D66() As String: D66 = Show(CDec("100") / CDec("0.25")): End Function
Function D67() As String: D67 = Show(CDec("1") / CDec("7") * CDec("7")): End Function
Function D68() As String: D68 = Show(CDec(3) * CDec("0.3333333333333333333333333333")): End Function
Function D69() As String: D69 = Show(CDec("12345678901234567890") * CDec("1000000000")): End Function
Function D70() As String: D70 = Show(CDec("0.5") + CDec("0.5") = 1): End Function
