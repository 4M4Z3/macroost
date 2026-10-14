Option Explicit

Function Show(ByVal v As Variant) As String
    If IsArray(v) Then
        Show = "Array(" & LBound(v) & " To " & UBound(v) & ")"
    ElseIf IsNull(v) Then
        Show = "Null"
    ElseIf IsError(v) Then
        Show = "Error:" & CStr(CLng(v))
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function

Function C001() As String
    C001 = Show(1 + 1)
End Function

Function C002() As String
    C002 = Show(32767 + 1)
End Function

Function C003() As String
    C003 = Show(32767& + 1)
End Function

Function C004() As String
    C004 = Show(300 * 200)
End Function

Function C005() As String
    C005 = Show(300& * 200)
End Function

Function C006() As String
    C006 = Show(2147483647 + 1)
End Function

Function C007() As String
    C007 = Show(7 / 2)
End Function

Function C008() As String
    C008 = Show(7 \ 2)
End Function

Function C009() As String
    C009 = Show(-7 \ 2)
End Function

Function C010() As String
    C010 = Show(7 Mod 3)
End Function

Function C011() As String
    C011 = Show(-7 Mod 3)
End Function

Function C012() As String
    C012 = Show(7.5 Mod 2)
End Function

Function C013() As String
    C013 = Show(6.5 \ 2)
End Function

Function C014() As String
    C014 = Show(2 ^ 10)
End Function

Function C015() As String
    C015 = Show(-2 ^ 2)
End Function

Function C016() As String
    C016 = Show((-8) ^ (1 / 3))
End Function

Function C017() As String
    C017 = Show(1 / 0)
End Function

Function C018() As String
    C018 = Show(0 / 0)
End Function

Function C019() As String
    C019 = Show(5 \ 0)
End Function

Function C020() As String
    C020 = Show(CSng(1) / CSng(3))
End Function

Function C021() As String
    C021 = Show(CInt(1) / CInt(3))
End Function

Function C022() As String
    C022 = Show(1 / 3)
End Function

Function C023() As String
    C023 = Show(0.1 + 0.2)
End Function

Function C024() As String
    C024 = Show(1E+15)
End Function

Function C025() As String
    C025 = Show(1E+16)
End Function

Function C026() As String
    C026 = Show(123456789012345)
End Function

Function C027() As String
    C027 = Show(1234567890123456)
End Function

Function C028() As String
    C028 = Show(0.0001)
End Function

Function C029() As String
    C029 = Show(0.00001)
End Function

Function C030() As String
    C030 = Show(0.000001)
End Function

Function C031() As String
    C031 = Show(1E-10)
End Function

Function C032() As String
    C032 = Show(-0.5)
End Function

Function C033() As String
    C033 = Show(CDbl("1,234.5"))
End Function

Function C034() As String
    C034 = Show(CDbl(" 12 "))
End Function

Function C035() As String
    C035 = Show(CDbl("&H10"))
End Function

Function C036() As String
    C036 = Show(CDbl("1e3"))
End Function

Function C037() As String
    C037 = Show(CDbl("$5"))
End Function

Function C038() As String
    C038 = Show(CDbl("abc"))
End Function

Function C039() As String
    C039 = Show(CDbl(""))
End Function

Function C040() As String
    C040 = Show(CInt(2.5))
End Function

Function C041() As String
    C041 = Show(CInt(3.5))
End Function

Function C042() As String
    C042 = Show(CInt(-2.5))
End Function

Function C043() As String
    C043 = Show(CInt(32767.5))
End Function

Function C044() As String
    C044 = Show(CLng(2.5))
End Function

Function C045() As String
    C045 = Show(CLng("1.5"))
End Function

Function C046() As String
    C046 = Show(Int(-2.5))
End Function

Function C047() As String
    C047 = Show(Fix(-2.5))
End Function

Function C048() As String
    C048 = Show(Round(2.5))
End Function

Function C049() As String
    C049 = Show(Round(3.5))
End Function

Function C050() As String
    C050 = Show(Round(2.345, 2))
End Function

Function C051() As String
    C051 = Show(Round(2.355, 2))
End Function

Function C052() As String
    C052 = Show(Round(-2.5))
End Function

Function C053() As String
    C053 = Show(Val("12abc"))
End Function

Function C054() As String
    C054 = Show(Val("  1 2 3"))
End Function

Function C055() As String
    C055 = Show(Val("&HFF"))
End Function

Function C056() As String
    C056 = Show(Val("1e2x"))
End Function

Function C057() As String
    C057 = Show(Val(""))
End Function

Function C058() As String
    C058 = Show("10" + 5)
End Function

Function C059() As String
    C059 = Show("10" & 5)
End Function

Function C060() As String
    C060 = Show("a" + "b")
End Function

Function C061() As String
    C061 = Show("10" + "5")
End Function

Function C062() As String
    C062 = Show(5 + Empty)
End Function

Function C063() As String
    C063 = Show("x" & Null)
End Function

Function C064() As String
    C064 = Show(Null + 1)
End Function

Function C065() As String
    C065 = Show(True + 1)
End Function

Function C066() As String
    C066 = Show(True And 3)
End Function

Function C067() As String
    C067 = Show(5 And 3)
End Function

Function C068() As String
    C068 = Show(5 Or 3)
End Function

Function C069() As String
    C069 = Show(5 Xor 3)
End Function

Function C070() As String
    C070 = Show(Not 0)
End Function

Function C071() As String
    C071 = Show(Not 5)
End Function

Function C072() As String
    C072 = Show(True = -1)
End Function

Function C073() As String
    C073 = Show("abc" < "abd")
End Function

Function C074() As String
    C074 = Show("a" < "B")
End Function

Function C075() As String
    C075 = Show("10" < "9")
End Function

Function C076() As String
    C076 = Show(10 < "9")
End Function

Function C077() As String
    C077 = Show(Empty = 0)
End Function

Function C078() As String
    C078 = Show(Empty = "")
End Function

Function C079() As String
    C079 = Show("abc" Like "a*")
End Function

Function C080() As String
    C080 = Show("abc" Like "A*")
End Function

Function C081() As String
    C081 = Show("a1c" Like "a#c")
End Function

Function C082() As String
    C082 = Show("abc" Like "[a-c]bc")
End Function

Function C083() As String
    C083 = Show("xbc" Like "[!a]bc")
End Function

Function C084() As String
    C084 = Show(Len("hello"))
End Function

Function C085() As String
    C085 = Show(Len(12345))
End Function

Function C086() As String
    C086 = Show(Mid("Hello", 2, 3))
End Function

Function C087() As String
    C087 = Show(Mid("Hello", 10))
End Function

Function C088() As String
    C088 = Show(Left("Hello", 2))
End Function

Function C089() As String
    C089 = Show(Right("Hello", 10))
End Function

Function C090() As String
    C090 = Show(InStr("abcabc", "c"))
End Function

Function C091() As String
    C091 = Show(InStr(4, "abcabc", "c"))
End Function

Function C092() As String
    C092 = Show(InStr(1, "abc", "B", vbTextCompare))
End Function

Function C093() As String
    C093 = Show(InStr("abc", ""))
End Function

Function C094() As String
    C094 = Show(InStrRev("abcabc", "b"))
End Function

Function C095() As String
    C095 = Show(Replace("aaa", "a", "bb"))
End Function

Function C096() As String
    C096 = Show(Replace("Hello", "L", "x", , , vbTextCompare))
End Function

Function C097() As String
    C097 = Show(Trim("  a b  "))
End Function

Function C098() As String
    C098 = Show(UCase("héllo"))
End Function

Function C099() As String
    C099 = Show(StrComp("a", "B"))
End Function

Function C100() As String
    C100 = Show(StrComp("a", "B", vbTextCompare))
End Function

Function C101() As String
    C101 = Show(StrReverse("abc"))
End Function

Function C102() As String
    C102 = Show(Space(3) & "|")
End Function

Function C103() As String
    C103 = Show(String(3, "ab"))
End Function

Function C104() As String
    C104 = Show(Chr(65))
End Function

Function C105() As String
    C105 = Show(Asc("a"))
End Function

Function C106() As String
    C106 = Show(Split("a,b,,c", ",")(2))
End Function

Function C107() As String
    C107 = Show(UBound(Split("a,b,,c", ",")))
End Function

Function C108() As String
    C108 = Show(Join(Array(1, 2, 3), "-"))
End Function

Function C109() As String
    C109 = Show(UBound(Array()))
End Function

Function C110() As String
    C110 = Show(Hex(255))
End Function

Function C111() As String
    C111 = Show(Hex(-1))
End Function

Function C112() As String
    C112 = Show(Oct(8))
End Function

Function C113() As String
    C113 = Show(Str(5))
End Function

Function C114() As String
    C114 = Show(Str(-5))
End Function

Function C115() As String
    C115 = Show(CStr(5))
End Function

Function C116() As String
    C116 = Show(CStr(True))
End Function

Function C117() As String
    C117 = Show(CStr(#1/2/2026#))
End Function

Function C118() As String
    C118 = Show(CStr(#1/2/2026 3:04:05 PM#))
End Function

Function C119() As String
    C119 = Show(CStr(#3:04:05 AM#))
End Function

Function C120() As String
    C120 = Show(CStr(CDate("2026-10-14")))
End Function

Function C121() As String
    C121 = Show(#10/14/2026# + 1)
End Function

Function C122() As String
    C122 = Show(#10/14/2026# - #10/1/2026#)
End Function

Function C123() As String
    C123 = Show(CDbl(#10/14/2026#))
End Function

Function C124() As String
    C124 = Show(DateSerial(2026, 13, 1))
End Function

Function C125() As String
    C125 = Show(DateSerial(2026, 2, 30))
End Function

Function C126() As String
    C126 = Show(DateAdd("m", 1, #1/31/2026#))
End Function

Function C127() As String
    C127 = Show(DateAdd("d", -1, #3/1/2024#))
End Function

Function C128() As String
    C128 = Show(DateDiff("d", #1/1/2026#, #12/31/2026#))
End Function

Function C129() As String
    C129 = Show(DateDiff("m", #1/31/2026#, #2/1/2026#))
End Function

Function C130() As String
    C130 = Show(Weekday(#10/14/2026#))
End Function

Function C131() As String
    C131 = Show(Year(#10/14/2026#) & "/" & Month(#10/14/2026#) & "/" & Day(#10/14/2026#))
End Function

Function C132() As String
    C132 = Show(MonthName(10))
End Function

Function C133() As String
    C133 = Show(WeekdayName(4))
End Function

Function C134() As String
    C134 = Show(Format(1234.5, "#,##0.00"))
End Function

Function C135() As String
    C135 = Show(Format(1234.5, "0"))
End Function

Function C136() As String
    C136 = Show(Format(0.256, "0.0%"))
End Function

Function C137() As String
    C137 = Show(Format(-1234.5, "$#,##0.00"))
End Function

Function C138() As String
    C138 = Show(Format(1234.5, "Currency"))
End Function

Function C139() As String
    C139 = Show(Format(1234.5, "Standard"))
End Function

Function C140() As String
    C140 = Show(Format(0.5, "Percent"))
End Function

Function C141() As String
    C141 = Show(Format(1234.5))
End Function

Function C142() As String
    C142 = Show(Format(#10/14/2026#, "yyyy-mm-dd"))
End Function

Function C143() As String
    C143 = Show(Format(#10/14/2026#, "mmm d, yyyy"))
End Function

Function C144() As String
    C144 = Show(Format(#10/14/2026#, "dddd"))
End Function

Function C145() As String
    C145 = Show(Format(#10/14/2026 3:04:05 PM#, "hh:nn AM/PM"))
End Function

Function C146() As String
    C146 = Show(Format(#10/14/2026#, "Long Date"))
End Function

Function C147() As String
    C147 = Show(Format(#10/14/2026#, "Short Date"))
End Function

Function C148() As String
    C148 = Show(Format(5, "00000"))
End Function

Function C149() As String
    C149 = Show(Format(True, "Yes/No"))
End Function

Function C150() As String
    C150 = Show(Format("abc", ">"))
End Function

Function C151() As String
    C151 = Show(FormatNumber(1234.567, 2))
End Function

Function C152() As String
    C152 = Show(FormatPercent(0.1234))
End Function

Function C153() As String
    C153 = Show(FormatCurrency(-5))
End Function

Function C154() As String
    C154 = Show(IsNumeric("12.5"))
End Function

Function C155() As String
    C155 = Show(IsNumeric("1,000"))
End Function

Function C156() As String
    C156 = Show(IsNumeric("abc"))
End Function

Function C157() As String
    C157 = Show(IsNumeric(""))
End Function

Function C158() As String
    C158 = Show(IsNumeric(Empty))
End Function

Function C159() As String
    C159 = Show(IsDate("10/14/2026"))
End Function

Function C160() As String
    C160 = Show(IsDate("hello"))
End Function

Function C161() As String
    C161 = Show(IsEmpty(Empty))
End Function

Function C162() As String
    C162 = Show(VarType(1))
End Function

Function C163() As String
    C163 = Show(VarType(1&))
End Function

Function C164() As String
    C164 = Show(VarType(1.5))
End Function

Function C165() As String
    C165 = Show(VarType("a"))
End Function

Function C166() As String
    C166 = Show(VarType(Null))
End Function

Function C167() As String
    C167 = Show(TypeName(1))
End Function

Function C168() As String
    C168 = Show(TypeName(1&))
End Function

Function C169() As String
    C169 = Show(TypeName(1.5))
End Function

Function C170() As String
    C170 = Show(TypeName(1@))
End Function

Function C171() As String
    C171 = Show(TypeName(Empty))
End Function

Function C172() As String
    C172 = Show(TypeName(Null))
End Function

Function C173() As String
    C173 = Show(TypeName(Array(1)))
End Function

Function C174() As String
    C174 = Show(TypeName(&HFFFF))
End Function

Function C175() As String
    C175 = Show(&HFFFF)
End Function

Function C176() As String
    C176 = Show(&HFFFF&)
End Function

Function C177() As String
    C177 = Show(&H8000)
End Function

Function C178() As String
    C178 = Show(IIf(1 > 2, "a", "b"))
End Function

Function C179() As String
    C179 = Show(Choose(2, "a", "b", "c"))
End Function

Function C180() As String
    C180 = Show(Switch(1 > 2, "a", 2 > 1, "b"))
End Function

Function C181() As String
    C181 = Show(CBool("True"))
End Function

Function C182() As String
    C182 = Show(CBool("yes"))
End Function

Function C183() As String
    C183 = Show(CBool(0.1))
End Function

Function C184() As String
    C184 = Show(CByte(255))
End Function

Function C185() As String
    C185 = Show(CByte(256))
End Function

Function C186() As String
    C186 = Show(CCur(1.23456))
End Function

Function C187() As String
    C187 = Show(CCur(1) / 3)
End Function

Function C188() As String
    C188 = Show(CDate(45000))
End Function

Function C189() As String
    C189 = Show(CDate("3:00 PM"))
End Function

Function C190() As String
    C190 = Show(Abs(-3))
End Function

Function C191() As String
    C191 = Show(Sgn(-3))
End Function

Function C192() As String
    C192 = Show(Sqr(16))
End Function

Function C193() As String
    C193 = Show(Exp(1))
End Function

Function C194() As String
    C194 = Show(Log(Exp(2)))
End Function

Function C195() As String
    C195 = Show(Atn(1) * 4)
End Function
