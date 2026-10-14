Option Explicit

' Application.WorksheetFunction and Application.<Function> (error values instead of 1004).

' Late bound, so a function the host lacks is a run-time error 438, not a compile error.
Private Function WF() As Object
    Set WF = Application.WorksheetFunction
End Function

Private Sub Data()
    With Worksheets("Sheet1")
        .Range("A1:C1").Value = Array("Item", "Cat", "Qty")
        .Range("A2:C2").Value = Array("pen", "supply", 10)
        .Range("A3:C3").Value = Array("ink", "supply", 5)
        .Range("A4:C4").Value = Array("tea", "food", 7)
        .Range("A5:C5").Value = Array("cup", "food", 3)
        .Range("A6:C6").Value = Array("pad", "supply", 12)
    End With
End Sub

Function W001() As String
    Reset: Data
    With WF()
        W001 = .Sum(Range("C2:C6")) & "|" & .Sum(1, 2, 3) & "|" & .Sum(Range("C2:C3"), 100) & "|" & .Max(Range("C2:C6")) & "|" & .Min(Range("C2:C6")) & "|" & .Average(Range("C2:C6")) & "|" & .Count(Range("A1:C6")) & "|" & .CountA(Range("A1:C6"))
    End With
End Function

Function W002() As String
    Reset: Data
    With WF()
        W002 = .SumIf(Range("B2:B6"), "supply", Range("C2:C6")) & "|" & .SumIf(Range("C2:C6"), ">5") & "|" & .CountIf(Range("B2:B6"), "food") & "|" & .CountIf(Range("C2:C6"), "<=7") & "|" & .CountIf(Range("A2:A6"), "p*")
    End With
End Function

Function W003() As String
    Reset: Data
    With WF()
        W003 = .SumIfs(Range("C2:C6"), Range("B2:B6"), "supply", Range("C2:C6"), ">6") & "|" & .CountIfs(Range("B2:B6"), "food", Range("C2:C6"), "<5") & "|" & .AverageIf(Range("B2:B6"), "supply", Range("C2:C6"))
    End With
End Function

Function W004() As String
    Reset: Data
    With WF()
        W004 = .VLookup("tea", Range("A2:C6"), 3, False) & "|" & .VLookup("TEA", Range("A2:C6"), 2, False) & "|" & .Match("cup", Range("A2:A6"), 0) & "|" & .Index(Range("A2:C6"), 3, 2) & "|" & .Index(Range("A2:C6"), .Match("pad", Range("A2:A6"), 0), 3)
    End With
End Function

Function W005() As String
    Reset: Data
    On Error Resume Next
    Dim v As Variant
    v = WF().VLookup("zzz", Range("A2:C6"), 3, False)
    W005 = Err.Number & "|" & Err.Description
    Err.Clear
    v = WF().Match("zzz", Range("A2:A6"), 0)
    W005 = W005 & "|" & Err.Number & "|" & Err.Description
    Err.Clear
    v = WF().Average(Range("E1:E5"))
    W005 = W005 & "|" & Err.Number & "|" & Err.Description
End Function

Function W006() As String
    Dim v As Variant
    Reset: Data
    v = Application.VLookup("zzz", Range("A2:C6"), 3, False)
    W006 = IsError(v) & "|" & ErrCode(v)
    v = Application.Match("zzz", Range("A2:A6"), 0)
    W006 = W006 & "|" & ErrCode(v)
    v = Application.VLookup("ink", Range("A2:C6"), 3, False)
    W006 = W006 & "|" & v & "|" & ErrCode(Application.Average(Range("E1:E5")))
End Function

Function W007() As String
    Reset
    Range("A1:E1").Value = Array(1, 5, 10, 50, 100)
    Range("A2:E2").Value = Array("a", "b", "c", "d", "e")
    With WF()
        W007 = .HLookup(10, Range("A1:E2"), 2, False) & "|" & .HLookup(60, Range("A1:E2"), 2, True) & "|" & .Match(7, Range("A1:E1"), 1) & "|" & .Match(50, Range("A1:E1")) & "|" & .VLookup(4, WF().Transpose(Range("A1:E2")), 2, True)
    End With
End Function

Function W008() As String
    With WF()
        W008 = .Round(2.345, 2) & "|" & .Round(-2.5, 0) & "|" & .Round(1234.567, -2) & "|" & .RoundUp(2.341, 2) & "|" & .RoundDown(-2.349, 2) & "|" & .RoundUp(-2.1, 0) & "|" & .Round(0.5, 0)
    End With
End Function

Function W009() As String
    With WF()
        W009 = "[" & .Trim("  a   b  ") & "]|" & .Proper("hELLO wORLD o'neil") & "|" & .Text(0.5, "0%") & "|" & .Text(45000, "yyyy-mm-dd") & "|" & .Text(1234.5, "$#,##0.00")
    End With
End Function

Function W010() As String
    Reset
    Range("A1:A6").Value = WF().Transpose(Array(30, 10, 50, 20, 50, 40))
    With WF()
        W010 = .Large(Range("A1:A6"), 1) & "|" & .Large(Range("A1:A6"), 3) & "|" & .Small(Range("A1:A6"), 2) & "|" & .Rank(20, Range("A1:A6")) & "|" & .Rank(20, Range("A1:A6"), 1) & "|" & .Rank(50, Range("A1:A6"))
    End With
End Function

Function W011() As String
    Reset
    On Error Resume Next
    Dim v As Variant
    v = WF().Large(Range("A1:A3"), 1)
    W011 = Err.Number
    Err.Clear
    v = WF().Rank(5, Range("A1:A3"))
    W011 = W011 & "|" & Err.Number
    Err.Clear
    v = WF().Index(Range("A1:B2"), 5, 1)
    W011 = W011 & "|" & Err.Number & "|" & Err.Description
End Function

Function W012() As String
    Reset
    Range("A1").Value = 1
    Range("A2").Formula = "=1/0"
    On Error Resume Next
    Dim v As Variant
    v = WF().Sum(Range("A1:A2"))
    W012 = Err.Number & "|" & Err.Description
    Err.Clear
    v = WF().IfError(Range("A2").Value, "fallback")
    W012 = W012 & "|" & Err.Number & "|" & v
    Err.Clear
    v = WF().IfError(5, "fallback")
    W012 = W012 & "|" & v & "|" & ErrCode(Application.Sum(Range("A1:A2")))
End Function

Function W013() As String
    Reset
    Range("A1:A5").Value = WF().Transpose(Array(1, "2", True, "", 4))
    With WF()
        W013 = .Sum(Range("A1:A5")) & "|" & .Count(Range("A1:A5")) & "|" & .CountA(Range("A1:A5")) & "|" & .CountBlank(Range("A1:A5")) & "|" & .Sum("2", True) & "|" & .Average(Range("A1:A5"))
    End With
End Function

Function W014() As String
    Reset
    Range("A1:B3").Value = 2
    With WF()
        W014 = .SumProduct(Range("A1:A3"), Range("B1:B3")) & "|" & .Product(Range("A1:B3")) & "|" & .Median(1, 5, 2, 8) & "|" & .Power(2, 10)
    End With
End Function

Function W015() As String
    With WF()
        W015 = .Max(3, 9, -2) & "|" & .Min(3, 9, -2) & "|" & .Ceiling(4.2, 1) & "|" & .Floor(4.8, 1) & "|" & .Even(3) & "|" & .Odd(4)
    End With
End Function

Function W016() As String
    With WF()
        W016 = .Substitute("a-b-c", "-", "+") & "|" & .Substitute("a-b-c", "-", "+", 2) & "|" & .Rept("ab", 3) & "|" & .Find("b", "abcb") & "|" & .Search("B", "abcb", 3) & "|" & .Clean("a" & Chr(7) & "b")
    End With
End Function

Function W017() As String
    Reset
    Range("A1:A4").Value = WF().Transpose(Array("x", "y", "x", "z"))
    With WF()
        W017 = .CountIf(Range("A1:A4"), "x") & "|" & .CountIf(Range("A1:A4"), "<>x") & "|" & .CountIf(Range("A1:A10"), "") & "|" & .CountIf(Range("A1:A4"), "*") & "|" & .CountIf(Range("A1:A4"), "?")
    End With
End Function

Function W018() As String
    Reset
    Range("A1:A3").Value = WF().Transpose(Array(DateSerial(2026, 1, 1), DateSerial(2026, 6, 1), DateSerial(2026, 12, 1)))
    With WF()
        W018 = .CountIf(Range("A1:A3"), ">" & CLng(DateSerial(2026, 3, 1))) & "|" & .Max(Range("A1:A3")) & "|" & .EoMonth(DateSerial(2026, 2, 10), 0) & "|" & .NetworkDays(DateSerial(2026, 10, 1), DateSerial(2026, 10, 14)) & "|" & .WeekNum(DateSerial(2026, 10, 14)) & "|" & .Weekday(DateSerial(2026, 10, 14))
    End With
End Function

Function W019() As String
    Reset: Data
    W019 = Application.Sum(Range("C2:C6")) & "|" & Application.Max(Range("C2:C6")) & "|" & Application.CountIf(Range("B2:B6"), "supply") & "|" & Application.Index(Range("A2:A6"), 2) & "|" & Application.Match("tea", Range("A2:A6"), 0)
End Function

Function W020() As String
    Dim arr As Variant
    Reset
    arr = Array(4, 8, 15, 16, 23, 42)
    With WF()
        W020 = .Sum(arr) & "|" & .Max(arr) & "|" & .Average(arr) & "|" & .Match(15, arr, 0) & "|" & .Index(arr, 3) & "|" & .Count(arr) & "|" & (.StDev(arr) > 0)
    End With
End Function

Function W021() As String
    Reset: Data
    With WF()
        W021 = .VLookup("pen", Range("A:C"), 3, False) & "|" & .CountA(Range("A:A")) & "|" & .Sum(Range("C:C")) & "|" & .Match("pad", Range("A:A"), 0) & "|" & .CountIf(Range("B:B"), "supply")
    End With
End Function

Function W022() As String
    Reset
    Range("A1:A3").Value = WF().Transpose(Array(1, 2, 3))
    Range("A2").EntireRow.Hidden = True
    With WF()
        W022 = .Subtotal(9, Range("A1:A3")) & "|" & .Subtotal(109, Range("A1:A3")) & "|" & .Subtotal(2, Range("A1:A3")) & "|" & .Aggregate(9, 5, Range("A1:A3"))
    End With
End Function

Function W023() As String
    With WF()
        W023 = .And(True, 1) & "|" & .Or(False, 0) & "|" & .Xor(True, True) & "|" & .IsNumber(5) & "|" & .IsText("x") & "|" & .IsError(CVErr(2042)) & "|" & .IsNA(CVErr(2042))
    End With
End Function

Function W024() As String
    Reset
    On Error Resume Next
    Dim v As Variant
    v = WF().NoSuchFunction(1)
    W024 = Err.Number & "|" & Err.Description
    Err.Clear
    v = WF().Sum()
    W024 = W024 & "|" & Err.Number
End Function

Function W025() As String
    Reset: Data
    With WF()
        W025 = .Index(Range("C2:C6"), .Match(.Max(Range("C2:C6")), Range("C2:C6"), 0)) & "|" & .Index(Range("A2:A6"), .Match(.Min(Range("C2:C6")), Range("C2:C6"), 0)) & "|" & TypeName(.Index(Range("A2:C6"), 2, 0))
    End With
End Function

Function W026() As String
    With WF()
        W026 = (.Pi() > 3.14) & "|" & .Fact(5) & "|" & .Combin(5, 2) & "|" & .Ln(1) & "|" & .Log10(1000) & "|" & .Log(8, 2)
    End With
End Function

Function W027() As String
    With WF()
        W027 = .Dollar(1234.567, 2) & "|" & .Fixed(1234.567, 1) & "|" & .Fixed(1234.567, 1, True) & "|" & (.Pmt(0.05 / 12, 360, -200000) > 1073)
    End With
End Function

Function W028() As String
    Reset
    Range("A1:B3").Value = WF().Transpose(Array(Array(1, 2, 3), Array("a", "b", "c")))
    W028 = Range("A3").Value & "|" & Range("B2").Value & "|" & WF().VLookup(2, Range("A1:B3"), 2, False)
End Function

' Which members the late-bound WorksheetFunction object has (438 when it lacks one).
Function W029() As String
    Dim v As Variant
    On Error Resume Next
    v = WF().Concatenate("a", 1): W029 = Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Sqrt(16): W029 = W029 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Abs(-3): W029 = W029 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Mod(10, 3): W029 = W029 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Int(-2.5): W029 = W029 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Exp(0): W029 = W029 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Value("12"): W029 = W029 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Len("abc"): W029 = W029 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Left("abc", 1): W029 = W029 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().If(True, 1, 2): W029 = W029 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Upper("abc"): W029 = W029 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Date(2026, 1, 2): W029 = W029 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Year(46000): W029 = W029 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Sign(-4): W029 = W029 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
End Function

Function W030() As String
    Dim v As Variant
    On Error Resume Next
    v = WF().Concat("a", "b", 1): W030 = Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().TextJoin("-", True, "a", "", "c"): W030 = W030 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Exact("a", "A"): W030 = W030 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Replace("abcdef", 2, 3, "X"): W030 = W030 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Trunc(-2.7): W030 = W030 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Char(65): W030 = W030 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Code("A"): W030 = W030 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().IsBlank(Range("Z99")): W030 = W030 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Not(True): W030 = W030 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Lower("AB"): W030 = W030 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Right("abc", 2): W030 = W030 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Mid("abc", 2, 1): W030 = W030 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Choose(2, "a", "b"): W030 = W030 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().Days(46010, 46000): W030 = W030 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().StDev_S(1, 2, 3): W030 = W030 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
    v = WF().MaxIfs(Range("A1:A2"), Range("A1:A2"), ">0"): W030 = W030 & "|" & Err.Number & "=" & v: Err.Clear: v = Empty
End Function
