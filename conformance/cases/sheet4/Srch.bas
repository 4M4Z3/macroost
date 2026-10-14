Option Explicit

' Find / FindNext / FindPrevious loops and Replace.

Private Sub Seed()
    With Worksheets("Sheet1")
        .Range("A1").Value = "apple"
        .Range("B2").Value = "Apple pie"
        .Range("C3").Value = "pineapple"
        .Range("A4").Value = "banana"
        .Range("B5").Value = "APPLE"
        .Range("C6").Value = 42
        .Range("A7").Formula = "=C6*2"
    End With
End Sub

Function D001() As String
    Dim f As Range, first As String, s As String
    Reset: Seed
    Set f = Range("A1:C10").Find(What:="apple", LookIn:=xlValues, LookAt:=xlPart)
    If Not f Is Nothing Then
        first = f.Address
        Do
            s = s & f.Address(0, 0) & ","
            Set f = Range("A1:C10").FindNext(f)
        Loop While Not f Is Nothing And f.Address <> first
    End If
    D001 = s
End Function

Function D002() As String
    Dim f As Range
    Reset: Seed
    Set f = Range("A1:C10").Find(What:="apple", LookAt:=xlWhole, MatchCase:=False)
    D002 = f.Address(0, 0)
    Set f = Range("A1:C10").Find(What:="apple", After:=f, LookAt:=xlWhole, MatchCase:=False)
    D002 = D002 & "|" & f.Address(0, 0)
    Set f = Range("A1:C10").Find(What:="APPLE", LookAt:=xlWhole, MatchCase:=True)
    D002 = D002 & "|" & f.Address(0, 0)
End Function

Function D003() As String
    Dim f As Range
    Reset: Seed
    Set f = Range("A1:C10").Find(What:="apple", SearchOrder:=xlByColumns, LookAt:=xlPart)
    D003 = f.Address(0, 0)
    Set f = Range("A1:C10").FindNext(f)
    D003 = D003 & "|" & f.Address(0, 0)
    Set f = Range("A1:C10").FindNext(f)
    D003 = D003 & "|" & f.Address(0, 0)
End Function

Function D004() As String
    Dim f As Range
    Reset: Seed
    Set f = Range("A1:C10").Find(What:="apple", SearchDirection:=xlPrevious, LookAt:=xlPart)
    D004 = f.Address(0, 0)
    Set f = Range("A1:C10").FindPrevious(f)
    D004 = D004 & "|" & f.Address(0, 0)
    Set f = Range("A1:C10").FindPrevious(f)
    D004 = D004 & "|" & f.Address(0, 0)
End Function

Function D005() As String
    Dim f As Range
    Reset: Seed
    Set f = Cells.Find(What:="84", LookIn:=xlValues)
    D005 = f.Address(0, 0)
    Set f = Cells.Find(What:="84", LookIn:=xlFormulas)
    D005 = D005 & "|" & (f Is Nothing)
    Set f = Cells.Find(What:="C6", LookIn:=xlFormulas, LookAt:=xlPart)
    D005 = D005 & "|" & f.Address(0, 0)
End Function

Function D006() As String
    Dim f As Range
    Reset: Seed
    Set f = Cells.Find(What:="*an*", LookAt:=xlWhole)
    D006 = f.Address(0, 0)
    Set f = Cells.Find(What:="b?nana", LookAt:=xlWhole)
    D006 = D006 & "|" & f.Address(0, 0)
    Set f = Cells.Find(What:="4", LookAt:=xlPart)
    D006 = D006 & "|" & f.Address(0, 0)
End Function

Function D007() As String
    Dim f As Range
    Reset: Seed
    Set f = Range("A1:C10").Find("pine")
    D007 = f.Address(0, 0) & "|" & f.Value
    Set f = Range("A1:A1").Find("apple")
    D007 = D007 & "|" & f.Address(0, 0)
End Function

Function D008() As String
    Dim f As Range
    Reset: Seed
    Set f = Range("A1:C10").Find(What:="apple", After:=Range("C10"), LookAt:=xlWhole)
    D008 = f.Address(0, 0)
    Set f = Range("A1:C10").Find(What:="apple", After:=Range("A1"), LookAt:=xlWhole)
    D008 = D008 & "|" & f.Address(0, 0)
End Function

Function D009() As String
    Dim f As Range, n As Long, first As String
    Reset
    Range("A1:A20").Value = "x"
    Range("A5").Value = "y"
    Range("A15").Value = "y"
    Set f = Columns("A").Find("y", LookAt:=xlWhole)
    first = f.Address
    Do
        n = n + 1
        f.Offset(0, 1).Value = n
        Set f = Columns("A").FindNext(f)
    Loop Until f.Address = first
    D009 = n & "|" & Range("B5").Value & "|" & Range("B15").Value
End Function

Function D010() As String
    Reset: Seed
    Range("A1:C10").Replace What:="apple", Replacement:="pear", LookAt:=xlPart, MatchCase:=False
    D010 = Range("A1").Value & "|" & Range("B2").Value & "|" & Range("C3").Value & "|" & Range("B5").Value & "|" & Range("A4").Value
End Function

Function D011() As String
    Reset: Seed
    Range("A1:C10").Replace What:="apple", Replacement:="X", LookAt:=xlWhole, MatchCase:=True
    D011 = Range("A1").Value & "|" & Range("B2").Value & "|" & Range("B5").Value
End Function

Function D012() As String
    Reset
    Range("A1").Value = "2026-01-02"
    Range("A2").Value = "a.b.c"
    Range("A3").Formula = "=""x""&""y"""
    Range("A2").Replace ".", "-"
    Range("A3").Replace "x", "z"
    D012 = Range("A2").Value & "|" & Range("A3").Formula & "|" & Range("A3").Value
End Function

Function D013() As String
    Reset
    Range("A1").Value = 1234
    Range("A2").Value = "12-34"
    Cells.Replace What:="2", Replacement:="9", LookAt:=xlPart
    D013 = Range("A1").Value & "|" & TypeName(Range("A1").Value) & "|" & Range("A2").Value
End Function

Function D014() As String
    Reset
    Range("A1").Value = "a*b"
    Range("A2").Value = "axxb"
    Range("A1:A2").Replace What:="~*", Replacement:="+"
    D014 = Range("A1").Value & "|" & Range("A2").Value
    Range("A1:A2").Replace What:="a*b", Replacement:="Z"
    D014 = D014 & "|" & Range("A1").Value & "|" & Range("A2").Value
End Function

Function D015() As String
    Dim ok As Variant
    Reset
    Range("A1").Value = "abc"
    ok = Range("A1").Replace("zzz", "y")
    D015 = ok & "|" & Range("A1").Value
    ok = Range("A1").Replace("b", "")
    D015 = D015 & "|" & ok & "|" & Range("A1").Value
End Function

Function D016() As String
    Dim f As Range
    Reset
    Range("B2").Value = 3.5
    Range("B3").Value = 35
    Range("B4").NumberFormat = "0.0"
    Range("B4").Value = 3.5
    Set f = Cells.Find(What:="3.5", LookIn:=xlValues, LookAt:=xlWhole)
    D016 = f.Address(0, 0)
    Set f = Cells.FindNext(f)
    D016 = D016 & "|" & f.Address(0, 0)
End Function

Function D017() As String
    Dim f As Range
    Reset
    Range("A1").Value = DateSerial(2026, 5, 1)
    Set f = Cells.Find(What:="5/1/2026", LookIn:=xlValues)
    D017 = (f Is Nothing)
    Set f = Cells.Find(What:="5/1/2026", LookIn:=xlFormulas)
    D017 = D017 & "|" & (f Is Nothing)
End Function

Function D018() As String
    Dim f As Range
    Reset: Seed
    Set f = Range("A1:C10").Find(What:="ap", LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, After:=Range("B2"))
    D018 = f.Address(0, 0)
    Set f = Range("B1:C10").Find(What:="apple", LookAt:=xlPart)
    D018 = D018 & "|" & f.Address(0, 0)
End Function

Function D019() As String
    Dim f As Range
    Reset
    Range("A1").Value = "   spaced"
    Range("A2").Value = "spaced"
    Set f = Cells.Find("spaced", LookAt:=xlWhole)
    D019 = f.Address(0, 0)
    Set f = Cells.Find("spaced", LookAt:=xlPart)
    D019 = D019 & "|" & f.Address(0, 0)
End Function

Function D020() As String
    Dim f As Range
    Reset
    Range("A1").Value = "first"
    Range("A1").Font.Bold = True
    Set f = Range("A1").Find("first")
    D020 = f.Address(0, 0)
    Set f = Range("A2").Find("first")
    D020 = D020 & "|" & f.Address(0, 0)
End Function
