Option Explicit

' Insert/Delete with every Shift, Copy/Cut/PasteSpecial, Clear*, formula adjustment.

Private Sub Grid3()
    Dim r As Long, c As Long
    For r = 1 To 3
        For c = 1 To 3
            Worksheets("Sheet1").Cells(r, c).Value = r * 10 + c
        Next
    Next
End Sub

Function E001() As String
    Reset: Grid3
    Range("B2").Insert Shift:=xlDown
    E001 = Dump(Range("A1:C4"))
End Function

Function E002() As String
    Reset: Grid3
    Range("B2").Insert Shift:=xlToRight
    E002 = Dump(Range("A1:D3"))
End Function

Function E003() As String
    Reset: Grid3
    Range("B2").Delete Shift:=xlUp
    E003 = Dump(Range("A1:C3"))
End Function

Function E004() As String
    Reset: Grid3
    Range("B2").Delete Shift:=xlToLeft
    E004 = Dump(Range("A1:C3"))
End Function

Function E005() As String
    Reset: Grid3
    Range("A2:C2").Insert
    E005 = Dump(Range("A1:C4"))
    Range("B1:B4").Insert
    E005 = E005 & "|" & Dump(Range("A1:D4"))
End Function

Function E006() As String
    Reset: Grid3
    Range("A2:B2").Delete
    E006 = Dump(Range("A1:C3"))
    Range("A1:A2").Delete
    E006 = E006 & "|" & Dump(Range("A1:C3"))
End Function

Function E007() As String
    Reset: Grid3
    Rows(2).Insert
    E007 = Dump(Range("A1:C4"))
    Columns("B").Insert
    E007 = E007 & "|" & Dump(Range("A1:D4"))
End Function

Function E008() As String
    Reset: Grid3
    Rows("1:2").Delete
    E008 = Dump(Range("A1:C2"))
    Columns("A:B").Delete
    E008 = E008 & "|" & Dump(Range("A1:A2"))
End Function

Function E009() As String
    Reset: Grid3
    Range("B2").EntireRow.Delete
    Range("B1").EntireColumn.Insert
    E009 = Dump(Range("A1:D2"))
End Function

Function E010() As String
    Reset
    Range("A1").Value = 1
    Range("A2").Value = 2
    Range("A3").Formula = "=SUM(A1:A2)"
    Range("B1").Formula = "=$A$2*2"
    Rows(2).Insert
    E010 = Range("A4").Formula & "|" & Range("A4").Value & "|" & Range("B1").Formula & "|" & Range("B1").Value
End Function

Function E011() As String
    Reset
    Range("A1:A5").Value = 1
    Range("B1").Formula = "=SUM(A1:A5)"
    Range("C1").Formula = "=A3"
    Rows(3).Delete
    E011 = Range("B1").Formula & "|" & Range("B1").Value & "|" & Range("C1").Formula & "|" & Range("C1").Text
End Function

Function E012() As String
    Reset
    Range("A1").Value = 1
    Range("C1").Formula = "=A1+B1"
    Columns("B").Insert
    E012 = Range("D1").Formula
    Columns("A").Delete
    E012 = E012 & "|" & Range("C1").Formula & "|" & Range("C1").Text
End Function

Function E013() As String
    Reset: Grid3
    Range("A1:C1").Copy Range("A5")
    E013 = Dump(Range("A5:C5"))
    Range("A1:A3").Copy Destination:=Range("E1")
    E013 = E013 & "|" & Dump(Range("E1:E3")) & "|" & (Application.CutCopyMode = False)
End Function

Function E014() As String
    Reset: Grid3
    Range("A1:B1").Cut Range("A6")
    E014 = Dump(Range("A1:C1")) & "|" & Dump(Range("A6:B6"))
End Function

Function E015() As String
    Reset
    Range("A1").Value = 2
    Range("B1").Formula = "=A1*10"
    Range("B1").Copy Range("B2")
    Range("A2").Value = 3
    E015 = Range("B2").Formula & "|" & Range("B2").Value
    Range("B1").Cut Range("C5")
    E015 = E015 & "|" & Range("C5").Formula & "|" & Range("C5").Value
End Function

Function E016() As String
    Reset
    Range("A1").Value = 5
    Range("B1").Formula = "=A1*2"
    Range("B1").Font.Bold = True
    Range("B1").Copy
    Range("D1").PasteSpecial Paste:=xlPasteValues
    Range("E1").PasteSpecial Paste:=xlPasteFormulas
    Range("F1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    E016 = Range("D1").Formula & "|" & Range("D1").Font.Bold & "|" & Range("E1").Formula & "|" & Range("E1").Font.Bold & "|" & IsEmpty(Range("F1").Value) & "|" & Range("F1").Font.Bold
End Function

Function E017() As String
    Reset: Grid3
    Range("A1:C2").Copy
    Range("E1").PasteSpecial Transpose:=True
    Application.CutCopyMode = False
    E017 = Dump(Range("E1:F3"))
End Function

Function E018() As String
    Reset
    Range("A1:A3").Value = 10
    Range("C1").Value = 5
    Range("C1").Copy
    Range("A1:A3").PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationAdd
    E018 = Dump(Range("A1:A3"))
    Range("A1:A3").PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationMultiply
    E018 = E018 & "|" & Dump(Range("A1:A3"))
    Range("A1:A3").PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationSubtract
    Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationDivide
    Application.CutCopyMode = False
    E018 = E018 & "|" & Dump(Range("A1:A3"))
End Function

Function E019() As String
    Reset
    Range("A1").Value = 1
    Range("A3").Value = 3
    Range("C1:C3").Value = "x"
    Range("A1:A3").Copy
    Range("C1").PasteSpecial Paste:=xlPasteValues, SkipBlanks:=True
    Application.CutCopyMode = False
    E019 = Dump(Range("C1:C3"))
End Function

Function E020() As String
    Reset
    Range("A1").Value = 7
    Range("A1").Copy
    Range("B1:D2").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    E020 = Dump(Range("B1:D2"))
End Function

Function E021() As String
    Reset
    Range("A1").Value = 1
    Range("A1").Copy
    E021 = Application.CutCopyMode
    Range("A1").Cut
    E021 = E021 & "|" & Application.CutCopyMode
    Application.CutCopyMode = False
    E021 = E021 & "|" & Application.CutCopyMode
End Function

Function E022() As String
    Reset
    On Error Resume Next
    Application.CutCopyMode = False
    Range("A1").PasteSpecial xlPasteValues
    E022 = Err.Number & "|" & Err.Description
End Function

Function E023() As String
    Reset: Grid3
    Range("A1:B2").Copy
    ActiveSheet.Paste Destination:=Range("D4")
    Application.CutCopyMode = False
    E023 = Dump(Range("D4:E5"))
    Range("A3").Copy
    Range("G1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    E023 = E023 & "|" & Range("G1").Value
End Function

Function E024() As String
    Reset: Grid3
    Range("A1:C3").Font.Bold = True
    Range("B2").ClearContents
    E024 = IsEmpty(Range("B2").Value) & "|" & Range("B2").Font.Bold
    Range("A1:C1").ClearFormats
    E024 = E024 & "|" & Range("A1").Value & "|" & Range("A1").Font.Bold
    Range("A3:C3").Clear
    E024 = E024 & "|" & IsEmpty(Range("A3").Value) & "|" & Range("A3").Font.Bold & "|" & Application.WorksheetFunction.CountA(Range("A1:C3"))
End Function

Function E025() As String
    Reset
    Range("A1:A3").Value = 1
    Range("A2").Font.Bold = True
    Rows(2).Insert
    E025 = Range("A3").Font.Bold & "|" & Range("A2").Font.Bold & "|" & IsEmpty(Range("A2").Value)
    Range("A3").EntireRow.Insert CopyOrigin:=xlFormatFromRightOrBelow
    E025 = E025 & "|" & Range("A3").Font.Bold
End Function

Function E026() As String
    Reset: Grid3
    Range("A1:C1").Copy
    Range("A3").Insert Shift:=xlDown
    Application.CutCopyMode = False
    E026 = Dump(Range("A1:C4"))
End Function

Function E027() As String
    Reset
    Range("A1").Value = 1
    Range("A2").Value = 2
    Range("B1").Formula = "=A1+A2"
    Range("A1:A2").Cut Range("D1")
    E027 = Range("B1").Formula & "|" & Range("B1").Value
End Function

Function E028() As String
    Reset
    Range("A1").Value = "x"
    Range("A1").Copy Range("A2:A4")
    E028 = Dump(Range("A1:A5"))
    Range("A1:A2").Copy Range("C1:C6")
    E028 = E028 & "|" & Dump(Range("C1:C7"))
End Function

Function E029() As String
    Reset: Grid3
    Range("A1:C3").Delete Shift:=xlUp
    E029 = Application.WorksheetFunction.CountA(Cells) & "|" & ActiveSheet.UsedRange.Address
End Function

Function E030() As String
    Reset
    Range("A1").Value = 1
    Range("A2").Formula = "=A1*2"
    Range("A1").Delete Shift:=xlUp
    E030 = Range("A1").Formula & "|" & Range("A1").Text
End Function

Function E031() As String
    Reset
    Range("A1:A3").Value = 1
    Range("B1").Formula = "=SUM(A1:A3)"
    Range("A4").EntireRow.Insert
    Range("A2").EntireRow.Insert
    E031 = Range("B1").Formula & "|" & Range("B1").Value
End Function

Function E032() As String
    Reset
    Range("A1").Formula = "=B1+C1"
    Range("A1").Copy Range("A2:A3")
    E032 = Range("A3").Formula
    Range("A1").Copy Range("D1")
    E032 = E032 & "|" & Range("D1").Formula
End Function

Function E033() As String
    Reset
    Range("B2").Formula = "=$A$1+A1"
    Range("B2").Copy Range("C3")
    E033 = Range("C3").Formula
    On Error Resume Next
    Range("B2").Copy Range("A2")
    E033 = E033 & "|" & Range("A2").Formula & "|" & Range("A2").Text
End Function

Function E034() As String
    Reset: Grid3
    Range("B1:B3").Insert Shift:=xlToRight
    Range("A2:D2").Delete Shift:=xlUp
    E034 = Dump(Range("A1:D3"))
End Function

Function E035() As String
    Reset
    Range("A1").Value = "a"
    Range("A1").Copy
    Range("B1").PasteSpecial
    Application.CutCopyMode = False
    E035 = Range("B1").Value
    Range("A1").Copy Range("A1")
    E035 = E035 & "|" & Range("A1").Value
End Function

Function E036() As String
    Reset
    Range("A1").Value = 1
    Range("A1").NumberFormat = "0.00"
    Range("A1").Copy
    Range("B1").PasteSpecial xlPasteValuesAndNumberFormats
    Range("C1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    E036 = Range("B1").Text & "|" & Range("C1").Text
End Function

Function E037() As String
    Reset
    Range("A1:B1").Value = Array(1, 2)
    Range("A1:B1").Copy
    Range("D1").PasteSpecial xlPasteAll, Transpose:=True
    Application.CutCopyMode = False
    E037 = Range("D1").Value & "|" & Range("D2").Value & "|" & IsEmpty(Range("E1").Value)
End Function

Function E038() As String
    Reset
    Range("A1:A3").Value = 4
    Range("A1:A3").Value = Range("A1:A3").Value
    Range("B1:B3").Value = Range("A1:A3").Value
    Range("C1:C2").Value = Range("A1:A3").Value
    E038 = Dump(Range("B1:C3"))
End Function

Function E039() As String
    Reset
    Range("A1").Formula = "=1+1"
    Range("A1").Value = Range("A1").Value
    E039 = Range("A1").HasFormula & "|" & Range("A1").Value
    Range("A2").Formula = "=3*3"
    Range("A2").Formula = Range("A2").Value
    E039 = E039 & "|" & Range("A2").HasFormula
End Function

Function E040() As String
    Reset
    Range("A1:A2").Value = 1
    On Error Resume Next
    Range("A1").Insert Shift:=5
    E040 = Err.Number
    Err.Clear
    Range("A1,C1").Copy
    E040 = E040 & "|" & Err.Number
End Function
