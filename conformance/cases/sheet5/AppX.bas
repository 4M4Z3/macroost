Option Explicit

' Application properties, window no-ops and error numbers of invalid operations.

Function Y001() As String
    Reset
    Application.ScreenUpdating = False
    Y001 = Application.ScreenUpdating
    Application.ScreenUpdating = True
    Y001 = Y001 & "|" & Application.ScreenUpdating
    Application.EnableEvents = False
    Y001 = Y001 & "|" & Application.EnableEvents
    Application.EnableEvents = True
    Y001 = Y001 & "|" & Application.EnableEvents
End Function

Function Y002() As String
    Reset
    Application.DisplayAlerts = False
    Y002 = Application.DisplayAlerts
    Application.DisplayAlerts = True
    Y002 = Y002 & "|" & Application.DisplayAlerts
    Application.DisplayAlerts = False
End Function

Function Y003() As String
    Reset
    Application.Calculation = xlCalculationManual
    Y003 = Application.Calculation
    Application.Calculation = xlCalculationSemiautomatic
    Y003 = Y003 & "|" & Application.Calculation
    Application.Calculation = xlCalculationAutomatic
    Y003 = Y003 & "|" & Application.Calculation & "|" & xlCalculationManual & "|" & xlCalculationAutomatic
    On Error Resume Next
    Application.Calculation = 99
    Y003 = Y003 & "|" & Err.Number & "|" & Application.Calculation
End Function

Function Y004() As String
    Reset
    Application.StatusBar = "Working..."
    Y004 = Application.StatusBar
    Application.StatusBar = False
    Y004 = Y004 & "|" & Application.StatusBar & "|" & TypeName(Application.StatusBar)
    Application.StatusBar = 42
    Y004 = Y004 & "|" & Application.StatusBar & "|" & TypeName(Application.StatusBar)
    Application.StatusBar = False
End Function

Function Y005() As String
    Reset
    Y005 = Application.CutCopyMode & "|" & TypeName(Application.CutCopyMode)
    Range("A1").Copy
    Y005 = Y005 & "|" & Application.CutCopyMode & "|" & (Application.CutCopyMode = xlCopy)
    Range("A1").Cut
    Y005 = Y005 & "|" & (Application.CutCopyMode = xlCut)
    Application.CutCopyMode = False
End Function

Function Y006() As String
    Reset
    Y006 = Application.International(xlDecimalSeparator) & "|" & Application.International(xlThousandsSeparator) & "|" & Application.International(xlListSeparator) & "|" & Application.International(xlDateOrder) & "|" & Application.DecimalSeparator & "|" & Application.UseSystemSeparators
End Function

Function Y007() As String
    Reset
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollColumn = 3
    Y007 = ActiveWindow.ScrollRow & "|" & ActiveWindow.ScrollColumn
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Function

Function Y008() As String
    Reset
    Range("B2").Select
    ActiveWindow.FreezePanes = True
    Y008 = ActiveWindow.FreezePanes & "|" & ActiveWindow.SplitRow & "|" & ActiveWindow.SplitColumn
    ActiveWindow.FreezePanes = False
    Y008 = Y008 & "|" & ActiveWindow.FreezePanes
End Function

Function Y009() As String
    Reset
    ActiveWindow.Zoom = 85
    Y009 = ActiveWindow.Zoom
    ActiveWindow.Zoom = 100
    ActiveWindow.DisplayGridlines = False
    Y009 = Y009 & "|" & ActiveWindow.DisplayGridlines
    ActiveWindow.DisplayGridlines = True
    Y009 = Y009 & "|" & TypeName(ActiveWindow) & "|" & ActiveWindow.ActiveSheet.Name & "|" & ActiveWindow.ActiveCell.Address
End Function

Function Y010() As String
    Reset
    Y010 = TypeName(Application) & "|" & Application.Name & "|" & TypeName(Application.ActiveWorkbook) & "|" & (Application.ThisWorkbook Is Nothing) & "|" & Application.Workbooks.Count
End Function

Function Y011() As String
    Reset
    On Error Resume Next
    Range("A1").Value = 1
    Range("A1").Offset(-1).Value = 2
    Y011 = Err.Number & "|" & Err.Description
    Err.Clear
    Cells(0, 1).Value = 1
    Y011 = Y011 & "|" & Err.Number & "|" & Err.Description
    Err.Clear
    Cells(1, 16385).Value = 1
    Y011 = Y011 & "|" & Err.Number
End Function

Function Y012() As String
    Reset
    On Error Resume Next
    Dim o As Object
    Set o = Range("A1")
    o.NoSuchMethod
    Y012 = Err.Number & "|" & Err.Description
    Err.Clear
    Dim x As Variant
    x = ActiveSheet.NoSuchProperty
    Y012 = Y012 & "|" & Err.Number & "|" & Err.Description
End Function

Function Y013() As String
    Reset
    On Error Resume Next
    Dim r As Range
    r.Value = 1
    Y013 = Err.Number & "|" & Err.Description
    Err.Clear
    Set r = Range("A1:B2")
    r.Value = "ok"
    Y013 = Y013 & "|" & Err.Number & "|" & Range("B2").Value
End Function

Function Y014() As String
    Reset
    On Error Resume Next
    Range("A1").Select
    Worksheets.Add(After:=Worksheets(1)).Name = "Other"
    Worksheets("Sheet1").Range("A1").Select
    Y014 = Err.Number & "|" & Err.Description
    Err.Clear
    Worksheets("Sheet1").Range("A1").Activate
    Y014 = Y014 & "|" & Err.Number & "|" & ActiveSheet.Name
End Function

Function Y015() As String
    Reset
    On Error Resume Next
    Range("A1:B2").Merge
    Range("A1").Value = "m"
    Range("A1:A3").Sort Key1:=Range("A1")
    Y015 = Err.Number
    Err.Clear
    Range("B1").Insert Shift:=xlDown
    Y015 = Y015 & "|" & Err.Number
End Function

Function Y016() As String
    Reset
    Application.Goto Range("C5")
    Y016 = ActiveCell.Address & "|" & Selection.Address
    Worksheets.Add(After:=Worksheets(1)).Name = "Jump"
    Application.Goto Worksheets("Sheet1").Range("B2:C3")
    Y016 = Y016 & "|" & ActiveSheet.Name & "|" & Selection.Address & "|" & ActiveCell.Address
End Function

Function Y017() As String
    Reset
    Range("B2:C3").Select
    Range("C3").Activate
    Y017 = Selection.Address & "|" & ActiveCell.Address
    ActiveCell.Offset(1, 0).Select
    Y017 = Y017 & "|" & Selection.Address & "|" & ActiveCell.Address
    Selection.Value = "s"
    Y017 = Y017 & "|" & Range("C4").Value & "|" & Selection.Count
End Function

Function Y018() As String
    Reset
    Range("A1:B3").Select
    Selection.Value = 1
    Y018 = Application.WorksheetFunction.Sum(Selection) & "|" & Selection.Rows.Count & "|" & TypeName(Selection)
    ActiveCell.Value = 5
    Y018 = Y018 & "|" & Range("A1").Value
End Function

Function Y019() As String
    Reset
    Y019 = (Val(Application.Version) >= 16) & "|" & Application.ReferenceStyle & "|" & Application.ActiveSheet.Name
End Function

Function Y020() As String
    Reset
    Application.EnableCancelKey = xlDisabled
    Y020 = Application.EnableCancelKey
    Application.EnableCancelKey = xlInterrupt
    Application.Cursor = xlWait
    Y020 = Y020 & "|" & Application.Cursor
    Application.Cursor = xlDefault
    Y020 = Y020 & "|" & Application.Cursor & "|" & Application.Interactive
End Function

Function Y021() As String
    Reset
    Range("A1").Value = 3
    Y021 = Application.WorksheetFunction.Sum(Application.Range("A1")) & "|" & Application.Cells(1, 1).Value & "|" & Application.ActiveCell.Address & "|" & Application.Selection.Address
End Function

Function Y022() As String
    Reset
    On Error Resume Next
    Err.Raise 1004, , "custom"
    Y022 = Err.Number & "|" & Err.Description
    Err.Clear
    Dim v As Variant
    v = Worksheets("Sheet1").Range("A1:A2").Value(1)
    Y022 = Y022 & "|" & Err.Number
End Function

Function Y023() As String
    Reset
    Range("A1").Value = 1
    ThisWorkbook.Saved = True
    Y023 = ThisWorkbook.Saved
    Range("A1").Value = 2
    Y023 = Y023 & "|" & ThisWorkbook.Saved & "|" & ThisWorkbook.CodeName & "|" & ThisWorkbook.Worksheets.Count
End Function

Function Y024() As String
    Reset
    Y024 = Worksheets(1).CodeName & "|" & ThisWorkbook.Worksheets("Sheet1").Type & "|" & (Worksheets(1).Type = xlWorksheet)
End Function
