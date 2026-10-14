Option Explicit

' Worksheets: add, name, delete, move, copy, visibility, index, activation.

Function S001() As String
    Dim ws As Worksheet
    Reset
    Set ws = Worksheets.Add
    S001 = Worksheets.Count & "|" & (Worksheets(1).Name Like "Sheet#*") & "|" & Worksheets(2).Name & "|" & (ActiveSheet.Name = Worksheets(1).Name) & "|" & ws.Index
End Function

Function S002() As String
    Reset
    Worksheets.Add After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "Data"
    S002 = SheetList() & "|" & ActiveSheet.Index & "|" & Worksheets("data").Name
End Function

Function S003() As String
    Reset
    Worksheets.Add(After:=Worksheets("Sheet1")).Name = "B"
    Worksheets.Add(Before:=Worksheets("Sheet1")).Name = "A"
    Worksheets.Add(After:=Worksheets("B")).Name = "C"
    S003 = SheetList() & "|" & Worksheets("C").Index & "|" & ActiveSheet.Name
End Function

Function S004() As String
    Reset
    Worksheets.Add Count:=3
    S004 = Worksheets.Count & "|" & ActiveSheet.Index & "|" & Worksheets(Worksheets.Count).Name
End Function

Function S005() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Temp"
    Application.DisplayAlerts = False
    Worksheets("Temp").Delete
    Application.DisplayAlerts = True
    S005 = Worksheets.Count & "|" & SheetList() & "|" & ActiveSheet.Name
End Function

Function S006() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "X"
    Worksheets.Add(After:=Worksheets(2)).Name = "Y"
    Worksheets("X").Activate
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    S006 = SheetList() & "|" & ActiveSheet.Name
End Function

Function S007() As String
    Reset
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(1).Delete
    S007 = Err.Number & "|" & Err.Description & "|" & Worksheets.Count
    Application.DisplayAlerts = True
End Function

Function S008() As String
    Reset
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Worksheets("NoSuchSheet")
    S008 = Err.Number & "|" & Err.Description & "|" & (ws Is Nothing)
End Function

Function S009() As String
    Reset
    On Error Resume Next
    ActiveSheet.Name = "Bad/Name"
    S009 = Err.Number & "|" & ActiveSheet.Name
    Err.Clear
    ActiveSheet.Name = "This name is far too long for a sheet tab"
    S009 = S009 & "|" & Err.Number & "|" & ActiveSheet.Name
    Err.Clear
    ActiveSheet.Name = ""
    S009 = S009 & "|" & Err.Number & "|" & ActiveSheet.Name
End Function

Function S010() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Two"
    On Error Resume Next
    Worksheets("Two").Name = "sheet1"
    S010 = Err.Number & "|" & Err.Description & "|" & SheetList()
End Function

Function S011() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "A"
    Worksheets.Add(After:=Worksheets(2)).Name = "B"
    Worksheets.Add(After:=Worksheets(3)).Name = "C"
    Worksheets("A").Move After:=Worksheets("C")
    S011 = SheetList()
    Worksheets("C").Move Before:=Worksheets(1)
    S011 = S011 & "|" & SheetList() & "|" & Worksheets("Sheet1").Index
End Function

Function S012() As String
    Reset
    Range("A1").Value = 42
    Range("B2").Formula = "=A1*2"
    Worksheets("Sheet1").Copy After:=Worksheets("Sheet1")
    S012 = SheetList() & "|" & ActiveSheet.Name & "|" & ActiveSheet.Range("A1").Value & "|" & ActiveSheet.Range("B2").Formula & "|" & ActiveSheet.Range("B2").Value
End Function

Function S013() As String
    Reset
    Worksheets("Sheet1").Copy Before:=Worksheets(1)
    Worksheets("Sheet1").Copy After:=Worksheets(Worksheets.Count)
    S013 = SheetList()
End Function

Function S014() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Hid"
    Worksheets("Hid").Visible = xlSheetHidden
    S014 = Worksheets("Hid").Visible & "|" & Worksheets.Count & "|" & ActiveSheet.Name
    Worksheets("Hid").Visible = xlSheetVeryHidden
    S014 = S014 & "|" & Worksheets("Hid").Visible
    Worksheets("Hid").Visible = True
    S014 = S014 & "|" & Worksheets("Hid").Visible
End Function

Function S015() As String
    Reset
    On Error Resume Next
    Worksheets(1).Visible = xlSheetHidden
    S015 = Err.Number & "|" & Worksheets(1).Visible
End Function

Function S016() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "H"
    Worksheets("H").Visible = False
    On Error Resume Next
    Worksheets("H").Activate
    S016 = Err.Number & "|" & ActiveSheet.Name
    Err.Clear
    Worksheets("H").Select
    S016 = S016 & "|" & Err.Number & "|" & ActiveSheet.Name
End Function

Function S017() As String
    Dim ws As Worksheet, s As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "P"
    Worksheets.Add(After:=Worksheets(2)).Name = "Q"
    For Each ws In Worksheets
        s = s & ws.Index & ":" & ws.Name & ","
    Next
    S017 = s
End Function

Function S018() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "N2"
    S018 = Worksheets(1).Next.Name & "|" & (Worksheets(1).Previous Is Nothing) & "|" & (Worksheets(2).Next Is Nothing) & "|" & Worksheets(2).Previous.Name
End Function

Function S019() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Other"
    Worksheets("Sheet1").Activate
    Worksheets("Other").Range("C3").Value = "o"
    S019 = ActiveSheet.Name & "|" & Worksheets("Other").Range("C3").Value & "|" & IsEmpty(Range("C3").Value)
End Function

Function S020() As String
    Reset
    S020 = TypeName(Worksheets) & "|" & TypeName(Sheets) & "|" & TypeName(ActiveSheet) & "|" & TypeName(Worksheets(1)) & "|" & Sheets.Count
End Function

Function S021() As String
    Reset
    Worksheets.Add After:=Worksheets(1)
    S021 = (ActiveSheet.Name Like "Sheet#*")
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Worksheets.Add After:=Worksheets(1)
    S021 = S021 & "|" & (ActiveSheet.Name Like "Sheet#*") & "|" & ActiveSheet.Index
    Application.DisplayAlerts = True
End Function

Function S022() As String
    Reset
    S022 = Worksheets("Sheet1").Parent.Name = ThisWorkbook.Name
    S022 = S022 & "|" & TypeName(Worksheets("Sheet1").Parent) & "|" & Worksheets("Sheet1").Range("A1").Parent.Name & "|" & Range("B2").Worksheet.Name
End Function

Function S023() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "My Data"
    Worksheets("My Data").Range("A1").Value = 5
    S023 = Worksheets("My Data").Range("A1").Address(External:=True) & "|" & Range("'My Data'!A1").Value
End Function

Function S024() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "S2"
    Worksheets("S2").Select
    S024 = ActiveSheet.Name
    Worksheets("Sheet1").Select
    S024 = S024 & "|" & ActiveSheet.Name
End Function

Function S025() As String
    Reset
    On Error Resume Next
    Worksheets("Sheet1").Range("A1").Value = 1
    Worksheets.Add(After:=Worksheets(1)).Name = "Z"
    Worksheets("Sheet1").Range("B2").Select
    S025 = Err.Number & "|" & ActiveSheet.Name
End Function

Function S026() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Calc"
    Worksheets("Calc").Range("A1").Value = 10
    Worksheets("Sheet1").Range("A1").Formula = "=Calc!A1*3"
    Worksheets("Calc").Name = "Calc2"
    S026 = Worksheets("Sheet1").Range("A1").Formula & "|" & Worksheets("Sheet1").Range("A1").Value
End Function

Function S027() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Gone"
    Worksheets("Gone").Range("A1").Value = 7
    Worksheets("Sheet1").Range("A1").Formula = "=Gone!A1"
    Application.DisplayAlerts = False
    Worksheets("Gone").Delete
    Application.DisplayAlerts = True
    S027 = Worksheets("Sheet1").Range("A1").Formula & "|" & Worksheets("Sheet1").Range("A1").Text
End Function

Function S028() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Mv"
    Worksheets("Mv").Move Before:=Worksheets("Sheet1")
    S028 = SheetList() & "|" & Worksheets(1).Name & "|" & ActiveSheet.Name
End Function

Function S029() As String
    Dim i As Long
    Reset
    For i = 1 To 5
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "W" & i
    Next
    Application.DisplayAlerts = False
    For i = Worksheets.Count To 1 Step -1
        If Worksheets(i).Name Like "W[24]" Then Worksheets(i).Delete
    Next
    Application.DisplayAlerts = True
    S029 = SheetList()
End Function

Function S030() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Hid"
    Worksheets("Hid").Visible = xlSheetHidden
    S030 = Worksheets.Count & "|" & Worksheets("Hid").Index
    Worksheets("Hid").Range("A1").Value = "still writable"
    S030 = S030 & "|" & Worksheets("Hid").Range("A1").Value
End Function

Function S031() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Copied"
    Worksheets("Copied").Range("A1:B2").Value = 3
    Worksheets("Copied").Range("A1").Font.Bold = True
    Worksheets("Copied").Copy After:=Worksheets("Copied")
    S031 = ActiveSheet.Name & "|" & Application.WorksheetFunction.Sum(ActiveSheet.Range("A1:B2")) & "|" & ActiveSheet.Range("A1").Font.Bold & "|" & SheetList()
End Function

Function S032() As String
    Reset
    S032 = Worksheets(1).Name & "|" & Worksheets("SHEET1").Name & "|" & Sheets("sheet1").Index
End Function

Function S033() As String
    Reset
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Worksheets(5)
    S033 = Err.Number & "|" & Err.Description
    Err.Clear
    Set ws = Worksheets(0)
    S033 = S033 & "|" & Err.Number
End Function

Function S034() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Dup"
    On Error Resume Next
    Worksheets.Add(After:=Worksheets(2)).Name = "dup"
    S034 = Err.Number & "|" & Worksheets.Count & "|" & (ActiveSheet.Name Like "Sheet*")
End Function

Function S035() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Apos"
    On Error Resume Next
    Worksheets("Apos").Name = "'Quoted"
    S035 = Err.Number & "|" & Worksheets(2).Name
    Err.Clear
    Worksheets("Apos").Name = "Has 'inner' quote"
    S035 = S035 & "|" & Err.Number & "|" & Worksheets(2).Name & "|" & Worksheets(2).Range("A1").Address(External:=True)
End Function

Function S036() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Second"
    Worksheets("Sheet1").Activate
    Worksheets("Second").Range("B5").Value = 1
    Worksheets("Second").Activate
    S036 = ActiveCell.Address & "|" & Selection.Address
    Range("B5").Select
    Worksheets("Sheet1").Activate
    Worksheets("Second").Activate
    S036 = S036 & "|" & ActiveCell.Address
End Function

Function S037() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "T1"
    Worksheets.Add(After:=Worksheets(2)).Name = "T2"
    Worksheets(Array("T1", "T2")).Visible = True
    S037 = Worksheets(Array("T1", "T2")).Count & "|" & TypeName(Worksheets(Array("T1", "T2")))
End Function

Function S038() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Movable"
    Worksheets("Movable").Move After:=Worksheets(Worksheets.Count)
    Worksheets("Sheet1").Move After:=Worksheets("Movable")
    S038 = SheetList() & "|" & Worksheets("Sheet1").Index
End Function

Function S039() As String
    Reset
    Dim n As Long
    n = Worksheets.Count
    Worksheets.Add Before:=Worksheets(1), Count:=2
    S039 = Worksheets.Count - n & "|" & ActiveSheet.Index & "|" & Worksheets(Worksheets.Count).Name
End Function

Function S040() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Ren"
    Worksheets("Ren").Range("A1").Value = 1
    Worksheets("Sheet1").Range("A1").Formula = "=SUM(Ren!A1:A3)"
    Worksheets("Ren").Range("A2").Value = 2
    S040 = Worksheets("Sheet1").Range("A1").Value
    Worksheets("Ren").Name = "Ren 2"
    S040 = S040 & "|" & Worksheets("Sheet1").Range("A1").Formula
End Function
