Option Explicit

' Sheet and workbook protection, passwords, UserInterfaceOnly, Locked / FormulaHidden.

Function P001() As String
    Reset
    ActiveSheet.Protect Password:="pw"
    P001 = ActiveSheet.ProtectContents
    On Error Resume Next
    Range("A1").Value = 1
    P001 = P001 & "|" & Err.Number & "|" & Err.Description
    Err.Clear
    ActiveSheet.Unprotect "wrong"
    P001 = P001 & "|" & Err.Number & "|" & Err.Description & "|" & ActiveSheet.ProtectContents
    Err.Clear
    ActiveSheet.Unprotect "PW"
    P001 = P001 & "|" & Err.Number & "|" & ActiveSheet.ProtectContents
    Err.Clear
    ActiveSheet.Unprotect "pw"
    P001 = P001 & "|" & Err.Number & "|" & ActiveSheet.ProtectContents
End Function

Function P002() As String
    Reset
    Range("A1").Locked = False
    ActiveSheet.Protect "pw"
    On Error Resume Next
    Range("A1").Value = "ok"
    P002 = Err.Number & "|" & Range("A1").Value
    Err.Clear
    Range("A2").Value = "no"
    P002 = P002 & "|" & Err.Number & "|" & IsEmpty(Range("A2").Value)
    Err.Clear
    Range("A1:A2").Value = 5
    P002 = P002 & "|" & Err.Number & "|" & Range("A1").Value
    ActiveSheet.Unprotect "pw"
End Function

Function P003() As String
    Reset
    ' On a sheet of its own: the UserInterfaceOnly flag outlives Unprotect for the session.
    Worksheets.Add(After:=Worksheets(1)).Name = "UI"
    ActiveSheet.Protect Password:="pw", UserInterfaceOnly:=True
    On Error Resume Next
    Range("A1").Value = 3
    P003 = Err.Number & "|" & Range("A1").Value & "|" & ActiveSheet.ProtectContents & "|" & ActiveSheet.ProtectionMode
    ActiveSheet.Unprotect "pw"
    P003 = P003 & "|" & ActiveSheet.ProtectionMode
    ActiveSheet.Protect
    Range("A2").Value = 4
    P003 = P003 & "|" & Err.Number & "|" & ActiveSheet.ProtectionMode
    Err.Clear
    ActiveSheet.Unprotect
    ActiveSheet.Protect UserInterfaceOnly:=False
    Range("A3").Value = 5
    P003 = P003 & "|" & Err.Number & "|" & ActiveSheet.ProtectionMode
    ActiveSheet.Unprotect
End Function

Function P004() As String
    Reset
    ActiveSheet.Protect
    On Error Resume Next
    Range("A1").Value = 1
    P004 = Err.Description & "|"
    Err.Clear
    Range("A1").Font.Bold = True
    P004 = P004 & Err.Number
    Err.Clear
    Range("A1").Interior.Color = vbRed
    P004 = P004 & "|" & Err.Number
    Err.Clear
    Rows(2).Insert
    P004 = P004 & "|" & Err.Number
    Err.Clear
    Range("A1").ClearContents
    P004 = P004 & "|" & Err.Number
    Err.Clear
    ActiveSheet.Unprotect
    P004 = P004 & "|" & Err.Number & "|" & ActiveSheet.ProtectContents
End Function

Function P005() As String
    Reset
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False
    P005 = ActiveSheet.ProtectContents & "|" & ActiveSheet.ProtectDrawingObjects & "|" & ActiveSheet.ProtectScenarios
    ActiveSheet.Unprotect
    ActiveSheet.Protect Contents:=False
    P005 = P005 & "|" & ActiveSheet.ProtectContents & "|" & ActiveSheet.ProtectDrawingObjects
    ActiveSheet.Unprotect
End Function

Function P006() As String
    Reset
    ActiveSheet.Protect "pw"
    ActiveSheet.Protect "pw"
    On Error Resume Next
    ActiveSheet.Protect "other"
    P006 = Err.Number & "|" & ActiveSheet.ProtectContents
    Err.Clear
    ActiveSheet.Unprotect "pw"
    P006 = P006 & "|" & Err.Number & "|" & ActiveSheet.ProtectContents
    Err.Clear
    ActiveSheet.Unprotect "other"
    P006 = P006 & "|" & Err.Number & "|" & ActiveSheet.ProtectContents
End Function

Function P007() As String
    Reset
    ActiveSheet.Protect Password:="Secret"
    On Error Resume Next
    ActiveSheet.Unprotect "secret"
    P007 = Err.Number & "|" & ActiveSheet.ProtectContents
    Err.Clear
    ActiveSheet.Unprotect Password:="Secret"
    P007 = P007 & "|" & Err.Number & "|" & ActiveSheet.ProtectContents
End Function

Function P008() As String
    Reset
    Worksheets.Add(After:=Worksheets(1)).Name = "Locked"
    Worksheets("Locked").Protect "pw"
    Worksheets("Sheet1").Activate
    On Error Resume Next
    Worksheets("Locked").Range("B2").Value = 1
    P008 = Err.Number & "|" & Worksheets("Locked").ProtectContents
    Err.Clear
    Worksheets("Locked").Unprotect ("pw")
    P008 = P008 & "|" & Err.Number & "|" & Worksheets("Locked").ProtectContents
End Function

Function P009() As String
    Reset
    P009 = Range("A1").Locked & "|" & Range("A1").FormulaHidden & "|" & Range("A1:B2").Locked
    Range("A1:A3").Locked = False
    P009 = P009 & "|" & Range("A2").Locked & "|" & IsNull(Range("A3:A4").Locked) & "|" & Range("A4").Locked
    Range("B1").FormulaHidden = True
    P009 = P009 & "|" & Range("B1").FormulaHidden & "|" & IsNull(Range("B1:B2").FormulaHidden)
End Function

Function P010() As String
    Reset
    Range("A1").Formula = "=1+2"
    Range("A1").FormulaHidden = True
    ActiveSheet.Protect
    P010 = Range("A1").Formula & "|" & Range("A1").HasFormula & "|" & Range("A1").Value
    ActiveSheet.Unprotect
End Function

Function P011() As String
    Reset
    Range("B2").Locked = False
    ActiveSheet.Protect
    On Error Resume Next
    Range("B2").Locked = True
    P011 = Err.Number & "|" & Range("B2").Locked
    Err.Clear
    ActiveSheet.Unprotect
    Range("B2").Locked = True
    P011 = P011 & "|" & Err.Number & "|" & Range("B2").Locked
End Function

Function P012() As String
    Reset
    ThisWorkbook.Protect Password:="wb", Structure:=True
    P012 = ThisWorkbook.ProtectStructure
    On Error Resume Next
    Worksheets.Add
    P012 = P012 & "|" & Err.Number & "|" & Worksheets.Count
    Err.Clear
    ActiveSheet.Name = "Renamed"
    P012 = P012 & "|" & Err.Number & "|" & ActiveSheet.Name
    Err.Clear
    ThisWorkbook.Unprotect "bad"
    P012 = P012 & "|" & Err.Number & "|" & ThisWorkbook.ProtectStructure
    Err.Clear
    ThisWorkbook.Unprotect Password:="wb"
    P012 = P012 & "|" & Err.Number & "|" & ThisWorkbook.ProtectStructure
End Function

Function P013() As String
    Reset
    P013 = ThisWorkbook.ProtectStructure & "|" & ThisWorkbook.ProtectWindows
    ThisWorkbook.Protect
    P013 = P013 & "|" & ThisWorkbook.ProtectStructure
    Range("A1").Value = "cells still editable"
    P013 = P013 & "|" & Range("A1").Value
    ThisWorkbook.Unprotect
    P013 = P013 & "|" & ThisWorkbook.ProtectStructure
End Function

Function P014() As String
    Reset
    ThisWorkbook.Protect "wb"
    On Error Resume Next
    ThisWorkbook.Unprotect
    P014 = Err.Number & "|" & Err.Description & "|" & ThisWorkbook.ProtectStructure
    Err.Clear
    ThisWorkbook.Unprotect "wb"
    P014 = P014 & "|" & Err.Number & "|" & ThisWorkbook.ProtectStructure
End Function

Function P015() As String
    Reset
    ActiveSheet.Protect "pw", AllowFormattingCells:=True, AllowInsertingRows:=True
    P015 = ActiveSheet.Protection.AllowFormattingCells & "|" & ActiveSheet.Protection.AllowInsertingRows & "|" & ActiveSheet.Protection.AllowDeletingRows
    On Error Resume Next
    Range("A1").Font.Bold = True
    P015 = P015 & "|" & Err.Number & "|" & Range("A1").Font.Bold
    Err.Clear
    Range("A1").Value = 1
    P015 = P015 & "|" & Err.Number
    ActiveSheet.Unprotect "pw"
End Function

Function P016() As String
    Reset
    ActiveSheet.Protect "pw"
    On Error Resume Next
    Range("A1").Select
    P016 = Err.Number & "|" & ActiveCell.Address
    Err.Clear
    ActiveSheet.Range("A1").Copy
    P016 = P016 & "|" & Err.Number
    Application.CutCopyMode = False
    ActiveSheet.Unprotect "pw"
End Function

Function P017() As String
    Reset
    ActiveSheet.Protect ""
    On Error Resume Next
    ActiveSheet.Unprotect
    P017 = Err.Number & "|" & ActiveSheet.ProtectContents
    Err.Clear
    ActiveSheet.Protect Password:="Sample#42"
    ActiveSheet.Unprotect ("Sample#42")
    P017 = P017 & "|" & Err.Number & "|" & ActiveSheet.ProtectContents
End Function

Function P018() As String
    Reset
    Range("A1:B2").Locked = False
    ActiveSheet.Protect "pw"
    On Error Resume Next
    Range("A1:C1").Value = 9
    P018 = Err.Number & "|" & Range("A1").Value & "|" & Range("C1").Value
    Err.Clear
    Range("B2").ClearContents
    P018 = P018 & "|" & Err.Number
    ActiveSheet.Unprotect "pw"
End Function
