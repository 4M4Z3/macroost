Option Explicit

' Shared helpers. Every procedure here takes an argument so the oracle harness does not run it
' as a case.

Public Sub Reset(Optional ByVal keep As Long = 0)
    Dim i As Long, fresh As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    ThisWorkbook.Unprotect
    ThisWorkbook.Unprotect "wb"
    ' A brand-new Sheet1 for every case, so nothing a case did to a sheet (filters, protection
    ' modes, stale last cells) leaks into the next one.
    Set fresh = Worksheets.Add(Before:=Worksheets(1))
    For i = Worksheets.Count To 2 Step -1
        Worksheets(i).Visible = xlSheetVisible
        Worksheets(i).Delete
    Next
    fresh.Name = "Sheet1"
    fresh.Activate
    fresh.Range("A1").Select
    For i = ThisWorkbook.Names.Count To 1 Step -1
        ThisWorkbook.Names(i).Delete
    Next
    On Error GoTo 0
    Err.Clear
End Sub

Public Sub FillCol(ByVal n As Long, Optional ByVal col As Long = 1)
    Dim i As Long
    For i = 1 To n
        Worksheets("Sheet1").Cells(i, col).Value = i
    Next
End Sub

' Names of the worksheets in tab order, comma separated.
Public Function SheetList(Optional ByVal sep As String = ",") As String
    Dim ws As Worksheet, s As String
    For Each ws In ThisWorkbook.Worksheets
        If Len(s) > 0 Then s = s & sep
        s = s & ws.Name
    Next
    SheetList = s
End Function

' Err.Number of the last trapped error as text ("0" when none), then clears it.
Public Function ErrNo(Optional ByVal dummy As Long = 0) As String
    ErrNo = CStr(Err.Number)
    Err.Clear
End Function

' Addresses of every area, relative, separated by ";".
Public Function AreaList(ByVal r As Range) As String
    Dim a As Range, s As String
    For Each a In r.Areas
        If Len(s) > 0 Then s = s & ";"
        s = s & a.Address(False, False)
    Next
    AreaList = s
End Function

' Values of a block, row by row: "1,2;3,4".
Public Function Dump(ByVal r As Range) As String
    Dim i As Long, j As Long, s As String
    For i = 1 To r.Rows.Count
        If i > 1 Then s = s & ";"
        For j = 1 To r.Columns.Count
            If j > 1 Then s = s & ","
            s = s & CStr(r.Cells(i, j).Value)
        Next
    Next
    Dump = s
End Function

' The error code of an error value ("2042" for #N/A), or "noerr".
Public Function ErrCode(ByVal v As Variant) As String
    Dim c As Variant
    If Not IsError(v) Then
        ErrCode = "noerr"
        Exit Function
    End If
    For Each c In Array(2000, 2007, 2015, 2023, 2029, 2036, 2042)
        If v = CVErr(c) Then
            ErrCode = CStr(c)
            Exit Function
        End If
    Next
    ErrCode = "?"
End Function
