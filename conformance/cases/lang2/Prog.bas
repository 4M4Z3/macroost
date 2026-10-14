Option Explicit

Private Sub Inc(x As Variant)
    x = x + 1
End Sub

Private Sub IncLong(x As Long)
    x = x + 1
End Sub

Private Function Fails() As Long
    Dim c As New Collection
    Fails = c(1)
End Function

Private Function SumAll(ParamArray xs() As Variant) As Double
    Dim v As Variant
    For Each v In xs
        SumAll = SumAll + v
    Next
End Function

Private Function Opt(Optional a As Variant, Optional b As Long = 7) As String
    Opt = IsMissing(a) & "," & b
End Function

Private Function Counter() As Long
    Static n As Long
    n = n + 1
    Counter = n
End Function

Function P001() As String
    Dim i As Integer
    On Error Resume Next
    i = 32767
    i = i + 1
    P001 = Err.Number & "|" & i
End Function

Function P002() As String
    Dim v As Variant
    On Error Resume Next
    v = CInt(32767)
    v = v + 1
    P002 = Err.Number & "|" & TypeName(v) & "|" & v
End Function

Function P003() As String
    Dim i As Integer, n As Long
    On Error Resume Next
    For i = 32760 To 32767
        n = n + 1
    Next
    P003 = Err.Number & "|" & n & "|" & i
End Function

Function P004() As String
    Dim i As Long, s As String
    For i = 10 To 1 Step -3
        s = s & i & ","
    Next
    P004 = s & "after=" & i
End Function

Function P005() As String
    Dim x As Double, s As String
    For x = 0 To 1 Step 0.25
        s = s & x & ","
    Next
    P005 = s & "after=" & x
End Function

Function P006() As String
    Dim i As Long, s As String
    For i = 1 To 10
        If i = 3 Then i = 8
        s = s & i & ","
    Next
    P006 = s
End Function

Function P007() As String
    Dim i As Long, j As Long, s As String
    For i = 1 To 3
        For j = 1 To 3
            If j = 2 Then Exit For
            s = s & i & j & ","
        Next j
    Next i
    P007 = s
End Function

Function P008() As String
    Dim i As Long, s As String
    On Error Resume Next
    For i = 1 To 3
        s = s & (10 / (2 - i)) & ","
    Next
    P008 = s & Err.Number
End Function

Function P009() As String
    Dim s As String
    On Error GoTo H
    s = "a"
    s = s & Fails()
    s = s & "b"
    P009 = s
    Exit Function
H:
    s = s & "[" & Err.Number & "]"
    Resume Next
End Function

Function P010() As String
    On Error Resume Next
    Err.Raise 5
    P010 = Err.Number & "|"
    Err.Clear
    P010 = P010 & Err.Number & "|"
    On Error GoTo 0
    P010 = P010 & Err.Number
End Function

Function P011() As String
    Dim a(1 To 3) As Variant
    a(1) = 1
    Inc a(1)
    P011 = a(1)
End Function

Function P012() As String
    Dim x As Long
    x = 1
    IncLong x
    IncLong (x)
    Call IncLong(x)
    P012 = x
End Function

Function P013() As String
    P013 = SumAll(1, 2, 3.5) & "|" & SumAll() & "|" & Opt() & "|" & Opt(1, 2)
End Function

Function P014() As String
    Counter
    Counter
    P014 = Counter()
End Function

Function P015() As String
    Dim a() As Long, b() As Long
    ReDim a(1 To 2)
    a(1) = 5
    b = a
    b(1) = 9
    P015 = a(1) & "|" & b(1)
End Function

Function P016() As String
    Dim a(3) As Integer, m(1 To 2, 0 To 4) As String
    P016 = LBound(a) & UBound(a) & "|" & LBound(m, 2) & UBound(m, 2) & "|" & TypeName(m(1, 1)) & "|" & Len(m(2, 4))
End Function

Function P017() As String
    Dim a() As Long, e As Long
    ReDim a(2)
    a(2) = 4
    ReDim a(3)
    P017 = a(2) & "|"
    Erase a
    On Error Resume Next
    e = UBound(a)
    P017 = P017 & Err.Number
End Function

Function P018() As String
    Dim s As String
    Select Case "banana"
        Case "apple" To "cherry": s = "range"
        Case Else: s = "else"
    End Select
    Select Case 5
        Case Is < 3: s = s & "|lt3"
        Case 4 To 6, 9: s = s & "|4to6"
    End Select
    P018 = s
End Function

Function P019() As String
    Dim i As Integer, s As String, d As Date, b As Boolean
    i = "2.5"
    s = True
    d = "10/14/2026"
    b = "False"
    P019 = i & "|" & s & "|" & TypeName(d) & "|" & d & "|" & b
End Function

Function P020() As String
    Dim n As Long
    n = 1
    GoSub AddTen
    GoSub AddTen
    P020 = n
    Exit Function
AddTen:
    n = n + 10
    Return
End Function

Function P021() As String
    Dim s As String
    s = "Hello"
    Mid(s, 2, 3) = "EYY"
    P021 = s
End Function

Function P022() As String
    Dim c As New Collection, d As Object, v As Variant, s As String
    c.Add 3: c.Add 1: c.Add 2
    For Each v In c: s = s & v: Next
    Set d = CreateObject("Scripting.Dictionary")
    d("z") = 1: d("a") = 2
    s = s & "|" & Join(d.Keys, ",") & "|" & d.Count & "|" & d.Exists("Z")
    P022 = s
End Function

Function P023() As String
    Dim arr As Variant, inner As Variant
    arr = Array(Array(1, 2), Array(3, 4))
    inner = arr(1)
    P023 = arr(0)(1) & "|" & inner(0) & "|" & UBound(arr)
End Function

Function P024() As String
    Dim x As Variant
    P024 = TypeName(x) & "|" & IsEmpty(x) & "|" & (x = 0) & "|" & (x = "") & "|" & Len(x)
End Function

Function P025() As String
    Dim v As Variant, w As Variant
    v = Null
    w = v & "x"
    P025 = IsNull(v + 1) & "|" & w & "|" & IsNull(v = Null) & "|" & IsNull(Null And False) & "|" & (Null Or True)
End Function

Function P026() As String
    Dim s As String, i As Long
    For i = 1 To 5
        Select Case i Mod 3
            Case 0: s = s & "F"
            Case Else: s = s & i
        End Select
    Next
    P026 = s
End Function

Function P027() As String
    Dim x As Long, y As Long
    x = 5: y = 0
    On Error GoTo H
    x = x \ y
    P027 = "no error"
    Exit Function
H:
    P027 = Err.Number & ":" & Err.Description & ":" & Erl
End Function

Function P028() As String
    Dim t As Double
    t = 0.1 * 3
    P028 = (t = 0.3) & "|" & t & "|" & (CDec(0.1) * 3 = 0.3) & "|" & Format(t, "0.00000000000000000")
End Function

Function P029() As String
    Dim s As String
    s = "a" & vbTab & "b" & vbCrLf & "c"
    P029 = Len(s) & "|" & InStr(s, vbLf) & "|" & Asc(Mid(s, 2, 1)) & "|" & Replace(s, vbCrLf, "/")
End Function

Function P030() As String
    Dim d As Date
    d = #1/31/2026 11:59:59 PM#
    P030 = DateAdd("s", 1, d) & "|" & DatePart("q", d) & "|" & DatePart("ww", d) & "|" & DatePart("y", d) & "|" & Format(d, "h:mm:ss AM/PM")
End Function

Function P031() As String
    Dim i As Long, s As String
    i = 0
    Do
        i = i + 1
        If i = 2 Then GoTo Skip
        s = s & i
Skip:
    Loop While i < 4
    P031 = s
End Function

Function P032() As String
    Dim x As Integer, y As Long, z As Double
    x = 7: y = 2: z = 2
    P032 = TypeName(x / y) & "|" & TypeName(x \ y) & "|" & TypeName(x * y) & "|" & TypeName(x + z) & "|" & TypeName(x Mod y) & "|" & TypeName(x * x)
End Function

Function P033() As String
    Dim s As String
    s = String$(3, "z") & Left$("abc", 1) & UCase$("q") & Trim$("  t ")
    P033 = s
End Function

Function P034() As String
    Dim b As Byte
    On Error Resume Next
    b = 255
    b = b + 1
    P034 = Err.Number & "|" & b
End Function

Function P035() As String
    Dim c As Currency
    c = 1 / 3
    P035 = c & "|" & TypeName(c * 3) & "|" & (c * 3)
End Function

Function P036() As String
    Dim s As String
    s = Format(-0.5, "0") & "|" & Format(0.5, "0") & "|" & Format(1.5, "0") & "|" & Format(2.5, "0") & "|" & Format(-1234567.891, "#,##0.0") & "|" & Format(0, "0.00;(0.00);""zero""")
    P036 = s
End Function

Function P037() As String
    Dim a As Variant
    a = Split("x y  z")
    P037 = UBound(a) & "|" & a(2) & "|" & Join(Filter(Array("ab", "cd", "abc"), "ab"), ",")
End Function

Function P038() As String
    Dim i As Long, total As Long
    i = 1
    While i <= 100
        total = total + i
        i = i + 1
    Wend
    P038 = total
End Function

Function P039() As String
    P039 = IIf(True, "y", 1 / 0)
End Function

Function P040() As String
    Dim x As Long
    x = 3
    P040 = (x > 1 And x < 5) & "|" & (x > 1 Or 1 / 0 > 0)
End Function
