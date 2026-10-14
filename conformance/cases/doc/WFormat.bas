Option Explicit

Function F01() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "Hello world"
    d.Range(0, 5).Font.Bold = True
    F01 = d.Range(0, 5).Font.Bold & "|" & d.Range(6, 11).Font.Bold & "|" & d.Content.Font.Bold
End Function

Function F02() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "Hello world"
    d.Range(6, 11).Font.Italic = True
    F02 = d.Range(0, 5).Font.Italic & "|" & d.Range(6, 11).Font.Italic & "|" & d.Content.Font.Italic & "|" & d.Words(2).Italic
End Function

Function F03() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "Hello world"
    d.Range(0, 5).Font.Size = 14
    F03 = d.Range(0, 5).Font.Size & "|" & d.Range(6, 11).Font.Size & "|" & d.Content.Font.Size
End Function

Function F04() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "Hello world"
    d.Range(0, 5).Font.Name = "Arial"
    F04 = d.Range(0, 5).Font.Name & "|" & d.Range(6, 11).Font.Name & "|[" & d.Content.Font.Name & "]"
End Function

Function F05() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "One" & vbCr & "Two"
    d.Paragraphs(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    F05 = d.Paragraphs(1).Alignment & "|" & d.Paragraphs(2).Alignment & "|" & d.Content.ParagraphFormat.Alignment
End Function

Function F06() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "One" & vbCr & "Two"
    d.Paragraphs(2).Alignment = wdAlignParagraphRight
    d.Paragraphs(1).Format.Alignment = wdAlignParagraphJustify
    F06 = d.Paragraphs(1).Range.ParagraphFormat.Alignment & "|" & d.Paragraphs(2).Format.Alignment
End Function

Function F07() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "abc"
    F07 = d.Content.Font.Bold & "|" & d.Content.Font.Italic & "|" & d.Content.Font.Name & "|" & d.Content.Font.Size & "|" & d.Content.ParagraphFormat.Alignment
End Function

Function F08() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "Hello world"
    d.Range(0, 5).Font.Underline = wdUnderlineSingle
    F08 = d.Range(0, 5).Font.Underline & "|" & d.Range(6, 11).Font.Underline & "|" & d.Content.Font.Underline
End Function

Function F09() As String
    Dim d As Document, r As Range
    Set d = ActiveDocument
    d.Content.Text = "abc"
    d.Range(0, 3).Font.Bold = True
    Set r = d.Range(0, 3)
    r.InsertAfter "def"
    F09 = d.Range(3, 6).Font.Bold & "|" & r.Text
End Function

Function F10() As String
    Dim d As Document, s As String
    Set d = ActiveDocument
    d.Content.Text = "abc"
    d.Range(0, 3).Font.Bold = wdToggle
    s = d.Range(0, 3).Font.Bold
    d.Range(0, 3).Font.Bold = wdToggle
    F10 = s & "|" & d.Range(0, 3).Font.Bold
End Function

Function F11() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "Hello world"
    d.Words(1).Bold = True
    F11 = d.Characters(1).Font.Bold & "|" & d.Characters(5).Bold & "|" & d.Characters(6).Bold & "|[" & d.Words(1).Text & "]"
End Function

Function F12() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "Title" & vbCr & "Body"
    d.Paragraphs(1).Style = d.Styles(wdStyleHeading1)
    F12 = d.Paragraphs(1).Style & "|" & d.Paragraphs(2).Style & "|" & d.Paragraphs(1).Range.Style.NameLocal
End Function

Function F13() As String
    F13 = ActiveDocument.Styles("Normal").Font.Name & "|" & ActiveDocument.Styles("Normal").Font.Size & "|" & ActiveDocument.Styles(wdStyleNormal).NameLocal
End Function

Function F14() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "Hello world"
    d.Range(0, 5).Bold = True
    d.Range(6, 11).Italic = True
    F14 = d.Range(0, 5).Bold & "|" & d.Range(0, 5).Italic & "|" & d.Range(6, 11).Italic & "|" & d.Content.Bold
End Function

Function F15() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "Hello world"
    d.Range(0, 5).Font.Color = wdColorRed
    F15 = d.Range(0, 5).Font.Color & "|" & d.Range(6, 11).Font.Color & "|" & d.Content.Font.Color
End Function

Function F16() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "abc"
    d.Content.Font.Size = 10.5
    F16 = d.Content.Font.Size & "|" & TypeName(d.Content.Font.Size)
End Function

Function F17() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "One"
    d.Paragraphs(1).Alignment = wdAlignParagraphCenter
    d.Content.InsertParagraphAfter
    d.Content.InsertAfter "Two"
    F17 = d.Paragraphs.Count & "|" & d.Paragraphs(2).Alignment
End Function

Function F18() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "One" & vbCr & "Two"
    d.Paragraphs(1).LeftIndent = 36
    d.Paragraphs(2).Format.SpaceAfter = 12
    F18 = d.Paragraphs(1).LeftIndent & "|" & d.Paragraphs(2).LeftIndent & "|" & d.Paragraphs(2).SpaceAfter & "|" & d.Paragraphs(1).SpaceAfter
End Function

Function F19() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "abc def"
    With d.Range(0, 3).Font
        .Bold = True
        .Italic = True
        .Size = 20
    End With
    F19 = d.Range(0, 3).Font.Bold & d.Range(0, 3).Font.Italic & d.Range(0, 3).Font.Size & "|" & d.Range(4, 7).Font.Size
End Function

Function F20() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "abc"
    d.Content.Bold = True
    F20 = d.Content.Bold & "|" & d.Content.Font.Bold & "|" & d.Characters(4).Bold
End Function

Function F21() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "Hello world"
    d.Range(0, 5).Font.Bold = True
    d.Range(3, 8).Font.Bold = False
    F21 = d.Range(0, 3).Font.Bold & "|" & d.Range(3, 5).Font.Bold & "|" & d.Range(0, 5).Font.Bold
End Function

Function F22() As String
    Dim d As Document
    Set d = ActiveDocument
    d.Content.Text = "abc"
    F22 = TypeName(d.Content.Font) & "|" & TypeName(d.Content.ParagraphFormat) & "|" & TypeName(d.Paragraphs(1)) & "|" & TypeName(d.Content) & "|" & TypeName(d.Styles(wdStyleNormal))
End Function
