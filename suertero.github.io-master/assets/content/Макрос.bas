Attribute VB_Name = "NewMacros"
Sub ״נטפע()
Attribute ״נטפע.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.״נטפע"
'
' ״נטפע ּאךנמס
'
'
    Selection.Font.Name = "Times New Roman"
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Size = 14
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(-0.5)
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
    End With
    Selection.PageSetup.LeftMargin = CentimetersToPoints(2)
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
    End With
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.15)
End Sub
