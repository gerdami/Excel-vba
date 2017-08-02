Attribute VB_Name = "ModuleTestChart"
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    ActiveChart.ChartTitle.Text = "Graph X: yyyyyy"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Graph X: yyyyyy"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 15).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 15).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "Arial"
        .NameFarEast = "Arial"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 9.6
        .Italic = msoFalse
        .Kerning = 12
        .Name = "Arial"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    Range("G1").Select
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.ChartTitle.Select
    Selection.Format.TextFrame2.TextRange.Font.Size = 9
    Range("A1").Select
End Sub
