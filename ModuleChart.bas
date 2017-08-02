Attribute VB_Name = "ModuleChart"
Option Explicit

Sub FreeCharts()
' Uncheck AutoScaleFont and Size with cells
Dim wkSheet As Worksheet

    For Each wkSheet In ActiveWorkbook.Worksheets
        wkSheet.Activate
        ApplySettingsToActiveSheetCharts
    Next wkSheet
    Worksheets(1).Activate
End Sub

Sub ApplySettingsToActiveSheetCharts()
' Macro recorded 20/12/2010 by gerdami
'
Dim oChart As ChartObject

    For Each oChart In ActiveSheet.ChartObjects
        oChart.Placement = xlMove 'Move but don't size with cells
        'oChart.Placement = xlFreeFloating 'Prevents from resizing
'        oChart.Chart.Axes(xlValue).TickLabels.NumberFormatLinked = 0  'Do not link number format
'        oChart.Chart.Axes(xlValue, xlSecondary).TickLabels.NumberFormatLinked = 0 'Do not link number format
        
        With oChart.Chart.ChartArea
         '.Border.LineStyle = 0 'No line around the chart
         .AutoScaleFont = False 'Uncheck font autoscale
         '.Font.Size = 8 'Force font size to 8
        End With
    Next oChart
       
End Sub
Sub IDRSizeTheChart()
  Call IDRSizeTheChartHxW(8, 8.5)
End Sub

Sub IDRSizeTheChartLarge()
  Call IDRSizeTheChartHxW(8, 18)
End Sub

Sub IDRSizeTheChartHxW(Optional MyHeight As Single, Optional MyWidth As Single)
' Written by Michel Gerday on 22.12.2010
' This sub
' - resizes the chart
' - set font and font size
' - uncheck autoscale and allow move but dont size with cells
' - set tickmarks inside
' - set plotarea and chartarea to white (automatic)
'Reference: http://peltiertech.com/Excel/ChartsHowTo/ResizeAndMoveAChart.html


  Dim oChart As ChartObject
  Dim MySerie As Series
  Dim MyRange As Range
  'Dim MyWidth As Single, MyHeight As Single
  Dim MyFont As String, MyFontSize As Single
  Dim iCount As Double
  
  If IsMissing(MyWidth) = True Then MyWidth = 8.5 'cm
  If IsMissing(MyHeight) = True Then MyHeight = 8 'cm
  
  MyFont = "Arial"
  MyFontSize = 8

  
  With ActiveSheet
    ' Define the chart
    On Error GoTo ChartErrorHandler
    Set oChart = ActiveChart.Parent
   'On Error GoTo 0
    
   
    ' Resize the chart
    With oChart
      .Width = Application.CentimetersToPoints(MyWidth)
      .Height = Application.CentimetersToPoints(MyHeight)
        
      .Placement = xlMove  'Move don't size with cells
      '.Placement = xlFreeFloating 'Prevents from resizing
      .RoundedCorners = False
      .Shadow = False
    End With
    
    ' CHART AREA
    With oChart.Chart.ChartArea
        .Border.Weight = xlHairline '1        'not nessary if linestyle =0
        .Border.LineStyle = xlNone '0         'no border
        '.Interior.ColorIndex = xlAutomatic
        .Format.Fill.Visible = msoFalse
        
        .AutoScaleFont = False 'Uncheck font autoscale
        .Font.Name = MyFont
        .Font.Size = MyFontSize 'Force font size to MyFontsize
    End With 'chartarea

   On Error Resume Next 'in case of no legend
    With oChart.Chart.Legend.Format.TextFrame2.TextRange.Font
        .NameComplexScript = MyFont
        .NameFarEast = MyFont
        .Name = MyFont
        .Size = MyFontSize
    End With 'textFrame.font
    
    
    With oChart.Chart.PlotArea
        .Border.Weight = xlHairline
        .Border.LineStyle = xlNone
        '.Interior.ColorIndex = xlAutomatic
        .Format.Fill.Visible = msoFalse
    End With 'plotarea
    
    
    
    ' TITLE bold
    On Error Resume Next 'in case of no title
    ' REM oChart.Chart.ChartTitle.Font.Bold = True     ' Cf B1 guidelines: not all the title is bold
    oChart.Chart.ChartTitle.Font.Size = MyFontSize
    
    ' Y1
    On Error Resume Next 'in case of no value axis
    With oChart.Chart.Axes(xlValue)
        .Border.Weight = xlHairline
        .Border.LineStyle = xlAutomatic
        .MajorTickMark = xlNone ' xlInside
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNextToAxis
        .TickLabels.NumberFormatLinked = 0
        .MajorGridlines.Delete
        .MinorGridlines.Delete
        .AxisTitle.Format.TextFrame2.TextRange.Font.Bold = msoFalse
    End With
    ' Y2 axis
    '     ActiveChart.Axes(xlValue).AxisTitle.Select
    '    Selection.Format.TextFrame2.TextRange.Font.Bold = msoFalse
'    ActiveChart.Axes(xlCategory).Select
'    Selection.MajorTickMark = xlNone
'    ActiveChart.Axes(xlCategory, xlSecondary).Select
'    Selection.MajorTickMark = xlNone

    On Error Resume Next 'in case of no secondary axis
    With oChart.Chart.Axes(xlValue, xlSecondary)
        .Border.Weight = xlHairline
        .Border.LineStyle = xlAutomatic
        .MajorTickMark = xlNone ' xlInside
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNextToAxis
        .TickLabels.NumberFormatLinked = 0
        .MajorGridlines.Delete
        .MinorGridlines.Delete
        .AxisTitle.Format.TextFrame2.TextRange.Font.Bold = msoFalse
    End With
    ' X1 axis
    On Error Resume Next
    With oChart.Chart.Axes(xlCategory)
        .Border.Weight = xlHairline
        .Border.LineStyle = xlAutomatic
        .MajorTickMark = xlNone ' xlInside
        .MinorTickMark = xlNone
        .TickLabelPosition = xlLow
        .TickLabels.NumberFormatLinked = 0
        .MajorGridlines.Delete
        .MinorGridlines.Delete
    End With
    On Error Resume Next 'in case of no secondary axis
    With oChart.Chart.Axes(xlCategory, xlSecondary)
        .Border.Weight = xlHairline
        .Border.LineStyle = xlAutomatic
        .MajorTickMark = xlNone ' xlInside
        .MinorTickMark = xlNone
        '.TickLabelPosition = xlNextToAxis
        .TickLabels.NumberFormatLinked = 0
        .MajorGridlines.Delete
        .MinorGridlines.Delete
        .AxisTitle.Format.TextFrame2.TextRange.Font.Bold = msoFalse
    End With
    
    For Each MySerie In oChart.Chart.SeriesCollection
      'With MySerie.Border
      If MySerie.Format.Line.Visible = msoTrue Then
            MySerie.Format.Line.Weight = 2 'B1=2.25 'REM would also apply to Marker line
            ' ATTEMPTS to get rid of the marker line, without success
            ' MySerie.Border.Weight = 2
            ' MySerie.Format.Border.Weight = 2.25
            'For iCount = 1 To MySerie.Points.Count
                'MySerie.Points(iCount).Format.Line.Visible = msoTrue
                'MySerie.Points(iCount).Format.Line.Weight = 3
                'MySerie.Points(5).Border.Weight = 9
                
            'Next iCount
      Else
        ' do nothing   'would add a line when not needed
      End If
      
      
    Next
    
  
  End With 'activesheet
Set MyRange = Nothing
Exit Sub

ChartErrorHandler:
  MsgBox "Error " & Err.Number & ", " & Err.Description
  If Err.Number = 91 Then
    MsgBox "Please select a chart, then try again", _
        vbOKOnly, "Select a Chart"
End If
Exit Sub
End Sub

Sub PowerpointSizeTheChart()
  Call PowerpointSizeTheChartHxW(11, 11)
End Sub
Sub PowerpointSizeTheChart11x11()
  Call PowerpointSizeTheChartHxW(11, 11)
End Sub
Sub PowerpointSizeTheChart11x24()
  Call PowerpointSizeTheChartHxW(11, 24)
End Sub
Sub PowerpointSizeTheChartHxW(Optional MyHeight As Single, Optional MyWidth As Single)
' Written by Michel Gerday on 22.12.2010
' This sub
' - resizes the chart
' - set font and font size
' - uncheck autoscale and allow move but dont size with cells
' - set tickmarks inside
' - set plotarea and chartarea to white (automatic)
'Reference: http://peltiertech.com/Excel/ChartsHowTo/ResizeAndMoveAChart.html


  Dim oChart As ChartObject
  Dim MySerie As Series
  Dim MyRange As Range
  ' Dim MyWidth As Single, MyHeight As Single
  Dim MyFont As String, MyFontSize As Single
  Dim iCount As Double
  
  If IsMissing(MyWidth) = True Then MyWidth = 11 'cm
  If IsMissing(MyHeight) = True Then MyHeight = 11 'cm
  
  MyFont = "Arial"
  MyFontSize = 10

  
  With ActiveSheet
    ' Define the chart
    On Error GoTo ChartErrorHandler
    Set oChart = ActiveChart.Parent
   'On Error GoTo 0
    
   
    ' Resize the chart
    With oChart
      .Width = Application.CentimetersToPoints(MyWidth)
      .Height = Application.CentimetersToPoints(MyHeight)
        
      .Placement = xlMove  'Move don't size with cells
      '.Placement = xlFreeFloating 'Prevents from resizing
      .RoundedCorners = False
      .Shadow = False
    End With
    
    ' CHART AREA
    With oChart.Chart.ChartArea
        .Border.Weight = xlHairline '1        'not nessary if linestyle =0
        .Border.LineStyle = xlNone '0         'no border
        '.Interior.ColorIndex = xlAutomatic
        .Format.Fill.Visible = msoFalse
        
        .AutoScaleFont = False 'Uncheck font autoscale
        .Font.Name = MyFont
        .Font.Size = MyFontSize 'Force font size to MyFontsize
    End With 'chartarea

   On Error Resume Next 'in case of no legend
    With oChart.Chart.Legend.Format.TextFrame2.TextRange.Font
        .NameComplexScript = MyFont
        .NameFarEast = MyFont
        .Name = MyFont
        .Size = MyFontSize
    End With 'textFrame.font
    
    
    With oChart.Chart.PlotArea
        .Border.Weight = xlHairline
        .Border.LineStyle = xlNone
        '.Interior.ColorIndex = xlAutomatic
        .Format.Fill.Visible = msoFalse
    End With 'plotarea
    
    
    
    ' TITLE bold
    On Error Resume Next 'in case of no title
    ' REM oChart.Chart.ChartTitle.Font.Bold = True     ' Cf B1 guidelines: not all the title is bold
    oChart.Chart.ChartTitle.Font.Size = MyFontSize
    
    ' Y1
    On Error Resume Next 'in case of no value axis
    With oChart.Chart.Axes(xlValue)
        .Border.Weight = xlHairline
        .Border.LineStyle = xlAutomatic
        .MajorTickMark = xlNone ' xlInside
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNextToAxis
        .TickLabels.NumberFormatLinked = 0
        .MajorGridlines.Delete
        .MinorGridlines.Delete
        .AxisTitle.Format.TextFrame2.TextRange.Font.Bold = msoFalse
    End With
    ' Y2 axis
    '     ActiveChart.Axes(xlValue).AxisTitle.Select
    '    Selection.Format.TextFrame2.TextRange.Font.Bold = msoFalse
'    ActiveChart.Axes(xlCategory).Select
'    Selection.MajorTickMark = xlNone
'    ActiveChart.Axes(xlCategory, xlSecondary).Select
'    Selection.MajorTickMark = xlNone

    On Error Resume Next 'in case of no secondary axis
    With oChart.Chart.Axes(xlValue, xlSecondary)
        .Border.Weight = xlHairline
        .Border.LineStyle = xlAutomatic
        .MajorTickMark = xlNone ' xlInside
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNextToAxis
        .TickLabels.NumberFormatLinked = 0
        .MajorGridlines.Delete
        .MinorGridlines.Delete
        .AxisTitle.Format.TextFrame2.TextRange.Font.Bold = msoFalse
    End With
    ' X1 axis
    On Error Resume Next
    With oChart.Chart.Axes(xlCategory)
        .Border.Weight = xlHairline
        .Border.LineStyle = xlAutomatic
        .MajorTickMark = xlNone ' xlInside
        .MinorTickMark = xlNone
        .TickLabelPosition = xlLow
        .TickLabels.NumberFormatLinked = 0
        .MajorGridlines.Delete
        .MinorGridlines.Delete
        .AxisTitle.Font.Size = MyFontSize - 1
        .AxisTitle.Font.Name = MyFont
        .AxisTitle.Font.Bold = False
        .AxisTitle.Font.Italic = True
        .AxisTitle.Left = 1           'Source position from left
        .AxisTitle.Top = 9999         'Position from top  999 : bottom
        
    End With
        

    
    
    On Error Resume Next 'in case of no secondary axis
    With oChart.Chart.Axes(xlCategory, xlSecondary)
        .Border.Weight = xlHairline
        .Border.LineStyle = xlAutomatic
        .MajorTickMark = xlNone ' xlInside
        .MinorTickMark = xlNone
        '.TickLabelPosition = xlNextToAxis
        .TickLabels.NumberFormatLinked = 0
        .MajorGridlines.Delete
        .MinorGridlines.Delete
        .AxisTitle.Format.TextFrame2.TextRange.Font.Bold = msoFalse
    End With
    
    For Each MySerie In oChart.Chart.SeriesCollection
      'With MySerie.Border
      If MySerie.Format.Line.Visible = msoTrue Then
            MySerie.Format.Line.Weight = 2 'B1=2.25 'REM would also apply to Marker line
            ' ATTEMPTS to get rid of the marker line, without success
            ' MySerie.Border.Weight = 2
            ' MySerie.Format.Border.Weight = 2.25
            'For iCount = 1 To MySerie.Points.Count
                'MySerie.Points(iCount).Format.Line.Visible = msoTrue
                'MySerie.Points(iCount).Format.Line.Weight = 3
                'MySerie.Points(5).Border.Weight = 9
                
            'Next iCount
      Else
        ' do nothing   'would add a line when not needed
      End If
      
      
    Next
    
  
  End With 'activesheet
Set MyRange = Nothing
Exit Sub

ChartErrorHandler:
  MsgBox "Error " & Err.Number & ", " & Err.Description
  If Err.Number = 91 Then
    MsgBox "Please select a chart, then try again", _
        vbOKOnly, "Select a Chart"
End If
Exit Sub
End Sub



