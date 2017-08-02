Attribute VB_Name = "ModuleInsertUpdateTextBox"
Option Explicit
Function NamedRangeExists(strRangeName As String) As Boolean
'   Local Variables
    Dim rngExists  As Range
    On Error Resume Next
    Set rngExists = Range(strRangeName)
    NamedRangeExists = True
    If rngExists Is Nothing Then NamedRangeExists = False
    On Error GoTo 0
End Function
Sub DeleteMyTextBox()
  DeleteTextBox ("TextUpdate")
End Sub
Sub DeleteTextBox(Optional MyTextBoxName As String)
  Dim oChart As ChartObject
  If IsMissing(MyTextBoxName) = True Then MyTextBoxName = "TextUpdate"
  With ActiveSheet
    ' Define the chart
    On Error GoTo ChartErrorHandler
    Set oChart = ActiveChart.Parent
    
    On Error Resume Next
    oChart.Chart.TextBoxes(MyTextBoxName).Delete
    On Error GoTo 0
  End With
Exit Sub
ChartErrorHandler:
  If Err.Number = 91 Then
      MsgBox "Please select a chart, then try again", _
        vbOKOnly, "Select a Chart"
    Else
      MsgBox "Error " & Err.Number & ", " & Err.Description
    End If
Exit Sub
End Sub
Sub InsertUpdatedTextBox()
' This sub inserts a text box within a chart, with
' - the text found in named range "ChartUpdate", or if not exists
' - the current date

  Dim oChart As ChartObject
  Dim MyFont As String, MyFontSize As Single
  Dim MySheetName As String, MyUpdateName As String, MyTextFormula As String
  MyUpdateName = "ChartUpdate" 'Should match a named range
  MyFont = "Arial Narrow"
  MyFontSize = 8

  With ActiveSheet
    ' Define the chart
    On Error GoTo ChartErrorHandler
    MySheetName = ActiveSheet.Name
    
    Set oChart = ActiveChart.Parent
    
    On Error Resume Next
    oChart.Chart.TextBoxes("TextUpdate").Delete
    On Error GoTo 0
    
    ' Insert a text box, left=0, top=999, width=999, height=MyFontSize + 1
    oChart.Chart.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 999, 999, MyFontSize + 1).Name = "TextUpdate"
    oChart.Chart.TextBoxes("TextUpdate").Select
    
    ' Box alignment and internal margins
    With Selection.ShapeRange.TextFrame2
      .VerticalAnchor = msoAnchorMiddle   'Vertical alignment center
      .MarginLeft = Application.CentimetersToPoints(0.1)
      .MarginBottom = Application.CentimetersToPoints(0.1)
      .MarginTop = Application.CentimetersToPoints(0.1)
      .MarginRight = Application.CentimetersToPoints(0.1)
    End With
    
    If NamedRangeExists(MyUpdateName) Then
      MyTextFormula = "='" & MySheetName & "'!" & MyUpdateName
      oChart.Chart.TextBoxes("TextUpdate").Formula = MyTextFormula
    Else
      Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = "Updated:" & Format(Date, "dd.mm.yyyy")
    End If
    
    ' Box text fontname, fontsize, not bold
    With Selection.ShapeRange.TextFrame2.TextRange.Font
      .Size = MyFontSize - 1
      .NameComplexScript = "Arial Narrow"
      .NameFarEast = "Arial Narrow"
      .Name = "Arial Narrow"
      .Bold = msoFalse
    End With
    
    'Reset color text, usually black  (othewise it takes the colour of the source cell)
    With Selection.ShapeRange.TextFrame2.TextRange.Font.Fill
      .Visible = msoTrue
      .ForeColor.ObjectThemeColor = msoThemeColorText1 ' 2nd top left colour
      '.ForeColor.RGB = RGB(0, 0, 0) 'black
      .ForeColor.TintAndShade = 0
      .ForeColor.Brightness = 0
      .Transparency = 0
      .Solid
    End With
 
    

   'On Error GoTo 0
  End With 'activesheet
  Call InsertChartName

  
Exit Sub

ChartErrorHandler:
  
  If Err.Number = 91 Then
      MsgBox "Please select a chart, then try again", _
        vbOKOnly, "Select a Chart"
    Else
      MsgBox "Error " & Err.Number & ", " & Err.Description
    End If
Exit Sub
End Sub
Sub InsertChartName()
' This sub inserts the ChartName with
' - the text found in named range "ChartUpdate", or
' - if not found, the worksheet name

  Dim oChart As ChartObject
  Dim MyFont As String, MyFontSize As Single
  Dim MySheetName As String, MyChartName As String, MyTextFormula As String
  MyChartName = "ChartName" 'Should match a named range
  
  With ActiveWorkbook.ActiveSheet
    ' Define the chart
    On Error GoTo ChartErrorHandler
    
    Set oChart = ActiveChart.Parent
    
    'oChart.Chart.TextBoxes("TextChart").Delete
    'On Error GoTo 0
    
    On Error Resume Next 'i.e. do not change ChartName
    If NamedRangeExists(MyChartName) Then
      oChart.Name = ActiveWorkbook.ActiveSheet.Range(MyChartName).Value
    Else
      oChart.Name = ActiveWorkbook.ActiveSheet.Name
    End If
    oChart.Select
    
    
   'On Error GoTo 0
  End With 'ActiveWorkbook.ActiveSheet
  
Exit Sub

ChartErrorHandler:
  
  If Err.Number = 91 Then
      MsgBox "Please select a chart, then try again", _
        vbOKOnly, "Select a Chart"
    Else
      MsgBox "Error " & Err.Number & ", " & Err.Description
    End If
Exit Sub
End Sub


