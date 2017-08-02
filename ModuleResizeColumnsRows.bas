Attribute VB_Name = "ModuleResizeColumnsRows"
Option Explicit

Sub ResizeColumnWidthRowHeight()
Attribute ResizeColumnWidthRowHeight.VB_ProcData.VB_Invoke_Func = " \n14"

' See Change the column width and row height

' On a worksheet, you can specify a column width of 0 (zero) to 255.
' This value represents the number of characters that can be displayed in a cell
' that is formatted with the standard font. The default column width is 8.43 characters.
' If a column has a width of 0 (zero), the column is hidden.
'
' You can specify a row height of 0 (zero) to 409.
' This value represents the height measurement in points (1 point equals approximately 1/72 inch or 0.035 cm).
' The default row height is 12.75 points (approximately 1/6 inch or 0.4 cm).
' If a row has a height of 0 (zero), the row is hidden.

' https://support.office.com/en-us/article/Change-the-column-width-and-row-height-72f5e3cc-994d-43e8-ae58-9774a0905f46

' This sub is useful for clusters of B1 charts

    
Dim MyWidth As Double
Dim MyHeight As Double
Dim MyWidthPoints As Double

MyWidth = 8.5 'cm
MyHeight = 8 'cm
MyWidthPoints = 8.43 'points

    With Selection
      .ColumnWidth = MyWidthPoints
      .RowHeight = Application.CentimetersToPoints(MyHeight)
    End With
    
End Sub
