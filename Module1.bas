Attribute VB_Name = "Module1"
Sub SelectAllSheets()
Attribute SelectAllSheets.VB_Description = "Select All Sheets"
Attribute SelectAllSheets.VB_ProcData.VB_Invoke_Func = "S\n14"
    Sheets.Select
End Sub
Public Sub CloseBook1()
    On Error Resume Next
    Application.StatusBar = "Trying to close Book1"
    Workbooks("Book1").Close SaveChanges:=False
    Application.StatusBar = "Trying to close Book1.XLS"
    Workbooks("Book1.XLS").Close SaveChanges:=False
    Application.StatusBar = "Trying to close Book1.XLSX"
    Workbooks("Book1.XLSX").Close SaveChanges:=False
    Application.StatusBar = "Trying to close Book1.XLSM"
    Workbooks("Book1.XLSM").Close SaveChanges:=False
    Application.StatusBar = False

End Sub

Sub Trim_Cells_Array_Method()

Dim arrData() As Variant
Dim arrReturnData() As Variant
Dim rng As Excel.Range
Dim lRows As Long
Dim lCols As Long
Dim I As Long, J As Long

  lRows = Selection.Rows.Count
  lCols = Selection.Columns.Count

  ReDim arrData(1 To lRows, 1 To lCols)
  ReDim arrReturnData(1 To lRows, 1 To lCols)

  Set rng = Selection
  arrData = rng.Value

  For J = 1 To lCols
    For I = 1 To lRows
      arrReturnData(I, J) = Trim(arrData(I, J))
    Next I
  Next J

  rng.Value = arrReturnData

  Set rng = Nothing
End Sub
