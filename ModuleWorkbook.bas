Attribute VB_Name = "ModuleWorkbook"
Option Explicit

Sub CloseAllWorkbooksWithoutSave()
Dim WBs As Workbook
For Each WBs In Application.Workbooks
  If Not WBs.Name = ThisWorkbook.Name Then WBs.Close (False)
Next WBs
End Sub

