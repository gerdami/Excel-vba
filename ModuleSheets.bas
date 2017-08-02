Attribute VB_Name = "ModuleSheets"
Sub UnhideAllSheets()
Attribute UnhideAllSheets.VB_ProcData.VB_Invoke_Func = " \n14"
'
'   Inspired by http://excelribbon.tips.net/T009636_Unhiding_Multiple_Worksheets.html
'
    
    Dim wsSheet As Worksheet
    ActiveWindow.DisplayWorkbookTabs = True
    
    For Each wsSheet In ActiveWorkbook.Worksheets
        wsSheet.Visible = xlSheetVisible
    Next wsSheet
End Sub

Sub ViewAllRowColHeaders()
'
' ViewRowColHeaders Macro
'

    Dim wsSheet As Worksheet
    Dim wsCurrentWorksheetName As String
    
    ' Store current sheet name
    wsCurrentWorksheetName = ActiveSheet.Name
    
    Application.ScreenUpdating = False
    For Each wsSheet In ActiveWorkbook.Worksheets
      wsSheet.Activate
      ActiveWindow.DisplayHeadings = True
    Next wsSheet
    
    ' Restore current sheet name
    Sheets(wsCurrentWorksheetName).Select
    Application.ScreenUpdating = True
End Sub
