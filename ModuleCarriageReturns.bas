Attribute VB_Name = "ModuleCarriageReturns"
Option Explicit

Sub RemoveCarriageReturns()
'   Source: https://www.ablebits.com/office-addins-blog/2013/12/03/remove-carriage-returns-excel/#vba-macro-delete-carriage-returns

    Dim MyRange As Range
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
 
    For Each MyRange In ActiveSheet.UsedRange
        If 0 < InStr(MyRange, Chr(10)) Then
            MyRange = Replace(MyRange, Chr(10), " ")
        End If
    Next
 
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

