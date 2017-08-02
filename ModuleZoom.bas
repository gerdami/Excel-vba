Attribute VB_Name = "ModuleZoom"
Sub MyZoomAssignKeys()
  'Ron de Bruin, Disable key or key combination or run a macro if you use it
  'How do I use Application.Onkey
  'https://www.rondebruin.nl/win/s4/win012.htm
  'http://www.msofficeforums.com/excel-programming/14804-application-onkey-numeric-plus.html
  
  'Shift key = "+"  (plus sign)
  'Ctrl key = "^"   (caret)
  'Alt key = "%"    (percent sign)
  
  Application.StatusBar = "Assigning keys to MyZoom functions..."
  Application.OnKey "^{107}", "MyZoomIn"  'KeyPadPlus
  Application.OnKey "^{109}", "MyZoomOut" 'KeyPadMinus
  Application.OnKey "^{096}", "MyZoom100" 'KeyPad0
  Application.OnTime Now + TimeSerial(0, 0, 1), "MyZoomResetStatusbar"
  
End Sub
Sub MyZoomResetStatusbar()
  Application.StatusBar = False
End Sub
      
Sub MyZoomIn()
Attribute MyZoomIn.VB_Description = "Zoom in by 5%"
Attribute MyZoomIn.VB_ProcData.VB_Invoke_Func = " \n14"
   ' Zoom in by 5%
   ' see also http://excelribbon.tips.net/T012582_Zooming_With_the_Keyboard.html
   Dim ZP As Integer
   'ZP = Application.WorksheetFunction.MRound(ActiveWindow.Zoom * 1.1, 10)
   ZP = Application.WorksheetFunction.MRound(ActiveWindow.Zoom + 5, 5)
   If ZP > 400 Then ZP = 400
   ActiveWindow.Zoom = ZP
End Sub
Sub MyZoomOut()
Attribute MyZoomOut.VB_Description = "Zoom out by 5%"
Attribute MyZoomOut.VB_ProcData.VB_Invoke_Func = " \n14"
   ' Zoom out by 5%
   ' see also http://excelribbon.tips.net/T012582_Zooming_With_the_Keyboard.html
   Dim ZP As Integer
   'ZP = Application.WorksheetFunction.MRound(ActiveWindow.Zoom * 0.9, 10)
   ZP = Application.WorksheetFunction.MRound(ActiveWindow.Zoom - 5, 5)
   If ZP < 10 Then ZP = 10
   ActiveWindow.Zoom = ZP
End Sub
Sub MyZoom100()
   ' see also http://excelribbon.tips.net/T012582_Zooming_With_the_Keyboard.html
   Dim ZP As Integer
   'ZP = Application.WorksheetFunction.MRound(ActiveWindow.Zoom * 1.1, 10)
   ZP = 100
   ActiveWindow.Zoom = ZP
End Sub

