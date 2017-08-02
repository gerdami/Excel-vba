Attribute VB_Name = "ModuleStatusBar"
Option Explicit
Sub ResetStatusBar()
    Application.StatusBar = False
    Application.DisplayStatusBar = True
End Sub

Sub ToggleStatusBar()
    Application.DisplayStatusBar = Not Application.DisplayStatusBar
End Sub
