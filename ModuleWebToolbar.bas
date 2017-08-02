Attribute VB_Name = "ModuleWebToolbar"
Option Explicit

Public Sub EnableWebToolbar()
    Application.CommandBars("Web").Enabled = True
End Sub

Public Sub DisableWebToolbar()
    Application.CommandBars("Web").Enabled = False
End Sub

