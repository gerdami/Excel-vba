Attribute VB_Name = "ModuleFameRepo"
Option Explicit

Sub UnloadFameRepo()
Attribute UnloadFameRepo.VB_ProcData.VB_Invoke_Func = " \n14"
'
' UnloadFameRepo Macro
'

'
    AddIns("Fame_Repo").Installed = False
End Sub
Sub LoadFameRepo()
'
' UnloadFameRepo Macro
'

'
    AddIns("Fame_Repo").Installed = True
End Sub

