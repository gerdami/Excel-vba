Attribute VB_Name = "ModuleSpanishTestNumbers"
Option Explicit

Sub SpanishTextNumbersInSelectedRegion()
Attribute SpanishTextNumbersInSelectedRegion.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro10 Macro
'

'
    Selection.Replace What:=".", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=",", Replacement:=".", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
