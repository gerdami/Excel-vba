Attribute VB_Name = "ModuleUpdateServerURL"

Option Explicit

Sub UpdateAmecoServerURL()
' 06.12.2012
' Macro written by Michel Gerday (ECFIN.F4)
' Problem occurred when Ramiro changed the Ameco server location
'
Const OLDAMECOSERVER = "http://www.ecfin.cec/apps/ameco/Include/QueryPost.cfm"
Const NEWAMECOSERVER = "http://intragate.ec.europa.eu/ecfin/ameco/Include/QueryPost.cfm"

    Cells.Replace What:=OLDAMECOSERVER, Replacement:=NEWAMECOSERVER, _
                    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                    SearchFormat:=False, ReplaceFormat:=False
                    
                    
End Sub


