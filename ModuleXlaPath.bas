Attribute VB_Name = "ModuleXlaPath"
Option Explicit

Sub ResetPopulatorPath()
  ActiveWorkbook.ChangeLink Name:="Fame_Repo.xlam" _
    , NewName:="C:\Program Files (x86)\Microsoft Office\Office14\Library\populator.xlam" _
    , Type:=xlExcelLinks
End Sub
Sub RemoveXlaPath()
Attribute RemoveXlaPath.VB_Description = "Remove XLA path in current sheet"
Attribute RemoveXlaPath.VB_ProcData.VB_Invoke_Func = " \n14"
' 04.12.2012
' Macro written by Michel Gerday (ECFIN.F4)
' Problem occurred when exchanging workbooks between users
' that do not have the FamePopulator installation
'
' Goal: delete the path reference to the add-in, i.e. everything before and including the '!'
' ='C:\Program Files (x86)\Microsoft Office\Office14\LIBRARY\populator.xlam'!famedata(...)
' ='C:\Users\castrfr\AppData\Local\famerepo\xla\Fame_Repo.xlam'!famedata(...)
'
'   06.12.2012: Added dot before xla in 'C:\*.xla*'!
'   20.02.2014: Added network paths \\net1....
'   18.06.2014: browse each visible worksheet
'
Dim MySheet As Object
  For Each MySheet In Worksheets
    Debug.Print "Worksheet " & MySheet.Name & " , visible = " & MySheet.Visible
    If MySheet.Visible = True Then
      Application.StatusBar = "Find & Replace in sheet " & MySheet.Name
    
      MySheet.Cells.Replace What:="'*:\*.xla*'!", _
        Replacement:="", _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        MatchCase:=False, _
        SearchFormat:=False, _
        ReplaceFormat:=False
      
      ' Also for \\net1.cec.eu.int\ECFIN\Users\hildhal\AppData\Local\famerepo\xla\Fame_Repo.xlam
      MySheet.Cells.Replace What:="\\*xla*!", _
        Replacement:="", _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        MatchCase:=False, _
        SearchFormat:=False, _
        ReplaceFormat:=False
      
      End If
  Next MySheet
  Application.StatusBar = False
End Sub


Sub RemoveXlaPathOld()
' 04.12.2012
' Macro written by Michel Gerday (ECFIN.F4)
' Problem occurred when exchanging workbooks between users
' that do not have the FamePopulator installation
'
' Goal: delete the path reference to the add-in, i.e. everything before and including the '!'
' ='C:\Program Files (x86)\Microsoft Office\Office14\LIBRARY\populator.xlam'!famedata(...)
'
'   06.12.2012: Added dot before xla in 'C:\*.xla*'!
'
    Cells.Replace What:="'C:\*.xla*'!", _
      Replacement:="", _
      LookAt:=xlPart, _
      SearchOrder:=xlByRows, _
      MatchCase:=False, _
      SearchFormat:=False, _
      ReplaceFormat:=False
End Sub

