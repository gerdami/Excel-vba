Attribute VB_Name = "addInSaveMe"
Option Explicit

' ----------------------------------------------------------------------------
' --  Procedure: SaveMeNow
' --  Parameters :none
' --  Return : none
' --  Description : Artifice to skirt a bug in the File -> Save menu bar
' --
' ----------------------------------------------------------------------------
Private Sub SaveMeNow()
    ThisWorkbook.Save
End Sub
