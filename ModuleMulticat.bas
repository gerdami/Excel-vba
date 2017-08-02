Attribute VB_Name = "ModuleMulticat"
Option Explicit

Public Function Multicat( _
     ByRef rRng As Excel.Range, _
     Optional ByVal sDelim As String = "") _
     As String
' Purpose: Concatenate a range of cells with a specified delimiter
' Usage:  D1 = MultiCat(A1:C1," ")
' http://www.mcgimpsey.com/excel/udfs/multicat.html
' MG: added test on empty cell
' MG: added Trim to remove leading and trailing spaces
Dim rCell As Range
  For Each rCell In rRng
    If IsEmpty(rCell) Then
      'do nothing
    Else
      Multicat = Multicat & sDelim & Trim(rCell.Text)
    End If
  Next rCell
  ' Remove first delimiter
  Multicat = Mid(Multicat, Len(sDelim) + 1)
End Function

