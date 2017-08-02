Attribute VB_Name = "ModuleComments"


Public Function GetComment(rCell As Range) As String

    GetComment = rCell.Comment.Text
End Function

Function GetComments(rngTemp As Range)
If Not rngTemp.Comment Is Nothing Then
GetComments = rngTemp.Comment.Text
End If
End Function
Function MyComment(rng As Range)
    Application.Volatile
    Dim str As String
    str = Trim(rng.Comment.Text)
'// If you want to remove Chr(10) character from string, then
    str = Application.Substitute(str, vbLf, " ")
    MyComment = str
End Function
Sub RemoveCellComments()

  Cells.ClearComments

End Sub
