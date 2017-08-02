Attribute VB_Name = "ModuleR1C1"
Option Explicit

Sub SwitchR1C1()
' http://excelribbon.tips.net/T009960_Getting_Rid_of_Numbered_Columns.html
'
    With Application
        If .ReferenceStyle = xlR1C1 Then
            .ReferenceStyle = xlA1
        Else
            .ReferenceStyle = xlR1C1
        End If
    End With
End Sub
