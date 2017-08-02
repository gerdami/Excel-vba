Attribute VB_Name = "ModuleSplitCode"
Option Explicit

Public Function SplitCode(ByVal MyString As String, ByVal Pos As Integer, Optional ByVal Delim As String = ".") As String
  Dim MySplit As Variant
  
  MySplit = Split(MyString, Delim)
  SplitCode = MySplit(Pos - 1) 'zero based

End Function
