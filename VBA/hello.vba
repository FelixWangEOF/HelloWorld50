Sub MacroHelloWorld()
'
' MacroHelloWorld Macro
'
Dim Str As String
Str = "Hello World"
Debug.Print Str
MsgBox (Str)
Application.ActiveSheet.Range("A1") = Str
'
End Sub