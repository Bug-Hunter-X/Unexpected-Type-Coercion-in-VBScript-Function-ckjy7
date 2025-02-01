Function f(a, b)
  If IsEmpty(a) Or IsEmpty(b) Then
    f = "Error: Empty input"
    Exit Function
  End If
  If IsNumeric(a) And IsNumeric(b) Then
    f = a + b
  Else
    f = "Error: Non-numeric input"
  End If
End Function

' Example usage:
MsgBox f(10, 20)  ' Output: 30
MsgBox f(10, "abc") ' Output: Error: Non-numeric input
MsgBox f(Empty, 20) ' Output: Error: Empty input