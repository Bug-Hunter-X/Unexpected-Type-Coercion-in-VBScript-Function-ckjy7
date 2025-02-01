Function f(a, b)
  If IsEmpty(a) Then Exit Function
  If IsEmpty(b) Then Exit Function
  If IsNumeric(a) And IsNumeric(b) Then
    f = a + b
  Else
    f = "Error: Non-numeric input"
  End If
End Function

' Example usage:
MsgBox f(10, 20)  ' Output: 30
MsgBox f(10, "abc") ' Output: Error: Non-numeric input
MsgBox f(Empty, 20) ' Output: (Empty)