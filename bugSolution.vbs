Function f(a,b)
  'Explicitly convert inputs to numbers for consistent comparison
  Dim numA, numB
  On Error Resume Next
  numA = CDbl(a)
  numB = CDbl(b)
  On Error GoTo 0
  
  If IsNumeric(numA) And IsNumeric(numB) Then
    If numA > numB Then
      MsgBox "a is greater than b"
    ElseIf numA < numB Then
      MsgBox "a is less than b"
    Else
      MsgBox "a is equal to b"
    End If
  Else
    MsgBox "Invalid input. Please enter numbers only."
  End If
end function