Function MyFunction(param)
  If Len(param) = 0 Then
    Err.Raise 9999, , "Parameter cannot be empty!" 'Custom error
  End If
  ' ... function logic ...
End Function

On Error GoTo ErrorHandler
Dim result: result = MyFunction("")
Exit Sub

ErrorHandler:
  If Err.Number = 9999 Then
    MsgBox Err.Description
  Else
    MsgBox "An unexpected error occurred: " & Err.Number & " - " & Err.Description
  End If