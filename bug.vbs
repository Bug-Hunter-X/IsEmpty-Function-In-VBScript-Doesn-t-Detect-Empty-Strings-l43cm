Function MyFunction(param)
  If IsEmpty(param) Then
    Err.Raise 9999, , "Parameter cannot be empty!" 'Custom error
  End If
  ' ... function logic ...
End Function

' Calling the function with an empty string will not raise the custom error
Dim result: result = MyFunction("")