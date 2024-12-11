Function MyFunc(param1)
  Select Case TypeName(param1)
    Case "Empty"
      param1 = ""
    Case "Variant"
      param1 = ""
    Case Else
      'param1 is initialized with some value
  End Select
  ' ... rest of the function
End Function