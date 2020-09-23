Attribute VB_Name = "Module2"

Public Function IncrSerial(ByVal prevStr As String) As String
    Dim x As Integer
    Dim y As Integer
    Dim ch As String
        
    y = Len(prevStr)

    For x = y To 1 Step -1

        ch = Mid$(prevStr, x, 1)
        
        Select Case ch
        Case "0" To "8"
            ch = Chr$(Asc(ch) + 1)
            Mid$(prevStr, x, 1) = ch
            Exit For
            
        Case "9"
            ch = "0"
            Mid$(prevStr, x, 1) = ch
            
        Case "A" To "Y"
            ch = Chr$(Asc(ch) + 1)
            Mid$(prevStr, x, 1) = ch
            Exit For
      
        Case "Z"
            ch = "A"
            Mid$(prevStr, x, 1) = ch
      
        Case "a" To "y"
            ch = Chr$(Asc(ch) + 1)
            Mid$(prevStr, x, 1) = ch
            Exit For
      
        Case "z"
            ch = "a"
            Mid$(prevStr, x, 1) = ch
        End Select
    Next x

    IncrSerial = prevStr
End Function
