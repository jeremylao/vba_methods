Sub remove_empty_lines()
    
    Dim flag As Integer
    Dim counter As Integer
    Dim user_input As Range
    
    Set user_input = Application.InputBox("Beginning of Input Range", "Beginning of Input Range", Type:=8)
    
    flag = 500
    counter = 0
    
    user_input.Activate
    
    While counter < flag
    
        If ActiveCell.Value = "" Then
        
            ActiveCell.EntireRow.Delete
            counter = counter + 1
            
            
        Else
            ActiveCell.Offset(1, 0).Select
            counter = 0
            
        End If
               
    Wend
    
    
End Sub
