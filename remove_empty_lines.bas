Sub remove_empty_lines()
    
    Dim flag As Integer
    Dim counter As Integer
    
    flag = 500
    counter = 0
    
    Range("A2").Activate
    
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
