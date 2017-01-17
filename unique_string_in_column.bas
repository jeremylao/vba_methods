Sub uniqueness()

    Dim name_table(100) As String
    Dim count As Integer
    Dim next_cell As Range
    Dim temp As String
    Dim flag As Integer
    
    
    count = 0
    flag = 0
    
    Range("B5").Select
    
    Set next_cell = ActiveCell.Offset(1, 0)
    
    While Not IsEmpty(next_cell)
        
        temp = ActiveCell.Value
               
        For i = 0 To count
            
            If temp = name_table(i) Then
            
                flag = 1
                
            End If
            
        Next i
        
        If flag = 0 Then
        
            name_table(count) = temp
            ActiveCell.Offset(1, 0).Select
            Set next_cell = ActiveCell.Offset(1, 0)
            count = count + 1
            flag = 0
        
        Else
        
            ActiveCell.Offset(1, 0).Select
            Set next_cell = ActiveCell.Offset(1, 0)
            flag = 0
            
        End If
                         
    Wend
    
    Range("D5").Select
    
    'ActiveCell.Resize(6, 1) = name_table
    
    For i = 0 To count
       ActiveCell.Value = name_table(i)
       ActiveCell.Offset(1, 0).Activate
    Next i
      

End Sub
                        
                        

