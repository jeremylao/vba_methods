Sub regex_finder()

    Dim regEx_1 As New RegExp
    Dim regEx_2 As New RegExp
    Dim case_one As String: case_one = "[Hh][^OoDd]*[Oo][^HhDd]*[Dd][^HhOo]*"
    Dim case_two As String: case_two = "[Mm]"
    Dim str_input As String
    Dim str_input_2 As String
    
    Dim name_table(900) As String
    Dim range_number As Integer
    Dim count As Integer
    Dim two_case_flag As Integer
    Dim user_output As Range
    
    Set user_output = Application.InputBox("Select Output Range", "Select Output Range", Type:=8)
    
    case_one = Application.InputBox("Input RegEx", "Input Expression", Type:=2)
    case_two = Application.InputBox("Input RegEx, #2", "Input Expression", Type:=2)
    
    two_case_flag = 1
    Range("G2").Activate  'The beginning of the range of cells where the data is contained
    str_pattern = ""
    range_number = 800
    count = 0
    
    For i = 0 To range_number
    
        str_input = ActiveCell.Value
        str_input_2 = ActiveCell.Offset(0, 1).Value
                
        With regEx_1
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = case_one
        End With
    
        If regEx_1.Test(str_input) And two_case_flag = 1 Then
            
            With regEx_2
                .Global = True
                .MultiLine = True
                .IgnoreCase = False
                .Pattern = case_two
            End With
        
            If regEx_2.Test(str_input_2) Then
            
                name_table(count) = str_input
                count = count + 1
                ActiveCell.Offset(1, 0).Activate
            
            Else
            
                ActiveCell.Offset(1, 0).Activate
            
            End If
            
        
        ElseIf regEx_1.Test(str_input) And two_case_flag = 0 Then
                                
                name_table(count) = str_input
                count = count + 1
                ActiveCell.Offset(1, 0).Activate
            
        Else
            
            ActiveCell.Offset(1, 0).Activate
            
        End If
        
    Next i
    
   user_output.Activate   'Location where you want to print out the results, will print horizontal
    
    For i = 0 To count
    
        ActiveCell.Value = name_table(i)
        ActiveCell.Offset(0, 1).Activate
    
    Next i

End Sub
