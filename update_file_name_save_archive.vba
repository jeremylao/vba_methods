Sub update_filename_save_in_archive()

    Dim path_name, old_sheet_name, new_sheet_name, today_date, one_to_remove As String


    old_sheet_name = ActiveWorkbook.Name
    
    one_to_remove = ActiveWorkbook.Name
    
    old_sheet_name = "c:\test this path name\archive\" + old_sheet_name
    
    ActiveWorkbook.SaveAs Filename:=old_sheet_name
    
    path_name = "c:\test this path name\"
    
    today_date = Format(Date, "yyyy-mm-dd")
    
    fname = path_name + today_date + " " + "hello_world2"
    
    ActiveWorkbook.SaveAs Filename:=fname

    'Remove existing file
    Kill one_to_remove

End Sub


