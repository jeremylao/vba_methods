Dim strPattern As String: patter = "[Nn][^EeVv]*[Ee][^NnVv]*[Vv][^NnEe]*"
  -This Regex will extract a word with NnEeVv appearing in that order
  
  
...
        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
...
  
 
