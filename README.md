# main
' Code_Update_2 Macro
'
Sub Code2()
 
Sheets("bulk").Cells.replace _
        What:=Sheets("data").Range("B2").Value, _
        Replacement:=Sheets("data").Range("B3").Value, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        MatchCase:=False
        
Sheets("bulk").Columns("U").replace _
        What:=Sheets("data").Range("D2").Value, _
        Replacement:=Sheets("data").Range("D3").Value, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        MatchCase:=False
        

End Sub

