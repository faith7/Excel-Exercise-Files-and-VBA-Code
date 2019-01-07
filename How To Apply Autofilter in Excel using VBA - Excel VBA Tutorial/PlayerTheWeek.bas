Sub ApplyAutoFilter()
Dim wsTarget As Worksheet
Dim LastRow As Long

Set wsTarget = ThisWorkbook.Worksheets("nba player of the week")
With wsTarget
        
        .AutoFilterMode = False
        
        LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
        
        If 2 > LastRow Then Exit Sub
        
        '// Apply Autofilter
        
        '// Filter by season
        .Range("A1:L" & LastRow).AutoFilter 1, "2015-2016"
        
        '// AND Clause
        .Range("A1:L" & LastRow).AutoFilter Range("I1").Column, ">=30"
        
End With

'// Release Memory
Set wsTarget = Nothing

End Sub
