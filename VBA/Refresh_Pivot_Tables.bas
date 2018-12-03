Option Explicit

Sub Refresh_PivotTable_CurrentSheet()
Dim PTIndex As Long, PTCount As Long
Dim ws As Worksheet

Set ws = ThisWorkbook.ActiveSheet
With ws
        
        PTCount = .PivotTables.Count
        
        If PTCount = 0 Then Exit Sub
        
        For PTIndex = 1 To PTCount
                .PivotTables(PTIndex).PivotCache.Refresh
        Next PTIndex
        
End With

End Sub

Sub Refresh_PivotTables_EntireWorkbook()
Dim PTIndex As Long, PTCount As Long
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
        
        With ws
                
                PTCount = .PivotTables.Count
                
                If PTCount > 0 Then
                
                        For PTIndex = 1 To PTCount
                                .PivotTables(PTIndex).PivotCache.Refresh
                        Next PTIndex

                End If
        
        End With

Next ws

End Sub
