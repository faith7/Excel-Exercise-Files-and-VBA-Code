Sub Change_Tab_Color()
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
        
        Select Case ws.Name
        
                Case "North"
                        ws.Tab.Color = RGB(255, 0, 0)
                        
                Case "West"
                        ws.Tab.Color = RGB(255, 255, 0)
                        
                Case "East"
                        ws.Tab.Color = RGB(0, 0, 255)
                        
                Case "South"
                        ws.Tab.Color = RGB(122, 55, 40)
                
        End Select
        
Next ws

End Sub
