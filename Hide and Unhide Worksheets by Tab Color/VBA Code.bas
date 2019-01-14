Sub Hide_Worksheets_By_Color()
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
        If ws.Tab.Color = 3506772 Or ws.Tab.Color = 7884319 Then
                ws.Visible = xlSheetHidden
        End If
Next ws

End Sub

Sub Get_Tab_Color()

Debug.Print ThisWorkbook.Worksheets("Sheet1").Tab.Color
Debug.Print ThisWorkbook.Worksheets("Sheet2").Tab.Color
Debug.Print ThisWorkbook.Worksheets("Sheet4").Tab.Color
