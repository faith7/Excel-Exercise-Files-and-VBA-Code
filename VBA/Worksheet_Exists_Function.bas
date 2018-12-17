Public Function Worksheet_Exists(ByVal Worksheet_Name As String) As Boolean
Dim x As String

On Error GoTo Sheet_Missing
Worksheet_Exists = True

x = ThisWorkbook.Worksheets(Worksheet_Name).Name

Exit Function

Sheet_Missing:
Worksheet_Exists = False
End Function
