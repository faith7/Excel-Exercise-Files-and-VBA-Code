Option Explicit

Dim conn As ADODB.Connection
Dim rst As ADODB.Recordset

Sub Connect_To_SQLServer(ByVal Server_Name As String, ByVal Database_Name As String, ByVal SQL_Statement As String)
Dim strConn As String
Dim wsReport As Worksheet
Dim col As Integer

strConn = "Provider=SQLOLEDB;"
strConn = strConn & "Server=" & Server_Name & ";"
strConn = strConn & "Database=" & Database_Name & ";"
strConn = strConn & "Trusted_Connection=yes;"

Set conn = New ADODB.Connection
With conn
        .Open ConnectionString:=strConn
        .CursorLocation = adUseClient
End With

Set rst = New ADODB.Recordset
With rst
        .ActiveConnection = conn
        .Open Source:=SQL_Statement
        
End With

Set wsReport = ThisWorkbook.Worksheets.Add
With wsReport
                
        For col = 0 To rst.Fields.Count - 1
                .Cells(1, col + 1).Value = rst.Fields(col).Name
        Next col
        
        .Range("A2").CopyFromRecordset rst
        
End With

Set wsReport = Nothing

Call Close_Connections

End Sub

Private Sub Close_Connections()

If rst.State <> 0 Then rst.Close
If conn.State <> 0 Then conn.Close

'// Release Memory
Set rst = Nothing
Set conn = Nothing

End Sub

Sub Run_Report()
Dim Server_Name As String

Server_Name = "<Your Server Name>"

Call Connect_To_SQLServer(Server_Name, "AdventureWorks2017", "SELECT * FROM HumanResources.vEmployee")
End Sub
