Option Explicit

'---------------------------------------------------------------------------------------------------
' Creator: Jie Jenn
' Please ensure ActiveX Control Library reference is checked
'---------------------------------------------------------------------------------------------------
Sub QueryExcel(ByVal SQL_Statement As String)
On Error GoTo errHandle
Dim conn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim cmd As ADODB.Command
Dim sConnection As String, sSQL As String
Dim ws As Worksheet, i As Integer, iCheck As Integer
Dim num_records As Integer

sConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ActiveWorkbook.FullName & _
";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=0;ReadOnly=False"""

Set conn = CreateObject("ADODB.Connection")
Set rst = CreateObject("ADODB.Recordset")

sSQL = UCase(SQL_Statement)

conn.Open sConnection


'// Check if it is a SELECT query or an update action query
If Left(sSQL, 6) = "UPDATE" Or InStr(sSQL, "INSERT INTO") > 0 Then
    
    '// Action Queries
    
    conn.Execute sSQL, num_records
    MsgBox num_records & " records affected.", vbInformation
    
Else

    '// Retriving Recordset
    rst.Open sSQL, conn, adOpenDynamic, adLockOptimistic
    
    Set ws = Worksheets.Add
    With ws
        .Move ActiveWorkbook.Worksheets(Sheets.Count)
        
        
        .Range("A2").CopyFromRecordset rst
        
        '// Column Names
        For i = 0 To rst.Fields.Count - 1
            .Cells(1, i + 1) = rst.Fields(i).Name
        Next i
    End With
End If

Door:
If rst.State <> 0 Then rst.Close
If conn.State <> 0 Then conn.Close

Set cmd = Nothing
Set rst = Nothing
Set conn = Nothing
Set ws = Nothing
Exit Sub

errHandle:
MsgBox "Error: " & Err.Description, vbInformation, "JJ Excel SQL"
GoTo Door

End Sub

Sub ShowExcelSQLQuery()
frmExcelSQL.Show vbModeless
End Sub
