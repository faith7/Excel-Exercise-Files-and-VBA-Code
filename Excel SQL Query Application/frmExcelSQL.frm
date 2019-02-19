
Option Explicit

Private Sub cmdClear_Click()
    Me.txtSQL.Value = ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdRun_Click()
Dim sSQL As String

sSQL = Me.txtSQL.Value

Call QueryExcel(sSQL)

End Sub

Private Sub lstTables_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

With Me
    .txtSQL.Value = ""
    .txtSQL = "SELECT *" & vbNewLine & "FROM [" & .lstTables.Value & "$]"
End With

End Sub

Private Sub UserForm_Initialize()
Dim ws As Worksheet

With Me
    .Top = Application.Top
    .Left = Application.Left
    
    .lstTables.Clear
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            .lstTables.AddItem ws.Name
        End If
    Next ws
    
End With

End Sub
