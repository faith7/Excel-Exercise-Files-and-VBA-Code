Sub Test()
Dim s As String, sPattern As String
Dim regEx As RegExp
Dim Matches As IMatchCollection2
Dim i As Long

s = "hello Jiejenn@gmail.com, jie-jenn@learndataanalysis.org"
sPattern = "[\w-_]+@[\w-_.]+"

Set regEx = New RegExp
With regEx
    .Global = True
    .MultiLine = False
    .IgnoreCase = True
    .pattern = sPattern
    
    Set Matches = .Execute(s)
End With

If Matches.Count > 0 Then
    For i = 0 To Matches.Count - 1
        Debug.Print Matches.Item(i).Value
    
    Next i
End If

set regEx =Nothing
set Matches = Nothing
End Sub
