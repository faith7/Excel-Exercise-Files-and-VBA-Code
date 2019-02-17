Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

If Target.Column = 1 And Target.Value <> Chr(214) Then
    
    With Target
        .Value = Chr(214)
        .Font.Name = "Symbol"
        .Font.FontStyle = "Regular"
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With

ElseIf Target.Column = 1 And Target.Value = Chr(214) Then
    
    Target.ClearContents

End If

End Sub
