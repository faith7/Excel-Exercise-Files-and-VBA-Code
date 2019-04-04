Public Function Vlookup2(ByVal Lookup_Value As String, ByVal Cell_Range As Range, ByVal Column_Index As Integer) As Variant
Dim cell As Range
Dim Result_String As String

On Error GoTo errHandle

For Each cell In Cell_Range
    
    If cell.Value = Lookup_Value Then
    
        If cell.Offset(0, Column_Index - 1).Value <> "" Then
        
            If Not Result_String Like "*" & cell.Offset(0, Column_Index - 1).Value & "*" Then
                Result_String = Result_String & ", " & cell.Offset(0, Column_Index - 1).Value
            End If
            
        End If
        
    End If

Next cell

Vlookup2 = LTrim(Right(Result_String, Len(Result_String) - 1))

Exit Function

errHandle:
Vlookup2 = ""

End Function
