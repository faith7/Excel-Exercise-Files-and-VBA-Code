Option Explicit

Public Function Vlookup2(ByVal Lookup_Value As String, ByVal Cell_Range As Range, ByVal Column_Index As Integer) As Variant
On Error GoTo errHandle
Dim cell As Range
Dim Result_String As String

Result_String = ""

For Each cell In Cell_Range
    If cell.Value = Lookup_Value Then
        Result_String = Result_String & ", " & cell.Offset(0, Column_Index - 1).Value
    End If
Next cell

Vlookup2 = LTrim(Right(Result_String, Len(Result_String) - 1))
Exit Function

errHandle:
Vlookup2 = ""

End Function
