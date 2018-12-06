Public Function CONCATEIF(ByVal Lookup_Value As Variant, ByVal Column_Index_Number As Long, ByVal Allow_Duplicate As Boolean, ParamArray Cell_Range() As Variant) As String
Dim i As Long, arrBound As Long, arrIndex As Long
Dim rng As Range, RowRange As Range
Dim sOutput As String
Dim stringSplit As Variant
Dim collection_Unique_Strings As Collection

sOutput = ""

On Error GoTo invalid_input
For i = LBound(Cell_Range()) To UBound(Cell_Range())

    Set rng = Cell_Range(i)

    For Each RowRange In rng.Rows
        If RowRange.Cells(1, 1).Value = Lookup_Value Then
            sOutput = sOutput & ", " & RowRange.Cells(1, Column_Index_Number).Value
        End If
    Next RowRange

    Set rng = Nothing
    
Next i

If Allow_Duplicate = False Then
    Set collection_Unique_Strings = New Collection
    
    stringSplit = Split(sOutput, ",")
    
    On Error Resume Next
    For arrBound = LBound(stringSplit) To UBound(stringSplit)
        
        '/ Ensure it is not an empty string
        If Len(LTrim(stringSplit(arrBound))) > 0 Then
            collection_Unique_Strings.Add stringSplit(arrBound), CStr(stringSplit(arrBound))
        End If
        
    Next arrBound
    
    sOutput = ""
    'concate each item with collection
    For arrIndex = 1 To collection_Unique_Strings.Count
        sOutput = sOutput & ", " & collection_Unique_Strings.Item(arrIndex)
    Next arrIndex
    
End If

'Remove extra comma
CONCATEIF = LTrim(Right(sOutput, Len(sOutput) - 1))

Set collection_Unique_Strings = Nothing
Exit Function

invalid_input:
Set collection_Unique_Strings = Nothing
CONCATEIF = "#N/A"
End Function
