Option Explicit

Const FOLDER_PATH = "<Folder directory where your master Excel workbook is located>"
Const TARGET_FOLDER_PATH = "<Folder directory where your output files will be saved>"

Dim collection_GetList As Collection
Dim wsData As Worksheet
Dim LastRow As Long

Sub Main()
Dim wbMain As Workbook, wbTemp As Workbook
Dim itemNo As Long, column_IndividualReport As Long

Set wbMain = ThisWorkbook
Set collection_GetList = New Collection
Set wsData = wbMain.Worksheets("Boston Public School List")

column_IndividualReport = 6

Call Generate_List(collection_GetList, column_IndividualReport)

If collection_GetList.Count = 0 Then
        MsgBox "No data available.", vbInformation
        Exit Sub
End If

For itemNo = 1 To collection_GetList.Count

        With wsData
                
                Set wbTemp = Workbooks.Add
                
                .AutoFilterMode = False
                  
                '// Filtering data
               .Range("A1").CurrentRegion.AutoFilter column_IndividualReport, collection_GetList.Item(itemNo)
                        
                .Range("A1").CurrentRegion.Copy wbTemp.Worksheets(1).Range("A1")
                        
                wbTemp.SaveAs TARGET_FOLDER_PATH & wsData.Name & "_" & collection_GetList.Item(itemNo) & ".xlsx"
                wbTemp.Close False
                
                .AutoFilterMode = False
        
        End With
        
        Set wbTemp = Nothing

Next itemNo

MsgBox "All files are saved", vbInformation

End Sub

Private Sub Generate_List(ByVal collection_object As Collection, ByVal target_column As Long)
Dim RowNumber As Long

With wsData
LastRow = .Cells(Rows.Count, target_column).End(xlUp).Row
        
        If 2 > LastRow Then Exit Sub
        
        On Error Resume Next
        For RowNumber = 2 To LastRow
                collection_object.Add .Cells(RowNumber, target_column).Value, CStr(.Cells(RowNumber, target_column).Value)
        Next RowNumber

End With

End Sub















