Attribute VB_Name = "Module11"
Sub Production()
'


'Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim m As Integer
Dim n As Integer
Dim plotName As String
Dim sHape As sHape
Dim BOOKCnt As Integer
Dim HOMEWKBK As String
Dim RNG As String
Dim sheetCnt As Integer
 

BOOKCnt = Application.Workbooks.Count

Workbooks(1).Activate
sheetCnt = Application.sheets.Count
Application.DisplayAlerts = False


For k = sheetCnt To 1 Step -1
    sheets(k).Activate
       If Not (Left(sheets(k).Name, 4) = "Prod") Then
        sheets(k).Delete
       End If
Next k

Application.DisplayAlerts = True

For i = 2 To BOOKCnt
    Workbooks(i).Activate
    sheets(sheetCnt).Activate
    ActiveSheet.Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Workbooks(1).Activate
    n = ActiveSheet.UsedRange.Rows.Count
   ' RNG = "A" & Str(n + 1)
    cells(n + 1, 1).Select
    ActiveSheet.Paste
    

Next i
'
End Sub
