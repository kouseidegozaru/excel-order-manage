Attribute VB_Name = "ProductsSearchModule"
Sub Decide()
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Set targetWb = ThisWorkbook '発注入力.xlsm
    Set targetWs = targetWb.Sheets(OrderWb_SheetName)
    
    Call WriteExcelData(targetWb, targetWs, GetCheckedValue(SearchWb_ProductCodeColumnNumber, SearchWb_StateColumnNumber))
End Sub

'検索フォームの選択された行の商品コードを取得
Public Function GetCheckedValue(ColumnNumber As Integer, stateColumnNumber As Integer) As Collection
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook '商品マスター
    Set ws = wb.Sheets(SearchWb_SheetName)
    
    Dim checkedIDs As New Collection
    Dim i As Long
    
    For i = 1 To ws.Cells(ws.Rows.Count, ColumnNumber).End(xlUp).row
        If ws.Cells(i, stateColumnNumber).Value = True Then
            checkedIDs.Add ws.Cells(i, ColumnNumber).Value
        End If
    Next i
    
    Set GetCheckedValue = checkedIDs
End Function

'検索フォームの商品コードを発注入力に入力
Public Sub WriteExcelData(wb As Workbook, ws As Worksheet, writeData As Collection)
    Dim lastRow As Long
    Dim i As Long
    
    ' ワークシートの最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, OrderWb_ProductCodeColumnNumber).End(xlUp).row + 1
    
    ' Collectionの各要素をワークシートに追加
    For i = 1 To writeData.Count
        ws.Cells(lastRow, OrderWb_ProductCodeColumnNumber).Value = writeData(i)
        lastRow = lastRow + 1
    Next i
End Sub
