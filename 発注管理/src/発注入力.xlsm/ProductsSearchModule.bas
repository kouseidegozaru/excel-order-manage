Attribute VB_Name = "ProductsSearchModule"
Sub Decide()
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Set targetWb = ThisWorkbook '発注入力.xlsm
    Set targetWs = targetWb.Sheets(OrderWb_SheetName)
    
    Dim selectedProductsCD As Collection
    Set selectedProductsCD = GetCheckedValue(SearchWb_ProductCodeColumnNumber, SearchWb_StateColumnNumber)
    
    Call WriteExcelData(targetWb, targetWs, selectedProductsCD)
    
    ThisWorkbook.Activate
    
End Sub

'検索フォームの選択された行の商品コードを取得
Public Function GetCheckedValue(columnNumber As Integer, stateColumnNumber As Integer) As Collection
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook '商品マスター
    Set ws = wb.Sheets(SearchWb_SheetName)
    
    Dim checkedIDs As New Collection
    Dim i As Long
    
    For i = 1 To ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).row
        If ws.Cells(i, stateColumnNumber).Value = True Then
            checkedIDs.Add ws.Cells(i, columnNumber).Value
        End If
    Next i
    
    Set GetCheckedValue = checkedIDs
End Function

'検索フォームの商品コードを発注入力に入力
Public Sub WriteExcelData(wb As Workbook, ws As Worksheet, selectedData As Collection)
    Dim lastRow As Long
    Dim i As Long
    Dim writeData As Collection
    
    Set writeData = FilterCollection(selectedData, GetProductsCD)
    
    ' ワークシートの最終行を取得
    lastRow = OrderWb_NextProductsRow
    
    ' Collectionの各要素をワークシートに追加
    For i = 1 To writeData.Count
        ws.Cells(lastRow, OrderWb_ProductCodeColumnNumber).Value = writeData(i)
        lastRow = lastRow + 1
    Next i
End Sub

'collection型の変数を比べ重複する値を除外
Function FilterCollection(baseCol As Collection, filterCol As Collection) As Collection
    Dim resultCol As New Collection
    Dim itemBase As Variant
    Dim itemFilter As Variant
    Dim exists As Boolean
    
    ' baseColの値をループして、filterColに存在しないものだけresultColに追加
    For Each itemBase In baseCol
        exists = False
        For Each itemFilter In filterCol
            If itemBase = itemFilter Then
                exists = True
                Exit For
            End If
        Next itemFilter
        If Not exists Then
            resultCol.Add itemBase
        End If
    Next itemBase
    
    ' 結果のコレクションを返す
    Set FilterCollection = resultCol
End Function
