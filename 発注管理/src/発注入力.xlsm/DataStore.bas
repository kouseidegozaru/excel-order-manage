Attribute VB_Name = "DataStore"
'データの読み書き

'データの書き込み
Sub SaveData()
    
    Dim savePath As String
    
    ' 新しいワークブックを作成
    Dim newWorkbook As Workbook
    Set newWorkbook = Workbooks.Add
    
    
    ' 新しいワークブックの値を入力
    '発注入力のシート
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(OrderWb_SheetName)
    '商品コードデータ
    Dim productsData As Collection
    Set productsData = GetRangeValue(ws.range(OrderWb_InputProductsRange))
    writeData newWorkbook.Sheets(DataWb_SheetName), DataWb_ProductCodeRowNumber, DataWb_ProductCodeColumnNumber, productsData
    '数量データ
    Dim qtyData As Collection
    Set qtyData = GetRangeValue(ws.range(OrderWb_InpuQtyRange))
    writeData newWorkbook.Sheets(DataWb_SheetName), DataWb_ProductCodeRowNumber, DataWb_ProductQtyColumnNumber, qtyData
    
    
    ' 保存パスを指定（例：デスクトップに保存）
    savePath = GetSaveFilePath
    
    ' 上書き保存のために警告メッセージをオフにする
    Application.DisplayAlerts = False
    
    ' ワークブックを保存（既存のファイルがあれば上書き保存）
    newWorkbook.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    ' ワークブックを閉じる
    newWorkbook.Close
        
    ' 警告メッセージを再度有効にする
    Application.DisplayAlerts = True
    
End Sub

'シートに列を指定して入力
Sub writeData(ws As Worksheet, rowIndex As Long, colIndex As Integer, writeData As Collection)
    Dim item As Variant

        For Each item In writeData
            ws.Cells(rowIndex, colIndex).value = item
            rowIndex = rowIndex + 1
        Next item

    
End Sub

Sub LoadData()
    
    '発注入力の商品情報を全て削除
    Dim orderRng As range
    Set orderRng = ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_InputProductsRange)
    orderRng.EntireRow.Delete
    
    'ファイルが存在しない場合は処理終了
    If Dir(GetSaveFilePath) = "" Then
        End
    End If
    
    '発注入力のシート
    Dim wb As Workbook
    Set wb = DataWb
    
    Dim ws As Worksheet
    Set ws = wb.Sheets(DataWb_SheetName)
    '商品コードデータ
    Dim productsData As Collection
    Set productsData = GetRangeValue(ws.range(DataWb_ProductsRange))
    '数量データ
    Dim qtyData As Collection
    Set qtyData = GetRangeValue(ws.range(DataWb_QtyRange))
    
    'データ入力
    Dim OrderWs As Worksheet
    Set OrderWs = ThisWorkbook.Sheets(OrderWb_SheetName)
    
    writeData OrderWs, OrderWb_ProductCodeRowNumber, OrderWb_ProductCodeColumnNumber, productsData
    Dim target As range
    Set target = OrderWs.range(OrderWb_InputProductsRange)
    DisplayProductsInfo target
    writeData OrderWs, OrderWb_ProductCodeRowNumber, OrderWb_ProductQtyColumnNumber, qtyData
    
    wb.Close
End Sub

