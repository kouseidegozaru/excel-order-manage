Attribute VB_Name = "ButtonEvents"

'商品検索シートに関する処理
'確定ボタン(商品コードの反映)
Sub Decide()

    Dim order As New OrderSheetAccesser
    Dim Search As New SearchSheetAccesser
    Set Search.Workbook = ActiveWorkbook
    Search.InitWorkSheet
    
    '変更のイベントを無視
    IsIgnoreChangeEvents = True
    
    '重複する商品コードを排除
    Dim writeData As Collection
    Set writeData = FilterCollection(Search.GetCheckedProductsCode, _
                                     order.ProductsCode)
                                     
    Dim startRowIndex As Long
    Dim lastRowIndex As Long
    
    startRowIndex = order.DataNextRowNumber
    lastRowIndex = startRowIndex
    
    '発注入力に商品コード入力
    For i = 1 To writeData.Count
        order.Cells(lastRowIndex, order.ProductCodeColumnNumber) = writeData(i)
        lastRowIndex = lastRowIndex + 1
    Next i
    
    '発注入力に商品コードを入力した範囲
    Dim target As range
    Set target = order.Worksheet.range(order.ProductCodeColumn & startRowIndex & ":" & order.ProductCodeColumn & lastRowIndex)
    
    '商品情報表示
    DisplayProductsInfo target
    
    '保存
    SaveData
    
    IsIgnoreChangeEvents = False
    
    order.Workbook.Activate
    
End Sub

'発注入力シートに関する処理
'検索フォーム更新
Sub Update()

    Application.ScreenUpdating = False
    
    Dim order As New OrderSheetAccesser
    Dim Search As New SearchSheetAccesser
    
    Set Search.Workbook = order.Workbook
    Search.InitWorkSheet
    Search.Clear
    
    Dim DataBaseAccesser As New DataBaseAccesser
    Dim rs As ADODB.recordSet
    ' データベースからレコードセットを取得
    Set rs = DataBaseAccesser.GetAllProducts(order.BumonCode)
    
    Dim rowIndex As Long
    Dim columnIndex As Integer
    
    rowIndex = Search.DataStartRowNumber
    columnIndex = Search.DataStartColumnNumber
    
    ' データの書き込み
    rs.MoveFirst
    Do While Not rs.EOF
        
        For i = 0 To rs.Fields.Count - 1
            Search.Cells(rowIndex, i + columnIndex) = rs.Fields(i).value
        Next i
        
        ' チェックボックスの追加
        Search.AddCheckBox rowIndex
        
        rowIndex = rowIndex + 1
        rs.MoveNext
    Loop
    
    Search.Worksheet.Activate

    Application.ScreenUpdating = True
    
End Sub

'商品検索シートの表示
Sub Search()
    Dim order As New OrderSheetAccesser
    order.FormatSheet.Activate
End Sub
