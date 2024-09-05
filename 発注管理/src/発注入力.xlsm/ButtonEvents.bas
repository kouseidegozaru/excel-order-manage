Attribute VB_Name = "ButtonEvents"

'商品検索シートに関する処理
'確定ボタン(商品コードの反映)
Sub Decide()

    '画面更新しない
    Application.ScreenUpdating = False

    '発注入力シートアクセサをインスタンス化
    Dim order As New OrderSheetAccesser
    '商品検索シートアクセサをインスタンス化
    Dim search As New SearchSheetAccesser
    
    '変更のイベントを無視
    IsIgnoreChangeEvents = True
    
    '重複する商品コードを排除
    Dim writeData As Collection
    Set writeData = FilterCollection(search.GetCheckedProductsCode, _
                                     order.productsCode)
                                     
    Dim startRowIndex As Long
    Dim lastRowIndex As Long
    
    startRowIndex = order.DataNextRowIndex
    lastRowIndex = startRowIndex
    
    '発注入力に商品コード入力
    For i = 1 To writeData.count
        order.Cells(lastRowIndex, order.ProductCodeColumnIndex) = writeData(i)
        lastRowIndex = lastRowIndex + 1
    Next i
    
    '発注入力に商品コードを入力した範囲
    Dim target As Range
    Set target = order.Worksheet.Range(IndexToLetter(order.ProductCodeColumnIndex) & startRowIndex & _
                                       ":" & _
                                       IndexToLetter(order.ProductCodeColumnIndex) & lastRowIndex)
                                       
    order.Worksheet.Activate
    
    '商品情報表示
    DisplayProductsInfo target
    
    '変更のイベントを有効化
    IsIgnoreChangeEvents = False
    
    '画面更新有効化
    Application.ScreenUpdating = True
    
End Sub

'商品検索シートに関する処理
'検索フォーム更新
Sub Update()

    '画面更新しない
    Application.ScreenUpdating = False
    
    '発注入力シートアクセサをインスタンス化
    Dim order As New OrderSheetAccesser
    '商品検索シートアクセサをインスタンス化
    Dim search As New SearchSheetAccesser
    
    '商品検索シートのデータをクリア
    search.Clear
    
    'クエリ発行クラスをインスタンス化
    Dim DataBaseAccesser As New DataBaseAccesser
    ' データベースから対象部門の商品情報を取得
    Dim rs As ADODB.Recordset
    Set rs = DataBaseAccesser.GetAllProducts(order.bumonCode)
    
    Dim rowIndex As Long
    Dim columnIndex As Integer
    
    rowIndex = search.DataStartRowIndex
    columnIndex = search.DataStartColumnIndex
    
    ' データの書き込み
    Dim targetRange As Range
    Set targetRange = search.Worksheet.Cells(rowIndex, columnIndex)
    
    ' レコードセットを一括で貼り付ける
    targetRange.CopyFromRecordset rs
    
    ' 貼り付けたデータの行数を取得
    Dim pastedRows As Long
    pastedRows = DataBaseAccesser.GetAllProductsCount(order.bumonCode)
    
    ' チェックボックスの追加
    For i = 0 To pastedRows - 1
        search.AddCheckBox rowIndex + i
    Next i
    
    search.Worksheet.Activate

    Application.ScreenUpdating = True
    
End Sub

'商品検索シートに関する処理
'クリアボタン
Sub ClearCheckBoxes()
    'チェックボックスのクリア
    Dim search As New SearchSheetAccesser
    search.ClearCheckBoxes
End Sub

'商品検索シートの表示
Sub search()
    Dim search As New SearchSheetAccesser
    '商品検索シートをアクティブ化
    search.Worksheet.Activate
End Sub

'送信ボタン
Sub Post()

    Dim result As VbMsgBoxResult
    '合わせ数と数量のチェック
    If Not IsMatchQty Then
        result = MsgBox("合わせ数と一致しない数量があります。送信しますか?", vbYesNo + vbQuestion, "確認")
        If result = vbNo Then
            End
        End If
    End If
    
    'データを保存
    SaveData
    
    MsgBox "データを送信しました"
End Sub
