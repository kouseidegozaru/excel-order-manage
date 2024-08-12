Attribute VB_Name = "ButtonEvents"

'商品検索シートに関する処理
'確定ボタン(商品コードの反映)
Sub Decide()

    Dim order As New OrderSheetAccesser
    Dim search As New SearchSheetAccesser
    
    '変更のイベントを無視
    IsIgnoreChangeEvents = True
    
    '重複する商品コードを排除
    Dim writeData As Collection
    Set writeData = FilterCollection(search.GetCheckedProductsCode, _
                                     order.ProductsCode)
                                     
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
    
    '商品情報表示
    DisplayProductsInfo target
    
    IsIgnoreChangeEvents = False
    
    order.Worksheet.Activate
    
End Sub

'発注入力シートに関する処理
'検索フォーム更新
Sub Update()

    Application.ScreenUpdating = False
    
    Dim order As New OrderSheetAccesser
    Dim search As New SearchSheetAccesser
    
    search.Clear
    
    Dim DataBaseAccesser As New DataBaseAccesser
    Dim rs As ADODB.recordSet
    ' データベースからレコードセットを取得
    Set rs = DataBaseAccesser.GetAllProducts(order.bumonCode)
    
    Dim rowIndex As Long
    Dim columnIndex As Integer
    
    rowIndex = search.DataStartRowIndex
    columnIndex = search.DataStartColumnIndex
    
    ' データの書き込み
    rs.MoveFirst
    Do While Not rs.EOF
        
        For i = 0 To rs.Fields.count - 1
            search.Cells(rowIndex, i + columnIndex) = rs.Fields(i).value
        Next i
        
        ' チェックボックスの追加
        search.AddCheckBox rowIndex
        
        rowIndex = rowIndex + 1
        rs.MoveNext
    Loop
    
    search.Worksheet.Activate

    Application.ScreenUpdating = True
    
End Sub

'チェックボックスのクリア
Sub ClearCheckBoxes()
    Dim search As New SearchSheetAccesser
    search.ClearCheckBoxes
End Sub

'商品検索シートの表示
Sub search()
    Dim search As New SearchSheetAccesser
    search.Worksheet.Activate
End Sub

'送信
Sub Post()

    Dim result As VbMsgBoxResult
    '合わせ数と数量のチェック
    If Not IsMatchQty Then
        result = MsgBox("合わせ数と一致しない数量があります。送信しますか?", vbYesNo + vbQuestion, "確認")
        If result = vbNo Then
            End
        End If
    End If
        
    SaveData
    MsgBox "データを送信しました"
End Sub
