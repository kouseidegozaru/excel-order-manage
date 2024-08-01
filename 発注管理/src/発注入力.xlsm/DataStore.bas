Attribute VB_Name = "DataStore"
'データの読み書き

'データの書き込み
Sub SaveData()

    Application.ScreenUpdating = False

    Dim order As New OrderSheetAccesser
    Dim data As New DataSheetAccesser
    data.NewWorkbook
    data.InitWorkSheet
    
    '商品コードデータ
    data.WriteProductsCode order.ProductsCode
    '数量データ
    data.WriteQty order.Qty
    
    '保存
    data.Save
    data.CloseWorkBook
    
    Application.ScreenUpdating = True
    
End Sub


Sub LoadData()
    
    Application.ScreenUpdating = False
    
    Dim order As New OrderSheetAccesser
    Dim data As New DataSheetAccesser
    
    '発注入力の商品情報を全て削除
    order.ProductsCodeRange.EntireRow.Delete
    
    'ファイルが存在しない場合は処理終了
    If Dir(data.SaveFilePath) = "" Then
        End
    End If
    
    data.OpenWorkBook
    data.InitWorkSheet
    
    '商品コードを入力
    order.WriteProductsCode data.ProductsCode
    '商品情報表示
    DisplayProductsInfo order.ProductsCodeRange
    '数量を入力
    order.WriteQty data.Qty
    
    'データワークブックを閉じる
    data.CloseWorkBook
    
    Application.ScreenUpdating = True
End Sub

