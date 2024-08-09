Attribute VB_Name = "DataStore"
'データの読み書き

'データの書き込み
Sub SaveData()

    Application.ScreenUpdating = False

    Dim order As New OrderSheetAccesser
    Dim Data As New DataSheetAccesser
    Data.NewWorkbook
    Data.InitWorkSheet
    
    '商品コードデータ
    Data.WriteProductsCode order.ProductsCode
    '数量データ
    Data.WriteQty order.Qty
    
    '保存
    Data.Save
    Data.CloseWorkBook
    
    Application.ScreenUpdating = True
    
End Sub


Sub LoadData()
    
    Application.ScreenUpdating = False
    
    Dim order As New OrderSheetAccesser
    Dim Data As New DataSheetAccesser
    
    '発注入力の商品情報を全て削除
    order.ProductsCodeRange.EntireRow.Delete
    
    'ファイルが存在しない場合は処理終了
    If Dir(Data.SaveFilePath) = "" Then
        End
    End If
    
    Data.OpenWorkBook
    Data.InitWorkSheet
    
    '商品コードを入力
    order.WriteProductsCode Data.ProductsCode
    '商品情報表示
    DisplayProductsInfo order.ProductsCodeRange
    '数量を入力
    order.WriteQty Data.Qty
    
    'データワークブックを閉じる
    Data.CloseWorkBook
    
    Application.ScreenUpdating = True
End Sub

