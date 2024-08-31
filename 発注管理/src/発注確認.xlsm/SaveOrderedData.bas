Attribute VB_Name = "SaveOrderedData"
'発注済みとしてチェックされた商品コードを保存する
Sub SaveOrderedProductsCode()
    
    Application.ScreenUpdating = False
    
    '発注確認シートへのアクセサ
    Dim load As New LoadSheetAccesser
    
    '発注済み商品コードへのアクセサ
    Dim ordered As New OrderedDataSheetAccesser
    ordered.InitNewWorkbook
    ordered.InitWorkSheet
    ordered.InitStatus load.bumonCode, load.targetDate
    
    'チェックされた商品コードを入力
    ordered.WriteProductsCode load.GetCheckedProductsCode
    
    '保存して閉じる
    ordered.Save
    ordered.CloseWorkBook
    
    Application.ScreenUpdating = True
End Sub
