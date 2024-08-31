Attribute VB_Name = "OrderedData"
'発注済みとしてチェックされた商品コードを保存する
Sub SaveOrderedData()
    
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
Sub LoadOrderedData()

End Sub
Sub test() 'テスト用
    '発注確認シートへのアクセサ
    Dim load As New LoadSheetAccesser
    '発注済み商品コードへのアクセサ
    Dim ordered As New OrderedDataSheetAccesser
    ordered.InitStatus load.bumonCode, load.targetDate
    ordered.InitOpenWorkBook
    ordered.InitWorkSheet
    
    
    Dim aa As Variant
    Set aa = ordered.GetAllData_NoHead
    
    ordered.CloseWorkBook
End Sub
