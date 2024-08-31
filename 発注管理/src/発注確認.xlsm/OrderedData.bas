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

'発注済みの商品のチェックボックスをオンにする
Sub LoadOrderedData()
        
    Application.ScreenUpdating = False
    
    '発注確認シートへのアクセサ
    Dim load As New LoadSheetAccesser
    
    '発注済み商品コードへのアクセサ
    Dim ordered As New OrderedDataSheetAccesser
    ordered.InitStatus load.bumonCode, load.targetDate
    ordered.InitOpenWorkBook
    ordered.InitWorkSheet
    
    '発注済みデータがない場合は終了
    If Dir(ordered.SaveFilePath) = "" Then
        Exit Sub
    End If
    
    '発注済みの商品コードのコレクション
    Dim orderedProductsCodes As Collection
    Set orderedProductsCodes = ordered.GetAllData_NoHead
    
    '発注済みの商品コードごと
    For Each orderedProductsCode In orderedProductsCodes
        load.OrderedIsTrue orderedProductsCode(1)
    Next
    
    ordered.CloseWorkBook
    
End Sub
