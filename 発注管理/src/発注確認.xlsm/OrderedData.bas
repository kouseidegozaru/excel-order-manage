Attribute VB_Name = "OrderedData"
'発注済みとしてチェックされた商品コードを保存する
Sub SaveOrderedData()
    
    '画面の更新を無効化
    Application.ScreenUpdating = False
    
    '発注確認シートへのアクセサ
    Dim load As New LoadSheetAccesser
    
    '発注済み商品コードへのアクセサ
    Dim ordered As New OrderedDataSheetAccesser
    '部門コードと発注日を設定
    ordered.InitStatus load.bumonCode, load.targetDate
    ordered.InitNewWorkbook
    ordered.InitWorkSheet
    
    'チェックされた商品コードを入力
    ordered.WriteProductsCode load.GetCheckedProductsCode
    
    '保存して閉じる
    ordered.Save
    ordered.CloseWorkBook
    
    '画面の更新を有効化
    Application.ScreenUpdating = True
End Sub

'発注済みの商品のチェックボックスをオンにする
Sub LoadOrderedData()
    
    '画面の更新を無効化
    Application.ScreenUpdating = False
    
    '発注確認シートへのアクセサ
    Dim load As New LoadSheetAccesser
    
    '発注済み商品コードへのアクセサ
    Dim ordered As New OrderedDataSheetAccesser
    ordered.InitStatus load.bumonCode, load.targetDate
    
    '発注済みデータがない場合は終了
    If Dir(ordered.SaveFilePath) = "" Then
        Exit Sub
    End If
    
    ordered.InitOpenWorkBook
    ordered.InitWorkSheet
    
    '発注済みの商品コードをコレクションで取得
    Dim orderedProductsCodes As Collection
    '二次元のコレクションで取得
    Set orderedProductsCodes = ordered.GetAllData_NoHead
    
    '発注済みの商品コードごと
    For Each orderedProductsCode In orderedProductsCodes
        '対象の商品コードのチェックボックスをオンにする
        load.OrderedIsTrue orderedProductsCode(1) '二次元なのでインデックスを指定
    Next
    
    '発注済みデータを閉じる
    ordered.CloseWorkBook
    
    '画面の更新を有効化
    Application.ScreenUpdating = True
    
End Sub
