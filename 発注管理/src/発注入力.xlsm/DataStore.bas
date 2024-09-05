Attribute VB_Name = "DataStore"
'発注データの保存
Sub SaveData()

    '画面更新しない
    Application.ScreenUpdating = False

    '発注入力シートアクセサをインスタンス化
    Dim order As New OrderSheetAccesser
    
    'データシートアクセサをインスタンス化
    Dim data As New DataSheetAccesser
    
    '部門コード、担当者コード、発注日を設定
    data.InitStatus order.bumonCode, _
                    order.userCode, _
                    order.targetDate
    data.InitNewWorkbook
    data.InitWorkSheet
    
    '商品データ書き込み
    data.WriteTableData order.GetAllData
    
    '保存
    data.Save
    data.CloseWorkBook
    
    '画面更新有効化
    Application.ScreenUpdating = True
    
End Sub

'データ読み込み
Sub LoadData()
    
    '画面更新しない
    Application.ScreenUpdating = False
    
    '発注入力シートアクセサをインスタンス化
    Dim order As New OrderSheetAccesser
    'データシートアクセサをインスタンス化
    Dim data As New DataSheetAccesser
    
    '発注入力の商品情報を全て削除
    order.ProductsCodeRange.EntireRow.Delete
    
    '部門コード、担当者コード、発注日を設定
    data.InitStatus order.bumonCode, _
                    order.userCode, _
                    order.targetDate
        
    'ファイルが存在しない場合は処理終了
    If Dir(data.SaveFilePath) = "" Then
        End
    End If
    
    data.InitOpenWorkBook
    data.InitWorkSheet
    
    '商品情報を入力
    order.WriteAllData data.GetAllData_NoHead
    
    'データワークブックを閉じる
    data.CloseWorkBook
    
    '仕入れ金額計算式の入力
    ApplyAmountCalcFormulaToRange
    
    '画面更新有効化
    Application.ScreenUpdating = True
    
End Sub

