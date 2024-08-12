Attribute VB_Name = "DataStore"
'データの読み書き

'データの書き込み
Sub SaveData()

    Application.ScreenUpdating = False

    Dim order As New OrderSheetAccesser
    
    Dim data As New DataSheetAccesser
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
    
    data.InitStatus order.bumonCode, _
                    order.userCode, _
                    order.targetDate
    data.InitOpenWorkBook
    data.InitWorkSheet
    
    '商品情報を入力
    order.WriteAllData data.GetAllData_NoHead
    
    'データワークブックを閉じる
    data.CloseWorkBook
    
    Application.ScreenUpdating = True
End Sub

