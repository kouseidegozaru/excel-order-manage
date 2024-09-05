Attribute VB_Name = "LoadOrderData"
Sub Loads()
    Application.ScreenUpdating = False
    LoadFileProperty
    LoadData
    LoadOrderedData
    Application.ScreenUpdating = True
End Sub

Sub LoadFileProperty()
    Dim load As New LoadSheetAccesser
    Dim data As New DataSheetAccesser
    Dim DataStorage As New DataBaseAccesser
    Dim fileProperty As New FilePropertyManager
    Dim filter As New FileFilter
    Dim users As New UserCodeAccesser
    
    Dim bumonCodeFilter As String
    Dim userCodeFilter As String
    Dim targetDateFilter As String
    
    'プロパティをクリア
    load.ClearFileProperty
    
    'ファイルの抽出条件文字列の設定
    bumonCodeFilter = fileProperty.BumonCodeIdentifier & load.bumonCode & fileProperty.BreakIdentifier
    targetDateFilter = fileProperty.DateIdentifier & Format(load.targetDate, "yyyymmdd") & fileProperty.BreakIdentifier
    
    '探索するディレクトリの設定
    filter.DirPath = data.SaveDirPath
    
    '部門を指定して従業員コードの取得
    Dim userCodes As Collection
    Set userCodes = users.GetEmployeeCodes(load.bumonCode)
    
    
    For Each userCode In userCodes
    
        '従業員名の取得
        Dim userName As String
        userName = DataStorage.GetUserName(userCode)
        
        'ファイルの抽出条件文字列の設定
        userCodeFilter = fileProperty.UserCodeIdentifier & userCode & fileProperty.BreakIdentifier
        
        'フィルターの実行
        Dim filePathCollection As Collection
        Set filePathCollection = filter.AndFilter(bumonCodeFilter, targetDateFilter, userCodeFilter)
        
        'フィルターの結果が存在する場合
        If filePathCollection.Count > 0 Then
            'ファイル情報の取得準備
            fileProperty.InitFilePath data.SaveDirPath & "\" & filePathCollection(1)
            'ファイルプロパティの表示
            load.AddFileProperty userName, True, fileProperty.UpdatedDate
        Else
            'ファイルプロパティの表示
            load.AddFileProperty userName, False, Date
        End If
        
        
    Next userCode
    
End Sub

Sub LoadData()
    Dim load As New LoadSheetAccesser
    Dim data As New DataSheetAccesser
    Dim fileProperty As New FilePropertyManager
    Dim filter As New FileFilter
    
    load.ClearData
    
'''発注情報の取得'''
    Dim bumonCodeFilter As String
    Dim targetDateFilter As String
    
    'ファイルの抽出条件文字列の設定
    bumonCodeFilter = fileProperty.BumonCodeIdentifier & load.bumonCode & fileProperty.BreakIdentifier
    targetDateFilter = fileProperty.DateIdentifier & Format(load.targetDate, "yyyymmdd") & fileProperty.BreakIdentifier
    
    'フィルターの実行
    Dim filePathCollection As Collection
    filter.DirPath = data.SaveDirPath
    Set filePathCollection = filter.AndFilter(bumonCodeFilter, targetDateFilter)
    
    
    For Each fileName In filePathCollection
        
        'データの取得準備
        data.InitSaveFileName CStr(fileName)
        data.InitOpenWorkBook
        data.InitWorkSheet
        'ファイル情報の取得準備
        fileProperty.InitFilePath data.SaveFilePath
        '商品情報を入力
        load.WriteAllData data.GetAllData_NoHead
        'データワークブックを閉じる
        data.CloseWorkBook
        
    Next fileName
    
    'グループ化と集計をして書き込む
    Dim rs As ADODB.Recordset
    Set rs = load.AllGroupData
    load.ClearData
    load.WriteAllData RecordsetToCollection(rs)
    
    'チェックボックスを追加
    '開始行番号
    Dim rowIndex As Long
    rowIndex = load.DataStartRowIndex
    '終了行番号
    Dim endRowIndex As Long
    endRowIndex = load.DataEndRowIndex
    
    '受け取るチェックボックス
    Dim chkbox As Shape
    
    'チェックボックス
    Dim i As Long
    For i = rowIndex To endRowIndex
        'チェックボックスの追加
        Set chkbox = load.AddCheckBox(i)
        'チェックボックスにイベントの付与
        chkbox.OnAction = "SaveOrderedData"
    Next i
    
    '条件付き書式を設定
    load.ApplyConditionalFormatting
    
End Sub
