Attribute VB_Name = "LoadOrderData"
Sub Loads()
    Application.ScreenUpdating = False
    LoadFileProperty
    LoadData
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
    bumonCodeFilter = fileProperty.BumonCodeIdentifier & load.BumonCode & fileProperty.BreakIdentifier
    targetDateFilter = fileProperty.DateIdentifier & Format(load.TargetDate, "yyyymmdd") & fileProperty.BreakIdentifier
    
    '探索するディレクトリの設定
    filter.DirPath = data.SaveDirPath
    
    '部門を指定して従業員コードの取得
    Dim userCodes As Collection
    Set userCodes = users.GetEmployeeCodes(load.BumonCode)
    
    
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
    bumonCodeFilter = fileProperty.BumonCodeIdentifier & load.BumonCode & fileProperty.BreakIdentifier
    targetDateFilter = fileProperty.DateIdentifier & Format(load.TargetDate, "yyyymmdd") & fileProperty.BreakIdentifier
    
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
    
    Dim rs As ADODB.Recordset
    Set rs = load.AllGroupData
    load.ClearData
    load.WriteAllData RecordsetToCollection(rs)
    
End Sub
