Attribute VB_Name = "Module1"
Sub LoadFileProperty()
    Dim load As New LoadSheetAccesser
    Dim data As New DataSheetAccesser
    Dim dataStorage As New DataBaseAccesser
    Dim fileProperty As New FilePropertyManager
    Dim filter As New fileFilter
    Dim users As New UserCodeAccesser
    
    Dim BumonCodeFilter As String
    Dim UserCodeFilter As String
    Dim TargetDateFilter As String
    
    'プロパティをクリア
    load.ClearFileProperty
    
    'ファイルの抽出条件文字列の設定
    BumonCodeFilter = fileProperty.BumonCodeIdentifier & load.BumonCode & fileProperty.BreakIdentifier
    TargetDateFilter = fileProperty.DateIdentifier & Format(load.TargetDate, "yyyymmdd") & fileProperty.BreakIdentifier
    
    '探索するディレクトリの設定
    filter.DirPath = data.DataDirPath
    
    Dim UserCodes As Collection
    Set UserCodes = users.GetEmployeeCodes(load.BumonCode)
    
    
    For Each UserCode In UserCodes
    
        '従業員名の取得
        Dim userName As String
        userName = dataStorage.GetUserName(UserCode)
        
        'ファイルの抽出条件文字列の設定
        UserCodeFilter = fileProperty.UserCodeIdentifier & UserCode & fileProperty.BreakIdentifier
        
        'フィルターの実行
        Dim filePathCollection As Collection
        Set filePathCollection = filter.AndFilter(BumonCodeFilter, TargetDateFilter, UserCodeFilter)
        
        'フィルターの結果が存在する場合
        If filePathCollection.Count > 0 Then
            fileProperty.filePath = data.DataDirPath & "\" & filePathCollection(1)
            load.AddFileProperty userName, True, fileProperty.UpdatedDate
        Else
            load.AddFileProperty userName, False, Date
        End If
        
        
    Next UserCode
    
End Sub

