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
    
    '部門を指定して従業員コードの取得
    Dim UserCodes As Collection
    Set UserCodes = users.GetEmployeeCodes(load.BumonCode)
    
    
    For Each UserCode In UserCodes
    
        '従業員名の取得
        Dim userName As String
        userName = DataStorage.GetUserName(UserCode)
        
        'ファイルの抽出条件文字列の設定
        UserCodeFilter = fileProperty.UserCodeIdentifier & UserCode & fileProperty.BreakIdentifier
        
        'フィルターの実行
        Dim filePathCollection As Collection
        Set filePathCollection = filter.AndFilter(BumonCodeFilter, TargetDateFilter, UserCodeFilter)
        
        'フィルターの結果が存在する場合
        If filePathCollection.Count > 0 Then
            'ファイル情報の取得準備
            fileProperty.FilePath = data.DataDirPath & "\" & filePathCollection(1)
            'ファイルプロパティの表示
            load.AddFileProperty userName, True, fileProperty.UpdatedDate
        Else
            'ファイルプロパティの表示
            load.AddFileProperty userName, False, Date
        End If
        
        
    Next UserCode
    
End Sub

Sub LoadData()
    Dim load As New LoadSheetAccesser
    Dim data As New DataSheetAccesser
    Dim fileProperty As New FilePropertyManager
    Dim filter As New FileFilter
    
    load.ClearData
    
'''発注情報の取得'''
    Dim BumonCodeFilter As String
    Dim TargetDateFilter As String
    
    'ファイルの抽出条件文字列の設定
    BumonCodeFilter = fileProperty.BumonCodeIdentifier & load.BumonCode & fileProperty.BreakIdentifier
    TargetDateFilter = fileProperty.DateIdentifier & Format(load.TargetDate, "yyyymmdd") & fileProperty.BreakIdentifier
    
    'フィルターの実行
    Dim filePathCollection As Collection
    filter.DirPath = data.DataDirPath
    Set filePathCollection = filter.AndFilter(BumonCodeFilter, TargetDateFilter)
    
    '商品コードをキー値、数量を値とした辞書型変数
    Dim resultDict As New Scripting.Dictionary
    
    For Each FileName In filePathCollection
        
        'データの取得準備
        data.FileName = FileName
        'ファイル情報の取得準備
        fileProperty.FilePath = data.FilePath
        '商品コードと数量の合計を取得(辞書型)
        Set resultDict = MergeDictionaries(resultDict, data.DataDict)
        'データワークブックを閉じる
        data.CloseWorkBook
        
    Next FileName
    
    
'''発注確認シートへ入力'''
    Dim startRowIndex As Long
    Dim lastRowIndex As Long
    
    startRowIndex = load.DataNextRowNumber
    lastRowIndex = startRowIndex
    
    
    '商品コードを入力
    For Each productsCode In resultDict.Keys
        load.Cells(lastRowIndex, load.ProductCodeColumnNumber) = productsCode
        lastRowIndex = lastRowIndex + 1
    Next productsCode
    
    '発注確認に商品コードを入力した範囲
    Dim target As Range
    Set target = load.Worksheet.Range(load.ProductCodeColumn & startRowIndex & ":" & load.ProductCodeColumn & lastRowIndex)
    '商品情報の表示
    DisplayProductsInfo target
    
    '数量を入力
    lastRowIndex = startRowIndex
    For Each productsCode In resultDict.Keys
        load.Cells(lastRowIndex, load.QtyColumnNumber) = resultDict(productsCode)
        lastRowIndex = lastRowIndex + 1
    Next productsCode
End Sub
