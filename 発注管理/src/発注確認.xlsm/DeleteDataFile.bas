Attribute VB_Name = "DeleteDataFile"
'一か月前の発注日の発注データファイルを削除
Sub DeleteDataFiles()
    
    '発注データシートアクセサのインスタンス化
    Dim data As New DataSheetAccesser
    'データファイルの属性情報取得クラスをインスタンス化
    Dim fileProperty As New FilePropertyManager
    'データファイルをファイルごとにフィルターするクラスをインスタンス化
    Dim filter As New FileFilter
    
    Dim FilePath As String
    Dim fs As New Scripting.FileSystemObject
    
    '対象ディレクトリを設定
    filter.DirPath = data.SaveDirPath
    
    '全てのファイル名の取得
    Dim fileNames As Collection
    Set fileNames = filter.AndFilter()
    
    ' 今日の日付を取得
    Dim today As Date
    today = Date

    ' 1か月前の日付を計算
    Dim oneMonthAgo As Date
    oneMonthAgo = DateAdd("m", -1, today)
    
    For Each fileName In fileNames
    
        'ファイル情報取得準備
        FilePath = data.SaveDirPath & "\" & fileName
        fileProperty.InitFilePath FilePath
        
        '一か月前の場合ファイル削除
        If fileProperty.targetDate < oneMonthAgo Then
            fs.DeleteFile FilePath
        End If
        
    Next fileName
        
End Sub
'一か月前の発注日の発注済み商品コードデータファイルを削除
Sub DeleteOrderedDataFiles()
    
    '発注データシートアクセサのインスタンス化
    Dim ordered As New OrderedDataSheetAccesser
    'データファイルの属性情報取得クラスをインスタンス化
    Dim fileProperty As New FilePropertyManager
    'データファイルをファイルごとにフィルターするクラスをインスタンス化
    Dim filter As New FileFilter
    
    Dim FilePath As String
    Dim fs As New Scripting.FileSystemObject
    
    '対象ディレクトリを設定
    filter.DirPath = ordered.SaveDirPath
    
    '全てのファイル名の取得
    Dim fileNames As Collection
    Set fileNames = filter.AndFilter()
    
    ' 今日の日付を取得
    Dim today As Date
    today = Date

    ' 1か月前の日付を計算
    Dim oneMonthAgo As Date
    oneMonthAgo = DateAdd("m", -1, today)
    
    For Each fileName In fileNames
    
        'ファイル情報取得準備
        FilePath = ordered.SaveDirPath & "\" & fileName
        fileProperty.InitFilePath FilePath
        
        '一か月前の場合ファイル削除
        If fileProperty.targetDate < oneMonthAgo Then
            fs.DeleteFile FilePath
        End If
        
    Next fileName
        
End Sub
