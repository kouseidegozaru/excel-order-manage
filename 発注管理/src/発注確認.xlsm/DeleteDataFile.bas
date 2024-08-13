Attribute VB_Name = "DeleteDataFile"
Sub DeleteDataFiles()
    
    Dim data As New DataSheetAccesser
    Dim fileProperty As New FilePropertyManager
    Dim filter As New FileFilter
    Dim FilePath As String
    Dim fs As New Scripting.FileSystemObject
    
    '対象ディレクトリを設定
    filter.DirPath = data.DataDirPath
    
    'ファイル名の取得
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
        FilePath = data.DataDirPath & "\" & fileName
        fileProperty.InitFilePath = FilePath
        
        '一か月前の場合ファイル削除
        If fileProperty.TargetDate < oneMonthAgo Then
            fs.DeleteFile FilePath
        End If
        
    Next fileName
        
End Sub
