Attribute VB_Name = "Module1"
Sub LoadFileProperty()
    Dim load As New LoadSheetAccesser
    Dim data As New DataSheetAccesser
    Dim dataStorage As New DataBaseAccesser
    Dim fileProperty As New FilePropertyManager
    Dim filter As New fileFilter
    
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
    
    Dim rs As ADODB.Recordset
    Set rs = dataStorage.GetUserCodes(load.BumonCode)
    
    Do Until rs.EOF
        'ファイルの抽出条件文字列の設定
        UserCodeFilter = fileProperty.UserCodeIdentifier & rs.Fields("担当者CD").value & fileProperty.BreakIdentifier
        
        'フィルターの実行
        Dim filePathCollection As Collection
        Set filePathCollection = filter.AndFilter(BumonCodeFilter, TargetDateFilter, UserCodeFilter)
        
        'フィルターの結果が存在する場合
        If filePathCollection.Count > 0 Then
            fileProperty.filePath = data.DataDirPath & "\" & filePathCollection(1)
            load.AddFileProperty rs.Fields("担当者名").value, True, fileProperty.UpdatedDate
        Else
            load.AddFileProperty rs.Fields("担当者名").value, False, Date
        End If
        
        rs.MoveNext
    Loop
    
    
    
    
End Sub

Sub LoadFilePropertys()
Dim filter As New fileFilter
filter.DirPath = "C:\Users\mfh077_user.MEFUREDMN\Desktop\excel-order-manage\発注管理\data"

Dim c As Collection
Set c = filter.AndFilter("b40-", "d20240725-", "u30-")
End Sub

