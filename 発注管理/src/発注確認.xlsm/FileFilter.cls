VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'指定のディレクトリから条件にあうファイル名を取得するクラス

Private dirPath_ As String

'ディレクトリパスの設定
Property Let DirPath(targetDirPath As String)
    dirPath_ = targetDirPath
End Property

'ファイル名をAND条件で取得する
Public Function AndFilter(ParamArray filterNames() As Variant) As Collection
    Dim fileList As New Collection
    Dim fileName As String
    
    'ファイル名を抽出
    fileName = Dir(dirPath_ & "\*.*")
    
    Do While fileName <> ""
        '条件がすべて一致するかのフラグ
        Dim matchesAll As Boolean
        matchesAll = True
        
        For Each filterName In filterNames
            'フィルター条件がファイル名に存在しない場合
            If InStr(1, fileName, filterName, vbTextCompare) = 0 Then
                '不一致として抜ける
                matchesAll = False
                Exit For
            End If
        Next filterName
        
        'すべての条件が一致している場合
        If matchesAll Then
            'ファイルリストに追加
            fileList.Add fileName
        End If
        
        '次のファイル名へ
        fileName = Dir()
    Loop
    
    Set AndFilter = fileList
End Function

'ファイル名をOR条件で取得する
Public Function OrFilter(ParamArray filterNames() As Variant) As Collection
    Dim fileList As New Collection
    Dim fileName As String
    
    'ファイル名を抽出
    fileName = Dir(dirPath_ & "\*.*")
    Do While fileName <> ""
    
        '条件が一つでも一致するかのフラグ
        Dim matchesAny As Boolean
        matchesAny = False
        
        For Each filterName In filterNames
        
            'filterNameがAndFilterメソッドの戻り値の場合
            If IsObject(filterName) Then
                Dim subFilterResult As Collection
                Set subFilterResult = filterName
                Dim subFile As Variant
                'AndFilterメソッドの戻り値のファイル名を処理
                For Each subFile In subFilterResult
                    'ANDフィルターに該当のファイル名が含まれている場合
                    If subFile = fileName Then
                        '一致
                        matchesAny = True
                        Exit For
                    End If
                Next subFile
                
            'フィルター条件がファイル名に存在する場合
            ElseIf InStr(1, fileName, filterName, vbTextCompare) > 0 Then
                '一致
                matchesAny = True
                Exit For
            End If
        Next filterName
        
        '一つでも条件が一致している場合
        If matchesAny Then
            'ファイルリストに追加
            fileList.Add fileName
        End If
        
        '次のファイル名へ
        fileName = Dir()
    Loop
    
    Set OrFilter = fileList
End Function

