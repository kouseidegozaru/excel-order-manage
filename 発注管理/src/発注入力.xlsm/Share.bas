Attribute VB_Name = "Share"

'数字をアルファベットに変更
Function IndexToLetter(ByVal num As Integer) As String
    If num < 1 Or num > 26 Then
        IndexToLetter = "Out of Range"
    Else
        IndexToLetter = Chr(64 + num)
    End If
End Function

'collection型の変数を比べ重複する値を除外
Function FilterCollection(baseCol As Collection, filterCol As Collection) As Collection
    Dim resultCol As New Collection
    Dim itemBase As Variant
    Dim itemFilter As Variant
    Dim exists As Boolean
    
    ' baseColの値をループして、filterColに存在しないものだけresultColに追加
    For Each itemBase In baseCol
        exists = False
        For Each itemFilter In filterCol
            If itemBase = itemFilter Then
                exists = True
                Exit For
            End If
        Next itemFilter
        If Not exists Then
            resultCol.Add itemBase
        End If
    Next itemBase
    
    ' 結果のコレクションを返す
    Set FilterCollection = resultCol
End Function

'Numberが MultipleOfの倍数の場合にTrueを返す
Function IsMultiple(Number As Long, MultipleOf As Long) As Boolean
    If MultipleOf = 0 Then
        IsMultiple = True
    Else
        IsMultiple = (Number Mod MultipleOf = 0)
    End If
End Function

'ディレクトリへのアクセス権限のチェック
Function CheckDirectoryAccess(ByVal directoryPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Dir関数でディレクトリが存在し、アクセス可能かどうか確認
    If Dir(directoryPath, vbDirectory) <> "" Then
        CheckDirectoryAccess = True
    Else
        CheckDirectoryAccess = False
    End If
    
    Exit Function

ErrorHandler:
    ' エラーハンドリング: アクセスできない場合、Falseを返す
    CheckDirectoryAccess = False
End Function

'''以下はSheetAccesserのみで使用する項目'''

' rangeで指定した範囲が一行または一列の場合に一次元のCollectionに格納する
Public Function RangeToOneDimCollection(rng As Range) As Collection
    Dim arr As Variant
    Dim oneDimCollection As New Collection
    Dim i As Integer

    arr = rng.value
    
    '空の場合空のコレクションを返す
    If IsEmpty(arr) Then
        Set RangeToOneDimCollection = oneDimCollection
        Exit Function
    End If
    
    'range.valueで範囲が一つのセル番地のみを指す場合に配列ではなくなる
    If Not IsArray(arr) Then
        oneDimCollection.Add arr
    
    ' 一行か一列かを判定
    ElseIf rng.Rows.count = 1 Then
        ' 一行の場合
        For i = 1 To rng.Columns.count
            oneDimCollection.Add arr(1, i)
        Next i
        
    ElseIf rng.Columns.count = 1 Then
        ' 一列の場合
        For i = 1 To rng.Rows.count
            oneDimCollection.Add arr(i, 1)
        Next i
        
    End If

    Set RangeToOneDimCollection = oneDimCollection
End Function

'二次元コレクションから一行目を削除
Function RemoveFirstRow(ByVal col As Collection) As Collection
    Dim newCol As Collection
    Dim item As Variant
    Dim row As Collection
    Dim numRows As Long
    Dim numCols As Long
    Dim i As Long, j As Long
    
    ' 新しいCollectionを作成
    Set newCol = New Collection
    
    ' Collectionの最初の行を取り出す
    If col.count = 0 Then
        Set RemoveFirstRow = newCol
        Exit Function
    End If

    ' 最初の行を削除する
    numRows = col.count
    
    ' 最初の行を削除して新しいCollectionにコピー
    For i = 2 To numRows
        Set row = New Collection
        For j = 1 To col(i).count
            row.Add col(i)(j)
        Next j
        newCol.Add row
    Next i
    
    ' 新しいCollectionを返す
    Set RemoveFirstRow = newCol
End Function

'二次元配列を二次元のコレクションに変換する
Function ArrayToCollection(ByVal arr As Variant) As Collection
    Dim col As New Collection
    Dim innerCol As Collection
    Dim i As Long, j As Long

    ' 行のループ
    For i = LBound(arr, 1) To UBound(arr, 1)
        Set innerCol = New Collection
        
        ' 列のループ
        For j = LBound(arr, 2) To UBound(arr, 2)
            innerCol.Add arr(i, j)
        Next j
        
        col.Add innerCol
    Next i
    
    Set ArrayToCollection = col
End Function
