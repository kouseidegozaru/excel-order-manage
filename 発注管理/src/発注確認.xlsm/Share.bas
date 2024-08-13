Attribute VB_Name = "Share"

'数字をアルファベットに変更
Function IndexToLetter(ByVal num As Integer) As String
    If num < 1 Or num > 26 Then
        IndexToLetter = "Out of Range"
    Else
        IndexToLetter = Chr(64 + num)
    End If
End Function


'''以下はSheetAccesserのみで使用する項目'''

' rangeで指定した範囲が一行または一列の場合に一次元のCollectionに格納する
Public Function RangeToOneDimCollection(rng As Range) As Collection
    Dim arr As Variant
    Dim oneDimCollection As New Collection
    Dim i As Integer

    arr = rng.value
    
    If IsEmpty(arr) Then
        Set RangeToOneDimCollection = oneDimCollection
        Exit Function
    End If

    ' 一行か一列かを判定
    If rng.Rows.Count = 1 Then
        ' 一行の場合
        For i = 1 To rng.Columns.Count
            oneDimCollection.Add arr(1, i)
        Next i
    ElseIf rng.Columns.Count = 1 Then
        ' 一列の場合
        For i = 1 To rng.Rows.Count
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
    If col.Count = 0 Then
        Set RemoveFirstRow = newCol
        Exit Function
    End If

    ' 最初の行を削除する
    numRows = col.Count
    
    ' 最初の行を削除して新しいCollectionにコピー
    For i = 2 To numRows
        Set row = New Collection
        For j = 1 To col(i).Count
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


Function RecordsetToArray(rs As ADODB.Recordset) As Variant
    Dim arr As Variant
    Dim i As Long, j As Long
    Dim rowCount As Long
    Dim colCount As Long
    
    ' レコードセットの列数を取得
    colCount = rs.Fields.Count
    
    ' レコードセットの行数を取得
    rs.MoveLast
    rowCount = rs.RecordCount
    rs.MoveFirst
    
    ' 二次元配列を初期化
    ReDim arr(0 To rowCount, 0 To colCount - 1)
    
    ' ヘッダーを配列に格納
    For i = 0 To colCount - 1
        arr(0, i) = rs.Fields(i).name
    Next i
    
    ' データを配列に格納
    i = 1
    Do Until rs.EOF
        For j = 0 To colCount - 1
            arr(i, j) = rs.Fields(j).value
        Next j
        rs.MoveNext
        i = i + 1
    Loop
    
    RecordsetToArray = arr
End Function

