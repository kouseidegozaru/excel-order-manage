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


Function RecordsetToCollection(rs As ADODB.Recordset) As Collection
    Dim col As Collection
    Dim rowCol As Collection
    Dim i As Long
    
    ' コレクションを初期化
    Set col = New Collection
    
    ' レコードセットが空でないことを確認
    If Not rs.EOF Then
        rs.MoveFirst
        
        ' データをコレクションに格納
        Do Until rs.EOF
            Set rowCol = New Collection
            For i = 0 To rs.Fields.Count - 1
                rowCol.Add rs.Fields(i).value, rs.Fields(i).name
            Next i
            col.Add rowCol
            rs.MoveNext
        Loop
    End If
    
    Set RecordsetToCollection = col
End Function


