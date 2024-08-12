Attribute VB_Name = "Share"

'数字をアルファベットに変更
Function IndexToLetter(ByVal num As Integer) As String
    If num < 1 Or num > 26 Then
        IndexToLetter = "Out of Range"
    Else
        IndexToLetter = Chr(64 + num)
    End If
End Function

' データをシートに書き込む
Sub writeData(ws As Worksheet, startRowIndex As Long, startColIndex As Integer, writeData As Variant)
    Dim i As Long, j As Long
    Dim item As Variant
    Dim rowCollection As Variant
    Dim rowCount As Long, colCount As Long

    ' writeDataが配列かコレクションかを確認
    If IsArray(writeData) Then
        ' 配列の場合
        For i = LBound(writeData, 1) To UBound(writeData, 1)
            For j = LBound(writeData, 2) To UBound(writeData, 2)
                ws.Cells(startRowIndex + i - 1, startColIndex + j - 1).value = writeData(i, j)
            Next j
        Next i
    ElseIf TypeName(writeData) = "Collection" Then
        ' コレクションの場合
        rowCount = 0
        For Each rowCollection In writeData
            colCount = 0
            If TypeName(rowCollection) = "Collection" Then
                ' 内部がさらにコレクションの場合
                For Each item In rowCollection
                    ws.Cells(startRowIndex + rowCount, startColIndex + colCount).value = item
                    colCount = colCount + 1
                Next item
            ElseIf IsArray(rowCollection) Then
                ' 内部が配列の場合
                For j = LBound(rowCollection, 1) To UBound(rowCollection, 1)
                    ws.Cells(startRowIndex + rowCount, startColIndex + j - LBound(rowCollection, 1)).value = rowCollection(j)
                Next j
            End If
            rowCount = rowCount + 1
        Next rowCollection
    Else
        ' エラーハンドリング
        Err.Raise vbObjectError + 9999, "writeData", "writeDataは配列またはコレクションでなければなりません。"
    End If
End Sub

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

