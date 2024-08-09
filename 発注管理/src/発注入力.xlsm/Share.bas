Attribute VB_Name = "Share"

Public Function GetRangeValue(rng As range) As Collection
    Dim cell As range
    Dim col As New Collection
    
    ' 範囲内の各セルをループ
    For Each cell In rng
        col.add cell.value
    Next cell
    
    Set GetRangeValue = col
    
End Function

'数字をアルファベットに変更
Function NumberToLetter(ByVal num As Integer) As String
    If num < 1 Or num > 26 Then
        NumberToLetter = "Out of Range"
    Else
        NumberToLetter = Chr(64 + num)
    End If
End Function

' データをシートに書き込む
Sub writeData(ws As Worksheet, startRowIndex As Long, startColIndex As Integer, writeData As Variant)
    Dim i As Long, j As Long
    Dim item As Variant
    Dim rowCount As Long, colCount As Long

    ' writeDataが配列かコレクションかを確認
    If IsArray(writeData) Then
        ' 配列の場合
        For i = LBound(writeData, 1) To UBound(writeData, 1)
            For j = LBound(writeData, 2) To UBound(writeData, 2)
                ws.Cells(startRowIndex + i, startColIndex + j).value = writeData(i, j)
            Next j
        Next i
    ElseIf TypeName(writeData) = "Collection" Then
        ' コレクションの場合
        rowCount = 0
        For Each item In writeData
            If IsArray(item) Then
                ' 内部配列の長さを取得
                colCount = UBound(item, 2) - LBound(item, 2) + 1
                For j = LBound(item, 2) To UBound(item, 2)
                    ws.Cells(startRowIndex + rowCount, startColIndex + j - LBound(item, 2)).value = item(j)
                Next j
                rowCount = rowCount + 1
            End If
        Next item
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
            resultCol.add itemBase
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


