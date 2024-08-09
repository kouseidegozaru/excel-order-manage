Attribute VB_Name = "Share"

Public Function GetRangeValue(rng As Range) As Collection
    Dim cell As Range
    Dim col As New Collection
    
    ' 範囲内の各セルをループ
    For Each cell In rng
        col.Add cell.value
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

'シートに列を指定して入力
Sub writeData(ws As Worksheet, rowIndex As Long, colIndex As Integer, writeData As Collection)
    Dim item As Variant

    For Each item In writeData
        ws.Cells(rowIndex, colIndex).value = item
        rowIndex = rowIndex + 1
    Next item

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

'辞書を比較しキーが同じものは加算
Function MergeDictionaries(dict1 As Scripting.Dictionary, dict2 As Scripting.Dictionary) As Scripting.Dictionary
    Dim resultDict As New Scripting.Dictionary
    Dim key As Variant

    ' dict1の内容をresultDictにコピー
    For Each key In dict1.Keys
        resultDict(key) = dict1(key)
    Next key

    ' dict2の内容をresultDictに追加
    For Each key In dict2.Keys
        If resultDict.exists(key) Then
            ' 両方の値が数値の場合は加算
            If IsNumeric(resultDict(key)) And IsNumeric(dict2(key)) Then
                resultDict(key) = resultDict(key) + dict2(key)
            ' 片方が数値で片方が数値でない場合は数値の方を結果に反映
            ElseIf IsNumeric(resultDict(key)) Then
                resultDict(key) = resultDict(key)
            ElseIf IsNumeric(dict2(key)) Then
                resultDict(key) = dict2(key)
            ' 両方とも数値でない場合は空文字を結果に反映
            Else
                resultDict(key) = ""
            End If
        Else
            resultDict(key) = dict2(key) ' キーが存在しない場合、新しく追加
        End If
    Next key

    Set MergeDictionaries = resultDict
End Function

'二次元配列から一行目を削除
Function RemoveFirstRow(ByVal arr As Variant) As Variant
    Dim newArr() As Variant
    Dim numRows As Long
    Dim numCols As Long
    Dim i As Long, j As Long
    
    ' 配列のサイズを取得
    numRows = UBound(arr, 1)
    numCols = UBound(arr, 2)
    
    ' 新しい配列のサイズを設定
    ReDim newArr(1 To numRows - 1, 1 To numCols)
    
    ' 一行目を削除して新しい配列にコピー
    For i = 2 To numRows
        For j = 1 To numCols
            newArr(i - 1, j) = arr(i, j)
        Next j
    Next i
    
    ' 新しい配列を返す
    RemoveFirstRow = newArr
End Function
