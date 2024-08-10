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

' データをシートに書き込む
Sub writeData(ws As Worksheet, startRowIndex As Long, startColIndex As Integer, writeData As Variant)
    Dim i As Long, j As Long
    Dim item As Variant
    Dim rowCount As Long, colCount As Long

    ' writeDataが配列かコレクションかを確認
    If IsArray(writeData) Then
        rowcnt = 0
        ' 配列の場合
        For i = LBound(writeData, 1) To UBound(writeData, 1)
            colcnt = 0
            For j = LBound(writeData, 2) To UBound(writeData, 2)
                ws.Cells(startRowIndex + rowcnt, startColIndex + colcnt).value = writeData(i, j)
                colcnt = colcnt + 1
            Next j
            rowcnt = rowcnt + 1
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
    Dim minRows As Long
    Dim minCols As Long
    Dim i As Long, j As Long
    
    ' 配列のサイズを取得
    numRows = UBound(arr, 1)
    numCols = UBound(arr, 2)
    minRows = LBound(arr, 1)
    minCols = LBound(arr, 2)
    
    ' 新しい配列のサイズを設定
    ReDim newArr(minRows To numRows - 1, minCols To numCols)
    
    ' 一行目を削除して新しい配列にコピー
    For i = minRows + 1 To numRows
        For j = minCols To numCols
            newArr(i - 1, j) = arr(i, j)
        Next j
    Next i
    
    ' 新しい配列を返す
    RemoveFirstRow = newArr
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

