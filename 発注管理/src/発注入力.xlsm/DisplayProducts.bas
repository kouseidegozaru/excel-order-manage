Attribute VB_Name = "DisplayProducts"
Sub DisplayProductsInfo(targetRng As range)

    Dim DataStorage As New dataAccesser
    Dim BumonCD As Integer
    BumonCD = GetBumonCD
    Dim cell As range
    
    ' 処理する列の指定
    Dim targetColumn As Integer
    targetColumn = OrderWb_ProductCodeColumnNumber
    
    ' 範囲内の指定した列の各行を処理
    For Each cell In targetRng.Columns(targetColumn).Cells
        ' 空白でないセルを処理
        If cell.Value <> "" Then
            If DataStorage.ExistsProducts(BumonCD, cell.Value) Then
                Dim rs As ADODB.recordSet
                Set rs = DataStorage.GetProduct(BumonCD, cell.Value)
                
                ' レコードセットをセルに貼り付ける
                If Not rs.EOF Then
                    Dim startRow As Long
                    Dim startCol As Long
                    Dim i As Integer
                    
                    ' 貼り付け開始セルを指定（セルの行と同じ行に）
                    startRow = cell.row
                    startCol = cell.Column + 1 ' 左のセルから1列右に貼り付け
                    
                    ' レコードセットをワークシートに貼り付け
'                    rs.MoveFirst
                    Do Until rs.EOF
                        For i = 0 To rs.Fields.Count - 1
                            cell.Offset(0, i + 1).Value = rs.Fields(i).Value
                        Next i
                        rs.MoveNext
                    Loop
                End If
                
                ' レコードセットを閉じる
                rs.Close
                Set rs = Nothing
            End If
        End If
    Next cell
    
End Sub

