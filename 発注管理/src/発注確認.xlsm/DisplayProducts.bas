Attribute VB_Name = "DisplayProducts"
'発注入力に入力された商品コードから商品情報を表示

Sub DisplayProductsInfo(targetRng As Range)

    Dim dataStorage As New DataBaseAccesser
    Dim order As New OrderSheetAccesser
    
    ' 処理する部門の指定
    Dim BumonCD As Integer
    BumonCD = order.BumonCode
    ' 処理する列の指定
    Dim targetColumn As Integer
    targetColumn = order.ProductCodeColumnNumber
    
    Dim cell As Range
    
    ' 範囲内の指定した列の各行を処理
    For Each cell In targetRng.Columns(targetColumn).Cells
        ' 空白でないセルを処理
        If cell.value <> "" Then
            If dataStorage.ExistsProducts(BumonCD, cell.value) Then
                '背景を白に
                Call ChangeBackColor(cell, 255, 255, 255)
                
                Dim rs As ADODB.Recordset
                Set rs = dataStorage.GetProduct(BumonCD, cell.value)
                
                ' レコードセットをセルに貼り付ける
                If Not rs.EOF Then
                    Dim startRow As Long
                    Dim startCol As Long
                    Dim i As Integer
                    
                    ' 貼り付け開始セルを指定（セルの行と同じ行に）
                    startRow = cell.Row
                    startCol = cell.Column + 1 ' 左のセルから1列右に貼り付け
                    
                    ' レコードセットをワークシートに貼り付け
'                    rs.MoveFirst
                    Do Until rs.EOF
                        For i = 0 To rs.Fields.Count - 1
                            cell.Offset(0, i + 1).value = rs.Fields(i).value
                        Next i
                        rs.MoveNext
                    Loop
                End If
                
                ' レコードセットを閉じる
                rs.Close
                Set rs = Nothing
            Else
                '商品コードが存在しない場合は背景を赤に
                Call ChangeBackColor(cell, 255, 0, 0)
            End If
        End If
    Next cell
    
End Sub

Private Sub ChangeBackColor(cell As Object, r As Integer, g As Integer, b As Integer)
        ' 背景色を赤に設定
        cell.Interior.Color = RGB(r, g, b)
        
        ' 既存の罫線を保持
        cell.Borders(xlEdgeLeft).LineStyle = xlContinuous
        cell.Borders(xlEdgeTop).LineStyle = xlContinuous
        cell.Borders(xlEdgeBottom).LineStyle = xlContinuous
        cell.Borders(xlEdgeRight).LineStyle = xlContinuous
        cell.Borders(xlInsideVertical).LineStyle = xlContinuous
        cell.Borders(xlInsideHorizontal).LineStyle = xlContinuous
End Sub


