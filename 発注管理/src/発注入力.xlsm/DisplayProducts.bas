Attribute VB_Name = "DisplayProducts"
'発注入力に入力された商品コードから商品情報を表示

Sub DisplayProductsInfo(targetRng As range)

    Dim DataStorage As New DataBaseAccesser
    Dim order As New OrderSheetAccesser
    
    ' 処理する部門の指定
    Dim BumonCD As Integer
    BumonCD = order.BumonCode
    ' 処理する列の指定
    Dim targetColumn As Integer
    targetColumn = order.ProductCodeColumnNumber
    '数量の列の指定
    Dim QtyColumn As Integer
    QtyColumn = order.QtyColumnNumber
    '仕入単価の列の指定
    Dim priceColumn As Integer
    priceColumn = order.PriceColumnNumber
    '仕入金額の列の指定
    Dim amountColumn As Integer
    amountColumn = order.AmountColumnNumber
    
    Dim cell As range
    
    ' 範囲内の指定した列の各行を処理
    For Each cell In targetRng.Columns(targetColumn).Cells
        ' 空白でないセルを処理
        If cell.value <> "" Then
            If DataStorage.ExistsProducts(BumonCD, cell.value) Then
                '背景を白に
                Call ChangeBackColor(cell, 255, 255, 255)
                
                Dim rs As ADODB.recordSet
                Set rs = DataStorage.GetProduct(BumonCD, cell.value)
                
                ' レコードセットをセルに貼り付ける
                If Not rs.EOF Then
                    Dim i As Integer
                    
                    ' レコードセットをワークシートに貼り付け
'                    rs.MoveFirst
                    Do Until rs.EOF
                        For i = 0 To rs.Fields.Count - 1
                            cell.Offset(0, i + 1).value = rs.Fields(i).value
                        Next i
                        '仕入金額の計算式を設定
                        cell.Offset(0, amountColumn - 1).value = GetAmountCalcFormula(QtyColumn, cell.Row, priceColumn, cell.Row)
                        
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
        cell.Interior.color = RGB(r, g, b)
        
        ' 既存の罫線を保持
        cell.Borders(xlEdgeLeft).LineStyle = xlContinuous
        cell.Borders(xlEdgeTop).LineStyle = xlContinuous
        cell.Borders(xlEdgeBottom).LineStyle = xlContinuous
        cell.Borders(xlEdgeRight).LineStyle = xlContinuous
        cell.Borders(xlInsideVertical).LineStyle = xlContinuous
        cell.Borders(xlInsideHorizontal).LineStyle = xlContinuous
End Sub

'仕入金額の計算式を返す
Private Function GetAmountCalcFormula(qtyColumnIndex As Integer, qtyRowIndex As Long, priceColumnIndex As Integer, priceRowIndex As Long) As String
    GetAmountCalcFormula = "=IFERROR(" & _
                            NumberToLetter(qtyColumnIndex) & _
                            qtyRowIndex & _
                            "*" & _
                            NumberToLetter(priceColumnIndex) & _
                            priceRowIndex & _
                            ",0)"
End Function
