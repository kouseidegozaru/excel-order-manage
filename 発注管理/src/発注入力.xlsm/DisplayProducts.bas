Attribute VB_Name = "DisplayProducts"
'発注入力に入力された商品コードから商品情報を表示

Sub DisplayProductsInfo(targetRng As Range)

    Dim dataStorage As New DataBaseAccesser
    Dim order As New OrderSheetAccesser
    
    ' 処理する部門の指定
    Dim bumonCD As Integer
    bumonCD = order.bumonCode
    ' 処理する列の指定
    Dim targetColumn As Integer
    targetColumn = order.ProductCodeColumnIndex
    '数量の列の指定
    Dim qtyColumn As Integer
    qtyColumn = order.qtyColumnIndex
    '仕入単価の列の指定
    Dim priceColumn As Integer
    priceColumn = order.PriceColumnIndex
    '仕入金額の列の指定
    Dim amountColumn As Integer
    amountColumn = order.AmountColumnIndex
    
    Dim cell As Range
    
    ' 範囲内の指定した列の各行を処理
    For Each cell In targetRng.Columns(targetColumn).Cells
        ' 空白でないセルを処理
        If cell.value <> "" Then
            If dataStorage.ExistsProducts(bumonCD, cell.value) Then
                '背景を白に
                Call ChangeBackColor(cell, 255, 255, 255)
                
                Dim rs As ADODB.recordSet
                Set rs = dataStorage.GetProduct(bumonCD, cell.value)
                
                ' レコードセットをセルに貼り付ける
                If Not rs.EOF Then
                    Dim i As Integer
                    
                    ' レコードセットをワークシートに貼り付け
'                    rs.MoveFirst
                    Do Until rs.EOF
                        For i = 0 To rs.Fields.count - 1
                            cell.Offset(0, i + 1).value = rs.Fields(i).value
                        Next i
                        '仕入金額の計算式を設定
                        cell.Offset(0, amountColumn - 1).value = GetAmountCalcFormula(qtyColumn, cell.row, priceColumn, cell.row)
                        
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
Private Function GetAmountCalcFormula(qtyColumnIndex As Integer, qtyRowIndex As Long, PriceColumnIndex As Integer, priceRowIndex As Long) As String
    GetAmountCalcFormula = "=IFERROR(" & _
                            IndexToLetter(qtyColumnIndex) & _
                            qtyRowIndex & _
                            "*" & _
                            IndexToLetter(PriceColumnIndex) & _
                            priceRowIndex & _
                            ",0)"
End Function
