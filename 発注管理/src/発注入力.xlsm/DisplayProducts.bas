Attribute VB_Name = "DisplayProducts"
'発注入力に入力された商品コードから商品情報を表示

Sub DisplayProductsInfo(targetRng As Range)

    Dim dataStorage As New DataBaseAccesser
    Dim order As New OrderSheetAccesser

    ' 処理する部門の指定
    Dim bumonCD As Integer: bumonCD = order.bumonCode
    '商品CDの列の指定
    Dim targetColumn As Integer: targetColumn = order.ProductCodeColumnIndex
    '数量の列の指定
    Dim qtyColumn As Integer: qtyColumn = order.qtyColumnIndex
    '仕入単価の列の指定
    Dim priceColumn As Integer: priceColumn = order.priceColumnIndex
    '仕入金額の列の指定
    Dim amountColumn As Integer: amountColumn = order.AmountColumnIndex
    
    Dim cell As Range
    
    ' 範囲内の指定した列の各行を処理
    For Each cell In targetRng.Columns(targetColumn).Cells
        ' 空白でないセルを処理
        If cell.value <> "" Then
            '商品が存在する場合
            If dataStorage.ExistsProducts(bumonCD, cell.value) Then
            
                DefaultCellDesign cell
                Call WriteRow(cell, bumonCD, qtyColumn, priceColumn, amountColumn)
                
            Else
            
                ErrorCellDesign cell
                
            End If
        End If
    Next cell
    
    '仕入金額計算式の入力
    ApplyAmountCalcFormulaToRange
    
End Sub
Private Sub WriteRow(cell As Object, bumonCD As Integer, qtyColumn As Integer, priceColumn As Integer, amountColumn As Integer)

    Dim dataStorage As New DataBaseAccesser
    
    Dim rs As ADODB.Recordset
    Set rs = dataStorage.GetProduct(bumonCD, cell.value)
    
    ' レコードセットをセルに貼り付ける
    If Not rs.EOF Then
        Dim i As Integer
        
        Do Until rs.EOF
        
            ' レコードセットをワークシートに貼り付け
            For i = 0 To rs.Fields.count - 1
                cell.Offset(0, i + 1).value = rs.Fields(i).value
            Next i
            
            rs.MoveNext
        Loop
    End If
    
    ' レコードセットを閉じる
    rs.Close
    Set rs = Nothing
    
End Sub
Private Sub ErrorCellDesign(cell As Object)
    Call ChangeBackColor(cell, 255, 0, 0)
End Sub
Private Sub DefaultCellDesign(cell As Object)
    Call ChangeBackColor(cell, 255, 255, 255)
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

'仕入金額計算式の入力
Public Sub ApplyAmountCalcFormulaToRange()
    Dim order As New OrderSheetAccesser
    
    
    Dim piecesColumnIndex As Integer
    Dim qtyColumnIndex As Integer
    Dim priceColumnIndex As Integer
    piecesColumnIndex = order.piecesColumnIndex
    qtyColumnIndex = order.qtyColumnIndex
    priceColumnIndex = order.priceColumnIndex
    
    Dim startRow As Long
    Dim endRow As Long
    Dim targetColumnIndex As Integer
    startRow = order.DataStartRowIndex
    endRow = order.DataEndRowIndex
    targetColumnIndex = order.AmountColumnIndex
    
    Dim row As Long
    Dim formula As String
    
    For row = startRow To endRow
        formula = GetAmountCalcFormula(row, piecesColumnIndex, qtyColumnIndex, priceColumnIndex)
        Cells(row, targetColumnIndex).formula = formula
    Next row
End Sub
'仕入金額の計算式を返す
Private Function GetAmountCalcFormula(rowIndex As Long, piecesColumnIndex As Integer, qtyColumnIndex As Integer, priceColumnIndex As Integer) As String
    GetAmountCalcFormula = "=IFERROR(" & _
                            IndexToLetter(piecesColumnIndex) & _
                            rowIndex & _
                            "*" & _
                            IndexToLetter(qtyColumnIndex) & _
                            rowIndex & _
                            "*" & _
                            IndexToLetter(priceColumnIndex) & _
                            rowIndex & _
                            ",0)"
End Function
