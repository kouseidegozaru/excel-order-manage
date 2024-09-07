Attribute VB_Name = "DisplayProducts"
'発注入力に入力された商品コードから商品情報を表示

Sub DisplayProductsInfo(targetRng As Range)

    'クエリ実行クラス
    Dim dataStorage As New DataBaseAccesser
    '発注入力シート
    Dim order As New OrderSheetAccesser

    ' 処理する部門の指定
    Dim bumonCD As Integer: bumonCD = order.bumonCode
    '商品CDの列の指定
    Dim targetColumn As Integer: targetColumn = order.ProductCodeColumnIndex
    '数量の列の指定
    Dim qtyColumn As Integer: qtyColumn = order.QtyColumnIndex
    '仕入単価の列の指定
    Dim priceColumn As Integer: priceColumn = order.PriceColumnIndex
    '仕入金額の列の指定
    Dim amountColumn As Integer: amountColumn = order.AmountColumnIndex
    
    Dim cell As Range
    
    ' 範囲内の指定した列の各行を処理
    For Each cell In targetRng.Columns(targetColumn).Cells
        ' 空白でないセルを処理
        If cell.value <> "" Then
            '商品が存在する場合
            If dataStorage.ExistsProducts(bumonCD, cell.value) Then
                
                '普通のセルのデザインを適用
                DefaultCellDesign cell
                '行の処理
                Call WriteRow(cell, bumonCD, qtyColumn, priceColumn, amountColumn)
                
            Else
                'エラーセルのデザインを適用
                ErrorCellDesign cell
                
            End If
        End If
    Next cell
    
    '仕入金額計算式の入力
    ApplyAmountCalcFormulaToRange
    
End Sub

'行ごとの処理を定義
Private Sub WriteRow(cell As Object, bumonCD As Integer, qtyColumn As Integer, priceColumn As Integer, amountColumn As Integer)

    'クエリ実行クラス
    Dim dataStorage As New DataBaseAccesser
    
    '商品コードから商品情報を取得
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

'商品コードにエラーがあるセルのデザイン
Private Sub ErrorCellDesign(cell As Object)
    Call ChangeBackColor(cell, 255, 0, 0)
End Sub
'普通のセルのデザイン
Private Sub DefaultCellDesign(cell As Object)
    Call ChangeBackColor(cell, 255, 255, 255)
End Sub

Private Sub ChangeBackColor(cell As Object, r As Integer, g As Integer, b As Integer)
        ' 背景色を設定
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
    
    '入数の列番号
    Dim piecesColumn As Integer
    piecesColumn = order.PiecesColumnIndex
    '数量の列番号
    Dim qtyColumn As Integer
    qtyColumn = order.QtyColumnIndex
    '単価の列番号
    Dim priceColumn As Integer
    priceColumn = order.PriceColumnIndex

    '開始行
    Dim startRow As Long
    startRow = order.DataStartRowIndex
    '終了行
    Dim endRow As Long
    endRow = order.DataEndRowIndex
    '計算式の入力列
    Dim targetColumn As Integer
    targetColumn = order.AmountColumnIndex
    
    Dim row As Long
    '式
    Dim formula As String
    
    For row = startRow To endRow
        '式の取得
        formula = GetAmountCalcFormula(row, piecesColumn, qtyColumn, priceColumn)
        '式の入力
        Cells(row, targetColumn).formula = formula
    Next row
End Sub
'仕入金額の計算式を返す
Private Function GetAmountCalcFormula(row As Long, piecesColumn As Integer, qtyColumn As Integer, priceColumn As Integer) As String
    '入数*数量*単価
    GetAmountCalcFormula = "=IFERROR(" & _
                            IndexToLetter(PiecesColumnIndex) & _
                            row & _
                            "*" & _
                            IndexToLetter(QtyColumnIndex) & _
                            row & _
                            "*" & _
                            IndexToLetter(PriceColumnIndex) & _
                            row & _
                            ",0)"
End Function
