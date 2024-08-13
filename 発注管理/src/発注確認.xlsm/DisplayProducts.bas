Attribute VB_Name = "DisplayProducts"
'発注入力に入力された商品コードから商品情報を表示

'仕入金額の計算式を返す
Private Function GetAmountCalcFormula(QtyColumnIndex As Integer, qtyRowIndex As Long, PriceColumnIndex As Integer, priceRowIndex As Long) As String
    GetAmountCalcFormula = "=IFERROR(" & _
                            NumberToLetter(QtyColumnIndex) & _
                            qtyRowIndex & _
                            "*" & _
                            NumberToLetter(PriceColumnIndex) & _
                            priceRowIndex & _
                            ",0)"
End Function

