Attribute VB_Name = "DisplayProducts"
'�������͂ɓ��͂��ꂽ���i�R�[�h���珤�i����\��

'�d�����z�̌v�Z����Ԃ�
Private Function GetAmountCalcFormula(QtyColumnIndex As Integer, qtyRowIndex As Long, PriceColumnIndex As Integer, priceRowIndex As Long) As String
    GetAmountCalcFormula = "=IFERROR(" & _
                            NumberToLetter(QtyColumnIndex) & _
                            qtyRowIndex & _
                            "*" & _
                            NumberToLetter(PriceColumnIndex) & _
                            priceRowIndex & _
                            ",0)"
End Function

