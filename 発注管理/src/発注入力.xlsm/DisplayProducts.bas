Attribute VB_Name = "DisplayProducts"
'�������͂ɓ��͂��ꂽ���i�R�[�h���珤�i����\��

Sub DisplayProductsInfo(targetRng As Range)

    Dim dataStorage As New DataBaseAccesser
    Dim order As New OrderSheetAccesser
    
    ' �������镔��̎w��
    Dim bumonCD As Integer
    bumonCD = order.bumonCode
    ' ���������̎w��
    Dim targetColumn As Integer
    targetColumn = order.ProductCodeColumnIndex
    '���ʂ̗�̎w��
    Dim qtyColumn As Integer
    qtyColumn = order.qtyColumnIndex
    '�d���P���̗�̎w��
    Dim priceColumn As Integer
    priceColumn = order.PriceColumnIndex
    '�d�����z�̗�̎w��
    Dim amountColumn As Integer
    amountColumn = order.AmountColumnIndex
    
    Dim cell As Range
    
    ' �͈͓��̎w�肵����̊e�s������
    For Each cell In targetRng.Columns(targetColumn).Cells
        ' �󔒂łȂ��Z��������
        If cell.value <> "" Then
            If dataStorage.ExistsProducts(bumonCD, cell.value) Then
                '�w�i�𔒂�
                Call ChangeBackColor(cell, 255, 255, 255)
                
                Dim rs As ADODB.recordSet
                Set rs = dataStorage.GetProduct(bumonCD, cell.value)
                
                ' ���R�[�h�Z�b�g���Z���ɓ\��t����
                If Not rs.EOF Then
                    Dim i As Integer
                    
                    ' ���R�[�h�Z�b�g�����[�N�V�[�g�ɓ\��t��
'                    rs.MoveFirst
                    Do Until rs.EOF
                        For i = 0 To rs.Fields.count - 1
                            cell.Offset(0, i + 1).value = rs.Fields(i).value
                        Next i
                        '�d�����z�̌v�Z����ݒ�
                        cell.Offset(0, amountColumn - 1).value = GetAmountCalcFormula(qtyColumn, cell.row, priceColumn, cell.row)
                        
                        rs.MoveNext
                    Loop
                End If
                
                ' ���R�[�h�Z�b�g�����
                rs.Close
                Set rs = Nothing
            Else
                '���i�R�[�h�����݂��Ȃ��ꍇ�͔w�i��Ԃ�
                Call ChangeBackColor(cell, 255, 0, 0)
            End If
        End If
    Next cell
    
End Sub

Private Sub ChangeBackColor(cell As Object, r As Integer, g As Integer, b As Integer)
        ' �w�i�F��Ԃɐݒ�
        cell.Interior.color = RGB(r, g, b)
        
        ' �����̌r����ێ�
        cell.Borders(xlEdgeLeft).LineStyle = xlContinuous
        cell.Borders(xlEdgeTop).LineStyle = xlContinuous
        cell.Borders(xlEdgeBottom).LineStyle = xlContinuous
        cell.Borders(xlEdgeRight).LineStyle = xlContinuous
        cell.Borders(xlInsideVertical).LineStyle = xlContinuous
        cell.Borders(xlInsideHorizontal).LineStyle = xlContinuous
End Sub

'�d�����z�̌v�Z����Ԃ�
Private Function GetAmountCalcFormula(qtyColumnIndex As Integer, qtyRowIndex As Long, PriceColumnIndex As Integer, priceRowIndex As Long) As String
    GetAmountCalcFormula = "=IFERROR(" & _
                            IndexToLetter(qtyColumnIndex) & _
                            qtyRowIndex & _
                            "*" & _
                            IndexToLetter(PriceColumnIndex) & _
                            priceRowIndex & _
                            ",0)"
End Function
