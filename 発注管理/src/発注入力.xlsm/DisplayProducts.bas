Attribute VB_Name = "DisplayProducts"
'�������͂ɓ��͂��ꂽ���i�R�[�h���珤�i����\��

Sub DisplayProductsInfo(targetRng As Range)

    Dim dataStorage As New DataBaseAccesser
    Dim order As New OrderSheetAccesser

    ' �������镔��̎w��
    Dim bumonCD As Integer: bumonCD = order.bumonCode
    '���iCD�̗�̎w��
    Dim targetColumn As Integer: targetColumn = order.ProductCodeColumnIndex
    '���ʂ̗�̎w��
    Dim qtyColumn As Integer: qtyColumn = order.qtyColumnIndex
    '�d���P���̗�̎w��
    Dim priceColumn As Integer: priceColumn = order.priceColumnIndex
    '�d�����z�̗�̎w��
    Dim amountColumn As Integer: amountColumn = order.AmountColumnIndex
    
    Dim cell As Range
    
    ' �͈͓��̎w�肵����̊e�s������
    For Each cell In targetRng.Columns(targetColumn).Cells
        ' �󔒂łȂ��Z��������
        If cell.value <> "" Then
            '���i�����݂���ꍇ
            If dataStorage.ExistsProducts(bumonCD, cell.value) Then
            
                DefaultCellDesign cell
                Call WriteRow(cell, bumonCD, qtyColumn, priceColumn, amountColumn)
                
            Else
            
                ErrorCellDesign cell
                
            End If
        End If
    Next cell
    
    '�d�����z�v�Z���̓���
    ApplyAmountCalcFormulaToRange
    
End Sub
Private Sub WriteRow(cell As Object, bumonCD As Integer, qtyColumn As Integer, priceColumn As Integer, amountColumn As Integer)

    Dim dataStorage As New DataBaseAccesser
    
    Dim rs As ADODB.Recordset
    Set rs = dataStorage.GetProduct(bumonCD, cell.value)
    
    ' ���R�[�h�Z�b�g���Z���ɓ\��t����
    If Not rs.EOF Then
        Dim i As Integer
        
        Do Until rs.EOF
        
            ' ���R�[�h�Z�b�g�����[�N�V�[�g�ɓ\��t��
            For i = 0 To rs.Fields.count - 1
                cell.Offset(0, i + 1).value = rs.Fields(i).value
            Next i
            
            rs.MoveNext
        Loop
    End If
    
    ' ���R�[�h�Z�b�g�����
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

'�d�����z�v�Z���̓���
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
'�d�����z�̌v�Z����Ԃ�
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
