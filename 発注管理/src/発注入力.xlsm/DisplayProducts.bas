Attribute VB_Name = "DisplayProducts"
'�������͂ɓ��͂��ꂽ���i�R�[�h���珤�i����\��

Sub DisplayProductsInfo(targetRng As Range)

    '�N�G�����s�N���X
    Dim dataStorage As New DataBaseAccesser
    '�������̓V�[�g
    Dim order As New OrderSheetAccesser

    ' �������镔��̎w��
    Dim bumonCD As Integer: bumonCD = order.bumonCode
    '���iCD�̗�̎w��
    Dim targetColumn As Integer: targetColumn = order.ProductCodeColumnIndex
    '���ʂ̗�̎w��
    Dim qtyColumn As Integer: qtyColumn = order.QtyColumnIndex
    '�d���P���̗�̎w��
    Dim priceColumn As Integer: priceColumn = order.PriceColumnIndex
    '�d�����z�̗�̎w��
    Dim amountColumn As Integer: amountColumn = order.AmountColumnIndex
    
    Dim cell As Range
    
    ' �͈͓��̎w�肵����̊e�s������
    For Each cell In targetRng.Columns(targetColumn).Cells
        ' �󔒂łȂ��Z��������
        If cell.value <> "" Then
            '���i�����݂���ꍇ
            If dataStorage.ExistsProducts(bumonCD, cell.value) Then
                
                '���ʂ̃Z���̃f�U�C����K�p
                DefaultCellDesign cell
                '�s�̏���
                Call WriteRow(cell, bumonCD, qtyColumn, priceColumn, amountColumn)
                
            Else
                '�G���[�Z���̃f�U�C����K�p
                ErrorCellDesign cell
                
            End If
        End If
    Next cell
    
    '�d�����z�v�Z���̓���
    ApplyAmountCalcFormulaToRange
    
End Sub

'�s���Ƃ̏������`
Private Sub WriteRow(cell As Object, bumonCD As Integer, qtyColumn As Integer, priceColumn As Integer, amountColumn As Integer)

    '�N�G�����s�N���X
    Dim dataStorage As New DataBaseAccesser
    
    '���i�R�[�h���珤�i�����擾
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

'���i�R�[�h�ɃG���[������Z���̃f�U�C��
Private Sub ErrorCellDesign(cell As Object)
    Call ChangeBackColor(cell, 255, 0, 0)
End Sub
'���ʂ̃Z���̃f�U�C��
Private Sub DefaultCellDesign(cell As Object)
    Call ChangeBackColor(cell, 255, 255, 255)
End Sub

Private Sub ChangeBackColor(cell As Object, r As Integer, g As Integer, b As Integer)
        ' �w�i�F��ݒ�
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
    
    '�����̗�ԍ�
    Dim piecesColumn As Integer
    piecesColumn = order.PiecesColumnIndex
    '���ʂ̗�ԍ�
    Dim qtyColumn As Integer
    qtyColumn = order.QtyColumnIndex
    '�P���̗�ԍ�
    Dim priceColumn As Integer
    priceColumn = order.PriceColumnIndex

    '�J�n�s
    Dim startRow As Long
    startRow = order.DataStartRowIndex
    '�I���s
    Dim endRow As Long
    endRow = order.DataEndRowIndex
    '�v�Z���̓��͗�
    Dim targetColumn As Integer
    targetColumn = order.AmountColumnIndex
    
    Dim row As Long
    '��
    Dim formula As String
    
    For row = startRow To endRow
        '���̎擾
        formula = GetAmountCalcFormula(row, piecesColumn, qtyColumn, priceColumn)
        '���̓���
        Cells(row, targetColumn).formula = formula
    Next row
End Sub
'�d�����z�̌v�Z����Ԃ�
Private Function GetAmountCalcFormula(row As Long, piecesColumn As Integer, qtyColumn As Integer, priceColumn As Integer) As String
    '����*����*�P��
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
