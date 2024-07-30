Attribute VB_Name = "DisplayProducts"
'�������͂ɓ��͂��ꂽ���i�R�[�h���珤�i����\��

Sub DisplayProductsInfo(targetRng As range)

    Dim DataStorage As New DataBaseAccesser
    Dim BumonCD As Integer
    BumonCD = GetBumonCD
    Dim cell As range
    
    ' ���������̎w��
    Dim targetColumn As Integer
    targetColumn = OrderWb_ProductCodeColumnNumber
    
    ' �͈͓��̎w�肵����̊e�s������
    For Each cell In targetRng.Columns(targetColumn).Cells
        ' �󔒂łȂ��Z��������
        If cell.value <> "" Then
            If DataStorage.ExistsProducts(BumonCD, cell.value) Then
                '�w�i�𔒂�
                Call ChangeBackColor(cell, 255, 255, 255)
                
                Dim rs As ADODB.recordSet
                Set rs = DataStorage.GetProduct(BumonCD, cell.value)
                
                ' ���R�[�h�Z�b�g���Z���ɓ\��t����
                If Not rs.EOF Then
                    Dim startRow As Long
                    Dim startCol As Long
                    Dim i As Integer
                    
                    ' �\��t���J�n�Z�����w��i�Z���̍s�Ɠ����s�Ɂj
                    startRow = cell.row
                    startCol = cell.Column + 1 ' ���̃Z������1��E�ɓ\��t��
                    
                    ' ���R�[�h�Z�b�g�����[�N�V�[�g�ɓ\��t��
'                    rs.MoveFirst
                    Do Until rs.EOF
                        For i = 0 To rs.Fields.Count - 1
                            cell.Offset(0, i + 1).value = rs.Fields(i).value
                        Next i
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

