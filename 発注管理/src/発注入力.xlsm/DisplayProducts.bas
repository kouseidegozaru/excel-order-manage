Attribute VB_Name = "DisplayProducts"
Sub DisplayProductsInfo(targetRng As range)

    Dim DataStorage As New dataAccesser
    Dim BumonCD As Integer
    BumonCD = GetBumonCD
    Dim cell As range
    
    ' ���������̎w��
    Dim targetColumn As Integer
    targetColumn = OrderWb_ProductCodeColumnNumber
    
    ' �͈͓��̎w�肵����̊e�s������
    For Each cell In targetRng.Columns(targetColumn).Cells
        ' �󔒂łȂ��Z��������
        If cell.Value <> "" Then
            If DataStorage.ExistsProducts(BumonCD, cell.Value) Then
                Dim rs As ADODB.recordSet
                Set rs = DataStorage.GetProduct(BumonCD, cell.Value)
                
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
                            cell.Offset(0, i + 1).Value = rs.Fields(i).Value
                        Next i
                        rs.MoveNext
                    Loop
                End If
                
                ' ���R�[�h�Z�b�g�����
                rs.Close
                Set rs = Nothing
            End If
        End If
    Next cell
    
End Sub

