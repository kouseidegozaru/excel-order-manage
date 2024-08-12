Attribute VB_Name = "ButtonEvents"

'���i�����V�[�g�Ɋւ��鏈��
'�m��{�^��(���i�R�[�h�̔��f)
Sub Decide()

    Dim order As New OrderSheetAccesser
    Dim search As New SearchSheetAccesser
    
    '�ύX�̃C�x���g�𖳎�
    IsIgnoreChangeEvents = True
    
    '�d�����鏤�i�R�[�h��r��
    Dim writeData As Collection
    Set writeData = FilterCollection(search.GetCheckedProductsCode, _
                                     order.ProductsCode)
                                     
    Dim startRowIndex As Long
    Dim lastRowIndex As Long
    
    startRowIndex = order.DataNextRowIndex
    lastRowIndex = startRowIndex
    
    '�������͂ɏ��i�R�[�h����
    For i = 1 To writeData.count
        order.Cells(lastRowIndex, order.ProductCodeColumnIndex) = writeData(i)
        lastRowIndex = lastRowIndex + 1
    Next i
    
    '�������͂ɏ��i�R�[�h����͂����͈�
    Dim target As Range
    Set target = order.Worksheet.Range(IndexToLetter(order.ProductCodeColumnIndex) & startRowIndex & _
                                       ":" & _
                                       IndexToLetter(order.ProductCodeColumnIndex) & lastRowIndex)
    
    '���i���\��
    DisplayProductsInfo target
    
    IsIgnoreChangeEvents = False
    
    order.Worksheet.Activate
    
End Sub

'�������̓V�[�g�Ɋւ��鏈��
'�����t�H�[���X�V
Sub Update()

    Application.ScreenUpdating = False
    
    Dim order As New OrderSheetAccesser
    Dim search As New SearchSheetAccesser
    
    search.Clear
    
    Dim DataBaseAccesser As New DataBaseAccesser
    Dim rs As ADODB.recordSet
    ' �f�[�^�x�[�X���烌�R�[�h�Z�b�g���擾
    Set rs = DataBaseAccesser.GetAllProducts(order.bumonCode)
    
    Dim rowIndex As Long
    Dim columnIndex As Integer
    
    rowIndex = search.DataStartRowIndex
    columnIndex = search.DataStartColumnIndex
    
    ' �f�[�^�̏�������
    rs.MoveFirst
    Do While Not rs.EOF
        
        For i = 0 To rs.Fields.count - 1
            search.Cells(rowIndex, i + columnIndex) = rs.Fields(i).value
        Next i
        
        ' �`�F�b�N�{�b�N�X�̒ǉ�
        search.AddCheckBox rowIndex
        
        rowIndex = rowIndex + 1
        rs.MoveNext
    Loop
    
    search.Worksheet.Activate

    Application.ScreenUpdating = True
    
End Sub

'�`�F�b�N�{�b�N�X�̃N���A
Sub ClearCheckBoxes()
    Dim search As New SearchSheetAccesser
    search.ClearCheckBoxes
End Sub

'���i�����V�[�g�̕\��
Sub search()
    Dim search As New SearchSheetAccesser
    search.Worksheet.Activate
End Sub

'���M
Sub Post()

    Dim result As VbMsgBoxResult
    '���킹���Ɛ��ʂ̃`�F�b�N
    If Not IsMatchQty Then
        result = MsgBox("���킹���ƈ�v���Ȃ����ʂ�����܂��B���M���܂���?", vbYesNo + vbQuestion, "�m�F")
        If result = vbNo Then
            End
        End If
    End If
        
    SaveData
    MsgBox "�f�[�^�𑗐M���܂���"
End Sub
