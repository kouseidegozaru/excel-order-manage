Attribute VB_Name = "ButtonEvents"

'���i�����V�[�g�Ɋւ��鏈��
'�m��{�^��(���i�R�[�h�̔��f)
Sub Decide()

    '��ʍX�V���Ȃ�
    Application.ScreenUpdating = False

    '�������̓V�[�g�A�N�Z�T���C���X�^���X��
    Dim order As New OrderSheetAccesser
    '���i�����V�[�g�A�N�Z�T���C���X�^���X��
    Dim search As New SearchSheetAccesser
    
    '�ύX�̃C�x���g�𖳎�
    IsIgnoreChangeEvents = True
    
    '�d�����鏤�i�R�[�h��r��
    Dim writeData As Collection
    Set writeData = FilterCollection(search.GetCheckedProductsCode, _
                                     order.productsCode)
                                     
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
                                       
    order.Worksheet.Activate
    
    '���i���\��
    DisplayProductsInfo target
    
    '�ύX�̃C�x���g��L����
    IsIgnoreChangeEvents = False
    
    '��ʍX�V�L����
    Application.ScreenUpdating = True
    
End Sub

'���i�����V�[�g�Ɋւ��鏈��
'�����t�H�[���X�V
Sub Update()

    '��ʍX�V���Ȃ�
    Application.ScreenUpdating = False
    
    '�������̓V�[�g�A�N�Z�T���C���X�^���X��
    Dim order As New OrderSheetAccesser
    '���i�����V�[�g�A�N�Z�T���C���X�^���X��
    Dim search As New SearchSheetAccesser
    
    '���i�����V�[�g�̃f�[�^���N���A
    search.Clear
    
    '�N�G�����s�N���X���C���X�^���X��
    Dim DataBaseAccesser As New DataBaseAccesser
    ' �f�[�^�x�[�X����Ώە���̏��i�����擾
    Dim rs As ADODB.Recordset
    Set rs = DataBaseAccesser.GetAllProducts(order.bumonCode)
    
    Dim rowIndex As Long
    Dim columnIndex As Integer
    
    rowIndex = search.DataStartRowIndex
    columnIndex = search.DataStartColumnIndex
    
    ' �f�[�^�̏�������
    Dim targetRange As Range
    Set targetRange = search.Worksheet.Cells(rowIndex, columnIndex)
    
    ' ���R�[�h�Z�b�g���ꊇ�œ\��t����
    targetRange.CopyFromRecordset rs
    
    ' �\��t�����f�[�^�̍s�����擾
    Dim pastedRows As Long
    pastedRows = DataBaseAccesser.GetAllProductsCount(order.bumonCode)
    
    ' �`�F�b�N�{�b�N�X�̒ǉ�
    For i = 0 To pastedRows - 1
        search.AddCheckBox rowIndex + i
    Next i
    
    search.Worksheet.Activate

    Application.ScreenUpdating = True
    
End Sub

'���i�����V�[�g�Ɋւ��鏈��
'�N���A�{�^��
Sub ClearCheckBoxes()
    '�`�F�b�N�{�b�N�X�̃N���A
    Dim search As New SearchSheetAccesser
    search.ClearCheckBoxes
End Sub

'���i�����V�[�g�̕\��
Sub search()
    Dim search As New SearchSheetAccesser
    '���i�����V�[�g���A�N�e�B�u��
    search.Worksheet.Activate
End Sub

'���M�{�^��
Sub Post()

    Dim result As VbMsgBoxResult
    '���킹���Ɛ��ʂ̃`�F�b�N
    If Not IsMatchQty Then
        result = MsgBox("���킹���ƈ�v���Ȃ����ʂ�����܂��B���M���܂���?", vbYesNo + vbQuestion, "�m�F")
        If result = vbNo Then
            End
        End If
    End If
    
    '�f�[�^��ۑ�
    SaveData
    
    MsgBox "�f�[�^�𑗐M���܂���"
End Sub
