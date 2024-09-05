Attribute VB_Name = "DataStore"
'�����f�[�^�̕ۑ�
Sub SaveData()

    '��ʍX�V���Ȃ�
    Application.ScreenUpdating = False

    '�������̓V�[�g�A�N�Z�T���C���X�^���X��
    Dim order As New OrderSheetAccesser
    
    '�f�[�^�V�[�g�A�N�Z�T���C���X�^���X��
    Dim data As New DataSheetAccesser
    
    '����R�[�h�A�S���҃R�[�h�A��������ݒ�
    data.InitStatus order.bumonCode, _
                    order.userCode, _
                    order.targetDate
    data.InitNewWorkbook
    data.InitWorkSheet
    
    '���i�f�[�^��������
    data.WriteTableData order.GetAllData
    
    '�ۑ�
    data.Save
    data.CloseWorkBook
    
    '��ʍX�V�L����
    Application.ScreenUpdating = True
    
End Sub

'�f�[�^�ǂݍ���
Sub LoadData()
    
    '��ʍX�V���Ȃ�
    Application.ScreenUpdating = False
    
    '�������̓V�[�g�A�N�Z�T���C���X�^���X��
    Dim order As New OrderSheetAccesser
    '�f�[�^�V�[�g�A�N�Z�T���C���X�^���X��
    Dim data As New DataSheetAccesser
    
    '�������͂̏��i����S�č폜
    order.ProductsCodeRange.EntireRow.Delete
    
    '����R�[�h�A�S���҃R�[�h�A��������ݒ�
    data.InitStatus order.bumonCode, _
                    order.userCode, _
                    order.targetDate
        
    '�t�@�C�������݂��Ȃ��ꍇ�͏����I��
    If Dir(data.SaveFilePath) = "" Then
        End
    End If
    
    data.InitOpenWorkBook
    data.InitWorkSheet
    
    '���i�������
    order.WriteAllData data.GetAllData_NoHead
    
    '�f�[�^���[�N�u�b�N�����
    data.CloseWorkBook
    
    '�d������z�v�Z���̓���
    ApplyAmountCalcFormulaToRange
    
    '��ʍX�V�L����
    Application.ScreenUpdating = True
    
End Sub

