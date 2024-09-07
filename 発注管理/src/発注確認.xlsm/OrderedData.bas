Attribute VB_Name = "OrderedData"
'�����ς݂Ƃ��ă`�F�b�N���ꂽ���i�R�[�h��ۑ�����
Sub SaveOrderedData()
    
    '��ʂ̍X�V�𖳌���
    Application.ScreenUpdating = False
    
    '�����m�F�V�[�g�ւ̃A�N�Z�T
    Dim load As New LoadSheetAccesser
    
    '�����ςݏ��i�R�[�h�ւ̃A�N�Z�T
    Dim ordered As New OrderedDataSheetAccesser
    '����R�[�h�Ɣ�������ݒ�
    ordered.InitStatus load.bumonCode, load.targetDate
    ordered.InitNewWorkbook
    ordered.InitWorkSheet
    
    '�`�F�b�N���ꂽ���i�R�[�h�����
    ordered.WriteProductsCode load.GetCheckedProductsCode
    
    '�ۑ����ĕ���
    ordered.Save
    ordered.CloseWorkBook
    
    '��ʂ̍X�V��L����
    Application.ScreenUpdating = True
End Sub

'�����ς݂̏��i�̃`�F�b�N�{�b�N�X���I���ɂ���
Sub LoadOrderedData()
    
    '��ʂ̍X�V�𖳌���
    Application.ScreenUpdating = False
    
    '�����m�F�V�[�g�ւ̃A�N�Z�T
    Dim load As New LoadSheetAccesser
    
    '�����ςݏ��i�R�[�h�ւ̃A�N�Z�T
    Dim ordered As New OrderedDataSheetAccesser
    ordered.InitStatus load.bumonCode, load.targetDate
    
    '�����ς݃f�[�^���Ȃ��ꍇ�͏I��
    If Dir(ordered.SaveFilePath) = "" Then
        Exit Sub
    End If
    
    ordered.InitOpenWorkBook
    ordered.InitWorkSheet
    
    '�����ς݂̏��i�R�[�h���R���N�V�����Ŏ擾
    Dim orderedProductsCodes As Collection
    '�񎟌��̃R���N�V�����Ŏ擾
    Set orderedProductsCodes = ordered.GetAllData_NoHead
    
    '�����ς݂̏��i�R�[�h����
    For Each orderedProductsCode In orderedProductsCodes
        '�Ώۂ̏��i�R�[�h�̃`�F�b�N�{�b�N�X���I���ɂ���
        load.OrderedIsTrue orderedProductsCode(1) '�񎟌��Ȃ̂ŃC���f�b�N�X���w��
    Next
    
    '�����ς݃f�[�^�����
    ordered.CloseWorkBook
    
    '��ʂ̍X�V��L����
    Application.ScreenUpdating = True
    
End Sub
