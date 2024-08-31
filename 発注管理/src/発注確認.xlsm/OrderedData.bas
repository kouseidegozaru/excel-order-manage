Attribute VB_Name = "OrderedData"
'�����ς݂Ƃ��ă`�F�b�N���ꂽ���i�R�[�h��ۑ�����
Sub SaveOrderedData()
    
    Application.ScreenUpdating = False
    
    '�����m�F�V�[�g�ւ̃A�N�Z�T
    Dim load As New LoadSheetAccesser
    
    '�����ςݏ��i�R�[�h�ւ̃A�N�Z�T
    Dim ordered As New OrderedDataSheetAccesser
    ordered.InitNewWorkbook
    ordered.InitWorkSheet
    ordered.InitStatus load.bumonCode, load.targetDate
    
    '�`�F�b�N���ꂽ���i�R�[�h�����
    ordered.WriteProductsCode load.GetCheckedProductsCode
    
    '�ۑ����ĕ���
    ordered.Save
    ordered.CloseWorkBook
    
    Application.ScreenUpdating = True
End Sub

'�����ς݂̏��i�̃`�F�b�N�{�b�N�X���I���ɂ���
Sub LoadOrderedData()
        
    Application.ScreenUpdating = False
    
    '�����m�F�V�[�g�ւ̃A�N�Z�T
    Dim load As New LoadSheetAccesser
    
    '�����ςݏ��i�R�[�h�ւ̃A�N�Z�T
    Dim ordered As New OrderedDataSheetAccesser
    ordered.InitStatus load.bumonCode, load.targetDate
    ordered.InitOpenWorkBook
    ordered.InitWorkSheet
    
    '�����ς݃f�[�^���Ȃ��ꍇ�͏I��
    If Dir(ordered.SaveFilePath) = "" Then
        Exit Sub
    End If
    
    '�����ς݂̏��i�R�[�h�̃R���N�V����
    Dim orderedProductsCodes As Collection
    Set orderedProductsCodes = ordered.GetAllData_NoHead
    
    '�����ς݂̏��i�R�[�h����
    For Each orderedProductsCode In orderedProductsCodes
        load.OrderedIsTrue orderedProductsCode(1)
    Next
    
    ordered.CloseWorkBook
    
End Sub
