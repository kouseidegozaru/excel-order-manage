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
Sub LoadOrderedData()

End Sub
Sub test() '�e�X�g�p
    '�����m�F�V�[�g�ւ̃A�N�Z�T
    Dim load As New LoadSheetAccesser
    '�����ςݏ��i�R�[�h�ւ̃A�N�Z�T
    Dim ordered As New OrderedDataSheetAccesser
    ordered.InitStatus load.bumonCode, load.targetDate
    ordered.InitOpenWorkBook
    ordered.InitWorkSheet
    
    
    Dim aa As Variant
    Set aa = ordered.GetAllData_NoHead
    
    ordered.CloseWorkBook
End Sub
