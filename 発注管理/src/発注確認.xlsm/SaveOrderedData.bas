Attribute VB_Name = "SaveOrderedData"
'�����ς݂Ƃ��ă`�F�b�N���ꂽ���i�R�[�h��ۑ�����
Sub SaveOrderedProductsCode()
    
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
