Attribute VB_Name = "DataStore"
'�f�[�^�̓ǂݏ���

'�f�[�^�̏�������
Sub SaveData()

    Application.ScreenUpdating = False

    Dim order As New OrderSheetAccesser
    
    Dim data As New DataSheetAccesser
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
    
    Application.ScreenUpdating = True
    
End Sub


Sub LoadData()
    
    Application.ScreenUpdating = False
    
    Dim order As New OrderSheetAccesser
    Dim data As New DataSheetAccesser
    
    '�������͂̏��i����S�č폜
    order.ProductsCodeRange.EntireRow.Delete
    
    '�t�@�C�������݂��Ȃ��ꍇ�͏����I��
    If Dir(data.SaveFilePath) = "" Then
        End
    End If
    
    data.InitStatus order.bumonCode, _
                    order.userCode, _
                    order.targetDate
    data.InitOpenWorkBook
    data.InitWorkSheet
    
    '���i�������
    order.WriteAllData data.GetAllData_NoHead
    
    '�f�[�^���[�N�u�b�N�����
    data.CloseWorkBook
    
    Application.ScreenUpdating = True
End Sub

