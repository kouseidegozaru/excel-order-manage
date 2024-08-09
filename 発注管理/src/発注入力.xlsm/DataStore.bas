Attribute VB_Name = "DataStore"
'�f�[�^�̓ǂݏ���

'�f�[�^�̏�������
Sub SaveData()

    Application.ScreenUpdating = False

    Dim order As New OrderSheetAccesser
    Dim data As New DataSheetAccesser
    data.NewWorkbook
    data.InitWorkSheet
    
    '���i�R�[�h�f�[�^
    data.WriteProductsCode order.ProductsCode
    '���ʃf�[�^
    data.WriteQty order.Qty
    
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
    
    data.OpenWorkBook
    data.InitWorkSheet
    
    '���i�R�[�h�����
    order.WriteProductsCode data.ProductsCode
    '���i���\��
    DisplayProductsInfo order.ProductsCodeRange
    '���ʂ����
    order.WriteQty data.Qty
    
    '�f�[�^���[�N�u�b�N�����
    data.CloseWorkBook
    
    Application.ScreenUpdating = True
End Sub

