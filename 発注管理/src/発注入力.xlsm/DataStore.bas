Attribute VB_Name = "DataStore"
'�f�[�^�̓ǂݏ���

'�f�[�^�̏�������
Sub SaveData()

    Application.ScreenUpdating = False

    Dim order As New OrderSheetAccesser
    Dim Data As New DataSheetAccesser
    Data.NewWorkbook
    Data.InitWorkSheet
    
    '���i�R�[�h�f�[�^
    Data.WriteProductsCode order.ProductsCode
    '���ʃf�[�^
    Data.WriteQty order.Qty
    
    '�ۑ�
    Data.Save
    Data.CloseWorkBook
    
    Application.ScreenUpdating = True
    
End Sub


Sub LoadData()
    
    Application.ScreenUpdating = False
    
    Dim order As New OrderSheetAccesser
    Dim Data As New DataSheetAccesser
    
    '�������͂̏��i����S�č폜
    order.ProductsCodeRange.EntireRow.Delete
    
    '�t�@�C�������݂��Ȃ��ꍇ�͏����I��
    If Dir(Data.SaveFilePath) = "" Then
        End
    End If
    
    Data.OpenWorkBook
    Data.InitWorkSheet
    
    '���i�R�[�h�����
    order.WriteProductsCode Data.ProductsCode
    '���i���\��
    DisplayProductsInfo order.ProductsCodeRange
    '���ʂ����
    order.WriteQty Data.Qty
    
    '�f�[�^���[�N�u�b�N�����
    Data.CloseWorkBook
    
    Application.ScreenUpdating = True
End Sub

