Attribute VB_Name = "SheetAccesser"
'���i�����t�H�[���̃Z���Ԓn�̐ݒ�
Public Const SearchWb_SheetName As String = "���i�}�X�^�[�t�H�[�}�b�g"
Public Const SearchWb_StateColumnNumber As Integer = 19
Public Const SearchWb_CheckBoxColumnNumber As Integer = 1
Public Const SearchWb_DataStartColumnNumber As Integer = 2
Public Const SearchWb_DataStartRowNumber As Integer = 2
Public Const SearchWb_ProductCodeColumnNumber As Integer = 3

'�������̓t�H�[���̃Z���Ԓn�̐ݒ�
Public Const OrderWb_SheetName As String = "��������"
Public Const OrderWb_ProductCodeColumnNumber As Integer = 1


Function GetBumonCD() As Integer
    GetBumonCD = 40
End Function

Function GetUserCD() As Integer
    GetUserCD = 70
End Function

Function GetDate() As Date
    Dim d As Date
    d = "2024/7/26"
    GetDate = d
End Function


