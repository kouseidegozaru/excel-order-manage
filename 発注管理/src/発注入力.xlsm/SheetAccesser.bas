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
Public Const OrderWb_InputBumonCDRange As String = "A2"
Public Const OrderWb_InputUserCDRange As String = "B2"
Public Const OrderWb_OutputUserCDRange As String = "C2"
Public Const OrderWb_InputDateRange As String = "D2"

'����R�[�h�̎擾
Function GetBumonCD() As Integer

    '�l�̎擾
    Dim Value As Integer
    Set Value = ThisWorkbook.Sheets(OrderWb_SheetName).Range(OrderWb_InputBumonCDRange).Value
    
    '����R�[�h�����݂��邩
    Dim DataStorage As New dataAccesser
    If DataStorage.ExistsBumon(Value) Then
        GetBumonCD = Value
    Else
        GetBumonCD = 0
        MsgBox ("����������R�[�h����͂��ĉ�����")
        End
    End If
    
End Function

'�S���҃R�[�h�̎擾
Function GetUserCD() As Integer
    
    '�l�̎擾
    Dim Value As Integer
    Set Value = ThisWorkbook.Sheets(OrderWb_SheetName).Range(OrderWb_InputUserCDRange).Value
    
     '�S���҃R�[�h�����݂��邩
    Dim DataStorage As New dataAccesser
    If DataStorage.ExistsUser(Value) Then
        GetUserCD = Value
    Else
        GetUserCD = 0
        MsgBox ("�������S���҃R�[�h����͂��ĉ�����")
        End
    End If
    
End Function

'�Ώۓ��t�̎擾
Function GetDate() As Date
    Dim Value As Date
    Set Value = ThisWorkbook.Sheets(OrderWb_SheetName).Range(OrderWb_InputDateRange).Value
    
    GetDate = Value
End Function


