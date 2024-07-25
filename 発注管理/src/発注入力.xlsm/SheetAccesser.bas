Attribute VB_Name = "SheetAccesser"
'�萔�������ňꊇ��`
'�萔���������߃J�v�Z����������

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
Public Const OrderWb_OutputBumonNameRange As String = "B2"
Public Const OrderWb_InputUserCDRange As String = "C2"
Public Const OrderWb_OutputUserNameRange As String = "D2"
Public Const OrderWb_InputDateRange As String = "E2"
Public Const OrderWb_InputProductsRange As String = "A5:A5000"

'����R�[�h�̎擾
Function GetBumonCD() As Integer

    '�l�̎擾
    Dim Value As Integer
    Value = ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_InputBumonCDRange).Value
    
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
    Value = ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_InputUserCDRange).Value
    
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
    Value = ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_InputDateRange).Value
    
    GetDate = Value
End Function

'�S���Җ��̕\��
Sub SetUserName(name As String)
    ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_OutputUserNameRange).Value = name
End Sub

'���喼�̕\��
Sub SetBumonName(name As String)
    ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_OutputBumonNameRange).Value = name
End Sub


