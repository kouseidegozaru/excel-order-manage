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
Public Const OrderWb_ProductCodeColumn As String = "A"
Public Const OrderWb_ProductCodeRowNumber As Integer = 5
Public Const OrderWb_InputBumonCDRange As String = "A2"
Public Const OrderWb_OutputBumonNameRange As String = "B2"
Public Const OrderWb_InputUserCDRange As String = "C2"
Public Const OrderWb_OutputUserNameRange As String = "D2"
Public Const OrderWb_InputDateRange As String = "E2"
Public Function OrderWb_InputProductsRange() As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim columnNumber As Long
    
    '�������͂ɏ��i�f�[�^���Ȃ��ꍇ�s�͈͂����炷(�����Ɣ͈͂Ƀw�b�_�[�s���܂܂�Ă��܂�)
    lastRow = OrderWb_NextProductsRow - 1
    If lastRow < OrderWb_ProductCodeRowNumber Then
        lastRow = OrderWb_ProductCodeRowNumber
    End If
    
    ' �Ώۂ̃V�[�g��ݒ�
    Set ws = ThisWorkbook.Sheets(OrderWb_SheetName)
    columnNumber = OrderWb_ProductCodeColumnNumber
    
    OrderWb_InputProductsRange = OrderWb_ProductCodeColumn & _
                                 OrderWb_ProductCodeRowNumber & _
                                 ":" & _
                                 OrderWb_ProductCodeColumn & _
                                 lastRow
End Function
'���ɓ��͂��鏤�i��񂪋󔒂̍s�ԍ�
Public Function OrderWb_NextProductsRow() As Long
    Dim ws As Worksheet
    Dim columnNumber As Long
    
    ' �Ώۂ̃V�[�g��ݒ�
    Set ws = ThisWorkbook.Sheets(OrderWb_SheetName)
    columnNumber = OrderWb_ProductCodeColumnNumber
    
    ' �Ώۂ̗�ōŌ�̍s���擾
    OrderWb_NextProductsRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).row + 1
End Function

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
'�������͂Ɋ��ɓ��͂���Ă��鏤�i�R�[�h�̎擾
Function GetProductsCD() As Collection
    ' �Ώۂ̃V�[�g��ݒ�
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(OrderWb_SheetName)
    
    Dim rng As range
    Set rng = ws.range(OrderWb_InputProductsRange)
    
    Set GetProductsCD = GetRangeValue(rng)
End Function

Private Function GetRangeValue(rng As range) As Collection
    Dim cell As range
    Dim col As New Collection
    
    ' �͈͓��̊e�Z�������[�v
    For Each cell In rng
        ' �󔒂łȂ��Z���̏ꍇ�ACollection�ɒǉ�
        If cell.Value <> "" Then
            col.Add cell.Value
        End If
    Next cell
    
    Set GetRangeValue = col
    
End Function

'�S���Җ��̕\��
Sub SetUserName(name As String)
    ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_OutputUserNameRange).Value = name
End Sub

'���喼�̕\��
Sub SetBumonName(name As String)
    ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_OutputBumonNameRange).Value = name
End Sub


