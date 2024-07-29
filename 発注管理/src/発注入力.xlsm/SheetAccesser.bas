Attribute VB_Name = "SheetAccesser"
'�萔�������ňꊇ��`
'�萔���������߃J�v�Z����������
'�������̓V�[�g�Ɋւ���f�[�^�̓ǂݏ����͊�{�I�ɂ�����ʂ�

'�����f�[�^�t�H���_�p�X
Public Const OrderDataDirPath As String = "C:\Users\mfh077_user.MEFUREDMN\Desktop\excel-order-manage\�����Ǘ�\data"

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
Public Const OrderWb_ProductQtyColumnNumber As Integer = 10
Public Const OrderWb_ProductQtyColumn As String = "J"
Public Const OrderWb_ProductCodeRowNumber As Integer = 5
Public Const OrderWb_InputBumonCDRange As String = "A2"
Public Const OrderWb_OutputBumonNameRange As String = "B2"
Public Const OrderWb_InputUserCDRange As String = "C2"
Public Const OrderWb_OutputUserNameRange As String = "D2"
Public Const OrderWb_InputDateRange As String = "E2"

Public Const OrderWb_IgnoreStateRange As String = "F1:F1"

'�ۑ��f�[�^�ݒ�
Public Const DataWb_SheetName As String = "Sheet1"
Public Const DataWb_ProductCodeColumnNumber As Integer = 1
Public Const DataWb_ProductQtyColumnNumber As Integer = 2
Public Const DataWb_ProductCodeColumn As String = "A"
Public Const DataWb_ProductQtyColumn As String = "B"
Public Const DataWb_ProductCodeRowNumber As Integer = 1


'�������͂ɂ��鏤�i�R�[�h�͈̔�
Public Function OrderWb_InputProductsRange() As String
    
    OrderWb_InputProductsRange = OrderWb_ProductCodeColumn & _
                                 OrderWb_ProductCodeRowNumber & _
                                 ":" & _
                                 OrderWb_ProductCodeColumn & _
                                 OrderWb_LastProductsRow
End Function
'�������͂ɂ��鏤�i�̐��ʂ͈̔�
Public Function OrderWb_InpuQtyRange() As String
    OrderWb_InpuQtyRange = OrderWb_ProductQtyColumn & _
                            OrderWb_ProductCodeRowNumber & _
                            ":" & _
                            OrderWb_ProductQtyColumn & _
                            OrderWb_LastProductsRow
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

'���i��񂪋L�ڂ���Ă���Ō�̍s�ԍ�
Public Function OrderWb_LastProductsRow() As Long
    Dim lastRow As Long
    '�������͂ɏ��i�f�[�^���Ȃ��ꍇ�s�͈͂����炷(�����Ɣ͈͂Ƀw�b�_�[�s���܂܂�Ă��܂�)
    lastRow = OrderWb_NextProductsRow - 1
    If lastRow < OrderWb_ProductCodeRowNumber Then
        lastRow = OrderWb_ProductCodeRowNumber
    End If
    OrderWb_LastProductsRow = lastRow
End Function

'����R�[�h�̎擾
Function GetBumonCD() As Integer

    '�l�̎擾
    Dim value As Integer
    value = ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_InputBumonCDRange).value
    
    '����R�[�h�����݂��邩
    Dim DataStorage As New dataAccesser
    If DataStorage.ExistsBumon(value) Then
        GetBumonCD = value
    Else
        GetBumonCD = 0
        MsgBox ("����������R�[�h����͂��ĉ�����")
        End
    End If
    
End Function

'�S���҃R�[�h�̎擾
Function GetUserCD() As Integer
    
    '�l�̎擾
    Dim value As Integer
    value = ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_InputUserCDRange).value
    
     '�S���҃R�[�h�����݂��邩
    Dim DataStorage As New dataAccesser
    If DataStorage.ExistsUser(value) Then
        GetUserCD = value
    Else
        GetUserCD = 0
        MsgBox ("�������S���҃R�[�h����͂��ĉ�����")
        End
    End If
    
End Function

'�Ώۓ��t�̎擾
Function GetDate() As Date
    Dim value As Date
    value = ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_InputDateRange).value
    
    GetDate = value
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


'�S���Җ��̕\��
Sub SetUserName(name As String)
    ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_OutputUserNameRange).value = name
End Sub

'���喼�̕\��
Sub SetBumonName(name As String)
    ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_OutputBumonNameRange).value = name
End Sub

'�ۑ��t�@�C����
Function GetSaveFileName() As String
    GetSaveFileName = "b" & GetBumonCD & "-" & _
                  "u" & GetUserCD & "-" & _
                  "d" & Format(GetDate, "yyyymmdd") & "-" & _
                  ".xlsx"
End Function

'�ۑ��t�@�C���p�X
Function GetSaveFilePath() As String
    GetSaveFilePath = OrderDataDirPath & "\" & GetSaveFileName
End Function

'�f�[�^�G�N�Z���t�@�C���̓ǂݍ���
Public Function DataWb() As Workbook
    Set DataWb = Workbooks.Open(GetSaveFilePath)
End Function

'���i��񂪋L�ڂ���Ă���Ō�̍s�ԍ�
Public Function DataWb_LastProductsRow() As Long
    Dim wb As Workbook
    Set wb = DataWb
    
    Dim ws As Worksheet
    Dim columnNumber As Long
    
    ' �Ώۂ̃V�[�g��ݒ�
    Set ws = wb.Sheets(DataWb_SheetName)
    columnNumber = OrderWb_ProductCodeColumnNumber
    
    ' �Ώۂ̗�ōŌ�̍s���擾
    DataWb_LastProductsRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).row
End Function

'�������͂ɂ��鏤�i�R�[�h�͈̔�
Public Function DataWb_ProductsRange() As String
    
    DataWb_ProductsRange = DataWb_ProductCodeColumn & _
                            DataWb_ProductCodeRowNumber & _
                            ":" & _
                            DataWb_ProductCodeColumn & _
                            DataWb_LastProductsRow
    aa = 1
End Function
'�������͂ɂ��鏤�i�̐��ʂ͈̔�
Public Function DataWb_QtyRange() As String
    DataWb_QtyRange = DataWb_ProductQtyColumn & _
                      DataWb_ProductCodeRowNumber & _
                      ":" & _
                      DataWb_ProductQtyColumn & _
                      DataWb_LastProductsRow
End Function



