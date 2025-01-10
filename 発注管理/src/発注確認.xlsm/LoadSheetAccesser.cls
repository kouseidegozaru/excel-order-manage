VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoadSheetAccesser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�����m�F�V�[�g�ւ̃A�N�Z�T

'SheetAccesser�N���X���^���I�Ɍp��
Private super As SheetAccesser

Private WorkSheetName_ As String

Private ProductCodeColumnIndex_ As Integer
Private PiecesColumnIndex_ As Integer
Private QtyColumnIndex_ As Integer
Private PriceColumnIndex_ As Integer
Private AmountColumnIndex_ As Integer

Private CheckStateColumnIndex_ As Integer
Private CheckBoxColumnIndex_ As Integer

Private InputBumonCodeRange_ As String
Private OutputBumonNameRange_ As String
Private InputTargetDateRange_ As String

Private FilePropertyStartColumnIndex_ As Integer
Private FilePropertyUserNameRowIndex_ As Long
Private FilePropertyCompleteRowIndex_ As Long
Private FilePropertyUpdateTimeRowIndex_ As Long

Private CompleteString_ As String

Private Sub Class_Initialize()

    '�V�[�g��
    WorkSheetName_ = "�����m�F"
    
    '�����f�[�^����͍ς݂̕\�����郆�[�U�[�t���O
    CompleteString_ = "��"
    
    '���i�R�[�h�̗�ԍ�
    ProductCodeColumnIndex_ = 3
    '�����̗�ԍ�
    PiecesColumnIndex_ = 11
    '���ʂ̗�ԍ�
    QtyColumnIndex_ = 13
    '�P���̗�ԍ�
    PriceColumnIndex_ = 15
    '�d�����z���v�̗�ԍ�
    AmountColumnIndex_ = 16
    
    '�`�F�b�N�{�b�N�X�̏�Ԃ̗�ԍ�
    CheckStateColumnIndex_ = 20
    '�`�F�b�N�{�b�N�X�̗�ԍ�
    CheckBoxColumnIndex_ = 1
    
    '�t�@�C�����
    '�J�n��
    FilePropertyStartColumnIndex_ = 2
    '�S���Җ��\���s
    FilePropertyUserNameRowIndex_ = 4
    '�������̓t���O�\���s
    FilePropertyCompleteRowIndex_ = 5
    '�ŏI�X�V���\���s
    FilePropertyUpdateTimeRowIndex_ = 6
    
    '����R�[�h���̓Z��
    InputBumonCodeRange_ = "A2"
    '���喼�\���Z��
    OutputBumonNameRange_ = "B2"
    '�������\���Z��
    InputTargetDateRange_ = "C2"
    
    '�^�����ۃN���X���C���X�^���X��
    Set super = New SheetAccesser
    '�f�[�^�̃w�b�_�[�s�A�J�n�s�A�J�n����w��
    super.InitStatus DataHeaderRow:=8, _
                     DataStartRow:=9, _
                     DataStartColumn:=3
    
    InitWorkbook
    InitWorkSheet
End Sub

'���[�N�u�b�N��ݒ�
Private Sub InitWorkbook()
    super.InitWorkbook ThisWorkbook
End Sub

'���[�N�V�[�g��ݒ�
Private Sub InitWorkSheet()
    super.InitWorkSheet super.Workbook.Sheets(WorkSheetName_)
End Sub

'���i�R�[�h�̗�ԍ�
Property Get ProductCodeColumnIndex() As Integer
    ProductCodeColumnIndex = ProductCodeColumnIndex_
End Property
'�����̗�ԍ�
Property Get PiecesColumnIndex() As Integer
    PiecesColumnIndex = PiecesColumnIndex_
End Property
'���ʂ̗�ԍ�
Property Get QtyColumnIndex() As Integer
    QtyColumnIndex = QtyColumnIndex_
End Property
'�P���̗�ԍ�
Property Get PriceColumnIndex() As Integer
    PriceColumnIndex = PriceColumnIndex_
End Property
'�d�����z�̗�ԍ�
Property Get AmountColumnIndex() As Integer
    AmountColumnIndex = AmountColumnIndex_
End Property

'�`�F�b�N�{�b�N�X�̏�Ԃ̗�ԍ�
Property Get CheckStateColumnIndex() As Integer
    CheckStateColumnIndex = CheckStateColumnIndex_
End Property
'�`�F�b�N�{�b�N�X�̗�ԍ�
Property Get CheckBoxColumnIndex() As Integer
    CheckBoxColumnIndex = CheckBoxColumnIndex_
End Property
'�t�@�C�����̊J�n��
Property Get FilePropertyStartColumnIndex() As Integer
    FilePropertyStartColumnIndex = FilePropertyStartColumnIndex_
End Property

'���ɓ��͂���t�@�C����񂪋󔒂̗�ԍ�
Public Function FilePropertyNextColumnIndex() As Integer
    FilePropertyNextColumnIndex = Me.Worksheet.Cells(FilePropertyUserNameRowIndex_, Me.Worksheet.Columns.Count).End(xlToLeft).Column + 1
End Function

'�t�@�C����񂪋L�ڂ���Ă���Ō�̗�ԍ�
Property Get FilePropertyEndColumnIndex() As Long
    Dim lastColumn As Long
    '�t�@�C�����Ƀf�[�^���Ȃ��ꍇ��͈͂����炷(�����Ɣ͈͂Ƀw�b�_�[����܂܂�Ă��܂�)
    lastColumn = Me.FilePropertyNextColumnIndex - 1
    If lastColumn < Me.FilePropertyStartColumnIndex Then
        lastColumn = Me.FilePropertyStartColumnIndex
    End If
    FilePropertyEndColumnIndex = lastColumn
End Property

'�t�@�C�����͈̔�
Property Get FilePropertyRange() As Range
    Set FilePropertyRange = super.GetRange(FilePropertyUserNameRowIndex_, _
                                           Me.FilePropertyStartColumnIndex, _
                                           FilePropertyUpdateTimeRowIndex_, _
                                           Me.FilePropertyEndColumnIndex)
End Property

'�t�@�C�������N���A
Public Sub ClearFileProperty()
    Me.FilePropertyRange.Delete Shift:=xlToLeft
End Sub

'�t�@�C������ǉ�
Public Sub AddFileProperty(userName As String, IsCompleted As Boolean, updatedTime As Date)
    
    '���͑Ώۗ�
    Dim lastColumn As Integer
    lastColumn = FilePropertyNextColumnIndex
    
    '�������̓t���O������
    Dim writeCompleteString As String
    writeCompleteString = ""
    
    '�������͂������S���҂̏ꍇ
    If IsCompleted Then
        writeCompleteString = CompleteString_
        
        '�X�V������������
        Me.Worksheet.Cells(FilePropertyUpdateTimeRowIndex_, lastColumn) = updatedTime
    End If
    
    '�S���Җ��\��
    Me.Worksheet.Cells(FilePropertyUserNameRowIndex_, lastColumn) = userName
    '�������̓t���O�\��
    Me.Worksheet.Cells(FilePropertyCompleteRowIndex_, lastColumn) = writeCompleteString
End Sub

'�S�Ẵ`�F�b�N�{�b�N�X�̍폜
Sub DeleteAllCheckboxes()
    Dim ws As Worksheet
    Dim cb As OLEObject
    Dim shp As Shape

    Set ws = Me.Worksheet

    ' Form Controls�̃`�F�b�N�{�b�N�X���폜
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            '�R���g���[�����`�F�b�N�{�b�N�X�̏ꍇ
            If shp.FormControlType = xlCheckBox Then
                '�폜
                shp.Delete
            End If
        End If
    Next shp

End Sub

'�\������Ă��锭���f�[�^�̍폜
Public Sub ClearData()
    Dim clearRange As Range
    Set clearRange = Me.ProductsCodeRange
    '�s�����݂���ꍇ
    If WorksheetFunction.CountA(clearRange) > 0 Then
        '�`�F�b�N�{�b�N�X�̍폜
        Me.DeleteAllCheckboxes
        '���i���̍폜
        clearRange.EntireRow.Delete
    End If
End Sub

'�������͂ɂ��鏤�i�R�[�h�͈̔�
Property Get ProductsCodeRange() As Range
    Set ProductsCodeRange = super.GetRange(Me.DataStartRowIndex, _
                                           Me.ProductCodeColumnIndex, _
                                           Me.DataEndRowIndex, _
                                           Me.ProductCodeColumnIndex)
End Property

'�������͂ɂ���`�F�b�N�{�b�N�X��Ԃ͈̔�
Property Get CheckStateRange() As Range
    Set CheckStateRange = super.GetRange(Me.DataStartRowIndex, _
                                           Me.CheckStateColumnIndex, _
                                           Me.DataEndRowIndex, _
                                           Me.CheckStateColumnIndex)
End Property

'����R�[�h�̎擾
Property Get bumonCode() As Integer
    '�l�̎擾
    Dim value As Variant
    value = Me.Worksheet.Range(InputBumonCodeRange_).value
    
    '�o���f�[�V�����`�F�b�N���s�����s������I��
    CheckExistsBumon value
    
    bumonCode = CInt(value)

End Property

'�Ώۓ��t�̎擾
Property Get targetDate() As Date
    '�l�̎擾
    Dim value As Variant
    value = Me.Worksheet.Range(InputTargetDateRange_).value
    
    '�o���f�[�V�����`�F�b�N���s�����s������I��
    CheckDateFormat value
    
    targetDate = CDate(value)
End Property

'����R�[�h�̓��̓Z��
Property Get BumonCodeRange() As Range
    Set BumonCodeRange = Me.Worksheet.Range(InputBumonCodeRange_)
End Property
'�������̓��̓Z��
Property Get TargetDateRange() As Range
    Set TargetDateRange = Me.Worksheet.Range(InputTargetDateRange_)
End Property

'���喼�̕\��
Property Let BumonName(name As String)
    Me.Worksheet.Range(OutputBumonNameRange_).value = name
End Property

'�V�[�g�ɓ��͂���Ă���e���ڂ��ƂɃO���[�v�����Đ��ʂƎd�����z�����Z
Public Function AllGroupData() As ADODB.Recordset
    Dim dataAccessor As New SqlExecutor
    Dim rs As ADODB.Recordset
    Dim FilePath As String
    Dim sheetName As String
    Dim rangeName As String
    Dim query As String
    
    ' Excel���
    FilePath = Me.Workbook.FullName
    sheetName = Me.Worksheet.name
    rangeName = Replace(Me.AllDataRange.Address(False, False), "$", "")
    
    
    ' ���s����N�G��
    query = "SELECT �iCD, ���׎��ԑ�, �d���於, �����於, ���[�J�[��, �i��, �K�i, �v��, ����, ���킹��, SUM(����) AS ���v����, �׎p, �d���P��, SUM(�d�����z) AS ���v�d�����z, LT, ���i���l, ���[�J�[�ǂ� " & _
           "FROM [" & sheetName & "$" & rangeName & "] " & _
           "GROUP BY �iCD, ���׎��ԑ�, �d���於, �����於, ���[�J�[��, �i��, �K�i, �v��, ����, ���킹��, �׎p, �d���P��, LT, ���i���l, ���[�J�[�ǂ�"
    
    ' Excel�t�@�C���ɃN�G�������s
    Set rs = dataAccessor.ExecuteExcelQuery(FilePath, query)
    
    Set AllGroupData = rs
End Function

'�Ώۂ̍s�̃`�F�b�N�{�b�N�X�̃X�e�[�^�X�ύX
Property Let CheckBoxState(rowIndex As Long, value As Boolean)
    Me.Cells(rowIndex, Me.CheckStateColumnIndex) = value
End Property

'�Ώۂ̏��i�̃`�F�b�N�{�b�N�X���I���ɂ���
Public Sub OrderedIsTrue(targetProductsCode As Variant)
    
    '�f�[�^�̏I���s�ƊJ�n�s
    Dim startRowIndex As Long
    startRowIndex = Me.DataStartRowIndex
    Dim endRowIndex As Long
    endRowIndex = Me.DataEndRowIndex
    
    '���ݕ\������Ă��鏤�i�R�[�h
    Dim displayProductsCode As Variant
    
    For i = startRowIndex To endRowIndex
        '���i�R�[�h
        displayProductsCode = Me.Worksheet.Cells(i, Me.ProductCodeColumnIndex).value
        '�Ώۂ̏��i�R�[�h�̏ꍇ�`�F�b�N�{�b�N�X���I��
        If targetProductsCode = displayProductsCode Then
            Me.CheckBoxState(CInt(i)) = True
        End If
    Next i
    
End Sub

'�s���w�肵�ă`�F�b�N�{�b�N�X��ǉ�
Public Function AddCheckBox(rowIndex As Long) As Shape
    Dim columnIndex As Integer
    Dim stateColumnIndex As Integer
    
    '�`�F�b�N�{�b�N�X�̗�
    columnIndex = Me.CheckBoxColumnIndex
    '�`�F�b�N�{�b�N�X�̏�Ԃ̗�
    stateColumnIndex = Me.CheckStateColumnIndex
    
    Dim ws As Worksheet
    Set ws = Me.Worksheet
    
    '�Z���̍�����ύX(�`�F�b�N�{�b�N�X�̍����𒲐߂��邽��)
    ws.Cells(rowIndex, columnIndex).RowHeight = 21
    
    ' �t�H�[���R���g���[���̃`�F�b�N�{�b�N�X��ǉ�
    Dim cb As Shape
    Set cb = ws.Shapes.AddFormControl(xlCheckBox, _
                                      Left:=ws.Cells(rowIndex, columnIndex).Left, _
                                      Top:=ws.Cells(rowIndex, columnIndex).Top, _
                                      Width:=ws.Cells(rowIndex, columnIndex).Width, _
                                      Height:=ws.Cells(rowIndex, columnIndex).Height)
    
    ' �`�F�b�N�{�b�N�X�̃L���v�V��������ɐݒ�
    cb.TextFrame.Characters.Text = ""
    
    ' �`�F�b�N�{�b�N�X�̃����N����Z����ݒ�
    cb.ControlFormat.LinkedCell = ws.Cells(rowIndex, stateColumnIndex).Address
    
    '�ǉ������`�F�b�N�{�b�N�X�I�u�W�F�N�g��ԋp
    Set AddCheckBox = cb
End Function

'�`�F�b�N���ꂽ���i�R�[�h�̎擾
Public Function GetCheckedProductsCode() As Collection
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Me.Workbook
    Set ws = Me.Worksheet
    
    '���i�R�[�h��
    Dim productsColumn As Integer
    '�`�F�b�N�{�b�N�X�̏�Ԃ̗�
    Dim checkedStateColumn As Integer
    
    '�`�F�b�N���ꂽ���i�R�[�h�̊i�[�p
    Dim checkedProductsCode As New Collection
    Dim i As Long
    
    productsColumn = Me.ProductCodeColumnIndex
    checkedStateColumn = Me.CheckStateColumnIndex
    
    '���i�R�[�h�����݂���s�ԍ��Ń��[�v
    For i = 1 To ws.Cells(ws.Rows.Count, productsColumn).End(xlUp).row
        '�`�F�b�N����Ă����ꍇ
        If ws.Cells(i, checkedStateColumn).value = True Then
            '���i�R�[�h��ǉ�
            checkedProductsCode.Add ws.Cells(i, productsColumn).value
        End If
    Next i
    
    Set GetCheckedProductsCode = checkedProductsCode
End Function

'�����t��������ݒ�
Sub ApplyConditionalFormatting()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim startRow As Long
    Dim startCol As Integer
    Dim endCol As Integer
    Dim checkCol As Integer
    Dim rng As Range
    Dim formatCondition As formatCondition
    
    ' �Ώۂ̃V�[�g��ݒ�
    Set ws = Me.Worksheet
    
    ' ��̎w��
    startCol = Me.DataStartColumnIndex
    endCol = Me.DataEndColumnIndex
    checkCol = Me.CheckStateColumnIndex
    
    ' �s�̎w��
    lastRow = Me.DataEndRowIndex
    startRow = Me.DataStartRowIndex
    
    ' �e�s���Ƃɏ����t��������ݒ�
    For i = startRow To lastRow
        ' S��True�̏ꍇ�AA����S��͈̔͂ɏ����t��������ݒ�
        Set rng = ws.Range(ws.Cells(i, startCol), ws.Cells(i, endCol))
        
        ' �����t���������N���A�i�����̏������폜�j
        rng.FormatConditions.Delete
        
        ' �����t��������ǉ�
        Set formatCondition = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & ws.Cells(i, checkCol).Address & "=TRUE")
        
        ' �w�i�F��ݒ�i���F�j
        formatCondition.Interior.Color = RGB(255, 255, 0) ' ���F
    Next i
End Sub

'�d�����z�̌v�Z����ݒ�
Sub ApplyAmountCalc()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim startRow As Long
    Dim piecesCol As Integer
    Dim qtyCol As Integer
    Dim priceCol As Integer
    Dim amountCol As Integer
    Dim rng As Range
    Dim formatCondition As formatCondition
    
    ' �Ώۂ̃V�[�g��ݒ�
    Set ws = Me.Worksheet
    
    ' ��̎w��
    checkCol = Me.CheckStateColumnIndex
    piecesCol = Me.PiecesColumnIndex
    qtyCol = Me.QtyColumnIndex
    priceCol = Me.PriceColumnIndex
    amountCol = Me.AmountColumnIndex
    
    ' �s�̎w��
    lastRow = Me.DataEndRowIndex
    startRow = Me.DataStartRowIndex
    
    ' �e�s���Ƃɂ�ݒ�
    For i = startRow To lastRow
        
        ' �d�����z�̌v�Z���̐ݒ�
        Set rng = ws.Range(ws.Cells(i, amountCol), ws.Cells(i, amountCol))
        rng.Formula = "=IFERROR(" & _
                            IndexToLetter(piecesCol) & _
                            i & _
                            "*" & _
                            IndexToLetter(qtyCol) & _
                            i & _
                            "*" & _
                            IndexToLetter(priceCol) & _
                            i & _
                            ",0)"
    Next i
End Sub

'''�p������'''

'�v���p�e�B
Property Get Workbook() As Workbook
    Set Workbook = super.Workbook
End Property

Property Get Worksheet() As Worksheet
    Set Worksheet = super.Worksheet
End Property

Property Get DataHeaderRowIndex() As Long
    DataHeaderRowIndex = super.DataHeaderRowIndex
End Property

Property Get DataStartRowIndex() As Long
    DataStartRowIndex = super.DataStartRowIndex
End Property

Property Get DataStartColumnIndex() As Integer
    DataStartColumnIndex = super.DataStartColumnIndex
End Property

Property Get DataEndRowIndex() As Long
    DataEndRowIndex = super.DataEndRowIndex
End Property

Property Get AllDataRange() As Range
    Set AllDataRange = super.AllDataRange
End Property

Property Let Cells(rowIndex As Long, columnIndex As Integer, value As Variant)
    super.Cells(rowIndex, columnIndex) = value
End Property

' ���\�b�h
Function DataEndColumnIndex() As Integer
    DataEndColumnIndex = super.DataEndColumnIndex
End Function

Function DataNextRowIndex() As Long
    DataNextRowIndex = super.DataNextRowIndex
End Function

Sub WriteAllData(targetData As Collection)
    Call super.WriteAllData(targetData)
End Sub

Function GetAllData() As Collection
    Set GetAllData = super.GetAllData
End Function

Function GetAllData_NoHead() As Collection
    Set GetAllData_NoHead = super.GetAllData_NoHead
End Function

