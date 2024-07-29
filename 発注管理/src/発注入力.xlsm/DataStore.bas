Attribute VB_Name = "DataStore"
'�f�[�^�̓ǂݏ���

'�f�[�^�̏�������
Sub SaveData()
    
    Dim savePath As String
    
    ' �V�������[�N�u�b�N���쐬
    Dim newWorkbook As Workbook
    Set newWorkbook = Workbooks.Add
    
    
    ' �V�������[�N�u�b�N�̒l�����
    '�������͂̃V�[�g
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(OrderWb_SheetName)
    '���i�R�[�h�f�[�^
    Dim productsData As Collection
    Set productsData = GetRangeValue(ws.range(OrderWb_InputProductsRange))
    writeData newWorkbook.Sheets(DataWb_SheetName), DataWb_ProductCodeRowNumber, DataWb_ProductCodeColumnNumber, productsData
    '���ʃf�[�^
    Dim qtyData As Collection
    Set qtyData = GetRangeValue(ws.range(OrderWb_InpuQtyRange))
    writeData newWorkbook.Sheets(DataWb_SheetName), DataWb_ProductCodeRowNumber, DataWb_ProductQtyColumnNumber, qtyData
    
    
    ' �ۑ��p�X���w��i��F�f�X�N�g�b�v�ɕۑ��j
    savePath = GetSaveFilePath
    
    ' �㏑���ۑ��̂��߂Ɍx�����b�Z�[�W���I�t�ɂ���
    Application.DisplayAlerts = False
    
    ' ���[�N�u�b�N��ۑ��i�����̃t�@�C��������Ώ㏑���ۑ��j
    newWorkbook.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    ' ���[�N�u�b�N�����
    newWorkbook.Close
        
    ' �x�����b�Z�[�W���ēx�L���ɂ���
    Application.DisplayAlerts = True
    
End Sub

'�V�[�g�ɗ���w�肵�ē���
Sub writeData(ws As Worksheet, rowIndex As Long, colIndex As Integer, writeData As Collection)
    Dim item As Variant

        For Each item In writeData
            ws.Cells(rowIndex, colIndex).value = item
            rowIndex = rowIndex + 1
        Next item

    
End Sub

Sub LoadData()
    
    '�������͂̏��i����S�č폜
    Dim orderRng As range
    Set orderRng = ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_InputProductsRange)
    orderRng.EntireRow.Delete
    
    '�t�@�C�������݂��Ȃ��ꍇ�͏����I��
    If Dir(GetSaveFilePath) = "" Then
        End
    End If
    
    '�������͂̃V�[�g
    Dim wb As Workbook
    Set wb = DataWb
    
    Dim ws As Worksheet
    Set ws = wb.Sheets(DataWb_SheetName)
    '���i�R�[�h�f�[�^
    Dim productsData As Collection
    Set productsData = GetRangeValue(ws.range(DataWb_ProductsRange))
    '���ʃf�[�^
    Dim qtyData As Collection
    Set qtyData = GetRangeValue(ws.range(DataWb_QtyRange))
    
    '�f�[�^����
    Dim OrderWs As Worksheet
    Set OrderWs = ThisWorkbook.Sheets(OrderWb_SheetName)
    
    writeData OrderWs, OrderWb_ProductCodeRowNumber, OrderWb_ProductCodeColumnNumber, productsData
    Dim target As range
    Set target = OrderWs.range(OrderWb_InputProductsRange)
    DisplayProductsInfo target
    writeData OrderWs, OrderWb_ProductCodeRowNumber, OrderWb_ProductQtyColumnNumber, qtyData
    
    wb.Close
End Sub

