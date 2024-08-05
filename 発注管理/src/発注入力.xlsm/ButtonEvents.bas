Attribute VB_Name = "ButtonEvents"

'���i�����V�[�g�Ɋւ��鏈��
'�m��{�^��(���i�R�[�h�̔��f)
Sub Decide()

    Dim order As New OrderSheetAccesser
    Dim Search As New SearchSheetAccesser
    Set Search.Workbook = ActiveWorkbook
    Search.InitWorkSheet
    
    '�ύX�̃C�x���g�𖳎�
    IsIgnoreChangeEvents = True
    
    '�d�����鏤�i�R�[�h��r��
    Dim writeData As Collection
    Set writeData = FilterCollection(Search.GetCheckedProductsCode, _
                                     order.ProductsCode)
                                     
    Dim startRowIndex As Long
    Dim lastRowIndex As Long
    
    startRowIndex = order.DataNextRowNumber
    lastRowIndex = startRowIndex
    
    '�������͂ɏ��i�R�[�h����
    For i = 1 To writeData.Count
        order.Cells(lastRowIndex, order.ProductCodeColumnNumber) = writeData(i)
        lastRowIndex = lastRowIndex + 1
    Next i
    
    '�������͂ɏ��i�R�[�h����͂����͈�
    Dim target As range
    Set target = order.Worksheet.range(order.ProductCodeColumn & startRowIndex & ":" & order.ProductCodeColumn & lastRowIndex)
    
    '���i���\��
    DisplayProductsInfo target
    
    '�ۑ�
    SaveData
    
    IsIgnoreChangeEvents = False
    
    order.Workbook.Activate
    
End Sub

'�������̓V�[�g�Ɋւ��鏈��
'�����t�H�[���X�V
Sub Update()

    Application.ScreenUpdating = False
    
    Dim order As New OrderSheetAccesser
    Dim Search As New SearchSheetAccesser
    
    Set Search.Workbook = order.Workbook
    Search.InitWorkSheet
    Search.Clear
    
    Dim DataBaseAccesser As New DataBaseAccesser
    Dim rs As ADODB.recordSet
    ' �f�[�^�x�[�X���烌�R�[�h�Z�b�g���擾
    Set rs = DataBaseAccesser.GetAllProducts(order.BumonCode)
    
    Dim rowIndex As Long
    Dim columnIndex As Integer
    
    rowIndex = Search.DataStartRowNumber
    columnIndex = Search.DataStartColumnNumber
    
    ' �f�[�^�̏�������
    rs.MoveFirst
    Do While Not rs.EOF
        
        For i = 0 To rs.Fields.Count - 1
            Search.Cells(rowIndex, i + columnIndex) = rs.Fields(i).value
        Next i
        
        ' �`�F�b�N�{�b�N�X�̒ǉ�
        Search.AddCheckBox rowIndex
        
        rowIndex = rowIndex + 1
        rs.MoveNext
    Loop
    
    Search.Worksheet.Activate

    Application.ScreenUpdating = True
    
End Sub

'���i�����V�[�g�̕\��
Sub Search()
    Dim order As New OrderSheetAccesser
    order.FormatSheet.Activate
End Sub
