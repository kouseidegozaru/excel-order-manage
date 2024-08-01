Attribute VB_Name = "ButtonEvents"

'���i�����V�[�g�Ɋւ��鏈��
'�m��{�^��(���i�R�[�h�̔��f)
Sub Decide()

    Dim order As New OrderSheetAccesser
    Dim search As New SearchSheetAccesser
    Set search.Workbook = ActiveWorkbook
    search.InitWorkSheet
    
    '�ύX�̃C�x���g�𖳎�
    IsIgnoreChangeEvents = True
    
    '�d�����鏤�i�R�[�h��r��
    Dim writeData As Collection
    Set writeData = FilterCollection(search.GetCheckedProductsCode, _
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
'�����t�H�[���Ăяo��
Sub search()

    Application.ScreenUpdating = False
    
    Dim order As New OrderSheetAccesser
    Dim search As New SearchSheetAccesser
    
    search.NewWorkbook
    search.CopyWorkSheet order.FormatSheet
    
    Dim DataBaseAccesser As New DataBaseAccesser
    Dim rs As ADODB.recordSet
    ' �f�[�^�x�[�X���烌�R�[�h�Z�b�g���擾
    Set rs = DataBaseAccesser.GetAllProducts(order.BumonCode)
    
    Dim rowIndex As Long
    Dim columnIndex As Integer
    
    rowIndex = search.DataStartRowNumber
    columnIndex = search.DataStartColumnNumber
    
    '�w�b�_�[��������
    For i = 0 To rs.Fields.Count - 1
        search.Cells(rowIndex, i + columnIndex) = rs.Fields(i).name
    Next i
    
    ' �f�[�^�̏�������
    rs.MoveFirst
    Do While Not rs.EOF
        rowIndex = rowIndex + 1
        For i = 0 To rs.Fields.Count - 1
            search.Cells(rowIndex, i + columnIndex) = rs.Fields(i).value
        Next i
        
        ' �`�F�b�N�{�b�N�X�̒ǉ�
        search.AddCheckBox rowIndex
        
        rs.MoveNext
    Loop

    Application.ScreenUpdating = True
    
End Sub
