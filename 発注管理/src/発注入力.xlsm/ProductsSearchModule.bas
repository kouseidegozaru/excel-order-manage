Attribute VB_Name = "ProductsSearchModule"
'���i�����V�[�g�Ɋւ��鏈��

'�m��{�^��(���i�R�[�h�̔��f)
Sub Decide()
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Set targetWb = ThisWorkbook '��������.xlsm
    Set targetWs = targetWb.Sheets(OrderWb_SheetName)
    
    Dim selectedProductsCD As Collection
    Set selectedProductsCD = GetCheckedValue(SearchWb_ProductCodeColumnNumber, SearchWb_StateColumnNumber)
    
    SetIgnoreState True
        
    Dim target As range
    Set target = WriteExcelData(targetWb, targetWs, selectedProductsCD)
    DisplayProductsInfo target
    SaveData
    SetIgnoreState False
    
    
    ThisWorkbook.Activate
    
End Sub

'�����t�H�[���̑I�����ꂽ�s�̏��i�R�[�h���擾
Public Function GetCheckedValue(columnNumber As Integer, stateColumnNumber As Integer) As Collection
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook '���i�}�X�^�[
    Set ws = wb.Sheets(SearchWb_SheetName)
    
    Dim checkedIDs As New Collection
    Dim i As Long
    
    For i = 1 To ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).row
        If ws.Cells(i, stateColumnNumber).value = True Then
            checkedIDs.add ws.Cells(i, columnNumber).value
        End If
    Next i
    
    Set GetCheckedValue = checkedIDs
End Function

'�����t�H�[���̏��i�R�[�h�𔭒����͂ɓ���
Public Function WriteExcelData(wb As Workbook, ws As Worksheet, selectedData As Collection) As range
    Dim lastRow As Long
    Dim startRow As Long
    
    Dim i As Long
    Dim writeData As Collection
    
    Set writeData = FilterCollection(selectedData, GetProductsCD)
    
    
    ' ���[�N�V�[�g�̍ŏI�s���擾
    startRow = OrderWb_NextProductsRow
    lastRow = startRow
    
    ' Collection�̊e�v�f�����[�N�V�[�g�ɒǉ�
    For i = 1 To writeData.Count
        ws.Cells(lastRow, OrderWb_ProductCodeColumnNumber).value = writeData(i)
        lastRow = lastRow + 1
    Next i
    
    '���͂����͈͂�Ԃ�
    
    Set WriteExcelData = ws.range(OrderWb_ProductCodeColumn & startRow & ":" & OrderWb_ProductCodeColumn & lastRow)
End Function

'collection�^�̕ϐ����׏d������l�����O
Function FilterCollection(baseCol As Collection, filterCol As Collection) As Collection
    Dim resultCol As New Collection
    Dim itemBase As Variant
    Dim itemFilter As Variant
    Dim exists As Boolean
    
    ' baseCol�̒l�����[�v���āAfilterCol�ɑ��݂��Ȃ����̂���resultCol�ɒǉ�
    For Each itemBase In baseCol
        exists = False
        For Each itemFilter In filterCol
            If itemBase = itemFilter Then
                exists = True
                Exit For
            End If
        Next itemFilter
        If Not exists Then
            resultCol.add itemBase
        End If
    Next itemBase
    
    ' ���ʂ̃R���N�V������Ԃ�
    Set FilterCollection = resultCol
End Function
