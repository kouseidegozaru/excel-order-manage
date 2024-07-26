Attribute VB_Name = "ProductsSearchModule"
Sub Decide()
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Set targetWb = ThisWorkbook '��������.xlsm
    Set targetWs = targetWb.Sheets(OrderWb_SheetName)
    
    Dim selectedProductsCD As Collection
    Set selectedProductsCD = GetCheckedValue(SearchWb_ProductCodeColumnNumber, SearchWb_StateColumnNumber)
    
    Call WriteExcelData(targetWb, targetWs, selectedProductsCD)
    
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
        If ws.Cells(i, stateColumnNumber).Value = True Then
            checkedIDs.Add ws.Cells(i, columnNumber).Value
        End If
    Next i
    
    Set GetCheckedValue = checkedIDs
End Function

'�����t�H�[���̏��i�R�[�h�𔭒����͂ɓ���
Public Sub WriteExcelData(wb As Workbook, ws As Worksheet, selectedData As Collection)
    Dim lastRow As Long
    Dim i As Long
    Dim writeData As Collection
    
    Set writeData = FilterCollection(selectedData, GetProductsCD)
    
    ' ���[�N�V�[�g�̍ŏI�s���擾
    lastRow = OrderWb_NextProductsRow
    
    ' Collection�̊e�v�f�����[�N�V�[�g�ɒǉ�
    For i = 1 To writeData.Count
        ws.Cells(lastRow, OrderWb_ProductCodeColumnNumber).Value = writeData(i)
        lastRow = lastRow + 1
    Next i
End Sub

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
            resultCol.Add itemBase
        End If
    Next itemBase
    
    ' ���ʂ̃R���N�V������Ԃ�
    Set FilterCollection = resultCol
End Function
