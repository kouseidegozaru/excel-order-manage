Attribute VB_Name = "ProductsSearchModule"
Sub Decide()
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Set targetWb = ThisWorkbook '��������.xlsm
    Set targetWs = targetWb.Sheets(OrderWb_SheetName)
    
    Call WriteExcelData(targetWb, targetWs, GetCheckedValue(SearchWb_ProductCodeColumnNumber, SearchWb_StateColumnNumber))
End Sub

'�����t�H�[���̑I�����ꂽ�s�̏��i�R�[�h���擾
Public Function GetCheckedValue(ColumnNumber As Integer, stateColumnNumber As Integer) As Collection
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook '���i�}�X�^�[
    Set ws = wb.Sheets(SearchWb_SheetName)
    
    Dim checkedIDs As New Collection
    Dim i As Long
    
    For i = 1 To ws.Cells(ws.Rows.Count, ColumnNumber).End(xlUp).row
        If ws.Cells(i, stateColumnNumber).Value = True Then
            checkedIDs.Add ws.Cells(i, ColumnNumber).Value
        End If
    Next i
    
    Set GetCheckedValue = checkedIDs
End Function

'�����t�H�[���̏��i�R�[�h�𔭒����͂ɓ���
Public Sub WriteExcelData(wb As Workbook, ws As Worksheet, writeData As Collection)
    Dim lastRow As Long
    Dim i As Long
    
    ' ���[�N�V�[�g�̍ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, OrderWb_ProductCodeColumnNumber).End(xlUp).row + 1
    
    ' Collection�̊e�v�f�����[�N�V�[�g�ɒǉ�
    For i = 1 To writeData.Count
        ws.Cells(lastRow, OrderWb_ProductCodeColumnNumber).Value = writeData(i)
        lastRow = lastRow + 1
    Next i
End Sub
