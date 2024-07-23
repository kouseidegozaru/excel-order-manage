VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProductsSearchSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.recordSet
Private wb As Workbook
Private ws As Worksheet
Private checkBoxes As Collection

Private stateColumnNumber As Integer
Private checkBoxColumnNumber As Integer
Private dataStartColumnNumber As Integer
Private dataStartRowNumber As Long


Public Sub Initialize(recordSet As ADODB.recordSet)

    Set rs = recordSet
    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
    Set checkBoxes = New Collection
    
    stateColumnNumber = rs.Fields.Count + 3
    checkBoxColumnNumber = 1
    dataStartColumnNumber = 2
    dataStartRowNumber = 1
End Sub

Public Sub ExportRecordSet()
    Dim i As Long, j As Long
    Dim fld As Object
    Dim row As Long: row = dataStartRowNumber
    
    ' ヘッダーの書き込み
    For i = 0 To rs.Fields.Count - 1
        ws.Cells(row, i + dataStartColumnNumber).Value = rs.Fields(i).Name
    Next i
    
    ' データの書き込み
    rs.MoveFirst
    Do While Not rs.EOF
        row = row + 1
        For i = 0 To rs.Fields.Count - 1
            ws.Cells(row, i + dataStartColumnNumber).Value = rs.Fields(i).Value
        Next i
        
        ' チェックボックスの追加
        Call AddCheckBox(row, checkBoxColumnNumber, stateColumnNumber)
        
        rs.MoveNext
    Loop
End Sub

Public Function GetCheckedValue(ColumnNumber As Integer) As Collection
    Dim checkedIDs As New Collection
    Dim i As Long
    
    For i = 0 To ws.Cells(ws.Rows.Count, ColumnNumber).End(xlUp).row
        If ws.Cells(i, stateColumnNumber).Value = True Then
            checkedIDs.Add ws.Cells(i, ColumnNumber).Value
        End If
    Next i
    
    Set GetCheckedIDs = checkedIDs
End Function

Private Sub AddCheckBox(RowNumber As Long, ColumnNumber As Integer, stateColumnNumber As Integer)
    ' チェックボックスの追加
    Dim cb As CheckBox
    Set cb = ws.checkBoxes.Add(Left:=ws.Cells(RowNumber, ColumnNumber).Left, _
                               Top:=ws.Cells(RowNumber, ColumnNumber).Top, _
                               Width:=ws.Cells(RowNumber, ColumnNumber).Width, _
                               Height:=ws.Cells(RowNumber, ColumnNumber).Height)
    cb.Caption = ""
    cb.LinkedCell = ws.Cells(RowNumber, stateColumnNumber).Address
End Sub

