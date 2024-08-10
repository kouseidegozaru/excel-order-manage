VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetAccesserInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' IWorksheetHandler インターフェース
' シートにアクセスするためのインターフェース

' プロパティ
Property Get Workbook() As Workbook
End Property
Property Get Worksheet() As Worksheet
End Property
Property Get DataHeaderRowIndex() As Long
End Property
Property Get DataStartRowIndex() As Long
End Property
Property Get DataStartColumnIndex() As Integer
End Property
Property Get DataEndRowIndex() As Long
End Property
Property Get AllDataRange() As Range
End Property
Property Let Cells(rowIndex As Long, columnIndex As Integer, value As Variant)
End Property

' メソッド
Function DataEndColumnIndex() As Integer
End Function
Function DataNextRowIndex() As Long
End Function
Sub WriteAllData(targetData As Variant)
End Sub
Function GetAllData() As Variant
End Function
Function GetAllDara_NoHead() As Variant
End Function

