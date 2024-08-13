VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSheetAccesser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'発注入力フォームのセル番地の設定

'SheetAccesserクラスを疑似的に継承
Private super As SheetAccesser

'発注データフォルダパス
Private SaveDirPath_ As String
Private SaveFileName_ As String
'保存データ設定
Private WorkSheetName_ As String

Private Sub Class_Initialize()

    SaveDirPath_ = "\\mfh008-mfre\FS事業部\新）発注ノート\発注ノートシステム\data"

    WorkSheetName_ = "Sheet1"
    
    ProductCodeColumnIndex_ = 1
    
    Set super = New SheetAccesser
    super.InitStatus DataHeaderRow:=1, _
                     DataStartRow:=2, _
                     DataStartColumn:=1
    
End Sub

Public Sub InitStatus(BumonCode As Integer, userCode As Integer, targetDate As Date)
    SaveFileName_ = "b" & BumonCode & "-" & _
                    "u" & userCode & "-" & _
                    "d" & Format(targetDate, "yyyymmdd") & "-" & _
                    ".xlsx"
End Sub

Public Sub InitNewWorkbook()
    ' 新規ワークブックの作成
    super.InitWorkbook Application.Workbooks.Add
End Sub

Public Sub InitOpenWorkBook()
    super.InitWorkbook Workbooks.Open(SaveFilePath)
End Sub

Public Sub InitWorkSheet()
    super.InitWorkSheet Me.Workbook.Sheets(WorkSheetName_)
End Sub

'保存ファイル名
Property Get SaveFileName() As String
    SaveFileName = SaveFileName_
End Property
'保存ファイル名
Property Get SaveDirPath() As String
    SaveDirPath = SaveDirPath_
End Property
'保存ファイルパス
Property Get SaveFilePath() As String
    SaveFilePath = Me.SaveDirPath & "\" & Me.SaveFileName
End Property

Property Get ProductCodeColumnIndex() As Integer
    ProductCodeColumnIndex = ProductCodeColumnIndex_
End Property

Sub WriteTableData(targetData As Collection)
    super.WriteMatrixCollection Me.DataHeaderRowIndex, Me.DataStartColumnIndex, targetData
End Sub

Public Sub Save()
    ' 上書き保存のために警告メッセージをオフにする
    Application.DisplayAlerts = False
    
    ' ワークブックを保存（既存のファイルがあれば上書き保存）
    Me.Workbook.SaveAs FileName:=Me.SaveFilePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    ' 警告メッセージを再度有効にする
    Application.DisplayAlerts = True
End Sub

Public Sub CloseWorkBook()
    Me.Workbook.Close SaveChanges:=False
End Sub


'''継承部分'''

'プロパティ
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
    Call super.Cells(rowIndex, columnIndex, value)
End Property

' メソッド
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


