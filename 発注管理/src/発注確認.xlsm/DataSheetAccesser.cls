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

    '発注データの保存ディレクトリパス
    SaveDirPath_ = "\\mfh008-mfre\FS事業部\新）発注ノート\発注ノートシステム\data\data"

    'シート名
    WorkSheetName_ = "Sheet1"
    
    '商品コードの列番号
    ProductCodeColumnIndex_ = 1
    
    '疑似抽象クラスをインスタンス化
    Set super = New SheetAccesser
    'データのヘッダー行、開始行、開始列を指定
    super.InitStatus DataHeaderRow:=1, _
                     DataStartRow:=2, _
                     DataStartColumn:=1
    
End Sub

'ファイル名を設定
Public Sub InitSaveFileName(fileName As String)
    SaveFileName_ = fileName
End Sub

'ワークブックを設定
Public Sub InitOpenWorkBook()
    super.InitWorkbook Workbooks.Open(SaveFilePath)
End Sub

'ワークシートを設定
Public Sub InitWorkSheet()
    super.InitWorkSheet Me.Workbook.Sheets(WorkSheetName_)
End Sub

'保存ファイル名の参照
Property Get SaveFileName() As String
    SaveFileName = SaveFileName_
End Property
'保存ディレクトリパスの参照
Property Get SaveDirPath() As String
    SaveDirPath = SaveDirPath_
End Property
'保存ファイルパスの参照
Property Get SaveFilePath() As String
    SaveFilePath = Me.SaveDirPath & "\" & Me.SaveFileName
End Property

'商品コードの列番号
Property Get ProductCodeColumnIndex() As Integer
    ProductCodeColumnIndex = ProductCodeColumnIndex_
End Property

'二次元のコレクションをヘッダーごと書き込む
Sub WriteTableData(targetData As Collection)
    super.WriteMatrixCollection Me.DataHeaderRowIndex, Me.DataStartColumnIndex, targetData
End Sub

'ワークブックを閉じる
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


