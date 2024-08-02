VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoadSheetAccesser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'発注入力フォームのセル番地の設定

Private WorkBookName_ As String
Private WorkBook_ As Workbook

Private WorkSheetName_ As String
Private WorkSheet_ As Worksheet

Private ProductCodeColumnNumber_ As Integer
Private QtyColumnNumber_ As Integer
Private DataStartRowNumber_ As Long

Private InputBumonCodeRange_ As String
Private OutputBumonNameRange_ As String
Private InputTargetDateRange_ As String
Private InputUserCodeRange_ As String
Private OutputUserNameRange_ As String

Private FilePropertyStartColumnNumber_ As Integer
Private FilePropertyUserNameRowNumber_ As Integer
Private FilePropertyCompleteRowNumber_ As Integer
Private FilePropertyUpdateTimeRowNumber_ As Integer

Private CompleteString_ As String

Private Sub Class_Initialize()
    WorkBookName_ = "発注確認.xlsm"
    WorkSheetName_ = "発注確認"
    
    CompleteString_ = "◎"
    
    ProductCodeColumnNumber_ = 1
    QtyColumnNumber_ = 10
    DataStartRowNumber_ = 9
    
    FilePropertyStartColumnNumber_ = 2
    FilePropertyUserNameRowNumber_ = 4
    FilePropertyCompleteRowNumber_ = 5
    FilePropertyUpdateTimeRowNumber_ = 6
    
    InputBumonCodeRange_ = "A2"
    OutputBumonNameRange_ = "B2"
    
    InputTargetDateRange_ = "C2"
    
    InitWorkBook
    InitWorkSheet
End Sub

Private Sub InitWorkBook()
    Set WorkBook_ = Workbooks(WorkBookName_)
End Sub

Private Sub InitWorkSheet()
    Set WorkSheet_ = WorkBook_.Sheets(WorkSheetName_)
End Sub

Property Get Workbook() As Workbook
    Set Workbook = WorkBook_
End Property

Property Get Worksheet() As Worksheet
    Set Worksheet = WorkSheet_
End Property

Property Get ProductCodeColumnNumber() As Integer
    ProductCodeColumnNumber = ProductCodeColumnNumber_
End Property

Property Get ProductCodeColumn() As String
    ProductCodeColumn = NumberToLetter(Me.ProductCodeColumnNumber)
End Property

Property Get QtyColumnNumber() As Integer
    QtyColumnNumber = QtyColumnNumber_
End Property

Property Get QtyColumn() As String
    QtyColumn = NumberToLetter(Me.QtyColumnNumber)
End Property

Property Get FilePropertyStartColumnNumber() As Integer
    FilePropertyStartColumnNumber = FilePropertyStartColumnNumber_
End Property

'次に入力するファイル情報が空白の列番号
Public Function FilePropertyNextColumnNumber() As Integer
    FilePropertyNextColumnNumber = Me.Worksheet.Cells(FilePropertyUserNameRowNumber_, Me.Worksheet.Columns.Count).End(xlToLeft).Column + 1
End Function

'ファイル情報が記載されている最後の列番号
Property Get FilePropertyEndColumnNumber() As Long
    Dim lastColumn As Long
    'ファイル情報にデータがない場合列範囲をずらす(無いと範囲にヘッダー列も含まれてしまう)
    lastColumn = Me.FilePropertyNextColumnNumber - 1
    If lastColumn < Me.FilePropertyStartColumnNumber Then
        lastColumn = Me.FilePropertyStartColumnNumber
    End If
    FilePropertyEndColumnNumber = lastColumn
End Property

'ファイル情報の範囲
Property Get FilePropertyRange() As Range
    Dim strRng As String
    strRng = NumberToLetter(Me.FilePropertyStartColumnNumber) & _
             FilePropertyUserNameRowNumber_ & _
             ":" & _
             NumberToLetter(Me.FilePropertyEndColumnNumber) & _
             FilePropertyUpdateTimeRowNumber_
    Set FilePropertyRange = Me.Worksheet.Range(strRng)
End Property

Public Sub ClearFileProperty()
    Me.FilePropertyRange.Delete Shift:=xlToLeft
End Sub

Public Sub AddFileProperty(UserName As String, IsCompleted As Boolean, updatedTime As Date)
    Dim lastColumn As Integer
    lastColumn = FilePropertyNextColumnNumber
    
    Dim writeCompleteString As String
    writeCompleteString = ""
    If IsCompleted Then
        writeCompleteString = CompleteString_
        
        '更新日時書き込み
        Me.Worksheet.Cells(FilePropertyUpdateTimeRowNumber_, lastColumn) = updatedTime
    End If
    
    Me.Worksheet.Cells(FilePropertyUserNameRowNumber_, lastColumn) = UserName
    Me.Worksheet.Cells(FilePropertyCompleteRowNumber_, lastColumn) = writeCompleteString
End Sub

Property Get DataStartRowNumber() As Long
    DataStartRowNumber = DataStartRowNumber_
End Property

'次に入力する商品情報が空白の行番号
Public Function DataNextRowNumber() As Long
    DataNextRowNumber = Me.Worksheet.Cells(Me.Worksheet.Rows.Count, Me.ProductCodeColumnNumber).End(xlUp).Row + 1
End Function

'商品情報が記載されている最後の行番号
Property Get DataEndRowNumber() As Long
    Dim lastRow As Long
    '発注入力に商品データがない場合行範囲をずらす(無いと範囲にヘッダー行も含まれてしまう)
    lastRow = Me.DataNextRowNumber - 1
    If lastRow < Me.DataStartRowNumber Then
        lastRow = Me.DataStartRowNumber
    End If
    DataEndRowNumber = lastRow
End Property

'発注入力にある商品コードの範囲
Property Get ProductsCodeRange() As Range
    Dim strRng As String
    strRng = Me.ProductCodeColumn & _
             Me.DataStartRowNumber & _
             ":" & _
             Me.ProductCodeColumn & _
             Me.DataEndRowNumber
    Set ProductsCodeRange = Me.Worksheet.Range(strRng)
End Property

'発注入力にある商品の数量の範囲
Property Get QtyRange() As Range
    Dim strRng As String
    strRng = Me.QtyColumn & _
             Me.DataStartRowNumber & _
             ":" & _
             Me.QtyColumn & _
             Me.DataEndRowNumber
    Set QtyRange = Me.Worksheet.Range(strRng)
End Property

'部門コードの取得
Property Get BumonCode() As Integer
    '値の取得
    Dim value As Integer
    value = Me.Worksheet.Range(InputBumonCodeRange_).value
    
    '部門コードが存在するか
    Dim dataStorage As New DataBaseAccesser
    If dataStorage.ExistsBumon(value) Then
        BumonCode = value
    Else
        BumonCode = 0
        MsgBox ("正しい部門コードを入力して下さい")
        End
    End If
End Property

Property Get BumonCodeRange() As Range
    Set BumonCodeRange = Me.Worksheet.Range(InputBumonCodeRange_)
End Property

'対象日付の取得
Property Get TargetDate() As Date
    Dim value As Date
    value = Me.Worksheet.Range(InputTargetDateRange_).value
    
    TargetDate = value
End Property

Property Get TargetDateRange() As Range
    Set TargetDateRange = Me.Worksheet.Range(InputTargetDateRange_)
End Property

'発注入力に既に入力されている商品コードの取得
Property Get ProductsCode() As Collection
    Set ProductsCode = GetRangeValue(Me.ProductsCodeRange)
End Property

'発注入力に既に入力されている数量の取得
Property Get Qty() As Collection
    Set Qty = GetRangeValue(Me.QtyRange)
End Property

'担当者名の表示
Property Let UserName(name As String)
    Me.Worksheet.Range(OutputUserNameRange_).value = name
End Property

'部門名の表示
Property Let BumonName(name As String)
    Me.Worksheet.Range(OutputBumonNameRange_).value = name
End Property

Property Let Cells(rowIndex As Long, columnIndex As Integer, value As Variant)
    Me.Worksheet.Cells(rowIndex, columnIndex).value = value
End Property

Public Sub WriteProductsCode(writeCollection As Collection)
    writeData Me.Worksheet, Me.DataStartRowNumber, Me.ProductCodeColumnNumber, writeCollection
End Sub

Public Sub WriteQty(writeCollection As Collection)
    writeData Me.Worksheet, Me.DataStartRowNumber, Me.QtyColumnNumber, writeCollection
End Sub


