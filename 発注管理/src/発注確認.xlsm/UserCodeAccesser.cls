VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserCodeAccesser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CSVファイルから対象の部門コードの従業員コードを取得
Private FilePath_ As String
Private Data_ As Collection
Private EmployeeCodeHeaderName As String
Private EmployeeCodeColumnIndex As Integer
Private departmentCodeHeaderName As String
Private departmentCodeColumnIndex As Integer

Private Sub Class_Initialize()
    '部門コードが記載された列のヘッダー名
    departmentCodeHeaderName = "部門コード"
    '部門コードが記載された列
    departmentCodeColumnIndex = 0
    
    '従業員コードが記載された列のヘッダー名
    EmployeeCodeHeaderName = "社員コード"
    '従業員コードが記載された列
    EmployeeCodeColumnIndex = 1
    
    'ファイルのパス
    FilePath_ = "\\mfh008-mfre\FS事業部\新）発注ノート\発注ノートシステム\data\従業員コード.csv"
    
    Set Data_ = New Collection
    Me.LoadData
End Sub

'部門コードから従業員コードを取得
Public Function GetEmployeeCodes(departmentCode As String) As Collection
    Dim employeeCodes As New Collection
    Dim item As Variant

    For Each item In Data_
        '該当の部門の場合
        If item(departmentCodeHeaderName) = departmentCode Then
            '従業員コードを追加
            employeeCodes.Add item(EmployeeCodeHeaderName)
        End If
    Next item

    Set GetEmployeeCodes = employeeCodes
End Function

'csvファイル読み込み
Private Sub LoadData()
    Dim fileContent As String
    Dim lines As Variant
    Dim line As Variant
    Dim i As Integer
    Dim headers As Variant
    Dim dataItem As Object

    'ファイル読み込み
    fileContent = ReadFileContent(FilePath_)
    '行ごとに分ける
    lines = Split(fileContent, vbCrLf)
    
    'データがない場合抜ける
    If UBound(lines) < 1 Then Exit Sub

    'ヘッダーを取得
    headers = Split(lines(0), ",")

    For i = 1 To UBound(lines)
        '行が空でない場合
        If lines(i) <> "" Then
            'コンマ区切りで分ける
            line = Split(lines(i), ",")
            '辞書型
            Set dataItem = CreateObject("Scripting.Dictionary")
            'ヘッダー名をキー値としてデータを格納
            '部門コード
            dataItem.Add departmentCodeHeaderName, line(departmentCodeColumnIndex)
            '従業員コード
            dataItem.Add EmployeeCodeHeaderName, line(EmployeeCodeColumnIndex)
            Data_.Add dataItem
        End If
    Next i
End Sub

'ファイル読み込み
Private Function ReadFileContent(FilePath As String) As String
    Dim fileContent As String
    Dim stream As New ADODB.stream
    
    On Error GoTo ErrorHandler
    
    With stream
        .Type = 2 ' adTypeText
        .Charset = "utf-8"
        .Open
        .LoadFromFile FilePath
        fileContent = .ReadText(-1) ' -1 = 全テキスト読み込み
        .Close
    End With
    
    ReadFileContent = fileContent
    Exit Function

ErrorHandler:
    MsgBox "ファイルを読み込む際にエラーが発生しました。エラーメッセージ: " & Err.Description
    ReadFileContent = ""
End Function


