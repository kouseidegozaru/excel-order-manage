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
Private pFilePath As String
Private pData As Collection
Private EmployeeCodeHeaderName As String
Private EmployeeCodeColumnIndex As Integer
Private departmentCodeHeaderName As String
Private departmentCodeColumnIndex As Integer

Private Sub Class_Initialize()
    EmployeeCodeHeaderName = "社員コード"
    EmployeeCodeColumnIndex = 1
    departmentCodeHeaderName = "部門コード"
    departmentCodeColumnIndex = 0
    pFilePath = "C:\Users\mfh077_user.MEFUREDMN\Desktop\excel-order-manage\発注管理\data\従業員コード.csv"
    
    Set pData = New Collection
    LoadData
End Sub

Public Function GetEmployeeCodes(departmentCode As String) As Collection
    Dim employeeCodes As New Collection
    Dim item As Variant

    For Each item In pData
        If item(departmentCodeHeaderName) = departmentCode Then
            employeeCodes.Add item(EmployeeCodeHeaderName)
        End If
    Next item

    Set GetEmployeeCodes = employeeCodes
End Function

Private Sub LoadData()
    Dim fileContent As String
    Dim lines As Variant
    Dim line As Variant
    Dim i As Integer
    Dim headers As Variant
    Dim dataItem As Object

    fileContent = ReadFileContent(pFilePath)
    lines = Split(fileContent, vbCrLf)
    
    If UBound(lines) < 1 Then Exit Sub

    headers = Split(lines(0), ",")

    For i = 1 To UBound(lines)
        If lines(i) <> "" Then
            line = Split(lines(i), ",")
            Set dataItem = CreateObject("Scripting.Dictionary")
            dataItem.Add departmentCodeHeaderName, line(departmentCodeColumnIndex)
            dataItem.Add EmployeeCodeHeaderName, line(EmployeeCodeColumnIndex)
            pData.Add dataItem
        End If
    Next i
End Sub

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


