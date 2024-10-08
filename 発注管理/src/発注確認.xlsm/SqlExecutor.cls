VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlExecutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'SQLを実行するための管理クラス

Private connectionString As String

' 接続文字列を設定するメソッド
Public Sub SetConnection(ByVal connString As String)
    connectionString = connString
End Sub

' SELECT文を実行してRecordsetを返すメソッド
Public Function ExecuteSelect(ByVal query As String, ByVal parameters As Object) As ADODB.Recordset
    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    Dim paramKey As Variant
    
    ' 接続を開く
    conn.Open connectionString
    Set cmd.ActiveConnection = conn
    cmd.CommandText = query
    
    ' パラメーターを追加
    If Not parameters Is Nothing Then
        For Each paramKey In parameters.Keys
            ' パラメーターの型とサイズは適宜調整する必要があります
            cmd.parameters.Append cmd.CreateParameter(paramKey, 3, 1, , parameters(paramKey)) ' 3はadInteger、1はadParamInputを意味する
        Next paramKey
    End If
    
    ' クエリを実行
    Set rs = cmd.Execute
    
    ' 結果を返す
    Set ExecuteSelect = rs
    
    'レコードセットが破棄されてしまうので閉じない
'    ' 接続を閉じる
'    conn.Close
'
'    ' オブジェクトを解放
'    Set cmd = Nothing
'    Set conn = Nothing
End Function


' Excelファイルにクエリを実行してRecordsetを返すメソッド
Public Function ExecuteExcelQuery(ByVal FilePath As String, ByVal query As String) As ADODB.Recordset
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim connectionString As String
    
    ' Excelファイルへの接続文字列を設定
    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FilePath & ";Extended Properties=""Excel 12.0 Xml;HDR=YES"""
    
    ' 接続を開く
    conn.Open connectionString
    
    ' クエリを実行
    rs.Open query, conn, adOpenStatic, adLockReadOnly
    
    ' 結果を返す
    Set ExecuteExcelQuery = rs
'
'    ' 接続を閉じる
'    conn.Close
'
'    ' オブジェクトを解放
'    Set conn = Nothing
End Function


