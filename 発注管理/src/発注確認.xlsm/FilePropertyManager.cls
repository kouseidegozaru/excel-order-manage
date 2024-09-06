VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilePropertyManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'データファイル名から各情報を取得するクラス

'ファイル名は次のような形式になる
'b41-u70-d20240101-.xlsx

Dim FilePath_ As String
Dim fileName_ As String

Dim BumonCodeIdentifier_ As String
Dim UserCodeIdentifier_ As String
Dim DateIdentifier_ As String
Dim BreakIdentifier_ As String

Private Sub Class_Initialize()
    'ファイル名の各区分の文字列を設定
    '部門コード
    BumonCodeIdentifier_ = "b"
    '担当者コード
    UserCodeIdentifier_ = "u"
    '発注日
    DateIdentifier_ = "d"
    '各区分の終了文字列
    BreakIdentifier_ = "-"
End Sub

Public Sub InitFilePath(targetFilePath As String)
    'ファイル名とファイルパスを設定
    FilePath_ = targetFilePath
    fileName_ = GetFileName(FilePath_)
End Sub

'部門コードの開始区分文字列を参照
Property Get BumonCodeIdentifier() As String
    BumonCodeIdentifier = BumonCodeIdentifier_
End Property
'担当者コードの開始区分文字列を参照
Property Get UserCodeIdentifier() As String
    UserCodeIdentifier = UserCodeIdentifier_
End Property
'発注日の開始区分文字列を参照
Property Get DateIdentifier() As String
    DateIdentifier = DateIdentifier_
End Property
'各区分の終了文字列を参照
Property Get BreakIdentifier() As String
    BreakIdentifier = BreakIdentifier_
End Property

'ファイル名から部門コードを参照
Property Get bumonCode() As Integer
    '部門区分と終了区分に挟まれた数値を取得
    bumonCode = CInt(ExtractBetweenCharacters(fileName_, BumonCodeIdentifier_, BreakIdentifier_))
End Property
'ファイル名から担当者コードを参照
Property Get userCode() As Integer
    '担当者区分と終了区分に挟まれた数値を取得
    userCode = CInt(ExtractBetweenCharacters(fileName_, UserCodeIdentifier_, BreakIdentifier_))
End Property
'ファイル名から発注日を参照
Property Get targetDate() As Date
    Dim strDate As String
    Dim year As Integer, month As Integer, day As Integer
    '部門区分と終了区分に挟まれた数値を取得
    strDate = ExtractBetweenCharacters(fileName_, DateIdentifier_, BreakIdentifier_)
    
    '取得した文字列を日付に変換
    If Len(strDate) = 8 Then
        year = CInt(Left(strDate, 4))
        month = CInt(Mid(strDate, 5, 2))
        day = CInt(Right(strDate, 2))
        targetDate = DateSerial(year, month, day)
    Else
        targetDate = Date ' デフォルト値
    End If
End Property
'ファイル更新日を参照
Property Get UpdatedDate() As Date
    UpdatedDate = FileDateTime(FilePath_)
End Property

'文字列から指定した文字列の間にある文字を取得
Private Function ExtractBetweenCharacters(ByVal str As String, ByVal startChar As String, ByVal endChar As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim searchPos As Long
    
    '開始文字列の位置
    startPos = InStr(str, startChar)
    
    '開始位置が見つからない場合空文字列を返却
    If startPos = 0 Then
        ExtractBetweenCharacters = ""
        Exit Function
    End If
    
    '抽出対象の開始位置
    searchPos = startPos + Len(startChar)
    '抽出対象の終了位置
    endPos = InStr(searchPos, str, endChar)
    
    '終了位置が見つかった場合
    If endPos > searchPos Then
        '抽出した文字列を返却
        ExtractBetweenCharacters = Mid(str, searchPos, endPos - searchPos)
    Else
        ExtractBetweenCharacters = ""
    End If
End Function
'ファイル名の取得
Private Function GetFileName(FilePath As String) As String
    Dim fileName As String
    fileName = Mid(FilePath, InStrRev(FilePath, "\") + 1)
    GetFileName = fileName
End Function

