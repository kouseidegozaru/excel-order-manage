VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilePropertyManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim FilePath_ As String
Dim fileName_ As String

Dim BumonCodeIdentifier_ As String
Dim UserCodeIdentifier_ As String
Dim DateIdentifier_ As String
Dim BreakIdentifier_ As String

Private Sub Class_Initialize()
    BumonCodeIdentifier_ = "b"
    UserCodeIdentifier_ = "u"
    DateIdentifier_ = "d"
    BreakIdentifier_ = "-"
End Sub

Public Sub InitFilePath(targetFilePath As String)
    FilePath_ = targetFilePath
    fileName_ = GetFileName(FilePath_)
End Sub

Property Get bumonCode() As Integer
    bumonCode = CInt(ExtractBetweenCharacters(fileName_, BumonCodeIdentifier_, BreakIdentifier_))
End Property
Property Get BumonCodeIdentifier() As String
    BumonCodeIdentifier = BumonCodeIdentifier_
End Property

Property Get userCode() As Integer
    userCode = CInt(ExtractBetweenCharacters(fileName_, UserCodeIdentifier_, BreakIdentifier_))
End Property
Property Get UserCodeIdentifier() As String
    UserCodeIdentifier = UserCodeIdentifier_
End Property

Property Get targetDate() As Date
    Dim strDate As String
    Dim year As Integer, month As Integer, day As Integer
    
    strDate = ExtractBetweenCharacters(fileName_, DateIdentifier_, BreakIdentifier_)
    
    If Len(strDate) = 8 Then
        year = CInt(Left(strDate, 4))
        month = CInt(Mid(strDate, 5, 2))
        day = CInt(Right(strDate, 2))
        targetDate = DateSerial(year, month, day)
    Else
        targetDate = Date ' デフォルト値
    End If
End Property
Property Get DateIdentifier() As String
    DateIdentifier = DateIdentifier_
End Property
Property Get BreakIdentifier() As String
    BreakIdentifier = BreakIdentifier_
End Property

Property Get UpdatedDate() As Date
    UpdatedDate = FileDateTime(FilePath_)
End Property

Private Function ExtractBetweenCharacters(ByVal str As String, ByVal startChar As String, ByVal endChar As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim searchPos As Long
    
    startPos = InStr(str, startChar)
    
    If startPos = 0 Then
        ExtractBetweenCharacters = ""
        Exit Function
    End If
    
    searchPos = startPos + Len(startChar)
    endPos = InStr(searchPos, str, endChar)
    
    If endPos > searchPos Then
        ExtractBetweenCharacters = Mid(str, searchPos, endPos - searchPos)
    Else
        ExtractBetweenCharacters = ""
    End If
End Function

Private Function GetFileName(FilePath As String) As String
    Dim fileName As String
    fileName = Mid(FilePath, InStrRev(FilePath, "\") + 1)
    GetFileName = fileName
End Function

