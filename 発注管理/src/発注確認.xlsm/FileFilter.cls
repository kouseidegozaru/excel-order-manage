VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private dirPath_ As String

Private Sub class_initialize()
    ' èâä˙âªéûÇ…çsÇ§èàóùÇ™Ç†ÇÍÇŒÇ±Ç±Ç…ãLèq
End Sub

Property Let DirPath(targetDirPath As String)
    dirPath_ = targetDirPath
End Property

Public Function AndFilter(ParamArray filterNames() As Variant) As Collection
    Dim fileList As Collection
    Set fileList = New Collection
    Dim fileName As String
    
    fileName = Dir(dirPath_ & "\*.*")
    Do While fileName <> ""
        Dim matchesAll As Boolean
        matchesAll = True
        For Each filterName In filterNames
            If InStr(1, fileName, filterName, vbTextCompare) = 0 Then
                matchesAll = False
                Exit For
            End If
        Next filterName
        
        If matchesAll Then
            fileList.Add fileName
        End If
        fileName = Dir()
    Loop
    
    Set AndFilter = fileList
End Function

Public Function OrFilter(ParamArray filterNames() As Variant) As Collection
    Dim fileList As Collection
    Set fileList = New Collection
    Dim fileName As String
    
    fileName = Dir(dirPath_ & "\*.*")
    Do While fileName <> ""
        Dim matchesAny As Boolean
        matchesAny = False
        For Each filterName In filterNames
            If IsObject(filterName) Then
                Dim subFilterResult As Collection
                Set subFilterResult = filterName
                Dim subFile As Variant
                For Each subFile In subFilterResult
                    If subFile = fileName Then
                        matchesAny = True
                        Exit For
                    End If
                Next subFile
            ElseIf InStr(1, fileName, filterName, vbTextCompare) > 0 Then
                matchesAny = True
                Exit For
            End If
        Next filterName
        
        If matchesAny Then
            fileList.Add fileName
        End If
        fileName = Dir()
    Loop
    
    Set OrFilter = fileList
End Function

