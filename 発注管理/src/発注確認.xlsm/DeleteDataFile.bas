Attribute VB_Name = "DeleteDataFile"
Sub DeleteDataFiles()
    
    Dim data As New DataSheetAccesser
    Dim fileProperty As New FilePropertyManager
    Dim filter As New FileFilter
    Dim FilePath As String
    Dim fs As New Scripting.FileSystemObject
    
    '�Ώۃf�B���N�g����ݒ�
    filter.DirPath = data.DataDirPath
    
    '�t�@�C�����̎擾
    Dim fileNames As Collection
    Set fileNames = filter.AndFilter()
    
    ' �����̓��t���擾
    Dim today As Date
    today = Date

    ' 1�����O�̓��t���v�Z
    Dim oneMonthAgo As Date
    oneMonthAgo = DateAdd("m", -1, today)
    
    For Each FileName In fileNames
    
        '�t�@�C�����擾����
        FilePath = data.DataDirPath & "\" & FileName
        fileProperty.FilePath = FilePath
        
        '�ꂩ���O�̏ꍇ�t�@�C���폜
        If fileProperty.targetDate < oneMonthAgo Then
            fs.DeleteFile FilePath
        End If
        
    Next FileName
        
End Sub
