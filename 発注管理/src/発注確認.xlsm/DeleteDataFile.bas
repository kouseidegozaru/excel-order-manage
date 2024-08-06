Attribute VB_Name = "DeleteDataFile"
Sub DeleteDataFilles()
    
    Dim data As New DataSheetAccesser
    Dim fileProperty As New FilePropertyManager
    Dim filter As New FileFilter
    Dim filePath As String
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
        filePath = data.DataDirPath & "\" & FileName
        fileProperty.filePath = filePath
        
        '�ꂩ���O�̏ꍇ�t�@�C���폜
        If fileProperty.TargetDate < oneMonthAgo Then
            fs.DeleteFile filePath
        End If
        
    Next FileName
        
End Sub
