Attribute VB_Name = "DeleteDataFile"
'3�����O�̔������̔����f�[�^�t�@�C�����폜
Sub DeleteDataFiles()
    
    '�����f�[�^�V�[�g�A�N�Z�T�̃C���X�^���X��
    Dim data As New DataSheetAccesser
    '�f�[�^�t�@�C���̑������擾�N���X���C���X�^���X��
    Dim fileProperty As New FilePropertyManager
    '�f�[�^�t�@�C�����t�@�C�����ƂɃt�B���^�[����N���X���C���X�^���X��
    Dim filter As New FileFilter
    
    Dim FilePath As String
    Dim fs As New Scripting.FileSystemObject
    
    '�Ώۃf�B���N�g����ݒ�
    filter.DirPath = data.SaveDirPath
    
    '�S�Ẵt�@�C�����̎擾
    Dim fileNames As Collection
    Set fileNames = filter.AndFilter()
    
    ' �����̓��t���擾
    Dim today As Date
    today = Date

    ' 3�����O�̓��t���v�Z
    Dim threeMonthAgo As Date
    threeMonthAgo = DateAdd("m", -3, today)
    
    For Each fileName In fileNames
    
        '�t�@�C�����擾����
        FilePath = data.SaveDirPath & "\" & fileName
        fileProperty.InitFilePath FilePath
        
        '�ꂩ���O�̏ꍇ�t�@�C���폜
        If fileProperty.targetDate < threeMonthAgo Then
            fs.DeleteFile FilePath
        End If
        
    Next fileName
        
End Sub
'�ꂩ���O�̔������̔����ςݏ��i�R�[�h�f�[�^�t�@�C�����폜
Sub DeleteOrderedDataFiles()
    
    '�����f�[�^�V�[�g�A�N�Z�T�̃C���X�^���X��
    Dim ordered As New OrderedDataSheetAccesser
    '�f�[�^�t�@�C���̑������擾�N���X���C���X�^���X��
    Dim fileProperty As New FilePropertyManager
    '�f�[�^�t�@�C�����t�@�C�����ƂɃt�B���^�[����N���X���C���X�^���X��
    Dim filter As New FileFilter
    
    Dim FilePath As String
    Dim fs As New Scripting.FileSystemObject
    
    '�Ώۃf�B���N�g����ݒ�
    filter.DirPath = ordered.SaveDirPath
    
    '�S�Ẵt�@�C�����̎擾
    Dim fileNames As Collection
    Set fileNames = filter.AndFilter()
    
    ' �����̓��t���擾
    Dim today As Date
    today = Date

    ' 1�����O�̓��t���v�Z
    Dim oneMonthAgo As Date
    oneMonthAgo = DateAdd("m", -1, today)
    
    For Each fileName In fileNames
    
        '�t�@�C�����擾����
        FilePath = ordered.SaveDirPath & "\" & fileName
        fileProperty.InitFilePath FilePath
        
        '�ꂩ���O�̏ꍇ�t�@�C���폜
        If fileProperty.targetDate < oneMonthAgo Then
            fs.DeleteFile FilePath
        End If
        
    Next fileName
        
End Sub
