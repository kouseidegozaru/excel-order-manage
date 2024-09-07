Attribute VB_Name = "Validations"
'�Z���Ƀo���f�[�V�����`�F�b�N��ݒ�
Sub SetValidations()
    SetBumonCD
    SetDate
End Sub

Private Sub SetBumonCD()
    Dim load As New LoadSheetAccesser
    Dim rng As Range
    Set rng = load.BumonCodeRange
    
    With rng.Validation
        .Delete ' �����̃o���f�[�V�������폜
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=1, Formula2:=10000 ' ���l�^�̃o���f�[�V������ǉ�
        .IgnoreBlank = True ' �󔒃Z���𖳎�
        .InCellDropdown = True ' �h���b�v�_�E�����X�g��\��
        .InputTitle = "����R�[�h"
        .ErrorTitle = "���̓G���["
        .InputMessage = "���l����͂��Ă��������B"
        .ErrorMessage = "���͒l�����l�ł͂���܂���B"
        .ShowInput = True ' ���̓��b�Z�[�W��\��
        .ShowError = True ' �G���[���b�Z�[�W��\��
    End With
End Sub

Private Sub SetDate()
    Dim load As New LoadSheetAccesser
    Dim rng As Range
    Set rng = load.TargetDateRange
    
    With rng.Validation
        .Delete ' �����̃o���f�[�V�������폜
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1/1/1900", Formula2:="12/31/2100" ' ���t�^�̃o���f�[�V������ǉ�
        .IgnoreBlank = True ' �󔒃Z���𖳎�
        .InCellDropdown = True ' �h���b�v�_�E�����X�g��\��
        .InputTitle = "�������t"
        .ErrorTitle = "���̓G���["
        .InputMessage = "���t����͂��Ă��������B"
        .ErrorMessage = "���͒l���L���ȓ��t�ł͂���܂���B"
        .ShowInput = True ' ���̓��b�Z�[�W��\��
        .ShowError = True ' �G���[���b�Z�[�W��\��
    End With
End Sub

'���L�t�H���_�ւ̃A�N�Z�X���������邩�`�F�b�N
Public Sub CheckDirPermission()
    Dim data As New DataSheetAccesser
    If Not CheckDirectoryAccess(data.SaveDirPath) Then
        MsgBox "���L�t�H���_�ւ̃A�N�Z�X����������܂���B�g�p����ɂ͏��ۂֈ˗����Ă�������"
        End
    End If
End Sub


'���͒l�̓��I�ȃo���f�[�V�����`�F�b�N

'����R�[�h�̓��͒l���`�F�b�N
Public Sub CheckExistsBumon(bumonCode As Variant)
    
    '��̏ꍇ
    If IsEmpty(bumonCode) Then
        End
    End If
    
    '���l�łȂ��ꍇ
    If Not IsNumeric(bumonCode) Then
        MsgBox ("���l����͂��ĉ�����")
        End
    End If

    '����R�[�h�����݂��邩
    Dim dataStorage As New DataBaseAccesser
    If Not dataStorage.ExistsBumon(bumonCode) Then
        MsgBox ("����������R�[�h����͂��ĉ�����")
        End
    End If
    
End Sub

'�������̓��͒l���`�F�b�N
Public Sub CheckDateFormat(targetDate As Variant)
    
    '��̏ꍇ
    If IsEmpty(targetDate) Then
        End
    End If
    
    '���t�łȂ��ꍇ
    If Not IsDate(targetDate) Then
        MsgBox ("���t����͂��ĉ�����")
        End
    End If
    
End Sub
