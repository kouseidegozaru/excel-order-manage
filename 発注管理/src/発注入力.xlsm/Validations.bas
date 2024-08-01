Attribute VB_Name = "Validations"
Sub SetValidations()
    SetBumonCD
    SetUserCD
    SetDate
End Sub

Private Sub SetBumonCD()
    Dim order As New OrderSheetAccesser
    Dim rng As range
    Set rng = order.BumonCodeRange
    
    With rng.Validation
        .Delete ' �����̃o���f�[�V�������폜
        .add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=1, Formula2:=10000 ' ���l�^�̃o���f�[�V������ǉ�
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

Private Sub SetUserCD()
    Dim order As New OrderSheetAccesser
    Dim rng As range
    Set rng = order.UserCodeRange
    
    With rng.Validation
        .Delete ' �����̃o���f�[�V�������폜
        .add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=1, Formula2:=10000 ' ���l�^�̃o���f�[�V������ǉ�
        .IgnoreBlank = True ' �󔒃Z���𖳎�
        .InCellDropdown = True ' �h���b�v�_�E�����X�g��\��
        .InputTitle = "�S���҃R�[�h"
        .ErrorTitle = "���̓G���["
        .InputMessage = "���l����͂��Ă��������B"
        .ErrorMessage = "���͒l�����l�ł͂���܂���B"
        .ShowInput = True ' ���̓��b�Z�[�W��\��
        .ShowError = True ' �G���[���b�Z�[�W��\��
    End With
End Sub

Private Sub SetDate()
    Dim order As New OrderSheetAccesser
    Dim rng As range
    Set rng = order.TargetDateRange
    
    With rng.Validation
        .Delete ' �����̃o���f�[�V�������폜
        .add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1/1/1900", Formula2:="12/31/2100" ' ���t�^�̃o���f�[�V������ǉ�
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
