Attribute VB_Name = "Validations"
Sub SetValidations()
    SetBumonCD
    SetUserCD
    SetDate
End Sub

Private Sub SetBumonCD()
    Dim order As New OrderSheetAccesser
    Dim rng As Range
    Set rng = order.BumonCodeRange
    
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

Private Sub SetUserCD()
    Dim order As New OrderSheetAccesser
    Dim rng As Range
    Set rng = order.UserCodeRange
    
    With rng.Validation
        .Delete ' �����̃o���f�[�V�������폜
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=1, Formula2:=10000 ' ���l�^�̃o���f�[�V������ǉ�
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
    Dim rng As Range
    Set rng = order.TargetDateRange
    
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

'���킹���̃o���f�[�V�����`�F�b�N
Public Function IsMatchQty() As Boolean
    Dim order As New OrderSheetAccesser
    Dim i As Long
    
    '����
    Dim qtyCol As Collection
    '���킹��
    Dim matchCol As Collection
    
    Set qtyCol = order.qty
    Set matchCol = order.match
    
    IsMatchQty = True
    
    For i = 1 To matchCol.count
        If Not IsMultiple(qtyCol(i), matchCol(i)) Then
            IsMatchQty = False
            Exit For
        End If
    Next i
    
End Function

'���L�t�H���_�ւ̃A�N�Z�X���������邩�`�F�b�N
Public Sub CheckDirPermission()
    Dim data As New DataSheetAccesser
    If Not CheckDirectoryAccess(data.SaveDirPath) Then
        MsgBox "���L�t�H���_�ւ̃A�N�Z�X����������܂���B�g�p����ɂ͏��ۂֈ˗����Ă�������"
        End
    End If
End Sub
