Attribute VB_Name = "Validations"
Sub SetValidations()
    '�Z���Ƀo���f�[�V�����`�F�b�N��ݒ�
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


'���͒l�̓��I�ȃo���f�[�V�����`�F�b�N

'�S���҃R�[�h�̓��͒l���`�F�b�N
Public Sub CheckExistsUser(userCode As Variant)
    
    '��̏ꍇ
    If IsEmpty(userCode) Then
        End
    End If
    
    '���l�łȂ��ꍇ
    If Not IsNumeric(userCode) Then
        MsgBox ("���l����͂��ĉ�����")
        End
    End If

    '�S���҃R�[�h�����݂��邩
    Dim dataStorage As New DataBaseAccesser
    If Not dataStorage.ExistsUser(userCode) Then
        MsgBox ("�������S���҃R�[�h����͂��ĉ�����")
        End
    End If
    
End Sub

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
    
    '���݂̎�����9�����߂��Ă���ꍇ
    If Time >= #9:00:00 AM# Then
        '�������������̓��t�̏ꍇ
        If DateValue(targetDate) = Date Then
            MsgBox ("���݂̎�����9�����߂��Ă��邽�߁A�����̓��t�̔����͂ł��܂���B")
            End
        End If
    End If
    
End Sub
