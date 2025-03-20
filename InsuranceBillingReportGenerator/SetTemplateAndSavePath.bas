Attribute VB_Name = "SetTemplateAndSavePath"
Sub SetTemplateAndSavePath()
    Dim ws As Worksheet
    Dim templatePath As String
    Dim saveFolder As String
    Dim storeName As String

    ' �V�[�g���w��i�K�v�ɉ����ĕύX�j
    Set ws = ThisWorkbook.Sheets(1)

    ' B1: �X�ܖ������[�U�[�ɓ��͂�����
    storeName = InputBox("�X�ܖ�����͂��Ă�������", "�X�ܖ��̐ݒ�")
    If storeName = "" Then
        MsgBox "�X�ܖ������͂���Ă��܂���B�����𒆎~���܂��B", vbExclamation
        Exit Sub
    End If
    ws.Range("B1").value = storeName

    ' B2: �e���v���[�g�ۑ��t�H���_��I��
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "�e���v���[�g�ۑ��t�H���_��I�����Ă�������"
        If .Show = -1 Then
            templatePath = .SelectedItems(1)
        Else
            MsgBox "�e���v���[�g�ۑ��t�H���_���I������Ă��܂���B�����𒆎~���܂��B", vbExclamation
            Exit Sub
        End If
    End With
    ws.Range("B2").value = templatePath

    ' B3: �V�K�t�@�C���̕ۑ��t�H���_��I��
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "�V�K�t�@�C���̕ۑ��t�H���_��I�����Ă�������"
        If .Show = -1 Then
            saveFolder = .SelectedItems(1)
        Else
            MsgBox "�V�K�t�@�C���̕ۑ��t�H���_���I������Ă��܂���B�����𒆎~���܂��B", vbExclamation
            Exit Sub
        End If
    End With
    ws.Range("B3").value = saveFolder

    ' �ݒ芮�����b�Z�[�W
    MsgBox "�ݒ肪�������܂����B" & vbCrLf & _
           "�X�ܖ�: " & storeName & vbCrLf & _
           "�e���v���[�g�ۑ���: " & templatePath & vbCrLf & _
           "�V�K�t�@�C���ۑ���: " & saveFolder, vbInformation
End Sub

