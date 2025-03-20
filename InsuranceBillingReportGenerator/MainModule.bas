Attribute VB_Name = "MainModule"
Option Explicit

' �萔��`
Public Const MAX_LINES_PER_SHEET As Long = 40
Public Const REQUIRED_SHEETS_COUNT As Integer = 6
Public Const BASE_DETAIL_ROWS As Integer = 4

' �e���v���[�g�E�ۑ���p�X
Public template_path As String
Public save_path As String

Sub CreateReportsFromCSV()
    On Error GoTo ErrorHandler
    
    ' �p�X�̐ݒ�
    template_path = ThisWorkbook.Sheets(1).Range("B2").value & "\�ی������Ǘ��񍐏��e���v���[�g20250320.xltm"
    save_path = ThisWorkbook.Sheets(1).Range("B3").value
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim csv_folder As String            ' CSV�t�H���_�p�X
    Dim file_system As Object          ' FileSystemObject
    Dim billing_year As String, billing_month As String  ' �����Ώۂ̐f�ÔN�E���i����j
    Dim fixf_files As New Collection, fmei_files As New Collection
    Dim henr_files As New Collection, zogn_files As New Collection
    Dim file_obj As Object

    ' 1. CSV�t�H���_�����[�U�[�ɑI��������
    csv_folder = SelectCsvFolder()
    If csv_folder = "" Then Exit Sub  ' ���[�U�[���L�����Z�������ꍇ

    ' 2. �t�H���_����Ȃ珈���𒆎~
    If IsFolderEmpty(csv_folder) Then
        MsgBox "�I�������t�H���_�ɂ�CSV�t�@�C��������܂���B�����𒆎~���܂��B", vbExclamation, "�G���["
        Exit Sub
    End If

    ' 3. �e���v���[�g�p�X�E�ۑ���t�H���_�̑��݊m�F
    If template_path = "" Or save_path = "" Then
        MsgBox "�e���v���[�g�p�X�܂��͕ۑ���t�H���_���ݒ肳��Ă��܂���B", vbExclamation, "�G���["
        Exit Sub
    End If

    ' 4. FileSystemObject�̗p��
    Set file_system = CreateObject("Scripting.FileSystemObject")

    ' 4. �t�H���_���̑SCSV�t�@�C������ޕʂɎ��W�ifixf, fmei, henr, zogn�j
    For Each file_obj In file_system.GetFolder(csv_folder).files
        If LCase(file_system.GetExtensionName(file_obj.Name)) = "csv" Then
            If InStr(LCase(file_obj.Name), "fixf") > 0 Then
                fixf_files.Add file_obj
                Set fixf_files = SortFileCollection(fixf_files, file_system, "fixf")
            ElseIf InStr(LCase(file_obj.Name), "fmei") > 0 Then
                fmei_files.Add file_obj
                Set fmei_files = SortFileCollection(fmei_files, file_system, "fmei")
            ElseIf InStr(LCase(file_obj.Name), "henr") > 0 Then
                henr_files.Add file_obj
                Set henr_files = SortFileCollection(henr_files, file_system, "henr")
            ElseIf InStr(LCase(file_obj.Name), "zogn") > 0 Then
                zogn_files.Add file_obj
                Set zogn_files = SortFileCollection(zogn_files, file_system, "zogn")
            End If
        End If
    Next file_obj

    ' 6. �Ώ�CSV�t�@�C��������Ȃ��ꍇ�A�����𒆎~
    If fixf_files.count = 0 And fmei_files.count = 0 And henr_files.count = 0 And zogn_files.count = 0 Then
        MsgBox "�I�������t�H���_�ɂ͏����Ώۂ�CSV�t�@�C��������܂���B�����𒆎~���܂��B", vbExclamation, "�G���["
        Exit Sub
    End If

    ' 7. fixf�t�@�C����fmei�t�@�C���̗L���ɂ�鏈������
    If fixf_files.count > 0 Then
        CreateReportFiles file_system, fixf_files, save_path, template_path
    End If
    If fmei_files.count > 0 Then
        CreateReportFiles file_system, fmei_files, save_path, template_path
    End If

    ' 8. �e�햾��CSV�ifmei, henr, zogn�j�̏���
    FileModule.ProcessCsvFilesByType file_system, fixf_files, "�����m���"
    FileModule.ProcessCsvFilesByType file_system, fmei_files, "�U���z���׏�"
    FileModule.ProcessCsvFilesByType file_system, henr_files, "�Ԗߓ���"
    FileModule.ProcessCsvFilesByType file_system, zogn_files, "�����_�A����"
    
    ' 9. �������b�Z�[�W
    MsgBox "CSV�t�@�C���̏������������܂����I", vbInformation, "����"

    ' �I�u�W�F�N�g�̉��������ǉ�
    Set file_system = Nothing
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error occurred in CreateReportsFromCSV"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    MsgBox "�������ɃG���[���������܂����B" & vbCrLf & _
           "�G���[�ԍ�: " & Err.Number & vbCrLf & _
           "�G���[���e: " & Err.Description, _
           vbCritical, "�G���["
           
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

' ���ܔN�����擾����֐��i�N�ƌ���ʁX�Ɏ擾�j
Private Sub GetDispensingYearMonth(file_path As String, ByRef year_out As String, ByRef month_out As String)
    Dim file_name As String
    Dim billing_year As Integer, billing_month As Integer
    
    file_name = Right(file_path, Len(file_path) - InStrRev(file_path, "\"))
    year_out = ""
    month_out = ""
    
    ' FIXF�t�@�C�����琿���N���𒊏o
    If Len(file_name) >= 23 Then
        billing_year = CInt(Mid(file_name, 18, 4))
        billing_month = CInt(Mid(file_name, 22, 2))
        
        ' ���������璲�ܔN�����v�Z�i1���O�j
        If billing_month = 1 Then
            year_out = CStr(billing_year - 1)
            month_out = "12"
        Else
            year_out = CStr(billing_year)
            month_out = Format(billing_month - 1, "00")
        End If
    End If
End Sub

