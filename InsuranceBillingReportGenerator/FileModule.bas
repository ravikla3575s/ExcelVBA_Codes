Attribute VB_Name = "FileModule"
Option Explicit

' �G���[������ێ�����O���[�o���ϐ�
Public error_response As VbMsgBoxResult

Function SelectCsvFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "CSV�t�H���_��I�����Ă�������"
        If .Show = -1 Then
            SelectCsvFolder = .SelectedItems(1) & "\"
        Else
            MsgBox "�t�H���_���I������܂���ł����B�����𒆎~���܂��B", vbExclamation, "�G���["
            SelectCsvFolder = ""
        End If
    End With
End Function

Function IsFolderEmpty(folder_path As String) As Boolean
    Dim fso_local As Object, folder_obj As Object
    Set fso_local = CreateObject("Scripting.FileSystemObject")
    If Not fso_local.FolderExists(folder_path) Then
        IsFolderEmpty = True
        Exit Function
    End If
    Set folder_obj = fso_local.GetFolder(folder_path)
    If folder_obj.files.count = 0 Then
        IsFolderEmpty = True   ' �t�@�C��������Ȃ�
    Else
        IsFolderEmpty = False  ' �t�@�C�������݂���
    End If
End Function

Function CreateReportFiles(file_system As Object, files As Collection, save_path As String, template_path As String)
    On Error GoTo ErrorHandler
    
    ' �ϐ��̐錾
    Dim file As Object
    Dim billing_year As String, billing_month As String
    Dim era_letter As String, era_year_val As Integer
    Dim report_file_name As String, report_file_path As String
    
    Debug.Print "Starting CreateReportFiles"
    Debug.Print "Template path: " & template_path
    Debug.Print "Save path: " & save_path
    
    ' �e���v���[�g�t�@�C���̑��݊m�F��ǉ�
    If Not file_system.FileExists(template_path) Then
        MsgBox "�e���v���[�g�t�@�C����������܂���B" & vbCrLf & _
               "�p�X: " & template_path & vbCrLf & _
               "�����ӏ�: CreateReportFiles", _
               vbCritical, "�G���["
        Exit Function
    End If
    
    For Each file In files
        On Error GoTo ErrorHandler
        
        Debug.Print "----------------------------------------"
        Debug.Print "Processing file: " & file.Name
        
        ' CSV����N�����擾
        billing_year = "": billing_month = ""
        
        ' �t�@�C����ނɂ���ĔN���擾���@��ς���
        If InStr(LCase(file.Name), "fixf") > 0 Then
            If Len(file.Name) < 25 Then
                MsgBox "FIXF�t�@�C���̌`�����s���ł��B" & vbCrLf & _
                       "�t�@�C����: " & file.Name & vbCrLf & _
                       "�K�v�Ȓ���: 25�����ȏ�", _
                       vbExclamation, "CreateReportFiles - �G���["
                GoTo NextFile
            End If
            billing_year = Mid(file.Name, 18, 4)
            billing_month = Mid(file.Name, 22, 2)
            
        ElseIf InStr(LCase(file.Name), "fmei") > 0 Then
            If Len(file.Name) < 22 Then
                MsgBox "FMEI�t�@�C���̌`�����s���ł��B" & vbCrLf & _
                       "�t�@�C����: " & file.Name & vbCrLf & _
                       "�K�v�Ȓ���: 22�����ȏ�", _
                       vbExclamation, "CreateReportFiles - �G���["
                GoTo NextFile
            End If
            
            Dim era_code As String
            ' �����R�[�h��ݒ�
            era_code = Mid(file.Name, 18, 1)
            era_year_val = CInt(Mid(file.Name, 19, 2))
            billing_month = Mid(file.Name, 21, 2)
            
            Select Case era_code
                Case "5"  ' �ߘa
                    era_letter = "R"
                    billing_year = CStr(2018 + era_year_val)
                Case "4"  ' ����
                    era_letter = "H"
                    billing_year = CStr(1988 + era_year_val)
                Case "3"  ' ���a
                    era_letter = "S"
                    billing_year = CStr(1925 + era_year_val)
                Case "2"  ' �吳
                    era_letter = "T"
                    billing_year = CStr(1911 + era_year_val)
                Case "1"  ' ����
                    era_letter = "M"
                    billing_year = CStr(1867 + era_year_val)
            End Select
        End If
        
        Debug.Print "File processing:"
        Debug.Print "File name: " & file.Name
        Debug.Print "Billing Year/Month: " & billing_year & "/" & billing_month
        
        If billing_year <> "" And billing_month <> "" Then
        
            Dim dispensing_year As Integer, dispensing_month As Integer
            
            dispensing_year = CInt(billing_year)
            If CInt(billing_month) < 2 Then
                dispensing_year = CInt(billing_year) - 1
                dispensing_month = 12
            Else
                dispensing_month = CInt(billing_month) - 1
            End If
            
            ' �񍐏��t�@�C�����𐶐��i�����N�����g�p�j
            report_file_name = GenerateReportFileName(billing_year, billing_month)
            Debug.Print "Generated report file name: " & report_file_name
            
            If report_file_name = "" Then
                MsgBox "�t�@�C�����̐����Ɏ��s���܂����B", vbExclamation, "�G���["
                GoTo NextFile
            End If
            
            report_file_path = save_path & "\" & report_file_name
            
            ' �t�@�C�������݂��Ȃ��ꍇ�̂ݐV�K�쐬
            If Not file_system.FileExists(report_file_path) Then
                Dim report_wb As Workbook
                Set report_wb = Workbooks.Add(template_path)
                
                If Not report_wb Is Nothing Then
                    ' �e���v���[�g����ݒ�i�����N����n���j
                    If SetTemplateInfo(report_wb, CInt(billing_year), CInt(billing_month)) Then
                        Application.DisplayAlerts = False
                        report_wb.SaveAs Filename:=report_file_path, _
                                       FileFormat:=xlOpenXMLWorkbookMacroEnabled, _
                                       Local:=True
                        Application.DisplayAlerts = True
                    End If
                    report_wb.Close SaveChanges:=True
                End If
            End If
        End If
NextFile:
    Next file
    Exit Function

ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in CreateReportFiles"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "Current file: " & IIf(Not file Is Nothing, file.Name, "Unknown")
    Debug.Print "Billing Year/Month: " & billing_year & "/" & billing_month
    Debug.Print "Report file name: " & report_file_name
    Debug.Print "=================================="
    
    error_response = MsgBox("�t�@�C���쐬���ɃG���[���������܂����B�ύX��ۑ����܂����H" & vbCrLf & _
                           "�G���[�ԍ�: " & Err.Number & vbCrLf & _
                           "�G���[���e: " & Err.Description & vbCrLf & _
                           "�t�@�C��: " & IIf(Not file Is Nothing, file.Name, "�s��"), _
                           vbYesNo + vbExclamation)
    
    If error_response = vbYes Then
        If Not report_wb Is Nothing Then
            Debug.Print "Cleaning up: Closing workbook"
            report_wb.Close SaveChanges:=True
            Set report_wb = Nothing
        End If
    End If
    Resume NextFile
End Function

Function SortFileCollection(files As Collection, file_system As Object, file_type As String) As Collection
    Dim sorted_files As New Collection
    Dim file_array() As Object
    Dim i As Long, count As Long
    
    ' Collection�̗v�f�����擾
    count = files.count
    If count = 0 Then
        Set SortFileCollection = sorted_files
        Exit Function
    End If
    
    ' �z���������
    ReDim file_array(1 To count)
    
    ' Collection��Array�ɃR�s�[
    For i = 1 To count
        Set file_array(i) = files(i)
    Next i
    
    ' �o�u���\�[�g�ŔN�����Ƀ\�[�g
    Dim j As Long
    For i = 1 To count - 1
        For j = 1 To count - i
            Dim year1 As Integer, month1 As Integer
            Dim year2 As Integer, month2 As Integer
            
            If GetYearMonthFromFile(file_array(j).Path, file_type, year1, month1) And _
               GetYearMonthFromFile(file_array(j + 1).Path, file_type, year2, month2) Then
                
                ' �N�����������Ĕ�r�i��F202402�j
                If (CStr(year1) & Format(month1, "00")) > (CStr(year2) & Format(month2, "00")) Then
                    ' �������t�̏ꍇ�A�v�f������
                    Dim temp_obj As Object
                    Set temp_obj = file_array(j)
                    Set file_array(j) = file_array(j + 1)
                    Set file_array(j + 1) = temp_obj
                End If
            End If
        Next j
    Next i
    
    ' �\�[�g���ꂽ�z���V����Collection�ɒǉ�
    For i = 1 To count
        sorted_files.Add file_array(i)
    Next i
    
    Set SortFileCollection = sorted_files
End Function

Function GetYearMonthFromFile(file_path As String, file_type As String, ByRef dispensing_year As Integer, ByRef dispensing_month As Integer) As Boolean
    Dim file_name As String, base_name As String
    dispensing_year = 0: dispensing_month = 0
    
    file_name = Right(file_path, Len(file_path) - InStrRev(file_path, "\"))
    base_name = file_name
    If InStr(file_name, ".") > 0 Then
        base_name = Left(file_name, InStrRev(file_name, ".") - 1)
    End If
    
    Debug.Print "Processing file: " & file_name
    
    Select Case file_type
        Case "�����m���"  ' fixf�t�@�C��
            If Len(file_name) >= 25 Then
                Dim billing_year As Integer, billing_month As Integer
                billing_year = CInt(Mid(file_name, 18, 4))
                billing_month = CInt(Mid(file_name, 22, 2))
                
                Debug.Print "Billing year/month from file: " & billing_year & "/" & billing_month
                
                ' ���܌��𐿋�����1�����O�ɐݒ�
                If billing_month = 1 Then
                    dispensing_year = billing_year - 1
                    dispensing_month = 12
                Else
                    dispensing_year = billing_year
                    dispensing_month = billing_month - 1
                End If
                
                Debug.Print "Set dispensing year/month to: " & dispensing_year & "/" & dispensing_month
                GetYearMonthFromFile = True
            End If
            
        Case "�U���z���׏�", "�Ԗߓ���", "�����_�A����"  ' fmei, henr, zogn�t�@�C��
            If Len(base_name) >= 5 Then
                Dim code_part As String
                code_part = Right(base_name, 5)
                If Len(code_part) = 5 And IsNumeric(code_part) Then
                    Dim era_code As String, era_year As Integer
                    era_code = Left(code_part, 1)
                    era_year = CInt(Mid(code_part, 2, 2))
                    billing_month = CInt(Right(code_part, 2))
                    
                    ' �����R�[�h���琼��N���v�Z
                    Select Case era_code
                        Case "5": billing_year = 2018 + era_year  ' �ߘa
                        Case "4": billing_year = 1988 + era_year  ' ����
                        Case "3": billing_year = 1925 + era_year  ' ���a
                        Case "2": billing_year = 1911 + era_year  ' �吳
                        Case "1": billing_year = 1867 + era_year  ' ����
                    End Select
                    
                    Debug.Print "Billing year/month from file: " & billing_year & "/" & billing_month
                    
                    ' ���܌��𐿋�����1�����O�ɐݒ�
                    If billing_month = 1 Then
                        dispensing_year = billing_year - 1
                        dispensing_month = 12
                    Else
                        dispensing_year = billing_year
                        dispensing_month = billing_month - 1
                    End If
                    
                    Debug.Print "Set dispensing year/month to: " & dispensing_year & "/" & dispensing_month
                    GetYearMonthFromFile = True
                End If
            End If
    End Select

    Debug.Print "Final dispensing year/month: " & dispensing_year & "/" & dispensing_month
End Function

Private Sub CreateBackup(file_path As String)
    ' �t�@�C���̃o�b�N�A�b�v���쐬
    ' TODO: �o�b�N�A�b�v�@�\�̎���
End Sub

Function GenerateReportFileName(ByVal billing_year As Integer, ByVal billing_month As Integer) As String
    On Error GoTo ErrorHandler
    
    GenerateReportFileName = ""
    
    ' ���͒l�̌���
    If billing_year < 1900 Or billing_year > 9999 Then
        MsgBox "�����N�������ł��B" & vbCrLf & _
               "�N: " & billing_year & vbCrLf & _
               "�����ӏ�: GenerateReportFileName", _
               vbExclamation, "�G���["
        GenerateReportFileName = ""
        Exit Function
    End If
    
    If billing_month < 1 Or billing_month > 12 Then
        MsgBox "�������������ł��B" & vbCrLf & _
               "��: " & billing_month & vbCrLf & _
               "�����ӏ�: GenerateReportFileName", _
               vbExclamation, "�G���["
        GenerateReportFileName = ""
        Exit Function
    End If
    
    Debug.Print "GenerateReportFileName input:"
    Debug.Print "Billing year: " & billing_year
    Debug.Print "Billing month: " & billing_month
    
    Dim dispensing_year As Integer, dispensing_month As Integer
    
    If billing_month < 2 Then
        dispensing_year = billing_year - 1
        dispensing_month = 12
    Else
        dispensing_year = billing_year
        dispensing_month = billing_month - 1
    End If
    
    ' ���������擾
    Dim era_info As Object
    Set era_info = DateConversionModule.ConvertEraYear(dispensing_year, True)
    
    If era_info Is Nothing Then
        MsgBox "�����̕ϊ��Ɏ��s���܂����B" & vbCrLf & _
               "�N: " & billing_year, _
               vbExclamation, "GenerateReportFileName - �G���["
        GenerateReportFileName = ""
        Exit Function
    End If
    
    
    ' �t�@�C�����𐶐�
    GenerateReportFileName = "�ی������Ǘ��񍐏�_" & _
                            era_info("era") & _
                            Format(era_info("year"), "00") & "�N" & _
                            Format(dispensing_month, "00") & "��.xlsm"
                            
    Debug.Print "Generated filename: " & GenerateReportFileName
    Exit Function
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error occurred in GenerateReportFileName"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    error_response = MsgBox("�t�@�C�����̐������ɃG���[���������܂����B�ύX��ۑ����܂����H" & vbCrLf & _
                           "�G���[�ԍ�: " & Err.Number & vbCrLf & _
                           "�G���[���e: " & Err.Description, _
                           vbYesNo + vbExclamation)
    
    If error_response = vbYes Then
        GenerateReportFileName = ""
    End If
End Function

Function ProcessCsvFilesByType(file_system As Object, csv_files As Collection, file_type_name As String)
    On Error GoTo ErrorHandler
    
    Dim file_obj As Object
    Dim report_file_name As String, report_file_path As String
    Dim base_name As String, sheet_name As String
    Dim report_wb As Workbook
    Dim sheet_exists As Boolean
    Dim dispensing_year As Integer, dispensing_month As Integer
    Dim insert_index As Long
    
    For Each file_obj In csv_files
        Dim save_successful As Boolean
        save_successful = False  ' �ۑ��t���O��������
        
        Debug.Print "----------------------------------------"
        Debug.Print "Processing file: " & file_obj.Name
        Debug.Print "File type: " & file_type_name
        Debug.Print "File path: " & file_obj.Path
        
        ' CSV�t�@�C�������璲�ܔN�����擾
        If Not GetYearMonthFromFile(file_obj.Path, file_type_name, dispensing_year, dispensing_month) Then
            Debug.Print "ERROR: Failed to get year/month from file"
            MsgBox "�t�@�C�� " & file_obj.Name & " ���璲�ܔN�����擾�ł��܂���ł����B", vbExclamation, "�G���["
            GoTo NextFile
        End If

        Debug.Print "Dispensing year/month: " & dispensing_year & "/" & dispensing_month
        
        Dim billing_year As Integer, billing_month As Integer
        
        If dispensing_month = 12 Then
            billing_year = dispensing_year + 1
            billing_month = 1
        Else
            billing_year = dispensing_year
            billing_month = dispensing_month + 1
        End If
        
        ' �񍐏��t�@�C�����𐶐�
        report_file_name = GenerateReportFileName(billing_year, billing_month)
        Debug.Print "Generated report file name: " & report_file_name
        
        If report_file_name = "" Then
            Debug.Print "ERROR: Failed to generate report file name"
            GoTo NextFile
        End If

        report_file_path = save_path & "\" & report_file_name
        Debug.Print "Full report file path: " & report_file_path
        
        ' �t�@�C���̑��݊m�F
        If Not file_system.FileExists(report_file_path) Then
            Debug.Print "ERROR: Report file does not exist: " & report_file_path
            GoTo NextFile
        End If
        
        ' ���[�N�u�b�N���J��
        On Error Resume Next
        Set report_wb = Workbooks.Open(report_file_path, ReadOnly:=False, UpdateLinks:=False)
        If Err.Number <> 0 Then
            Debug.Print "ERROR: Failed to open workbook"
            Debug.Print "Error number: " & Err.Number
            Debug.Print "Error description: " & Err.Description
            On Error GoTo ErrorHandler
            GoTo NextFile
        End If
        On Error GoTo ErrorHandler
        
        If report_wb Is Nothing Then
            Debug.Print "ERROR: Failed to open workbook (report_wb is Nothing)"
            GoTo NextFile
        End If
        
        Debug.Print "Successfully opened workbook"
        
        ' CSV�f�[�^���C���|�[�g���ĐV�K�V�[�g�ɓ]�L
        base_name = file_system.GetBaseName(file_obj.Name)
        sheet_name = base_name
        Debug.Print "Base sheet name: " & sheet_name
        
        ' �V�[�g���̏d���`�F�b�N�ƈ�ӂ̖��O����
        Dim sheet_index As Integer
        sheet_index = 1
        
        On Error Resume Next
        Do
            sheet_exists = False
            Dim test_ws As Worksheet
            Set test_ws = report_wb.Sheets(sheet_name)
            If Not test_ws Is Nothing Then
                sheet_exists = True
                sheet_name = base_name & "_" & Format(sheet_index, "00")
                sheet_index = sheet_index + 1
                Debug.Print "Sheet exists, trying new name: " & sheet_name
            End If
        Loop While sheet_exists
        On Error GoTo ErrorHandler
        
        Debug.Print "Final sheet name: " & sheet_name
        
        ' �V�K�V�[�g�̒ǉ�
        insert_index = Application.WorksheetFunction.Min(3, report_wb.Sheets.count)
        Debug.Print "Insert index: " & insert_index
        
        On Error Resume Next
        Dim ws_csv As Worksheet
        Set ws_csv = report_wb.Sheets.Add(After:=report_wb.Sheets(insert_index))
        If Err.Number <> 0 Then
            Debug.Print "ERROR: Failed to add new sheet"
            Debug.Print "Error number: " & Err.Number
            Debug.Print "Error description: " & Err.Description
            GoTo NextFile
        End If
        On Error GoTo ErrorHandler
        
        If ws_csv Is Nothing Then
            Debug.Print "ERROR: Failed to create new sheet (ws_csv is Nothing)"
            GoTo NextFile
        End If
        
        ws_csv.Name = sheet_name
        Debug.Print "Successfully created and named new sheet"
        
        ' �G���[����������\���̂��鏈���̑O�� On Error Resume Next
        On Error Resume Next
        
        ' �����������������ǂ������m�F
        Dim process_error As Boolean
        process_error = False
        
        ' CSV�f�[�^�̃C���|�[�g
        If file_type_name = "�����m���" Then
            ImportCsvData file_obj.Path, ws_csv, file_type_name, True
        Else
            ImportCsvData file_obj.Path, ws_csv, file_type_name, False
        End If
        
        If Err.Number <> 0 Then
            Debug.Print "ERROR in ImportCsvData: " & Err.Description
            process_error = True
        End If
        
        ' �G���[�����Z�b�g
        Err.Clear
        
        ' �ڍ׃f�[�^���ڍ׃V�[�g�ɔ��f
        If Not process_error Then
            Call TransferBillingDetails(report_wb, file_obj.Name, CStr(dispensing_year), _
                                      Format(dispensing_month, "00"), _
                                      (file_type_name = "�����m���"))
            
            If Err.Number <> 0 Then
                Debug.Print "ERROR in TransferBillingDetails: " & Err.Description
                process_error = True
            End If
        End If
        
        ' �G���[���������ɖ߂�
        On Error GoTo ErrorHandler
        
        ' ���������������ꍇ�̂ݕۑ�
        If Not process_error Then
            Debug.Print "Processing completed successfully, saving workbook"
            report_wb.Save
            save_successful = True
        Else
            Debug.Print "Processing encountered errors, changes will not be saved"
        End If

NextFile:
        If Not report_wb Is Nothing Then
            Debug.Print "Cleaning up: Closing workbook"
            ' �G���[�������������A�d�v�ȕύX������ꍇ�͕ۑ����邩�ǂ��������[�U�[�Ɋm�F
            error_response = MsgBox("�G���[���������܂����B�ύX��ۑ����܂����H" & vbCrLf & _
                                  "�G���[���e: " & Err.Description, _
                                  vbYesNo + vbExclamation)
            report_wb.Close SaveChanges:=(error_response = vbYes)
            Set report_wb = Nothing
        End If
        Set ws_csv = Nothing
        Debug.Print "----------------------------------------"
    Next file_obj
    Exit Function

ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error occurred in ProcessCsvFilesByType"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "Current file: " & IIf(Not file_obj Is Nothing, file_obj.Name, "Unknown")
    Debug.Print "Current report file: " & report_file_name
    Debug.Print "Current report path: " & report_file_path
    Debug.Print "Current sheet name: " & sheet_name
    Debug.Print "Insert index: " & insert_index
    Debug.Print "Dispensing year: " & dispensing_year
    Debug.Print "Dispensing month: " & dispensing_month
    Debug.Print "File type: " & file_type_name
    Debug.Print "=================================="
    
    error_response = MsgBox("�������ɃG���[���������܂����B�ύX��ۑ����܂����H" & vbCrLf & _
                           "�G���[�ԍ�: " & Err.Number & vbCrLf & _
                           "�G���[���e: " & Err.Description & vbCrLf & _
                           "�t�@�C��: " & IIf(Not file_obj Is Nothing, file_obj.Name, "�s��"), _
                           vbYesNo + vbExclamation)
    
    If error_response = vbYes Then
        If Not report_wb Is Nothing Then
            Debug.Print "Cleaning up: Closing workbook"
            report_wb.Close SaveChanges:=True
            Set report_wb = Nothing
        End If
    End If
    Resume NextFile
End Function

Private Function SetTemplateInfo(ByVal wb As Workbook, ByVal billing_year As Integer, ByVal billing_month As Integer) As Boolean
    On Error GoTo ErrorHandler
    
    ' ���C���V�[�g�i1���ځj�Əڍ׃V�[�g�i2���ځj���擾
    Dim ws_main As Worksheet, ws_sub As Worksheet
    Set ws_main = wb.Sheets(1)
    Set ws_sub = wb.Sheets(2)
    
    ' ���ܔN�����v�Z�i��������1�����O�����܌��j
    Dim dispensing_year As Integer, dispensing_month As Integer
    If billing_month = 1 Then
        dispensing_year = billing_year - 1
        dispensing_month = 12
    Else
        dispensing_year = billing_year
        dispensing_month = billing_month - 1
    End If
    
    ' �������̐ݒ�
    Dim send_date As String
    send_date = billing_month & "��10��������"
    
    ' �X�ܖ��̎擾
    Dim store_name As String
    On Error Resume Next
    ' ThisWorkbook�ɐݒ�V�[�g�����邩�m�F
    Dim ws_settings As Worksheet
    Set ws_settings = ThisWorkbook.Worksheets("�ݒ�")
    If Not ws_settings Is Nothing Then
        store_name = ws_settings.Range("B1").Value
    Else
        ' �ݒ�V�[�g���Ȃ��ꍇ�̓��C�����W���[������p�X���擾
        store_name = ""
    End If
    On Error GoTo ErrorHandler
    
    ' �ߘa�N���v�Z
    Dim era_info As Object
    Set era_info = DateConversionModule.ConvertEraYear(dispensing_year, True)
    
    ' �V�[�g����ݒ�
    Dim era_year As String
    era_year = CStr(era_info("year"))
    
    ' �V�[�g����ύX
    On Error Resume Next
    ws_main.Name = "R" & era_year & "." & dispensing_month
    ws_sub.Name = UtilityModule.ConvertToCircledNumber(dispensing_month)
    On Error GoTo ErrorHandler
    
    ' �e���v���[�g�̔N����ݒ�
    With ws_main
        ' �N���̐ݒ�
        .Range("G2").Value = dispensing_year & "�N" & dispensing_month & "�����ܕ�"
        .Range("I2").Value = send_date
        .Range("J2").Value = store_name
    End With
    
    With ws_sub
        .Range("H1").Value = dispensing_year & "�N" & dispensing_month & "�����ܕ�"
        .Range("J1").Value = send_date
        .Range("L1").Value = store_name
    End With
    
    ' �������V�[�g��ǉ�
    Dim ws_unclaimed As Worksheet
    On Error Resume Next
    Set ws_unclaimed = wb.Worksheets("�������ꗗ")
    If ws_unclaimed Is Nothing Then
        Set ws_unclaimed = wb.Worksheets.Add(After:=ws_sub)
        ws_unclaimed.Name = "�������ꗗ"
        
        ' �������V�[�g�̊�{�ݒ�
        With ws_unclaimed
            .Range("A1").Value = "���������Z�v�g�ꗗ"
            .Range("A1").Font.Size = 14
            .Range("A1").Font.Bold = True
            
            ' ��w�b�_�[
            .Range("A2").Value = "���Ҏ���"
            .Range("B2").Value = "���ܔN��"
            .Range("C2").Value = "��Ë@��"
            .Range("D2").Value = "���������R"
            .Range("E2").Value = "������"
            .Range("F2").Value = "�ی�����"
            .Range("G2").Value = "�����_��"
            .Range("H2").Value = "���l"
            
            ' ��w�b�_�[�̊�{�t�H�[�}�b�g
            .Range("A2:H2").Font.Bold = True
            
            ' �񕝂̊�{����
            .Columns("A:D").ColumnWidth = 15
            .Columns("E").ColumnWidth = 10
            .Columns("F:G").ColumnWidth = 12
            .Columns("H").ColumnWidth = 25
        End With
    End If
    On Error GoTo ErrorHandler
    
    SetTemplateInfo = True
    Exit Function

ErrorHandler:
    Debug.Print "Error in SetTemplateInfo"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "Billing year: " & billing_year
    Debug.Print "Billing month: " & billing_month
    Debug.Print "Dispensing year: " & dispensing_year
    Debug.Print "Dispensing month: " & dispensing_month
    
    error_response = MsgBox("�e���v���[�g���̐ݒ蒆�ɃG���[���������܂����B�ύX��ۑ����܂����H" & vbCrLf & _
                           "�G���[�ԍ�: " & Err.Number & vbCrLf & _
                           "�G���[���e: " & Err.Description, _
                           vbYesNo + vbExclamation)
    
    SetTemplateInfo = (error_response = vbYes)
End Function

