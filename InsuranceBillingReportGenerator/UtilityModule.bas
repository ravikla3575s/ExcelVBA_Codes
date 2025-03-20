Attribute VB_Name = "UtilityModule"
Option Explicit

' �����ԏ����̐i���\��
Private Sub UpdateProgress(current As Long, total As Long, message As String)
    Application.StatusBar = message & " - " & current & "/" & total
End Sub

' �I�u�W�F�N�g����p�̊֐�
Private Sub CleanupObjects(ParamArray objects() As Variant)
    Dim obj As Variant
    For Each obj In objects
        If Not obj Is Nothing Then
            If TypeName(obj) = "Workbook" Then
                obj.Close SaveChanges:=False
            End If
            Set obj = Nothing
        End If
    Next obj
End Sub

' �V�[�g�̑��݃`�F�b�N�p�̊֐�
Private Function SheetExists(wb As Workbook, sheet_name As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = wb.Sheets(sheet_name)
    On Error GoTo 0
    
    SheetExists = Not ws Is Nothing
End Function

' �������ەt�������ɕϊ�����֐�
Function ConvertToCircledNumber(ByVal month As Integer) As String
    Dim circled_numbers As Variant
    circled_numbers = Array("", "�@", "�A", "�B", "�C", "�D", "�E", "�F", "�G", "�H", "�I", "�J", "�K")  ' �C���f�b�N�X0�ɋ󕶎���ǉ�
    If month >= 1 And month <= 12 Then
        ConvertToCircledNumber = circled_numbers(month)  ' ���̂܂܂�month���g�p
    Else
        ConvertToCircledNumber = CStr(month)
    End If
End Function

' �J�e�S���̊J�n�s���擾����֐�
Public Function GetStartRow(ws As Worksheet, category_name As String) As Long
    Dim found_cell As Range
    Set found_cell = ws.Cells.Find(what:=category_name, LookAt:=xlWhole)
    If Not found_cell Is Nothing Then
        GetStartRow = found_cell.row
    Else
        GetStartRow = 0
    End If
End Function

' �}�[�L���O���ꂽ�J�n�s����������֐�
Public Function FindMarkedRow(ws As Worksheet, marker As String) As Long
    Dim found_cell As Range
    Dim search_marker As String
    
    ' �}�[�L���O�̃t�H�[�}�b�g���m�F�i����<<>>���܂܂�Ă��邩�ǂ����j
    If Left(marker, 2) <> "<<" Then
        search_marker = "<<" & marker & ">>"
    Else
        search_marker = marker
    End If
    
    ' �V�[�g�S�̂�����
    Set found_cell = ws.Cells.Find(what:=search_marker, LookAt:=xlPart, MatchCase:=False)
    
    If Not found_cell Is Nothing Then
        FindMarkedRow = found_cell.Row
        Debug.Print "�}�[�J�[ '" & search_marker & "' ���s " & FindMarkedRow & " �Ŕ������܂���"
    Else
        FindMarkedRow = 0
        Debug.Print "�}�[�J�[ '" & search_marker & "' �͌�����܂���ł���"
    End If
End Function

' �J�e�S���̊J�n�s���擾����֐� - �}�[�L���O�x�[�X�̐V�o�[�W����
Public Function GetCategoryStartRowsFromMarkers(ws As Worksheet, payer_type As String) As Object
    Dim start_row_dict As Object
    Set start_row_dict = CreateObject("Scripting.Dictionary")
    
    Debug.Print "�}�[�L���O����J�n�s���������Ă��܂�: " & payer_type
    
    ' �Еۊ֘A�̃}�[�J�[������
    Dim shaho_saiseikyu As Long
    Dim shaho_tsukiokure As Long
    Dim shaho_tsukinokuri As Long
    Dim shaho_henrei As Long
    Dim shaho_miseikyuu As Long
    
    ' ���ۊ֘A�̃}�[�J�[������
    Dim kokuho_saiseikyu As Long
    Dim kokuho_tsukiokure As Long
    Dim kokuho_tsukinokuri As Long
    Dim kokuho_henrei As Long
    Dim kokuho_miseikyuu As Long
    
    ' ���֘A�̃}�[�J�[������
    Dim kaigo_henrei As Long
    
    ' ���̑��̃}�[�J�[������
    Dim sonota As Long
    
    ' �e�}�[�J�[�̍s������
    shaho_saiseikyu = FindMarkedRow(ws, "�ЕۍĐ���")
    shaho_tsukiokure = FindMarkedRow(ws, "�Еی��x��")
    shaho_tsukinokuri = FindMarkedRow(ws, "�Еی�����")
    shaho_henrei = FindMarkedRow(ws, "�ЕەԖ�")
    shaho_miseikyuu = FindMarkedRow(ws, "�Еۖ���������")
    
    kokuho_saiseikyu = FindMarkedRow(ws, "���ۍĐ���")
    kokuho_tsukiokure = FindMarkedRow(ws, "���ی��x��")
    kokuho_tsukinokuri = FindMarkedRow(ws, "���ی�����")
    kokuho_henrei = FindMarkedRow(ws, "���ەԖ�")
    kokuho_miseikyuu = FindMarkedRow(ws, "���ۖ���������")
    
    kaigo_henrei = FindMarkedRow(ws, "���Ԗ�")
    
    sonota = FindMarkedRow(ws, "���̑�")
    
    ' ������^�C�v�Ɋ�Â��ăf�B�N�V���i�����\�z
    If payer_type = "�Е�" Then
        If shaho_saiseikyu > 0 Then
            start_row_dict.Add "�Đ���", shaho_saiseikyu
        End If
        
        If shaho_tsukiokure > 0 Then
            start_row_dict.Add "���x�ꐿ��", shaho_tsukiokure
        End If
        
        If shaho_tsukinokuri > 0 Then
            start_row_dict.Add "������", shaho_tsukinokuri
        End If
        
        If shaho_henrei > 0 Then
            start_row_dict.Add "�Ԗ߁E����", shaho_henrei
        End If
        
        If shaho_miseikyuu > 0 Then
            start_row_dict.Add "����������", shaho_miseikyuu
        End If
    ElseIf payer_type = "����" Then
        If kokuho_saiseikyu > 0 Then
            start_row_dict.Add "�Đ���", kokuho_saiseikyu
        End If
        
        If kokuho_tsukiokure > 0 Then
            start_row_dict.Add "���x�ꐿ��", kokuho_tsukiokure
        End If
        
        If kokuho_tsukinokuri > 0 Then
            start_row_dict.Add "������", kokuho_tsukinokuri
        End If
        
        If kokuho_henrei > 0 Then
            start_row_dict.Add "�Ԗ߁E����", kokuho_henrei
        End If
        
        If kokuho_miseikyuu > 0 Then
            start_row_dict.Add "����������", kokuho_miseikyuu
        End If
    ElseIf payer_type = "���" Then
        If kaigo_henrei > 0 Then
            start_row_dict.Add "�Ԗ�", kaigo_henrei
        End If
    End If
    
    ' ���̑��͋���
    If sonota > 0 Then
        start_row_dict.Add "���̑�", sonota
    End If
    
    ' �f�B�N�V���i������̏ꍇ�i�}�[�J�[��������Ȃ��ꍇ�j
    If start_row_dict.Count = 0 Then
        Debug.Print "WARNING: �}�[�L���O��������܂���ł����B�]���̕��@�ŃJ�e�S���J�n�s���擾���܂��B"
        Set start_row_dict = GetCategoryStartRows(ws, payer_type)
    End If
    
    Set GetCategoryStartRowsFromMarkers = start_row_dict
End Function

' �]���̃J�e�S���̊J�n�s���擾����֐��i�o�b�N�A�b�v�Ƃ��Ďc���j
Public Function GetCategoryStartRows(ws As Worksheet, payer_type As String) As Object
    Dim start_row_dict As Object
    Set start_row_dict = CreateObject("Scripting.Dictionary")
    
    Debug.Print "�]���̕��@�ŊJ�n�s���������Ă��܂�: " & payer_type
    
    If payer_type = "�Е�" Then
        Dim social_start_row As Long
        social_start_row = GetStartRow(ws, "�ЕەԖߍĐ���")
        
        If social_start_row > 0 Then
            start_row_dict.Add "�Đ���", social_start_row
            start_row_dict.Add "���x�ꐿ��", GetStartRow(ws, "�Еی��x�ꐿ��")
            start_row_dict.Add "�Ԗ߁E����", GetStartRow(ws, "�ЕەԖ߁E����")
            start_row_dict.Add "����������", GetStartRow(ws, "�Еۖ���������")
        Else
            ' ���o����������Ȃ��ꍇ�̃f�t�H���g�l��ݒ�
            Debug.Print "�Еۂ̌��o����������܂���B�f�t�H���g�l���g�p���܂��B"
            start_row_dict.Add "�Đ���", 3  ' �f�t�H���g�J�n�s
            start_row_dict.Add "���x�ꐿ��", 8
            start_row_dict.Add "�Ԗ߁E����", 13
            start_row_dict.Add "����������", 18
        End If
    ElseIf payer_type = "����" Then
        Dim kokuho_start_row As Long
        kokuho_start_row = GetStartRow(ws, "���ەԖߍĐ���")
        
        If kokuho_start_row > 0 Then
            start_row_dict.Add "�Đ���", kokuho_start_row
            start_row_dict.Add "���x�ꐿ��", GetStartRow(ws, "���ی��x�ꐿ��")
            start_row_dict.Add "�Ԗ߁E����", GetStartRow(ws, "���ەԖ߁E����")
            start_row_dict.Add "����������", GetStartRow(ws, "���ۖ���������")
        Else
            ' ���o����������Ȃ��ꍇ�̃f�t�H���g�l��ݒ�
            Debug.Print "���ۂ̌��o����������܂���B�f�t�H���g�l���g�p���܂��B"
            start_row_dict.Add "�Đ���", 23  ' �f�t�H���g�J�n�s
            start_row_dict.Add "���x�ꐿ��", 28
            start_row_dict.Add "�Ԗ߁E����", 33
            start_row_dict.Add "����������", 38
        End If
    End If
    
    ' �f�B�N�V���i������̏ꍇ�i�z��O�̐�����^�C�v�Ȃǁj
    If start_row_dict.Count = 0 Then
        Debug.Print "WARNING: �J�e�S���̊J�n�s���ݒ�ł��܂���ł����B������: " & payer_type
        start_row_dict.Add "�Đ���", 3
        start_row_dict.Add "���x�ꐿ��", 8
        start_row_dict.Add "�Ԗ߁E����", 13
        start_row_dict.Add "����������", 18
    End If
    
    Set GetCategoryStartRows = start_row_dict
End Function

' �ǉ��s�̑}������
Public Sub InsertAdditionalRows(ws As Worksheet, start_row_dict As Object, _
    rebill_count As Long, late_count As Long, assessment_count As Long)
    
    Dim a As Long, b As Long, c As Long
    
    If rebill_count > BASE_DETAIL_ROWS Then a = rebill_count - BASE_DETAIL_ROWS
    If late_count > BASE_DETAIL_ROWS Then b = late_count - BASE_DETAIL_ROWS
    If assessment_count > BASE_DETAIL_ROWS Then c = assessment_count - BASE_DETAIL_ROWS
    
    If a > 0 Then ws.rows(start_row_dict("�Đ���") + 1 & ":" & start_row_dict("�Đ���") + a).Insert Shift:=xlDown
    If b > 0 Then ws.rows(start_row_dict("���x�ꐿ��") + 1 & ":" & start_row_dict("���x�ꐿ��") + b).Insert Shift:=xlDown
    If c > 0 Then ws.rows(start_row_dict("�Ԗ߁E����") + 1 & ":" & start_row_dict("�Ԗ߁E����") + c).Insert Shift:=xlDown
End Sub

' �f�[�^���ڍ׃V�[�g�ɓ]�L����֐�
Function TransferData(dataDict As Object, ws As Worksheet, start_row As Long, payer_type As String) As Boolean
    If dataDict.count = 0 Then
        TransferData = False
        Exit Function
    End If

    Dim key As Variant, row_data As Variant
    Dim r As Long: r = start_row
    Dim payer_col As Long

    ' �Еۂ�H��(8), ���ۂ�I��(9)�Ɏ�ʂ��L��
    If payer_type = "�Е�" Then
        payer_col = 8
    ElseIf payer_type = "����" Then
        payer_col = 9
    Else
        TransferData = False  ' ���̑��i�J�Г��j�͑ΏۊO
        Exit Function
    End If

    ' �e���R�[�h����������
    For Each key In dataDict.Keys
        row_data = dataDict(key)
        ws.Cells(r, 4).value = row_data(0)          ' ���Ҏ���
        ws.Cells(r, 5).value = row_data(1)          ' ���ܔN�� (YY.MM�`��)
        ws.Cells(r, 6).value = row_data(2)          ' ��Ë@�֖�
        ws.Cells(r, payer_col).value = payer_type   ' �������� (�Е�/����)
        ws.Cells(r, payer_col).Font.Bold = True     ' �����\��
        ws.Cells(r, 10).value = row_data(3)         ' �����_��
        r = r + 1
    Next key
    
    TransferData = True
End Function

' �S�p�����𔼊p�����ɕϊ�����֐�
Public Function ConvertToHankaku(ByVal strText As String) As String
    Dim i As Long
    Dim result As String
    Dim c As String
    
    result = ""
    For i = 1 To Len(strText)
        c = Mid(strText, i, 1)
        Select Case c
            Case "�O": result = result & "0"
            Case "�P": result = result & "1"
            Case "�Q": result = result & "2"
            Case "�R": result = result & "3"
            Case "�S": result = result & "4"
            Case "�T": result = result & "5"
            Case "�U": result = result & "6"
            Case "�V": result = result & "7"
            Case "�W": result = result & "8"
            Case "�X": result = result & "9"
            Case Else: result = result & c
        End Select
    Next i
    
    ConvertToHankaku = result
End Function

