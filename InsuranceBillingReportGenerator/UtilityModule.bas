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

' �J�e�S���̊J�n�s���擾����֐�
Public Function GetCategoryStartRows(ws As Worksheet, payer_type As String) As Object
    Dim start_row_dict As Object
    Set start_row_dict = CreateObject("Scripting.Dictionary")
    
    Debug.Print "Getting category start rows for: " & payer_type
    
    If payer_type = "�Е�" Then
        Dim social_start_row As Long
        social_start_row = GetStartRow(ws, "�ЕەԖߍĐ���")
        
        If social_start_row > 0 Then
            start_row_dict.Add "�ԖߍĐ���", social_start_row
            start_row_dict.Add "���x�ꐿ��", GetStartRow(ws, "�Еی��x�ꐿ��")
            start_row_dict.Add "�Ԗ߁E����", GetStartRow(ws, "�ЕەԖ߁E����")
            start_row_dict.Add "����������", GetStartRow(ws, "�Еۖ���������")
        Else
            ' ���o����������Ȃ��ꍇ�̃f�t�H���g�l��ݒ�
            Debug.Print "�Еۂ̌��o����������܂���B�f�t�H���g�l���g�p���܂��B"
            start_row_dict.Add "�ԖߍĐ���", 3  ' �f�t�H���g�J�n�s
            start_row_dict.Add "���x�ꐿ��", 8
            start_row_dict.Add "�Ԗ߁E����", 13
            start_row_dict.Add "����������", 18
        End If
    ElseIf payer_type = "����" Then
        Dim kokuho_start_row As Long
        kokuho_start_row = GetStartRow(ws, "���ەԖߍĐ���")
        
        If kokuho_start_row > 0 Then
            start_row_dict.Add "�ԖߍĐ���", kokuho_start_row
            start_row_dict.Add "���x�ꐿ��", GetStartRow(ws, "���ی��x�ꐿ��")
            start_row_dict.Add "�Ԗ߁E����", GetStartRow(ws, "���ەԖ߁E����")
            start_row_dict.Add "����������", GetStartRow(ws, "���ۖ���������")
        Else
            ' ���o����������Ȃ��ꍇ�̃f�t�H���g�l��ݒ�
            Debug.Print "���ۂ̌��o����������܂���B�f�t�H���g�l���g�p���܂��B"
            start_row_dict.Add "�ԖߍĐ���", 23  ' �f�t�H���g�J�n�s
            start_row_dict.Add "���x�ꐿ��", 28
            start_row_dict.Add "�Ԗ߁E����", 33
            start_row_dict.Add "����������", 38
        End If
    End If
    
    ' �f�B�N�V���i������̏ꍇ�i�z��O�̐�����^�C�v�Ȃǁj
    If start_row_dict.Count = 0 Then
        Debug.Print "WARNING: �J�e�S���̊J�n�s���ݒ�ł��܂���ł����B������: " & payer_type
        start_row_dict.Add "�ԖߍĐ���", 3
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
    
    If a > 0 Then ws.rows(start_row_dict("�ԖߍĐ���") + 1 & ":" & start_row_dict("�ԖߍĐ���") + a).Insert Shift:=xlDown
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

' �}�[�L���O����J�e�S���̊J�n�s���擾����֐�
Public Function FindMarkedRow(ws As Worksheet, marker As String) As Long
    Dim found_cell As Range
    Dim search_text As String
    search_text = "<<" & marker & ">>"
    
    ' D��i4��ځj�Ɍ��肵�Č���
    Set found_cell = ws.Columns(4).Find(What:=search_text, LookIn:=xlValues, LookAt:=xlPart)
    
    If Not found_cell Is Nothing Then
        Debug.Print "Found marker '" & marker & "' at row " & found_cell.Row
        FindMarkedRow = found_cell.Row
    Else
        Debug.Print "WARNING: Marker '" & marker & "' not found"
        FindMarkedRow = 0
    End If
End Function

' �}�[�L���O����ɃJ�e�S���J�n�s�̎������쐬����֐�
Public Function GetMarkedCategoryRows(ws As Worksheet) As Object
    Dim category_dict As Object
    Set category_dict = CreateObject("Scripting.Dictionary")
    
    ' ���ׂẴJ�e�S���Ƃ���ɑΉ�����}�[�J�[���`
    Dim categories As Object
    Set categories = CreateObject("Scripting.Dictionary")
    
    categories.Add "�ЕۍĐ���", "�ЕۍĐ���"
    categories.Add "���ۍĐ���", "���ۍĐ���"
    categories.Add "�Еی��x��", "�Еی��x��"
    categories.Add "���ی��x��", "���ی��x��"
    categories.Add "�Еی�����", "�Еی�����"
    categories.Add "���ی�����", "���ی�����"
    categories.Add "�ЕەԖ�", "�ЕەԖ�"
    categories.Add "�Еۖ�����", "�Еۖ���������"
    categories.Add "���ەԖ�", "���ەԖ�"
    categories.Add "���ۖ�����", "���ۖ���������"
    categories.Add "���Ԗ�", "���Ԗ�"
    categories.Add "���̑�", "���̑�"
    
    ' �e�J�e�S���̃}�[�J�[������
    Dim cat_key As Variant
    Dim row_num As Long
    
    For Each cat_key In categories.Keys
        row_num = FindMarkedRow(ws, categories(cat_key))
        
        ' �}�[�J�[�����������ꍇ�̂ݎ����ɒǉ�
        If row_num > 0 Then
            category_dict.Add cat_key, row_num
            Debug.Print "Added category '" & cat_key & "' with row " & row_num
        End If
    Next cat_key
    
    ' �}�[�J�[�����������Ȃ������ꍇ�A�]���̕��@�Ō���
    If category_dict.Count = 0 Then
        Debug.Print "WARNING: No markers found, using default method"
        ' �ЕۃJ�e�S��
        category_dict.Add "�ЕۍĐ���", 3
        category_dict.Add "�Еی��x��", 8
        category_dict.Add "�ЕەԖ�", 13
        category_dict.Add "�Еۖ�����", 18
        ' ���ۃJ�e�S��
        category_dict.Add "���ۍĐ���", 23
        category_dict.Add "���ی��x��", 28
        category_dict.Add "���ەԖ�", 33
        category_dict.Add "���ۖ�����", 38
    End If
    
    Set GetMarkedCategoryRows = category_dict
End Function

' ������^�C�v�ʂɃJ�e�S���J�n�s���擾����֐��i�}�[�L���O�Ή��Łj
Public Function GetCategoryStartRowsByMarker(ws As Worksheet, payer_type As String) As Object
    Dim all_category_rows As Object
    Dim filtered_dict As Object
    
    Set all_category_rows = GetMarkedCategoryRows(ws)
    Set filtered_dict = CreateObject("Scripting.Dictionary")
    
    Dim cat_key As Variant
    
    If payer_type = "�Е�" Then
        ' �Еۊ֘A�̃J�e�S���݂̂𒊏o
        For Each cat_key In all_category_rows.Keys
            If InStr(cat_key, "�Е�") > 0 Then
                ' �L�[����W����
                If cat_key = "�ЕۍĐ���" Then
                    filtered_dict.Add "�ԖߍĐ���", all_category_rows(cat_key)
                ElseIf cat_key = "�Еی��x��" Then
                    filtered_dict.Add "���x�ꐿ��", all_category_rows(cat_key)
                ElseIf cat_key = "�ЕەԖ�" Then
                    filtered_dict.Add "�Ԗ߁E����", all_category_rows(cat_key)
                ElseIf cat_key = "�Еۖ�����" Then
                    filtered_dict.Add "����������", all_category_rows(cat_key)
                End If
            End If
        Next cat_key
    ElseIf payer_type = "����" Then
        ' ���ۊ֘A�̃J�e�S���݂̂𒊏o
        For Each cat_key In all_category_rows.Keys
            If InStr(cat_key, "����") > 0 Then
                ' �L�[����W����
                If cat_key = "���ۍĐ���" Then
                    filtered_dict.Add "�ԖߍĐ���", all_category_rows(cat_key)
                ElseIf cat_key = "���ی��x��" Then
                    filtered_dict.Add "���x�ꐿ��", all_category_rows(cat_key)
                ElseIf cat_key = "���ەԖ�" Then
                    filtered_dict.Add "�Ԗ߁E����", all_category_rows(cat_key)
                ElseIf cat_key = "���ۖ�����" Then
                    filtered_dict.Add "����������", all_category_rows(cat_key)
                End If
            End If
        Next cat_key
    End If
    
    ' �K�v�ȃJ�e�S����������Ȃ������ꍇ�͏]���̕��@�Ŏ擾
    If filtered_dict.Count = 0 Then
        Debug.Print "WARNING: No " & payer_type & " categories found, using default method"
        Set filtered_dict = GetCategoryStartRows(ws, payer_type)
    End If
    
    Set GetCategoryStartRowsByMarker = filtered_dict
End Function

