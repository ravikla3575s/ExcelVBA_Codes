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
        End If
    ElseIf payer_type = "����" Then
        Dim kokuho_start_row As Long
        kokuho_start_row = GetStartRow(ws, "���ەԖߍĐ���")
        
        If kokuho_start_row > 0 Then
            start_row_dict.Add "�ԖߍĐ���", kokuho_start_row
            start_row_dict.Add "���x�ꐿ��", GetStartRow(ws, "���ی��x�ꐿ��")
            start_row_dict.Add "�Ԗ߁E����", GetStartRow(ws, "���ەԖ߁E����")
            start_row_dict.Add "����������", GetStartRow(ws, "���ۖ���������")
        End If
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

