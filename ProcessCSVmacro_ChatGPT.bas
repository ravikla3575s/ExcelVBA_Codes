Option Explicit

' �����Х��ѿ��ʥ桼�����ե���������ѡ�
Dim gRebillForm As Object          ' ưŪ�˺��������������������ե�����
Dim gUnclaimedForm As Object       ' ưŪ�˺�������̤����쥻�ץ�����ե�����
Dim gOlderList As Object           ' ���쥻�ץȥǡ������������������/���٤������ѡ�
Dim gUnclaimedList As Object       ' ����̤����ǡ���������̤����쥻�ץ������ѡ�
Dim gRebillData As Object          ' �桼���������̡�����������ʬ�ह��ǡ���
Dim gLateData As Object            ' �桼���������̡����٤������ʬ�ह��ǡ���
Dim gSelectedUnclaimed As Object   ' �桼���������̡�����̤���ᤫ���ɲä���ǡ���

Sub ProcessCSV()
    Dim csvFolder As String
    Dim fso As Object
    Dim targetYear As String
    Dim targetMonth As String
    Dim savePath As String
    Dim templatePath As String
    Dim newBook As Workbook
    Dim targetFile As String
    Dim fixfFile As String
    Dim fixfFiles As Object
    Dim file As Object

    ' 1. CSV�ե������桼���������򤵤���
    csvFolder = SelectCSVFolder()
    If csvFolder = "" Then Exit Sub

    ' 1.1 �ե���������ʤ���������
    If IsFolderEmpty(csvFolder) Then
        MsgBox "���򤷤��ե�����ˤ�CSV�ե����뤬����ޤ��󡣽�������ߤ��ޤ���", vbExclamation, "���顼"
        Exit Sub
    End If

    ' 2. �ƥ�ץ졼�ȥѥ�����¸�ե��������
    templatePath = GetTemplatePath()
    savePath = GetSavePath()
    If templatePath = "" Or savePath = "" Then Exit Sub

    ' 3. �ե����륷���ƥ४�֥������Ȥκ���
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 4. �ե������Τ��٤ƤΡ�fixf�ץե���������
    Set fixfFiles = FindAllFixfFiles(fso, csvFolder)

    ' 5. fixf�ե����뤬�ʤ������̾��CSV�������ڤ��ؤ�
    If fixfFiles Is Nothing Or fixfFiles.Count = 0 Then
        ProcessWithoutFixf fso, csvFolder, savePath, templatePath
        Exit Sub
    End If

    ' 6. ʣ����fixf�ե��������֤˽���
    For Each file In fixfFiles
        fixfFile = file.Path

        ' 7. �о�ǯ������
        targetYear = "": targetMonth = ""
        GetYearMonthFromFixf fixfFile, targetYear, targetMonth

        ' �о�ǯ������Ǥ��ʤ��ä����ϥ����å�
        If targetYear = "" Or targetMonth = "" Then
            MsgBox "�ե����� " & fixfFile & " �������ǯ�������Ǥ��ޤ���Ǥ�����", vbExclamation, "���顼"
            GoTo NextFile
        End If

        ' 8. �о�Excel�ե����뤬����¸�ߤ��뤫��ǧ��¸�ߤ�����ϥ����åס�
        Dim csvYYMM As String, reportName As String
        csvYYMM = Format(CInt(targetYear) - 2018, "00") & targetMonth
        reportName = "�ݸ������������_R" & csvYYMM & ".xlsx"
        If fso.FileExists(savePath & "\" & reportName) Then
            MsgBox "����ǯ�� " & targetYear & "ǯ" & targetMonth & "�� ������ϴ��˽����ѤߤǤ���", vbInformation, "�����Ѥ�"
            GoTo NextFile
        End If

        ' �о�Excel�ե�����������¸�ߤ��ʤ���Хƥ�ץ졼�Ȥ��鿷��������
        targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)
        If targetFile = "" Then
            MsgBox "����ǯ�� " & targetYear & "ǯ" & targetMonth & "�� ��Excel�ե����������Ǥ��ޤ���Ǥ�����", vbExclamation, "���顼"
            GoTo NextFile
        End If

        ' 9. Excel�򳫤�
        On Error Resume Next
        Set newBook = Workbooks.Open(targetFile)
        On Error GoTo 0
        If newBook Is Nothing Then
            MsgBox "�ե����� " & targetFile & " �򳫤��ޤ���Ǥ�����", vbExclamation, "���顼"
            GoTo NextFile
        End If

        ' 10. fixf�ե���������Ƥ򥷡���1�˼�����
        ImportCSVData fixfFile, newBook.Sheets(1), "����������"

        ' 11. �ƥ�ץ졼�Ⱦ��������ʥ�����1�ؤδ���ž���ϥ����åס�
        SetTemplateInfo newBook, targetYear, targetMonth, True

        ' 12. �ե�������¾��CSV�ե�������˽�����fmei��henr��zogn��
        ProcessAllCSVFiles fso, newBook, csvFolder

        ' 13. ��¸���ƥ֥å����Ĥ���
        newBook.Save
        newBook.Close
NextFile:
    Next file

    ' 14. ������λ��å�����
    MsgBox "���٤Ƥ� fixf �ե�����ν�������λ���ޤ�����", vbInformation, "������λ"
End Sub

Sub ProcessWithoutFixf(fso As Object, csvFolder As String, savePath As String, templatePath As String)
    Dim targetYear As String, targetMonth As String
    Dim targetFile As String, newBook As Workbook

    ' �о�ǯ���CSV���Ƥ������
    targetYear = "": targetMonth = ""
    GetYearMonthFromCSV fso, csvFolder, targetYear, targetMonth
    If targetYear = "" Or targetMonth = "" Then
        MsgBox "CSV�ե����뤫�����ǯ�������Ǥ��ޤ���Ǥ�����", vbExclamation, "���顼"
        Exit Sub
    End If

    ' �о�Excel�ե����뤬����¸�ߤ�����ϥ����å�
    Dim csvYYMM As String, reportName As String, fsoLocal As Object
    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    csvYYMM = Format(CInt(targetYear) - 2018, "00") & targetMonth
    reportName = "�ݸ������������_R" & csvYYMM & ".xlsx"
    If fsoLocal.FileExists(savePath & "\" & reportName) Then
        MsgBox "����ǯ�� " & targetYear & "ǯ" & targetMonth & "�� ������ϴ��˽����ѤߤǤ���", vbInformation, "�����Ѥ�"
        Exit Sub
    End If

    ' �о�Excel�ե�����������¸�ߤ��ʤ���п���������
    targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)
    If targetFile = "" Then
        MsgBox "����ǯ�� " & targetYear & "ǯ" & targetMonth & "�� ��Excel�ե����������Ǥ��ޤ���Ǥ�����", vbExclamation, "���顼"
        Exit Sub
    End If

    ' �֥å��򳫤�
    On Error Resume Next
    Set newBook = Workbooks.Open(targetFile)
    On Error GoTo 0
    If newBook Is Nothing Then
        MsgBox "�ե����� " & targetFile & " �򳫤��ޤ���Ǥ�����", vbExclamation, "���顼"
        Exit Sub
    End If

    ' �ƥ�ץ졼�Ⱦ����������̾��̤������
    SetTemplateInfo newBook, targetYear, targetMonth, False

    ' CSV�ե������缡������fixf�ʤ��Ǥ�¾��CSV�������ǽ��
    ProcessAllCSVFiles fso, newBook, csvFolder

    ' ��¸���ƥ֥å����Ĥ���
    newBook.Save
    newBook.Close

    MsgBox "CSV�ե�����ν�������λ���ޤ�����", vbInformation, "������λ"
End Sub

Sub SetTemplateInfo(newBook As Workbook, targetYear As String, targetMonth As String, Optional skipSheet1Info As Boolean = False)
    Dim wsTemplate As Worksheet, wsTemplate2 As Worksheet
    Dim receiptYear As Integer, receiptMonth As Integer
    Dim sendMonth As Integer, sendDate As String

    ' ����ǯ��Ĵ�޷�η׻�
    receiptYear = CInt(targetYear)
    receiptMonth = CInt(targetMonth)

    ' ������Ĵ�޷�����ˤη׻�
    sendMonth = receiptMonth + 1
    If sendMonth = 13 Then sendMonth = 1
    sendDate = sendMonth & "��10������ʬ"

    ' ������1(A), ������2(B)�����
    Set wsTemplate = newBook.Sheets(1)
    Set wsTemplate2 = newBook.Sheets(2)

    ' ������̾�ѹ��ʥ�����1�� "R{����YY}.{M}", ������2��ݿ����η���ѹ���
    wsTemplate.Name = "R" & (receiptYear - 2018) & "." & receiptMonth
    wsTemplate2.Name = ConvertToCircledNumber(receiptMonth)

    ' ����ž��
    If Not skipSheet1Info Then
        wsTemplate.Range("G2").Value = targetYear & "ǯ" & targetMonth & "��Ĵ��ʬ"
        wsTemplate.Range("I2").Value = sendDate
        wsTemplate.Range("J2").Value = ThisWorkbook.Sheets(1).Range("B1").Value  ' ���ᵡ�ءʻ���̾����
    End If
    wsTemplate2.Range("H1").Value = targetYear & "ǯ" & targetMonth & "��Ĵ��ʬ"
    wsTemplate2.Range("J1").Value = sendDate
    wsTemplate2.Range("L1").Value = ThisWorkbook.Sheets(1).Range("B1").Value     ' ���ᵡ�ءʻ���̾����
End Sub

Sub ProcessAllCSVFiles(fso As Object, newBook As Workbook, csvFolder As String)
    Dim csvFile As Object
    Dim fileType As String
    Dim wsDetails As Worksheet
    Dim wsCSV As Worksheet
    Dim sheetName As String
    Dim sheetIndex As Integer

    ' ������2�ʾܺ٥ǡ����ѡˤ����
    Set wsDetails = newBook.Sheets(2)

    ' 1. ���������ٽ��fmei�ˤν���
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" And InStr(csvFile.Name, "fmei") > 0 Then
            fileType = "���������ٽ�"
            sheetName = fso.GetBaseName(csvFile.Name)
            sheetName = GetUniqueSheetName(newBook, sheetName)
            sheetIndex = Application.WorksheetFunction.Min(3, newBook.Sheets.Count + 1)
            Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(sheetIndex - 1))
            wsCSV.Name = sheetName
            ImportCSVData csvFile.Path, wsCSV, fileType
            Call TransferBillingDetails(newBook, sheetName, csvFile.Name)
        End If
    Next csvFile

    ' 2. �����������henr�ˤν���
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" And InStr(csvFile.Name, "henr") > 0 Then
            fileType = "����������"
            sheetName = fso.GetBaseName(csvFile.Name)
            sheetName = GetUniqueSheetName(newBook, sheetName)
            sheetIndex = Application.WorksheetFunction.Min(3, newBook.Sheets.Count + 1)
            Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(sheetIndex - 1))
            wsCSV.Name = sheetName
            ImportCSVData csvFile.Path, wsCSV, fileType
            Call TransferBillingDetails(newBook, sheetName, csvFile.Name)
        End If
    Next csvFile

    ' 3. ������Ϣ����zogn�ˤν���
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" And InStr(csvFile.Name, "zogn") > 0 Then
            fileType = "������Ϣ���"
            sheetName = fso.GetBaseName(csvFile.Name)
            sheetName = GetUniqueSheetName(newBook, sheetName)
            sheetIndex = Application.WorksheetFunction.Min(3, newBook.Sheets.Count + 1)
            Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(sheetIndex - 1))
            wsCSV.Name = sheetName
            ImportCSVData csvFile.Path, wsCSV, fileType
            Call TransferBillingDetails(newBook, sheetName, csvFile.Name)
        End If
    Next csvFile
End Sub

Sub TransferBillingDetails(newBook As Workbook, sheetName As String, csvFileName As String)
    Dim wsBilling As Worksheet, wsDetails As Worksheet, wsCSV As Worksheet
    Dim lastRowBilling As Long, lastRowDetails As Long
    Dim i As Long, j As Long
    Dim dispensingMonth As String, convertedMonth As String
    Dim payerCode As String, payerType As String
    Dim startRowDict As Object
    Dim rebillDict As Object, lateDict As Object, unpaidDict As Object, assessmentDict As Object
    Dim rowData As Variant
    Dim a As Long, b As Long, c As Long

    ' ���������������ǡ��������ȤȾܺ٥����ȡ�
    Set wsBilling = newBook.Sheets(1)
    Set wsDetails = newBook.Sheets(2)

    ' ����ǯ���YYMM�����ˤ����
    Dim csvYYMM As String
    csvYYMM = Right(CStr(wsBilling.Cells(2, 2).Value), 4)

    ' CSV�ե�����̾�����������ʬ��Ƚ��
    payerCode = Mid(sheetName, 7, 1)
    Select Case payerCode
        Case "1": payerType = "����"
        Case "2": payerType = "����"
        Case Else: payerType = "ϫ��"
    End Select

    ' ���Ϲ԰��֤μ��������ʥ�����2�γƥ��ƥ��긫�Ф��Ԥ������
    Set startRowDict = CreateObject("Scripting.Dictionary")
    If payerType = "����" Then
        startRowDict.Add "���������", GetStartRow(wsDetails, "�������������")
        startRowDict.Add "���٤�����", GetStartRow(wsDetails, "���ݷ��٤�����")
        startRowDict.Add "���ᡦ����", GetStartRow(wsDetails, "�������ᡦ����")
        startRowDict.Add "̤���᰷��", GetStartRow(wsDetails, "����̤���᰷��")
    ElseIf payerType = "����" Then
        startRowDict.Add "���������", GetStartRow(wsDetails, "�������������")
        startRowDict.Add "���٤�����", GetStartRow(wsDetails, "���ݷ��٤�����")
        startRowDict.Add "���ᡦ����", GetStartRow(wsDetails, "�������ᡦ����")
        startRowDict.Add "̤���᰷��", GetStartRow(wsDetails, "����̤���᰷��")
    End If

    ' �ƥ��ƥ����ѤΥǥ�������ʥ�����
    Set rebillDict = CreateObject("Scripting.Dictionary")    ' ���������
    Set lateDict = CreateObject("Scripting.Dictionary")      ' ���٤�����
    Set unpaidDict = CreateObject("Scripting.Dictionary")    ' ̤���᰷��
    Set assessmentDict = CreateObject("Scripting.Dictionary") ' ���ᡦ����

    ' ����ǡ��������Ȥκǽ��Ԥ����
    lastRowBilling = wsBilling.Cells(Rows.Count, "D").End(xlUp).Row

    ' fixf�ե����뤬�ʤ���������ǡ�����������1�ˤʤ����ˡ���CSV����ܺ٥ǡ�����ž��
    If lastRowBilling < 2 Then
        If InStr(csvFileName, "henr") > 0 Then
            Set wsCSV = newBook.Sheets(sheetName)
            lastRowDetails = wsCSV.Cells(Rows.Count, 1).End(xlUp).Row
            For j = 2 To lastRowDetails
                dispensingMonth = CStr(wsCSV.Cells(j, 1).Value)
                If dispensingMonth <> "" Then
                    If Len(dispensingMonth) = 4 Then dispensingMonth = "5" & dispensingMonth
                    convertedMonth = ConvertToWesternDate(dispensingMonth)
                    rowData = Array(wsCSV.Cells(j, 2).Value, convertedMonth, wsCSV.Cells(j, 5).Value, wsCSV.Cells(j, 14).Value)  ' �����ֹ�, ���ŷ�, ��������, ��ͳ������
                    assessmentDict.Add CStr(wsCSV.Cells(j, 2).Value) & "_" & j, rowData
                End If
            Next j
        ElseIf InStr(csvFileName, "zogn") > 0 Then
            Set wsCSV = newBook.Sheets(sheetName)
            lastRowDetails = wsCSV.Cells(Rows.Count, 1).End(xlUp).Row
            For j = 2 To lastRowDetails
                dispensingMonth = CStr(wsCSV.Cells(j, 1).Value)
                If dispensingMonth <> "" Then
                    If Len(dispensingMonth) = 4 Then dispensingMonth = "5" & dispensingMonth
                    convertedMonth = ConvertToWesternDate(dispensingMonth)
                    rowData = Array(wsCSV.Cells(j, 2).Value, convertedMonth, wsCSV.Cells(j, 6).Value, wsCSV.Cells(j, 7).Value)  ' �����ֹ�, Ĵ�޷�, ��������, ��ͳ
                    unpaidDict.Add CStr(wsCSV.Cells(j, 2).Value) & "_" & j, rowData
                End If
            Next j
        End If
    End If

    ' ����ǡ�����fixf�ˤ�ǥ�������ʥ�˳�Ǽ��fixf�ե����뤬������Τ߳�����
    Dim dispGYM As String
    For i = 2 To lastRowBilling
        dispensingMonth = wsBilling.Cells(i, 2).Value       ' GYYMM�����ο��ŷ�
        convertedMonth = ConvertToWesternDate(dispensingMonth)
        rowData = Array(wsBilling.Cells(i, 4).Value, convertedMonth, wsBilling.Cells(i, 5).Value, wsBilling.Cells(i, 10).Value)
        ' �оݿ��ŷ��csvYYMM�ˤȰۤʤ���Τ߳ƥ��ƥ�����ɲ�
        If Right(dispensingMonth, 4) <> csvYYMM Then
            If InStr(csvFileName, "fixf") > 0 Then
                ' fixf����ȥ�ξ�硢�桼��������������ᤫ���٤����ᤫ���򤵤���
                If ShowRebillSelectionForm(rowData) Then
                    rebillDict.Add wsBilling.Cells(i, 1).Value, rowData   ' ���������
                Else
                    lateDict.Add wsBilling.Cells(i, 1).Value, rowData    ' ���٤�����
                End If
            ElseIf InStr(csvFileName, "zogn") > 0 Then
                unpaidDict.Add wsBilling.Cells(i, 1).Value, rowData      ' ̤���᰷��
            ElseIf InStr(csvFileName, "henr") > 0 Then
                assessmentDict.Add wsBilling.Cells(i, 1).Value, rowData  ' ���ᡦ����
            End If
        End If
    Next i

    ' �ƥ��ƥ�����ɲùԿ���׻��ʳƥ��ƥ���4�Ԥ�Ķ����ʬ��
    a = 0: b = 0: c = 0
    If rebillDict.Count > 4 Then a = rebillDict.Count - 4
    If lateDict.Count > 4 Then b = lateDict.Count - 4
    If assessmentDict.Count > 4 Then c = assessmentDict.Count - 4

    ' �ƥ��ƥ���γ��ϹԤ�Ĵ��
    Dim lateStartRow As Long, assessmentStartRow As Long, unpaidStartRow As Long
    lateStartRow = startRowDict("���٤�����") + 1 + a
    assessmentStartRow = startRowDict("���ᡦ����") + 1 + a + b
    unpaidStartRow = startRowDict("̤���᰷��") + 1 + a + b + c

    ' ɬ�פ˱����ƹԤ����������Ȥ����
    If a + b + c > 0 Then
        wsDetails.Rows(lateStartRow & ":" & lateStartRow + a).Insert Shift:=xlDown
        wsDetails.Rows(assessmentStartRow & ":" & assessmentStartRow + b).Insert Shift:=xlDown
        wsDetails.Rows(unpaidStartRow & ":" & unpaidStartRow + c).Insert Shift:=xlDown
    End If

    ' �ƥǥ�������ʥ�Υǡ����򥷡���2��ž���ʥǥ�������ʥ꤬���ξ��ϥ����åס�
    If rebillDict.Count > 0 Then
        j = startRowDict("���������")
        TransferData rebillDict, wsDetails, j, payerType
    End If
    If lateDict.Count > 0 Then
        j = startRowDict("���٤�����")
        TransferData lateDict, wsDetails, j, payerType
    End If
    If unpaidDict.Count > 0 Then
        j = startRowDict("̤���᰷��")
        TransferData unpaidDict, wsDetails, j, payerType
    End If
    If assessmentDict.Count > 0 Then
        j = startRowDict("���ᡦ����")
        TransferData assessmentDict, wsDetails, j, payerType
    End If

    MsgBox payerType & " �Υǡ���ž������λ���ޤ�����", vbInformation, "������λ"
End Sub

Function SelectCSVFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "CSV�ե���������򤷤Ƥ�������"
        If .Show = -1 Then
            SelectCSVFolder = .SelectedItems(1) & "\"
        Else
            MsgBox "�ե���������򤵤�ޤ���Ǥ�������������ߤ��ޤ���", vbExclamation, "��ǧ"
            SelectCSVFolder = ""
        End If
    End With
End Function

Function IsFolderEmpty(folderPath As String) As Boolean
    Dim fso As Object, folder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        IsFolderEmpty = True
        Exit Function
    End If
    Set folder = fso.GetFolder(folderPath)
    If folder.Files.Count = 0 Then
        IsFolderEmpty = True   ' **�ե�����˥ե����뤬�ʤ���� True**
    Else
        IsFolderEmpty = False
    End If
End Function

Function GetTemplatePath() As String
    ' �ƥ�ץ졼�ȥե�����Υѥ��򥷡���1�Υ���B2�������
    GetTemplatePath = ThisWorkbook.Sheets(1).Range("B2").Value & "\�ݸ������������ƥ�ץ졼��.xltm"
End Function

Function GetSavePath() As String
    ' ��¸��ե�����Υѥ��򥷡���1�Υ���B3�������
    GetSavePath = ThisWorkbook.Sheets(1).Range("B3").Value
End Function

Function FindAllFixfFiles(fso As Object, csvFolder As String) As Collection
    Dim csvFile As Object
    Dim fixfFiles As New Collection
    ' **�ե������Τ��٤ƤΥե����������å�**
    For Each csvFile In fso.GetFolder(csvFolder).Files
        ' **��ĥ�Ҥ� "csv" �Ǥ��ꡢ̾���� "fixf" ��ޤ���**
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" And InStr(LCase(csvFile.Name), "fixf") > 0 Then
            fixfFiles.Add csvFile  ' **fixf�ե������ꥹ�Ȥ��ɲ�**
        End If
    Next csvFile
    Set FindAllFixfFiles = fixfFiles
End Function

Sub GetYearMonthFromFixf(fixfFile As String, ByRef targetYear As String, ByRef targetMonth As String)
    Dim fso As Object, fileName As String, baseName As String
    Dim code As String, yrCode As String, monCode As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = fso.GetFileName(fixfFile)
    baseName = fso.GetBaseName(fixfFile)
    ' **fixf�ե�����̾����ǯ����ɤ����**
    code = ""
    '  ¾��CSV�ե�����̾���� GYYMM ���������: ���������ٽ�ʤɡ� 
    Dim csvFile As Object, folderPath As String
    folderPath = fso.GetFile(fixfFile).ParentFolder.Path
    For Each csvFile In fso.GetFolder(folderPath).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" Then
            If InStr(LCase(csvFile.Name), "fmei") > 0 Or InStr(LCase(csvFile.Name), "zogn") > 0 Or InStr(LCase(csvFile.Name), "henr") > 0 Then
                ' ̾������4�夬�����ʤ餽���ǯ����ɤȤ���
                Dim nameNoExt As String
                nameNoExt = fso.GetBaseName(csvFile.Name)
                If Len(nameNoExt) >= 4 Then
                    If IsNumeric(Right(nameNoExt, 4)) Then
                        code = Right(nameNoExt, 4)
                        Exit For
                    End If
                End If
            End If
        End If
    Next csvFile
    If code <> "" Then
        yrCode = Left(code, 2)
        monCode = Right(code, 2)
        targetYear = CStr(2018 + CInt(yrCode))    ' **����ǯ�����ɤ�����ǯ���Ѵ�**
        targetMonth = CStr(CInt(monCode))         ' **����ɡ���Ƭ0�ޤ�ˤ�������**
    Else
        ' **fallback: fixf�ե����뤫�����ǯ������**�ʴʰס�
        Dim ts As Object, lineText As String
        On Error Resume Next
        Set ts = fso.OpenTextFile(fixfFile, 1, False, -2)
        On Error GoTo 0
        If Not ts Is Nothing Then
            ' ��Ƭ���Ԥ��ɤ߿���ǯ���ޤ�Ԥ�õ��
            Dim i As Integer
            For i = 1 To 5
                If ts.AtEndOfStream Then Exit For
                lineText = ts.ReadLine
                If InStr(lineText, "G") > 0 And InStr(lineText, ",") = 0 Then
                    ' **��: "5XXXX" ������ʸ�����ޤ���**
                    Dim matchStr As Variant
                    matchStr = lineText
                    matchStr = Replace(matchStr, """", "")
                    If Len(matchStr) >= 5 And IsNumeric(matchStr) Then
                        yrCode = Left(matchStr, 2)
                        monCode = Right(matchStr, 2)
                        targetYear = CStr(2018 + CInt(yrCode))
                        targetMonth = CStr(CInt(monCode))
                        Exit For
                    End If
                End If
            Next i
            ts.Close
        End If
        ' **�������Ի����桼���������Ϥ�¥��**
        If targetYear = "" Or targetMonth = "" Then
            MsgBox "����ǯ���ư�����Ǥ��ޤ���Ǥ��������ꤷ�Ƥ���������", vbExclamation, "��ǧ"
            targetYear = InputBox("����ǯ�����Ϥ��Ƥ�����������: 2023��:", "����ǯ")
            targetMonth = InputBox("������Ϥ��Ƥ���������1���12��:", "���ŷ�")
            If targetYear = "" Or targetMonth = "" Then
                ' �桼����̤���Ϥξ��
                Exit Sub
            End If
        End If
    End If
End Sub

Function FindOrCreateReport(savePath As String, targetYear As String, targetMonth As String, templatePath As String) As String
    Dim fso As Object, existingFile As Object
    Dim fileName As String, filePath As String
    Dim csvYYMM As String
    csvYYMM = Format(CInt(targetYear) - 2018, "00") & Format(CInt(targetMonth), "00")  ' **����ǯ+�����**
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' **��¸�ե������˴�¸�� RYYMM �ե����뤬���뤫��ǧ**
    For Each existingFile In fso.GetFolder(savePath).Files
        If LCase(fso.GetExtensionName(existingFile.Name)) = "xlsm" Or LCase(fso.GetExtensionName(existingFile.Name)) = "xlsx" Then
            If InStr(existingFile.Name, "�ݸ������������_R" & csvYYMM) > 0 Then
                FindOrCreateReport = existingFile.Path  ' **��¸�ե�����Υѥ����֤�**
                Exit Function
            End If
        End If
    Next existingFile
    ' **��������ե����뤬�ʤ���С���������**
    fileName = "�ݸ������������_R" & csvYYMM & ".xlsm"   ' **xlsm��������¸**�ʲ�����1��
    filePath = savePath & "\" & fileName
    ' **�ƥ�ץ졼�Ȥ򸵤˿����֥å�����**
    On Error Resume Next
    Dim tmplWb As Workbook
    Set tmplWb = Workbooks.Open(templatePath)   ' **�ƥ�ץ졼�ȥ֥å��򳫤�**
    On Error GoTo 0
    If tmplWb Is Nothing Then
        MsgBox "�ƥ�ץ졼�Ȥ򳫤��ޤ���Ǥ���: " & templatePath, vbCritical, "���顼"
        FindOrCreateReport = ""
        Exit Function
    End If
    On Error Resume Next
    tmplWb.SaveAs filePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled  ' **xlsm��������¸**
    If Err.Number <> 0 Then
        MsgBox "�ե��������¸�Ǥ��ޤ���Ǥ���: " & filePath, vbCritical, "���顼"
        FindOrCreateReport = ""
        tmplWb.Close SaveChanges:=False
        Exit Function
    End If
    On Error GoTo 0
    tmplWb.Close SaveChanges:=True
    FindOrCreateReport = filePath
End Function

Sub ImportCSVData(csvFile As String, ws As Worksheet, fileType As String)
    Dim colMap As Object
    Dim fso As Object, ts As Object
    Dim lineText As String
    Dim dataArray As Variant
    Dim i As Long, j As Long, key
    Dim isHeader As Boolean
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ���ܥޥåԥ󥰤����
    Set colMap = GetColumnMapping(fileType)
    ' �����Ȥ򥯥ꥢ���ƹ���̾��1���ܤ�����
    ws.Cells.Clear
    j = 1
    For Each key In colMap.Keys
        ws.Cells(1, j).Value = colMap(key)
        j = j + 1
    Next key

    ' CSV�ե������UTF-8�ƥ����ȤȤ��Ƴ���
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(csvFile, 1, False, -2)  ' -2: UTF-8

    ' �ǡ�����ʬ��ž��
    i = 2
    isHeader = True
    Do While Not ts.AtEndOfStream
        lineText = ts.ReadLine
        dataArray = Split(lineText, ",")
        If isHeader Then
            ' ����ܡʥإå����ԡˤϥ����å�
            isHeader = False
        Else
            j = 1
            For Each key In colMap.Keys
                If key - 1 <= UBound(dataArray) Then
                    ws.Cells(i, j).Value = Trim(dataArray(key - 1))
                End If
                j = j + 1
            Next key
            i = i + 1
        End If
    Loop
    ts.Close

    ws.Cells.EntireColumn.AutoFit

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "CSV�ɹ���˥��顼��ȯ�����ޤ���: " & Err.Description, vbCritical, "���顼"
    If Not ts Is Nothing Then ts.Close
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Function GetColumnMapping(fileType As String) As Object
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    Select Case fileType
        Case "���������ٽ�"
            colMap.Add 2, "����ǯ��"
            colMap.Add 3, "�������"
            colMap.Add 4, "������"
            ' ��ɬ�פ�����ɲá�
        Case "������Ϣ���"
            colMap.Add 2, "Ĵ��ǯ��"
            colMap.Add 4, "�����ֹ�"
            colMap.Add 11, "��ʬ"
            colMap.Add 14, "Ϸ�͸��ȶ�ʬ"
            colMap.Add 15, "��̾"
            colMap.Add 21, "���������ʶ�ۡ�"
            colMap.Add 22, "��ͳ"
        Case "����������"
            colMap.Add 2, "Ĵ��ǯ��(YYMM����)"
            colMap.Add 3, "�����ֹ�"
            colMap.Add 4, "�ݸ����ֹ�"
            colMap.Add 7, "��̾"
            colMap.Add 9, "��������"
            colMap.Add 10, "���ް�����ô��"
            colMap.Add 12, "������ô���"
            colMap.Add 13, "������ô��ۡʸ����"
            colMap.Add 14, "��ͳ������"
        Case Else
            ' ����¾��ɬ�פ˱������ɲá�
            colMap.Add 1, "����1"
    End Select
    Set GetColumnMapping = colMap
End Function

Function GetStartRow(ws As Worksheet, category As String) As Long
    ' �ܺ٥����Ȥ�����ꥫ�ƥ���ι��ֹ�����
    Dim rng As Range
    Set rng = ws.Cells.Find(what:=category, LookAt:=xlWhole)
    If rng Is Nothing Then
        MsgBox "�ܺ٥����Ⱦ�ǥ��ƥ��� """ & category & """ �򸫤Ĥ����ޤ���Ǥ�����", vbExclamation, "���顼"
        GetStartRow = 0
    Else
        GetStartRow = rng.Row
    End If
End Function

Sub TransferData(dataDict As Object, ws As Worksheet, startRow As Long, payerType As String)
    Dim key As Variant, rowData As Variant
    Dim j As Long, payerColumn As Long
    ' **Dictionary�����ʤ�������ʤ�**
    If dataDict.Count = 0 Then Exit Sub
    ' **payerType�˱�����ž��������**
    If payerType = "����" Then
        payerColumn = 8   ' ���ݤ�H���������ޡ���
    ElseIf payerType = "����" Then
        payerColumn = 9   ' ���ݤ�I���������ޡ���
    Else
        payerColumn = 8   ' ��ϫ�����ϼ�����˲������
    End If
    j = startRow
    For Each key In dataDict.Keys
        rowData = dataDict(key)
        ws.Cells(j, 4).Value = rowData(0)    ' ���Ի�̾
        ws.Cells(j, 5).Value = rowData(1)    ' Ĵ��ǯ�������ɽ����
        ws.Cells(j, 6).Value = rowData(2)    ' ���ŵ���̾
        ws.Cells(j, payerColumn).Value = payerType   ' �������ʬ�ʼ���/���ݡ�
        ws.Cells(j, payerColumn).Font.Bold = True    ' **��Ĵɽ��**
        ws.Cells(j, 10).Value = rowData(3)   ' ��������
        j = j + 1
    Next key
End Sub

Sub ShowRebillSelectionForm()
    ' ����쥻�ץȤΰ�����桼������ɽ��������������᤹���Τ����򤷤Ƥ�餦
    Dim uf As Object, listData As Object
    Set listData = gOlderList
    If listData Is Nothing Or listData.Count = 0 Then Exit Sub
    ' �桼�����ե����������ɽ��
    Set uf = CreateRebillSelectionForm(listData)
    Set gRebillForm = uf  ' �����Х뻲����¸
    gRebillForm.Show vbModal
    ' �ե������Ĥ����塢�����̤� gRebillData �� gLateData �˳�Ǽ�Ѥߡ�ProcessRebillSelection�������
End Sub

Function CreateRebillSelectionForm(listData As Object) As Object
    Dim uf As Object, listBox As Object, btnOK As Object
    Dim i As Long, rowData As Variant
    ' **UserForm ��ưŪ�˺���**
    Set uf = VBA.UserForms.Add()  ' ����UserForm
    uf.Caption = "��������������"
    uf.Width = 400
    uf.Height = 500
    ' **ListBox���ɲ�**
    Set listBox = uf.Controls.Add("Forms.ListBox.1", "listBox", True)
    listBox.Left = 20
    listBox.Top = 20
    listBox.Width = 350
    listBox.Height = 350
    listBox.MultiSelect = 1  ' **ʣ�������ǽ**
    ' **�ꥹ�Ȥ˥ǡ������ɲá�Ĵ��ǯ�� | ���Ի�̾ | ���ŵ���̾ | ������**
    For i = 0 To listData.Count - 1
        rowData = listData.Items()(i)
        listBox.AddItem rowData(1) & " | " & rowData(0) & " | " & rowData(2) & " | " & rowData(3)
    Next i
    ' **OK�ܥ�����ɲ�**
    Set btnOK = uf.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    btnOK.Caption = "����"
    btnOK.Left = 150
    btnOK.Top = 400
    btnOK.Width = 100
    btnOK.Height = 30
    ' **�ܥ��󥯥�å����ν���������**
    btnOK.OnClick = "ProcessRebillSelection"
    Set CreateRebillSelectionForm = uf
End Function

Sub ProcessRebillSelection()
    ' �������������ե������OK�ܥ�����������򤵤줿���ܤ�ʬ���
    Dim uf As Object, listBox As Object
    Dim i As Long
    ' ưŪ�ե����प���ListBox�����
    Set uf = gRebillForm
    Set listBox = uf.Controls("listBox")
    ' �����Dictionary������
    Set gRebillData = CreateObject("Scripting.Dictionary")
    Set gLateData = CreateObject("Scripting.Dictionary")
    ' **������֤˱����ƿ���ʬ��**
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) Then
            ' ���򤵤줿��� -> ���������
            gRebillData.Add gOlderList.Keys()(i), gOlderList.Items()(i)
        Else
            ' ���򤵤�ʤ��ä���� -> ���٤�����
            gLateData.Add gOlderList.Keys()(i), gOlderList.Items()(i)
        End If
    Next i
    ' �ե�����򥢥���ɤ����Ĥ���
    Unload uf
    Set gRebillForm = Nothing
End Sub

Function AddUnclaimedRecords(payerType As String, targetYear As String, targetMonth As String) As Object
    Dim prevYear As String, prevMonth As String
    Dim prevFileName As String, prevFilePath As String
    Dim prevBook As Workbook, wsPrevDetails As Worksheet
    Dim startRow As Long, endRow As Long, row As Long
    ' ����򻻽�
    If CInt(targetMonth) = 1 Then
        prevYear = CStr(CInt(targetYear) - 1)
        prevMonth = "12"
    Else
        prevYear = targetYear
        prevMonth = CStr(CInt(targetMonth) - 1)
    End If
    ' ���������ե�����̾
    Dim prevYYMM As String
    prevYYMM = Format(CInt(prevYear) - 2018, "00") & Format(CInt(prevMonth), "00")
    prevFileName = "�ݸ������������_R" & prevYYMM & ".xlsm"
    prevFilePath = GetSavePath() & "\" & prevFileName
    If Dir(prevFilePath) = "" Then
        ' �ե����뤬¸�ߤ��ʤ����
        Set AddUnclaimedRecords = Nothing
        Exit Function
    End If
    ' ����ե�����򳫤����ɤ߼�����ѡ�
    On Error Resume Next
    Set prevBook = Workbooks.Open(prevFilePath, ReadOnly:=True)
    On Error GoTo 0
    If prevBook Is Nothing Then
        Set AddUnclaimedRecords = Nothing
        Exit Function
    End If
    ' ����ܺ٥����Ȥ�����ʥ�����̾�ϼ���/���ݶ��̤�"B"�����ȤȲ����
    Set wsPrevDetails = prevBook.Sheets(2)
    ' �оݥ��ƥ���γ��ϹԤ����
    Dim categoryLabel As String
    If payerType = "����" Then
        categoryLabel = "����̤���᰷��"
    Else
        categoryLabel = "����̤���᰷��"
    End If
    startRow = GetStartRow(wsPrevDetails, categoryLabel)
    If startRow = 0 Then
        ' ��٥뤬���Ĥ���ʤ����Ͻ�λ
        prevBook.Close SaveChanges:=False
        Set AddUnclaimedRecords = Nothing
        Exit Function
    End If
    ' ���ϹԤ��鲼�����˥ǡ��������
    Set gUnclaimedList = CreateObject("Scripting.Dictionary")
    endRow = startRow + 3  ' �����Ȥ�4��
    ' �ǡ������ɲä���Ƥ����硢����Ԥ��Ф�ޤǱ�Ĺ
    Do While wsPrevDetails.Cells(endRow, 4).Value <> "" Or wsPrevDetails.Cells(endRow + 1, 4).Value <> ""
        endRow = endRow + 1
        If endRow > wsPrevDetails.Rows.Count Then Exit Do
    Loop
    For row = startRow + 1 To endRow
        If wsPrevDetails.Cells(row, 4).Value <> "" Then   ' ���Ի�̾�󤬶��Ǥʤ���Хǡ�������
            Dim prevRowData As Variant
            prevRowData = Array(wsPrevDetails.Cells(row, 4).Value, wsPrevDetails.Cells(row, 5).Value, wsPrevDetails.Cells(row, 6).Value, wsPrevDetails.Cells(row, 10).Value)
            gUnclaimedList.Add row, prevRowData
        End If
    Next row
    ' ����֥å����Ĥ���
    prevBook.Close SaveChanges:=False
    ' �桼����������̤�����ɽ�������ɲä����Τ����򤵤���
    If gUnclaimedList.Count > 0 Then
        ShowUnclaimedSelectionForm
        Set AddUnclaimedRecords = gSelectedUnclaimed
    Else
        Set AddUnclaimedRecords = Nothing
    End If
End Function

Sub ShowUnclaimedSelectionForm()
    If gUnclaimedList Is Nothing Or gUnclaimedList.Count = 0 Then Exit Sub
    Dim uf As Object
    Set uf = CreateUnclaimedSelectionForm(gUnclaimedList)
    Set gUnclaimedForm = uf
    gUnclaimedForm.Show vbModal
    ' �ե����ब�Ĥ���줿�塢gSelectedUnclaimed�˷�̤���Ǽ�����
End Sub

Function CreateUnclaimedSelectionForm(listData As Object) As Object
    Dim uf As Object, listBox As Object, btnOK As Object
    Dim i As Long, rowData As Variant
    Set uf = VBA.UserForms.Add()
    uf.Caption = "���� ̤����쥻�ץȤ��ɲ�����"
    uf.Width = 400
    uf.Height = 500
    Set listBox = uf.Controls.Add("Forms.ListBox.1", "listBox", True)
    listBox.Left = 20
    listBox.Top = 20
    listBox.Width = 350
    listBox.Height = 350
    listBox.MultiSelect = 1
    For i = 0 To listData.Count - 1
        rowData = listData.Items()(i)
        listBox.AddItem rowData(1) & " | " & rowData(0) & " | " & rowData(2) & " | " & rowData(3)
    Next i
    Set btnOK = uf.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    btnOK.Caption = "�ɲ�"
    btnOK.Left = 150
    btnOK.Top = 400
    btnOK.Width = 100
    btnOK.Height = 30
    btnOK.OnClick = "ProcessUnclaimedSelection"
    Set CreateUnclaimedSelectionForm = uf
End Function

Sub ProcessUnclaimedSelection()
    Dim uf As Object, listBox As Object
    Dim i As Long
    Set uf = gUnclaimedForm
    Set listBox = uf.Controls("listBox")
    Set gSelectedUnclaimed = CreateObject("Scripting.Dictionary")
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) Then
            gSelectedUnclaimed.Add gUnclaimedList.Keys()(i), gUnclaimedList.Items()(i)
        End If
    Next i
    Unload uf
    Set gUnclaimedForm = Nothing
End Sub

Function ConvertToCircledNumber(month As Integer) As String
    Dim circledNumbers As Variant
    circledNumbers = Array("��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��")
    If month >= 1 And month <= 12 Then
        ConvertToCircledNumber = circledNumbers(month - 1)
    Else
        ConvertToCircledNumber = CStr(month)
    End If
End Function

Function ConvertToWesternDate(dispensingMonth As String) As String
    ' GYYMM����������ˤ�����ǯ��2��.��������Ѵ�
    Dim eraCode As String, yearPart As Integer, westernYear As Integer, monthPart As String
    eraCode = Left(dispensingMonth, 1)
    yearPart = CInt(Mid(dispensingMonth, 2, 2))
    monthPart = Right(dispensingMonth, 2)
    Select Case eraCode
        Case "5": westernYear = 2018 + yearPart   ' ���� (2019ǯ=����1ǯ)
        Case "4": westernYear = 1988 + yearPart   ' ʿ�� (1989ǯ=ʿ��1ǯ) - �����ǡ����б�
        Case Else: westernYear = 2018 + yearPart  ' �ʥǥե����:���¤Ȥ��Ʒ׻���
    End Select
    ConvertToWesternDate = Right(CStr(westernYear), 2) & "." & monthPart
End Function

' **Ⱦǯ���Ȥ���ݥǡ�����ӡ���ʬ�ϵ�ǽ**�ʲ�����6��
Sub CompareHalfYearData()
    Dim inputYear As String, half As String
    inputYear = InputBox("ʬ�Ϥ���ǯ�����Ϥ��Ƥ��������������:", "Ⱦǯ������")
    If inputYear = "" Then Exit Sub
    half = InputBox("���=1 �ޤ��� ����=2 �����Ϥ��Ƥ�������:", "Ⱦǯ��ʬ")
    If half = "" Then Exit Sub
    If half <> "1" And half <> "2" Then
        MsgBox "Ⱦ����ʬ��1�ޤ���2�����Ϥ��Ƥ���������", vbExclamation, "���ϥ��顼"
        Exit Sub
    End If
    Dim startMonth As Integer, endMonth As Integer
    If half = "1" Then
        startMonth = 1: endMonth = 6
    Else
        startMonth = 7: endMonth = 12
    End If
    Dim analysisWb As Workbook
    Set analysisWb = ThisWorkbook  ' ��̽������ޥ���֥å�������
    Dim outSheet As Worksheet
    On Error Resume Next
    Set outSheet = analysisWb.Sheets("HalfYearAnalysis")
    On Error GoTo 0
    If outSheet Is Nothing Then
        Set outSheet = analysisWb.Sheets.Add
        outSheet.Name = "HalfYearAnalysis"
    Else
        outSheet.Cells.Clear
    End If
    outSheet.Range("A1:E1").Value = Array("��", "�����׾�����", "�����������", "������(��)", "��������")
    Dim m As Integer, rowIndex As Integer
    rowIndex = 2
    For m = startMonth To endMonth
        Dim yy As String, mm As String, fileCode As String
        yy = Format(CInt(inputYear) - 2018, "00")
        mm = Format(m, "00")
        fileCode = "R" & yy & mm
        Dim reportName As String
        reportName = "�ݸ������������_" & fileCode & ".xlsm"
        Dim reportPath As String
        reportPath = GetSavePath() & "\" & reportName
        If Dir(reportPath) <> "" Then
            Dim repWb As Workbook
            Set repWb = Workbooks.Open(reportPath, ReadOnly:=True)
            Dim wsA As Worksheet, wsCSV As Worksheet, wsCSV2 As Worksheet
            Set wsA = repWb.Sheets(1)   ' �����ǡ���������
            ' �����ǡ����������������������Ȥ��ƥ�����A�Υ���J���ʤɤ˽��פ�����Ȳ����
            Dim dailyTotal As Long
            dailyTotal = 0
            On Error Resume Next
            dailyTotal = CLng(wsA.Range("J50").Value) ' ��Ŭ�ڤʥ��뻲�Ȥ��׽���
            On Error GoTo 0
            ' �������������fixf�ǡ������������˼����ʥ�����A�˷פޤ��ϥ�����B��������������J50�Ȥ����
            Dim billedTotal As Long
            billedTotal = 0
            On Error Resume Next
            billedTotal = CLng(wsA.Range("J50").Value)
            On Error GoTo 0
            ' ���������٤����׶�ۼ�����CSV������̾��fmei�ޤ������
            Dim payAmount As Long
            payAmount = 0
            For Each wsCSV In repWb.Worksheets
                If InStr(wsCSV.Name, "fmei") > 0 Then
                    On Error Resume Next
                    payAmount = CLng(wsCSV.Cells(wsCSV.Rows.Count, 3).End(xlUp).Value)
                    On Error GoTo 0
                    Exit For
                End If
            Next wsCSV
            repWb.Close SaveChanges:=False
            ' �������۷׻�
            Dim pointDiff As Long
            pointDiff = dailyTotal - billedTotal
            ' ��̤����
            outSheet.Cells(rowIndex, 1).Value = inputYear & "ǯ" & m & "��"
            outSheet.Cells(rowIndex, 2).Value = dailyTotal
            outSheet.Cells(rowIndex, 3).Value = billedTotal
            outSheet.Cells(rowIndex, 4).Value = payAmount
            outSheet.Cells(rowIndex, 5).Value = pointDiff
            rowIndex = rowIndex + 1
        Else
            ' �ե����뤬�ʤ����϶��Ԥޤ���0����
            outSheet.Cells(rowIndex, 1).Value = inputYear & "ǯ" & m & "��"
            outSheet.Cells(rowIndex, 2).Value = "N/A"
            outSheet.Cells(rowIndex, 3).Value = "N/A"
            outSheet.Cells(rowIndex, 4).Value = "N/A"
            outSheet.Cells(rowIndex, 5).Value = "N/A"
            rowIndex = rowIndex + 1
        End If
    Next m
    MsgBox inputYear & "ǯ " & IIf(half = "1", "���", "����") & " ����ݥǡ�����Ӥ���λ���ޤ�����" & vbCrLf & _
            "������[" & outSheet.Name & "]�˷�̤���Ϥ��ޤ�����", vbInformation, "ʬ�ϴ�λ"
End Sub