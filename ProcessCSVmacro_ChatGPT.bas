Option Explicit

' ¥°¥í¡¼¥Ð¥ëÊÑ¿ô¡Ê¥æ¡¼¥¶¡¼¥Õ¥©¡¼¥à´ÉÍýÍÑ¡Ë
Dim gRebillForm As Object          ' Æ°Åª¤ËºîÀ®¤·¤¿ÊÖÌáºÆÀÁµáÁªÂò¥Õ¥©¡¼¥à
Dim gUnclaimedForm As Object       ' Æ°Åª¤ËºîÀ®¤·¤¿Ì¤ÀÁµá¥ì¥»¥×¥ÈÁªÂò¥Õ¥©¡¼¥à
Dim gOlderList As Object           ' ²áµî¥ì¥»¥×¥È¥Ç¡¼¥¿°ìÍ÷¡ÊÊÖÌáºÆÀÁµá/·îÃÙ¤ìÁªÂòÍÑ¡Ë
Dim gUnclaimedList As Object       ' Á°·îÌ¤ÀÁµá¥Ç¡¼¥¿°ìÍ÷¡ÊÌ¤ÀÁµá¥ì¥»¥×¥ÈÁªÂòÍÑ¡Ë
Dim gRebillData As Object          ' ¥æ¡¼¥¶¡¼ÁªÂò·ë²Ì¡§ÊÖÌáºÆÀÁµá¤ËÊ¬Îà¤¹¤ë¥Ç¡¼¥¿
Dim gLateData As Object            ' ¥æ¡¼¥¶¡¼ÁªÂò·ë²Ì¡§·îÃÙ¤ìÀÁµá¤ËÊ¬Îà¤¹¤ë¥Ç¡¼¥¿
Dim gSelectedUnclaimed As Object   ' ¥æ¡¼¥¶¡¼ÁªÂò·ë²Ì¡§Á°·îÌ¤ÀÁµá¤«¤éÄÉ²Ã¤¹¤ë¥Ç¡¼¥¿

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

    ' 1. CSV¥Õ¥©¥ë¥À¤ò¥æ¡¼¥¶¡¼¤ËÁªÂò¤µ¤»¤ë
    csvFolder = SelectCSVFolder()
    If csvFolder = "" Then Exit Sub

    ' 1.1 ¥Õ¥©¥ë¥À¤¬¶õ¤Ê¤é½èÍý¤òÃæ»ß
    If IsFolderEmpty(csvFolder) Then
        MsgBox "ÁªÂò¤·¤¿¥Õ¥©¥ë¥À¤Ë¤ÏCSV¥Õ¥¡¥¤¥ë¤¬¤¢¤ê¤Þ¤»¤ó¡£½èÍý¤òÃæ»ß¤·¤Þ¤¹¡£", vbExclamation, "¥¨¥é¡¼"
        Exit Sub
    End If

    ' 2. ¥Æ¥ó¥×¥ì¡¼¥È¥Ñ¥¹¡¦ÊÝÂ¸¥Õ¥©¥ë¥À¼èÆÀ
    templatePath = GetTemplatePath()
    savePath = GetSavePath()
    If templatePath = "" Or savePath = "" Then Exit Sub

    ' 3. ¥Õ¥¡¥¤¥ë¥·¥¹¥Æ¥à¥ª¥Ö¥¸¥§¥¯¥È¤ÎºîÀ®
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 4. ¥Õ¥©¥ë¥ÀÆâ¤Î¤¹¤Ù¤Æ¤Î¡Öfixf¡×¥Õ¥¡¥¤¥ë¤ò¼èÆÀ
    Set fixfFiles = FindAllFixfFiles(fso, csvFolder)

    ' 5. fixf¥Õ¥¡¥¤¥ë¤¬¤Ê¤¤¾ì¹ç¤ÏÄÌ¾ï¤ÎCSV½èÍý¤ËÀÚ¤êÂØ¤¨
    If fixfFiles Is Nothing Or fixfFiles.Count = 0 Then
        ProcessWithoutFixf fso, csvFolder, savePath, templatePath
        Exit Sub
    End If

    ' 6. Ê£¿ô¤Îfixf¥Õ¥¡¥¤¥ë¤ò½çÈÖ¤Ë½èÍý
    For Each file In fixfFiles
        fixfFile = file.Path

        ' 7. ÂÐ¾ÝÇ¯·î¤ò¼èÆÀ
        targetYear = "": targetMonth = ""
        GetYearMonthFromFixf fixfFile, targetYear, targetMonth

        ' ÂÐ¾ÝÇ¯·î¤¬¼èÆÀ¤Ç¤­¤Ê¤«¤Ã¤¿¾ì¹ç¤Ï¥¹¥­¥Ã¥×
        If targetYear = "" Or targetMonth = "" Then
            MsgBox "¥Õ¥¡¥¤¥ë " & fixfFile & " ¤«¤é¿ÇÎÅÇ¯·î¤ò¼èÆÀ¤Ç¤­¤Þ¤»¤ó¤Ç¤·¤¿¡£", vbExclamation, "¥¨¥é¡¼"
            GoTo NextFile
        End If

        ' 8. ÂÐ¾ÝExcel¥Õ¥¡¥¤¥ë¤¬´û¤ËÂ¸ºß¤¹¤ë¤«³ÎÇ§¡ÊÂ¸ºß¤¹¤ë¾ì¹ç¤Ï¥¹¥­¥Ã¥×¡Ë
        Dim csvYYMM As String, reportName As String
        csvYYMM = Format(CInt(targetYear) - 2018, "00") & targetMonth
        reportName = "ÊÝ¸±ÀÁµá´ÉÍýÊó¹ð½ñ_R" & csvYYMM & ".xlsx"
        If fso.FileExists(savePath & "\" & reportName) Then
            MsgBox "¿ÇÎÅÇ¯·î " & targetYear & "Ç¯" & targetMonth & "·î ¤ÎÊó¹ð½ñ¤Ï´û¤Ë½èÍýºÑ¤ß¤Ç¤¹¡£", vbInformation, "½èÍýºÑ¤ß"
            GoTo NextFile
        End If

        ' ÂÐ¾ÝExcel¥Õ¥¡¥¤¥ë¤ò¼èÆÀ¡ÊÂ¸ºß¤·¤Ê¤±¤ì¤Ð¥Æ¥ó¥×¥ì¡¼¥È¤«¤é¿·µ¬ºîÀ®¡Ë
        targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)
        If targetFile = "" Then
            MsgBox "¿ÇÎÅÇ¯·î " & targetYear & "Ç¯" & targetMonth & "·î ¤ÎExcel¥Õ¥¡¥¤¥ë¤òºîÀ®¤Ç¤­¤Þ¤»¤ó¤Ç¤·¤¿¡£", vbExclamation, "¥¨¥é¡¼"
            GoTo NextFile
        End If

        ' 9. Excel¤ò³«¤¯
        On Error Resume Next
        Set newBook = Workbooks.Open(targetFile)
        On Error GoTo 0
        If newBook Is Nothing Then
            MsgBox "¥Õ¥¡¥¤¥ë " & targetFile & " ¤ò³«¤±¤Þ¤»¤ó¤Ç¤·¤¿¡£", vbExclamation, "¥¨¥é¡¼"
            GoTo NextFile
        End If

        ' 10. fixf¥Õ¥¡¥¤¥ë¤ÎÆâÍÆ¤ò¥·¡¼¥È1¤Ë¼è¤ê¹þ¤à
        ImportCSVData fixfFile, newBook.Sheets(1), "ÀÁµá³ÎÄê¾õ¶·"

        ' 11. ¥Æ¥ó¥×¥ì¡¼¥È¾ðÊó¤òÀßÄê¡Ê¥·¡¼¥È1¤Ø¤Î´ûÄêÅ¾µ­¤Ï¥¹¥­¥Ã¥×¡Ë
        SetTemplateInfo newBook, targetYear, targetMonth, True

        ' 12. ¥Õ¥©¥ë¥ÀÆâ¤ÎÂ¾¤ÎCSV¥Õ¥¡¥¤¥ë¤ò½ç¤Ë½èÍý¡Êfmei¢ªhenr¢ªzogn¡Ë
        ProcessAllCSVFiles fso, newBook, csvFolder

        ' 13. ÊÝÂ¸¤·¤Æ¥Ö¥Ã¥¯¤òÊÄ¤¸¤ë
        newBook.Save
        newBook.Close
NextFile:
    Next file

    ' 14. ½èÍý´°Î»¥á¥Ã¥»¡¼¥¸
    MsgBox "¤¹¤Ù¤Æ¤Î fixf ¥Õ¥¡¥¤¥ë¤Î½èÍý¤¬´°Î»¤·¤Þ¤·¤¿¡ª", vbInformation, "½èÍý´°Î»"
End Sub

Sub ProcessWithoutFixf(fso As Object, csvFolder As String, savePath As String, templatePath As String)
    Dim targetYear As String, targetMonth As String
    Dim targetFile As String, newBook As Workbook

    ' ÂÐ¾ÝÇ¯·î¤òCSVÆâÍÆ¤«¤é¼èÆÀ
    targetYear = "": targetMonth = ""
    GetYearMonthFromCSV fso, csvFolder, targetYear, targetMonth
    If targetYear = "" Or targetMonth = "" Then
        MsgBox "CSV¥Õ¥¡¥¤¥ë¤«¤é¿ÇÎÅÇ¯·î¤ò¼èÆÀ¤Ç¤­¤Þ¤»¤ó¤Ç¤·¤¿¡£", vbExclamation, "¥¨¥é¡¼"
        Exit Sub
    End If

    ' ÂÐ¾ÝExcel¥Õ¥¡¥¤¥ë¤¬´û¤ËÂ¸ºß¤¹¤ë¾ì¹ç¤Ï¥¹¥­¥Ã¥×
    Dim csvYYMM As String, reportName As String, fsoLocal As Object
    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    csvYYMM = Format(CInt(targetYear) - 2018, "00") & targetMonth
    reportName = "ÊÝ¸±ÀÁµá´ÉÍýÊó¹ð½ñ_R" & csvYYMM & ".xlsx"
    If fsoLocal.FileExists(savePath & "\" & reportName) Then
        MsgBox "¿ÇÎÅÇ¯·î " & targetYear & "Ç¯" & targetMonth & "·î ¤ÎÊó¹ð½ñ¤Ï´û¤Ë½èÍýºÑ¤ß¤Ç¤¹¡£", vbInformation, "½èÍýºÑ¤ß"
        Exit Sub
    End If

    ' ÂÐ¾ÝExcel¥Õ¥¡¥¤¥ë¤ò¼èÆÀ¡ÊÂ¸ºß¤·¤Ê¤±¤ì¤Ð¿·µ¬ºîÀ®¡Ë
    targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)
    If targetFile = "" Then
        MsgBox "¿ÇÎÅÇ¯·î " & targetYear & "Ç¯" & targetMonth & "·î ¤ÎExcel¥Õ¥¡¥¤¥ë¤òºîÀ®¤Ç¤­¤Þ¤»¤ó¤Ç¤·¤¿¡£", vbExclamation, "¥¨¥é¡¼"
        Exit Sub
    End If

    ' ¥Ö¥Ã¥¯¤ò³«¤¯
    On Error Resume Next
    Set newBook = Workbooks.Open(targetFile)
    On Error GoTo 0
    If newBook Is Nothing Then
        MsgBox "¥Õ¥¡¥¤¥ë " & targetFile & " ¤ò³«¤±¤Þ¤»¤ó¤Ç¤·¤¿¡£", vbExclamation, "¥¨¥é¡¼"
        Exit Sub
    End If

    ' ¥Æ¥ó¥×¥ì¡¼¥È¾ðÊó¤òÀßÄê¡ÊÄÌ¾ïÄÌ¤êÀßÄê¡Ë
    SetTemplateInfo newBook, targetYear, targetMonth, False

    ' CSV¥Õ¥¡¥¤¥ë¤ò½ç¼¡½èÍý¡Êfixf¤Ê¤·¤Ç¤âÂ¾¤ÎCSV¤ò½èÍý²ÄÇ½¡Ë
    ProcessAllCSVFiles fso, newBook, csvFolder

    ' ÊÝÂ¸¤·¤Æ¥Ö¥Ã¥¯¤òÊÄ¤¸¤ë
    newBook.Save
    newBook.Close

    MsgBox "CSV¥Õ¥¡¥¤¥ë¤Î½èÍý¤¬´°Î»¤·¤Þ¤·¤¿¡£", vbInformation, "½èÍý´°Î»"
End Sub

Sub SetTemplateInfo(newBook As Workbook, targetYear As String, targetMonth As String, Optional skipSheet1Info As Boolean = False)
    Dim wsTemplate As Worksheet, wsTemplate2 As Worksheet
    Dim receiptYear As Integer, receiptMonth As Integer
    Dim sendMonth As Integer, sendDate As String

    ' À¾ÎñÇ¯¤ÈÄ´ºÞ·î¤Î·×»»
    receiptYear = CInt(targetYear)
    receiptMonth = CInt(targetMonth)

    ' ÀÁµá·î¡ÊÄ´ºÞ·î¤ÎÍâ·î¡Ë¤Î·×»»
    sendMonth = receiptMonth + 1
    If sendMonth = 13 Then sendMonth = 1
    sendDate = sendMonth & "·î10ÆüÀÁµáÊ¬"

    ' ¥·¡¼¥È1(A), ¥·¡¼¥È2(B)¤ò¼èÆÀ
    Set wsTemplate = newBook.Sheets(1)
    Set wsTemplate2 = newBook.Sheets(2)

    ' ¥·¡¼¥ÈÌ¾ÊÑ¹¹¡Ê¥·¡¼¥È1¤ò "R{ÎáÏÂYY}.{M}", ¥·¡¼¥È2¤ò´Ý¿ô»ú¤Î·î¤ËÊÑ¹¹¡Ë
    wsTemplate.Name = "R" & (receiptYear - 2018) & "." & receiptMonth
    wsTemplate2.Name = ConvertToCircledNumber(receiptMonth)

    ' ¾ðÊóÅ¾µ­
    If Not skipSheet1Info Then
        wsTemplate.Range("G2").Value = targetYear & "Ç¯" & targetMonth & "·îÄ´ºÞÊ¬"
        wsTemplate.Range("I2").Value = sendDate
        wsTemplate.Range("J2").Value = ThisWorkbook.Sheets(1).Range("B1").Value  ' ÀÁµáµ¡´Ø¡Ê»ÜÀßÌ¾Åù¡Ë
    End If
    wsTemplate2.Range("H1").Value = targetYear & "Ç¯" & targetMonth & "·îÄ´ºÞÊ¬"
    wsTemplate2.Range("J1").Value = sendDate
    wsTemplate2.Range("L1").Value = ThisWorkbook.Sheets(1).Range("B1").Value     ' ÀÁµáµ¡´Ø¡Ê»ÜÀßÌ¾Åù¡Ë
End Sub

Sub ProcessAllCSVFiles(fso As Object, newBook As Workbook, csvFolder As String)
    Dim csvFile As Object
    Dim fileType As String
    Dim wsDetails As Worksheet
    Dim wsCSV As Worksheet
    Dim sheetName As String
    Dim sheetIndex As Integer

    ' ¥·¡¼¥È2¡Ê¾ÜºÙ¥Ç¡¼¥¿ÍÑ¡Ë¤ò¼èÆÀ
    Set wsDetails = newBook.Sheets(2)

    ' 1. ¿¶¹þ³ÛÌÀºÙ½ñ¡Êfmei¡Ë¤Î½èÍý
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" And InStr(csvFile.Name, "fmei") > 0 Then
            fileType = "¿¶¹þ³ÛÌÀºÙ½ñ"
            sheetName = fso.GetBaseName(csvFile.Name)
            sheetName = GetUniqueSheetName(newBook, sheetName)
            sheetIndex = Application.WorksheetFunction.Min(3, newBook.Sheets.Count + 1)
            Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(sheetIndex - 1))
            wsCSV.Name = sheetName
            ImportCSVData csvFile.Path, wsCSV, fileType
            Call TransferBillingDetails(newBook, sheetName, csvFile.Name)
        End If
    Next csvFile

    ' 2. ÊÖÌáÆâÌõ½ñ¡Êhenr¡Ë¤Î½èÍý
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" And InStr(csvFile.Name, "henr") > 0 Then
            fileType = "ÊÖÌáÆâÌõ½ñ"
            sheetName = fso.GetBaseName(csvFile.Name)
            sheetName = GetUniqueSheetName(newBook, sheetName)
            sheetIndex = Application.WorksheetFunction.Min(3, newBook.Sheets.Count + 1)
            Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(sheetIndex - 1))
            wsCSV.Name = sheetName
            ImportCSVData csvFile.Path, wsCSV, fileType
            Call TransferBillingDetails(newBook, sheetName, csvFile.Name)
        End If
    Next csvFile

    ' 3. Áý¸ºÅÀÏ¢Íí½ñ¡Êzogn¡Ë¤Î½èÍý
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" And InStr(csvFile.Name, "zogn") > 0 Then
            fileType = "Áý¸ºÅÀÏ¢Íí½ñ"
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

    ' ¥·¡¼¥ÈÀßÄê¡ÊÀÁµá¥Ç¡¼¥¿¥·¡¼¥È¤È¾ÜºÙ¥·¡¼¥È¡Ë
    Set wsBilling = newBook.Sheets(1)
    Set wsDetails = newBook.Sheets(2)

    ' ¿ÇÎÅÇ¯·î¡ÊYYMM·Á¼°¡Ë¤ò¼èÆÀ
    Dim csvYYMM As String
    csvYYMM = Right(CStr(wsBilling.Cells(2, 2).Value), 4)

    ' CSV¥Õ¥¡¥¤¥ëÌ¾¤«¤éÀÁµáÀè¶èÊ¬¤òÈ½ÊÌ
    payerCode = Mid(sheetName, 7, 1)
    Select Case payerCode
        Case "1": payerType = "¼ÒÊÝ"
        Case "2": payerType = "¹ñÊÝ"
        Case Else: payerType = "Ï«ºÒ"
    End Select

    ' ³«»Ï¹Ô°ÌÃÖ¤Î¼­½ñ¤òºîÀ®¡Ê¥·¡¼¥È2¤Î³Æ¥«¥Æ¥´¥ê¸«½Ð¤·¹Ô¤ò¼èÆÀ¡Ë
    Set startRowDict = CreateObject("Scripting.Dictionary")
    If payerType = "¼ÒÊÝ" Then
        startRowDict.Add "ÊÖÌáºÆÀÁµá", GetStartRow(wsDetails, "¼ÒÊÝÊÖÌáºÆÀÁµá")
        startRowDict.Add "·îÃÙ¤ìÀÁµá", GetStartRow(wsDetails, "¼ÒÊÝ·îÃÙ¤ìÀÁµá")
        startRowDict.Add "ÊÖÌá¡¦ººÄê", GetStartRow(wsDetails, "¼ÒÊÝÊÖÌá¡¦ººÄê")
        startRowDict.Add "Ì¤ÀÁµá°·¤¤", GetStartRow(wsDetails, "¼ÒÊÝÌ¤ÀÁµá°·¤¤")
    ElseIf payerType = "¹ñÊÝ" Then
        startRowDict.Add "ÊÖÌáºÆÀÁµá", GetStartRow(wsDetails, "¹ñÊÝÊÖÌáºÆÀÁµá")
        startRowDict.Add "·îÃÙ¤ìÀÁµá", GetStartRow(wsDetails, "¹ñÊÝ·îÃÙ¤ìÀÁµá")
        startRowDict.Add "ÊÖÌá¡¦ººÄê", GetStartRow(wsDetails, "¹ñÊÝÊÖÌá¡¦ººÄê")
        startRowDict.Add "Ì¤ÀÁµá°·¤¤", GetStartRow(wsDetails, "¹ñÊÝÌ¤ÀÁµá°·¤¤")
    End If

    ' ³Æ¥«¥Æ¥´¥êÍÑ¤Î¥Ç¥£¥¯¥·¥ç¥Ê¥ê¤òºîÀ®
    Set rebillDict = CreateObject("Scripting.Dictionary")    ' ÊÖÌáºÆÀÁµá
    Set lateDict = CreateObject("Scripting.Dictionary")      ' ·îÃÙ¤ìÀÁµá
    Set unpaidDict = CreateObject("Scripting.Dictionary")    ' Ì¤ÀÁµá°·¤¤
    Set assessmentDict = CreateObject("Scripting.Dictionary") ' ÊÖÌá¡¦ººÄê

    ' ÀÁµá¥Ç¡¼¥¿¥·¡¼¥È¤ÎºÇ½ª¹Ô¤ò¼èÆÀ
    lastRowBilling = wsBilling.Cells(Rows.Count, "D").End(xlUp).Row

    ' fixf¥Õ¥¡¥¤¥ë¤¬¤Ê¤¤¾ì¹ç¡ÊÀÁµá¥Ç¡¼¥¿¤¬¥·¡¼¥È1¤Ë¤Ê¤¤¾ì¹ç¡Ë¡¢³ÆCSV¤«¤é¾ÜºÙ¥Ç¡¼¥¿¤òÅ¾µ­
    If lastRowBilling < 2 Then
        If InStr(csvFileName, "henr") > 0 Then
            Set wsCSV = newBook.Sheets(sheetName)
            lastRowDetails = wsCSV.Cells(Rows.Count, 1).End(xlUp).Row
            For j = 2 To lastRowDetails
                dispensingMonth = CStr(wsCSV.Cells(j, 1).Value)
                If dispensingMonth <> "" Then
                    If Len(dispensingMonth) = 4 Then dispensingMonth = "5" & dispensingMonth
                    convertedMonth = ConvertToWesternDate(dispensingMonth)
                    rowData = Array(wsCSV.Cells(j, 2).Value, convertedMonth, wsCSV.Cells(j, 5).Value, wsCSV.Cells(j, 14).Value)  ' ¼õÉÕÈÖ¹æ, ¿ÇÎÅ·î, ÀÁµáÅÀ¿ô, »öÍ³¥³¡¼¥É
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
                    rowData = Array(wsCSV.Cells(j, 2).Value, convertedMonth, wsCSV.Cells(j, 6).Value, wsCSV.Cells(j, 7).Value)  ' ¼õÉÕÈÖ¹æ, Ä´ºÞ·î, Áý¸ºÅÀ¿ô, »öÍ³
                    unpaidDict.Add CStr(wsCSV.Cells(j, 2).Value) & "_" & j, rowData
                End If
            Next j
        End If
    End If

    ' ÀÁµá¥Ç¡¼¥¿¡Êfixf¡Ë¤ò¥Ç¥£¥¯¥·¥ç¥Ê¥ê¤Ë³ÊÇ¼¡Êfixf¥Õ¥¡¥¤¥ë¤¬¤¢¤ë¾ì¹ç¤Î¤ß³ºÅö¡Ë
    Dim dispGYM As String
    For i = 2 To lastRowBilling
        dispensingMonth = wsBilling.Cells(i, 2).Value       ' GYYMM·Á¼°¤Î¿ÇÎÅ·î
        convertedMonth = ConvertToWesternDate(dispensingMonth)
        rowData = Array(wsBilling.Cells(i, 4).Value, convertedMonth, wsBilling.Cells(i, 5).Value, wsBilling.Cells(i, 10).Value)
        ' ÂÐ¾Ý¿ÇÎÅ·î¡ÊcsvYYMM¡Ë¤È°Û¤Ê¤ë¾ì¹ç¤Î¤ß³Æ¥«¥Æ¥´¥ê¤ËÄÉ²Ã
        If Right(dispensingMonth, 4) <> csvYYMM Then
            If InStr(csvFileName, "fixf") > 0 Then
                ' fixf¥¨¥ó¥È¥ê¤Î¾ì¹ç¡¢¥æ¡¼¥¶¡¼¤ËÊÖÌáºÆÀÁµá¤«·îÃÙ¤ìÀÁµá¤«ÁªÂò¤µ¤»¤ë
                If ShowRebillSelectionForm(rowData) Then
                    rebillDict.Add wsBilling.Cells(i, 1).Value, rowData   ' ÊÖÌáºÆÀÁµá
                Else
                    lateDict.Add wsBilling.Cells(i, 1).Value, rowData    ' ·îÃÙ¤ìÀÁµá
                End If
            ElseIf InStr(csvFileName, "zogn") > 0 Then
                unpaidDict.Add wsBilling.Cells(i, 1).Value, rowData      ' Ì¤ÀÁµá°·¤¤
            ElseIf InStr(csvFileName, "henr") > 0 Then
                assessmentDict.Add wsBilling.Cells(i, 1).Value, rowData  ' ÊÖÌá¡¦ººÄê
            End If
        End If
    Next i

    ' ³Æ¥«¥Æ¥´¥ê¤ÎÄÉ²Ã¹Ô¿ô¤ò·×»»¡Ê³Æ¥«¥Æ¥´¥ê4¹Ô¤òÄ¶¤¨¤ëÊ¬¡Ë
    a = 0: b = 0: c = 0
    If rebillDict.Count > 4 Then a = rebillDict.Count - 4
    If lateDict.Count > 4 Then b = lateDict.Count - 4
    If assessmentDict.Count > 4 Then c = assessmentDict.Count - 4

    ' ³Æ¥«¥Æ¥´¥ê¤Î³«»Ï¹Ô¤òÄ´À°
    Dim lateStartRow As Long, assessmentStartRow As Long, unpaidStartRow As Long
    lateStartRow = startRowDict("·îÃÙ¤ìÀÁµá") + 1 + a
    assessmentStartRow = startRowDict("ÊÖÌá¡¦ººÄê") + 1 + a + b
    unpaidStartRow = startRowDict("Ì¤ÀÁµá°·¤¤") + 1 + a + b + c

    ' É¬Í×¤Ë±þ¤¸¤Æ¹Ô¤òÁÞÆþ¤·¤ÆÏÈ¤ò³ÎÊÝ
    If a + b + c > 0 Then
        wsDetails.Rows(lateStartRow & ":" & lateStartRow + a).Insert Shift:=xlDown
        wsDetails.Rows(assessmentStartRow & ":" & assessmentStartRow + b).Insert Shift:=xlDown
        wsDetails.Rows(unpaidStartRow & ":" & unpaidStartRow + c).Insert Shift:=xlDown
    End If

    ' ³Æ¥Ç¥£¥¯¥·¥ç¥Ê¥ê¤Î¥Ç¡¼¥¿¤ò¥·¡¼¥È2¤ËÅ¾µ­¡Ê¥Ç¥£¥¯¥·¥ç¥Ê¥ê¤¬¶õ¤Î¾ì¹ç¤Ï¥¹¥­¥Ã¥×¡Ë
    If rebillDict.Count > 0 Then
        j = startRowDict("ÊÖÌáºÆÀÁµá")
        TransferData rebillDict, wsDetails, j, payerType
    End If
    If lateDict.Count > 0 Then
        j = startRowDict("·îÃÙ¤ìÀÁµá")
        TransferData lateDict, wsDetails, j, payerType
    End If
    If unpaidDict.Count > 0 Then
        j = startRowDict("Ì¤ÀÁµá°·¤¤")
        TransferData unpaidDict, wsDetails, j, payerType
    End If
    If assessmentDict.Count > 0 Then
        j = startRowDict("ÊÖÌá¡¦ººÄê")
        TransferData assessmentDict, wsDetails, j, payerType
    End If

    MsgBox payerType & " ¤Î¥Ç¡¼¥¿Å¾µ­¤¬´°Î»¤·¤Þ¤·¤¿¡ª", vbInformation, "½èÍý´°Î»"
End Sub

Function SelectCSVFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "CSV¥Õ¥©¥ë¥À¤òÁªÂò¤·¤Æ¤¯¤À¤µ¤¤"
        If .Show = -1 Then
            SelectCSVFolder = .SelectedItems(1) & "\"
        Else
            MsgBox "¥Õ¥©¥ë¥À¤¬ÁªÂò¤µ¤ì¤Þ¤»¤ó¤Ç¤·¤¿¡£½èÍý¤òÃæ»ß¤·¤Þ¤¹¡£", vbExclamation, "³ÎÇ§"
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
        IsFolderEmpty = True   ' **¥Õ¥©¥ë¥À¤Ë¥Õ¥¡¥¤¥ë¤¬¤Ê¤¤¾ì¹ç True**
    Else
        IsFolderEmpty = False
    End If
End Function

Function GetTemplatePath() As String
    ' ¥Æ¥ó¥×¥ì¡¼¥È¥Õ¥¡¥¤¥ë¤Î¥Ñ¥¹¤ò¥·¡¼¥È1¤Î¥»¥ëB2¤«¤é¼èÆÀ
    GetTemplatePath = ThisWorkbook.Sheets(1).Range("B2").Value & "\ÊÝ¸±ÀÁµá´ÉÍýÊó¹ð½ñ¥Æ¥ó¥×¥ì¡¼¥È.xltm"
End Function

Function GetSavePath() As String
    ' ÊÝÂ¸Àè¥Õ¥©¥ë¥À¤Î¥Ñ¥¹¤ò¥·¡¼¥È1¤Î¥»¥ëB3¤«¤é¼èÆÀ
    GetSavePath = ThisWorkbook.Sheets(1).Range("B3").Value
End Function

Function FindAllFixfFiles(fso As Object, csvFolder As String) As Collection
    Dim csvFile As Object
    Dim fixfFiles As New Collection
    ' **¥Õ¥©¥ë¥ÀÆâ¤Î¤¹¤Ù¤Æ¤Î¥Õ¥¡¥¤¥ë¤ò¥Á¥§¥Ã¥¯**
    For Each csvFile In fso.GetFolder(csvFolder).Files
        ' **³ÈÄ¥»Ò¤¬ "csv" ¤Ç¤¢¤ê¡¢Ì¾Á°¤Ë "fixf" ¤ò´Þ¤à¾ì¹ç**
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" And InStr(LCase(csvFile.Name), "fixf") > 0 Then
            fixfFiles.Add csvFile  ' **fixf¥Õ¥¡¥¤¥ë¤ò¥ê¥¹¥È¤ËÄÉ²Ã**
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
    ' **fixf¥Õ¥¡¥¤¥ëÌ¾¤«¤éÇ¯·î¥³¡¼¥É¤ò¿äÄê**
    code = ""
    '  Â¾¤ÎCSV¥Õ¥¡¥¤¥ëÌ¾¤«¤é GYYMM ¤ò¼èÆÀ¡ÊÎã: ¿¶¹þ³ÛÌÀºÙ½ñ¤Ê¤É¡Ë 
    Dim csvFile As Object, folderPath As String
    folderPath = fso.GetFile(fixfFile).ParentFolder.Path
    For Each csvFile In fso.GetFolder(folderPath).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" Then
            If InStr(LCase(csvFile.Name), "fmei") > 0 Or InStr(LCase(csvFile.Name), "zogn") > 0 Or InStr(LCase(csvFile.Name), "henr") > 0 Then
                ' Ì¾Á°ËöÈø4·å¤¬¿ô»ú¤Ê¤é¤½¤ì¤òÇ¯·î¥³¡¼¥É¤È¤¹¤ë
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
        targetYear = CStr(2018 + CInt(yrCode))    ' **ÏÂÎñÇ¯¥³¡¼¥É¤òÀ¾ÎñÇ¯¤ËÊÑ´¹**
        targetMonth = CStr(CInt(monCode))         ' **·î¥³¡¼¥É¡ÊÀèÆ¬0´Þ¤à¡Ë¤òÀ°¿ô²½**
    Else
        ' **fallback: fixf¥Õ¥¡¥¤¥ë¤«¤é¿ÇÎÅÇ¯·î¤òÃê½Ð**¡Ê´Ê°×¡Ë
        Dim ts As Object, lineText As String
        On Error Resume Next
        Set ts = fso.OpenTextFile(fixfFile, 1, False, -2)
        On Error GoTo 0
        If Not ts Is Nothing Then
            ' ÀèÆ¬¿ô¹Ô¤òÆÉ¤ß¿ÇÎÅÇ¯·î¤ò´Þ¤à¹Ô¤òÃµ¤¹
            Dim i As Integer
            For i = 1 To 5
                If ts.AtEndOfStream Then Exit For
                lineText = ts.ReadLine
                If InStr(lineText, "G") > 0 And InStr(lineText, ",") = 0 Then
                    ' **Îã: "5XXXX" ·Á¼°¤ÎÊ¸»úÎó¤ò´Þ¤à¾ì¹ç**
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
        ' **¼èÆÀ¼ºÇÔ»þ¡¢¥æ¡¼¥¶¡¼¤ËÆþÎÏ¤òÂ¥¤¹**
        If targetYear = "" Or targetMonth = "" Then
            MsgBox "¿ÇÎÅÇ¯·î¤ò¼«Æ°¼èÆÀ¤Ç¤­¤Þ¤»¤ó¤Ç¤·¤¿¡£»ØÄê¤·¤Æ¤¯¤À¤µ¤¤¡£", vbExclamation, "³ÎÇ§"
            targetYear = InputBox("À¾ÎñÇ¯¤òÆþÎÏ¤·¤Æ¤¯¤À¤µ¤¤¡ÊÎã: 2023¡Ë:", "¿ÇÎÅÇ¯")
            targetMonth = InputBox("·î¤òÆþÎÏ¤·¤Æ¤¯¤À¤µ¤¤¡Ê1¢·12¡Ë:", "¿ÇÎÅ·î")
            If targetYear = "" Or targetMonth = "" Then
                ' ¥æ¡¼¥¶¡¼Ì¤ÆþÎÏ¤Î¾ì¹ç
                Exit Sub
            End If
        End If
    End If
End Sub

Function FindOrCreateReport(savePath As String, targetYear As String, targetMonth As String, templatePath As String) As String
    Dim fso As Object, existingFile As Object
    Dim fileName As String, filePath As String
    Dim csvYYMM As String
    csvYYMM = Format(CInt(targetYear) - 2018, "00") & Format(CInt(targetMonth), "00")  ' **ÏÂÎñÇ¯+·î¥³¡¼¥É**
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' **ÊÝÂ¸¥Õ¥©¥ë¥ÀÆâ¤Ë´ûÂ¸¤Î RYYMM ¥Õ¥¡¥¤¥ë¤¬¤¢¤ë¤«³ÎÇ§**
    For Each existingFile In fso.GetFolder(savePath).Files
        If LCase(fso.GetExtensionName(existingFile.Name)) = "xlsm" Or LCase(fso.GetExtensionName(existingFile.Name)) = "xlsx" Then
            If InStr(existingFile.Name, "ÊÝ¸±ÀÁµá´ÉÍýÊó¹ð½ñ_R" & csvYYMM) > 0 Then
                FindOrCreateReport = existingFile.Path  ' **´ûÂ¸¥Õ¥¡¥¤¥ë¤Î¥Ñ¥¹¤òÊÖ¤¹**
                Exit Function
            End If
        End If
    Next existingFile
    ' **³ºÅö¤¹¤ë¥Õ¥¡¥¤¥ë¤¬¤Ê¤±¤ì¤Ð¡¢¿·µ¬ºîÀ®**
    fileName = "ÊÝ¸±ÀÁµá´ÉÍýÊó¹ð½ñ_R" & csvYYMM & ".xlsm"   ' **xlsm·Á¼°¤ÇÊÝÂ¸**¡Ê²þÎÉÅÀ1¡Ë
    filePath = savePath & "\" & fileName
    ' **¥Æ¥ó¥×¥ì¡¼¥È¤ò¸µ¤Ë¿·µ¬¥Ö¥Ã¥¯ºîÀ®**
    On Error Resume Next
    Dim tmplWb As Workbook
    Set tmplWb = Workbooks.Open(templatePath)   ' **¥Æ¥ó¥×¥ì¡¼¥È¥Ö¥Ã¥¯¤ò³«¤¯**
    On Error GoTo 0
    If tmplWb Is Nothing Then
        MsgBox "¥Æ¥ó¥×¥ì¡¼¥È¤ò³«¤±¤Þ¤»¤ó¤Ç¤·¤¿: " & templatePath, vbCritical, "¥¨¥é¡¼"
        FindOrCreateReport = ""
        Exit Function
    End If
    On Error Resume Next
    tmplWb.SaveAs filePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled  ' **xlsm·Á¼°¤ÇÊÝÂ¸**
    If Err.Number <> 0 Then
        MsgBox "¥Õ¥¡¥¤¥ë¤òÊÝÂ¸¤Ç¤­¤Þ¤»¤ó¤Ç¤·¤¿: " & filePath, vbCritical, "¥¨¥é¡¼"
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

    ' ¹àÌÜ¥Þ¥Ã¥Ô¥ó¥°¤ò¼èÆÀ
    Set colMap = GetColumnMapping(fileType)
    ' ¥·¡¼¥È¤ò¥¯¥ê¥¢¤·¤Æ¹àÌÜÌ¾¤ò1¹ÔÌÜ¤ËÀßÄê
    ws.Cells.Clear
    j = 1
    For Each key In colMap.Keys
        ws.Cells(1, j).Value = colMap(key)
        j = j + 1
    Next key

    ' CSV¥Õ¥¡¥¤¥ë¤òUTF-8¥Æ¥­¥¹¥È¤È¤·¤Æ³«¤¯
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(csvFile, 1, False, -2)  ' -2: UTF-8

    ' ¥Ç¡¼¥¿ÉôÊ¬¤òÅ¾µ­
    i = 2
    isHeader = True
    Do While Not ts.AtEndOfStream
        lineText = ts.ReadLine
        dataArray = Split(lineText, ",")
        If isHeader Then
            ' °ì¹ÔÌÜ¡Ê¥Ø¥Ã¥À¡¼¹Ô¡Ë¤Ï¥¹¥­¥Ã¥×
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
    MsgBox "CSVÆÉ¹þÃæ¤Ë¥¨¥é¡¼¤¬È¯À¸¤·¤Þ¤·¤¿: " & Err.Description, vbCritical, "¥¨¥é¡¼"
    If Not ts Is Nothing Then ts.Close
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Function GetColumnMapping(fileType As String) As Object
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    Select Case fileType
        Case "¿¶¹þ³ÛÌÀºÙ½ñ"
            colMap.Add 2, "¿¶¹þÇ¯·î"
            colMap.Add 3, "¿¶¹þ¶â³Û"
            colMap.Add 4, "¿¶¹þÆü"
            ' ¡ÊÉ¬Í×¤ÊÎó¤òÄÉ²Ã¡Ë
        Case "Áý¸ºÅÀÏ¢Íí½ñ"
            colMap.Add 2, "Ä´ºÞÇ¯·î"
            colMap.Add 4, "¼õÉÕÈÖ¹æ"
            colMap.Add 11, "¶èÊ¬"
            colMap.Add 14, "Ï·¿Í¸ºÌÈ¶èÊ¬"
            colMap.Add 15, "»áÌ¾"
            colMap.Add 21, "Áý¸ºÅÀ¿ô¡Ê¶â³Û¡Ë"
            colMap.Add 22, "»öÍ³"
        Case "ÊÖÌáÆâÌõ½ñ"
            colMap.Add 2, "Ä´ºÞÇ¯·î(YYMM·Á¼°)"
            colMap.Add 3, "¼õÉÕÈÖ¹æ"
            colMap.Add 4, "ÊÝ¸±¼ÔÈÖ¹æ"
            colMap.Add 7, "»áÌ¾"
            colMap.Add 9, "ÀÁµáÅÀ¿ô"
            colMap.Add 10, "ÌôºÞ°ìÉôÉéÃ´¶â"
            colMap.Add 12, "°ìÉôÉéÃ´¶â³Û"
            colMap.Add 13, "´µ¼ÔÉéÃ´¶â³Û¡Ê¸øÈñ¡Ë"
            colMap.Add 14, "»öÍ³¥³¡¼¥É"
        Case Else
            ' ¤½¤ÎÂ¾¡ÊÉ¬Í×¤Ë±þ¤¸¤ÆÄÉ²Ã¡Ë
            colMap.Add 1, "¹àÌÜ1"
    End Select
    Set GetColumnMapping = colMap
End Function

Function GetStartRow(ws As Worksheet, category As String) As Long
    ' ¾ÜºÙ¥·¡¼¥È¤«¤é»ØÄê¥«¥Æ¥´¥ê¤Î¹ÔÈÖ¹æ¤ò¼èÆÀ
    Dim rng As Range
    Set rng = ws.Cells.Find(what:=category, LookAt:=xlWhole)
    If rng Is Nothing Then
        MsgBox "¾ÜºÙ¥·¡¼¥È¾å¤Ç¥«¥Æ¥´¥ê """ & category & """ ¤ò¸«¤Ä¤±¤é¤ì¤Þ¤»¤ó¤Ç¤·¤¿¡£", vbExclamation, "¥¨¥é¡¼"
        GetStartRow = 0
    Else
        GetStartRow = rng.Row
    End If
End Function

Sub TransferData(dataDict As Object, ws As Worksheet, startRow As Long, payerType As String)
    Dim key As Variant, rowData As Variant
    Dim j As Long, payerColumn As Long
    ' **Dictionary¤¬¶õ¤Ê¤é½èÍý¤·¤Ê¤¤**
    If dataDict.Count = 0 Then Exit Sub
    ' **payerType¤Ë±þ¤¸¤¿Å¾µ­Îó¤ò·èÄê**
    If payerType = "¼ÒÊÝ" Then
        payerColumn = 8   ' ¼ÒÊÝ¤ÏHÎó¤ËÀÁµáÀè¥Þ¡¼¥¯
    ElseIf payerType = "¹ñÊÝ" Then
        payerColumn = 9   ' ¹ñÊÝ¤ÏIÎó¤ËÀÁµáÀè¥Þ¡¼¥¯
    Else
        payerColumn = 8   ' ¡ÊÏ«ºÒÅù¤Ï¼ÒÊÝÎó¤Ë²¾ÀßÄê¡Ë
    End If
    j = startRow
    For Each key In dataDict.Keys
        rowData = dataDict(key)
        ws.Cells(j, 4).Value = rowData(0)    ' ´µ¼Ô»áÌ¾
        ws.Cells(j, 5).Value = rowData(1)    ' Ä´ºÞÇ¯·î¡ÊÀ¾ÎñÉ½µ­¡Ë
        ws.Cells(j, 6).Value = rowData(2)    ' °åÎÅµ¡´ØÌ¾
        ws.Cells(j, payerColumn).Value = payerType   ' ÀÁµáÀè¶èÊ¬¡Ê¼ÒÊÝ/¹ñÊÝ¡Ë
        ws.Cells(j, payerColumn).Font.Bold = True    ' **¶¯Ä´É½¼¨**
        ws.Cells(j, 10).Value = rowData(3)   ' ÀÁµáÅÀ¿ô
        j = j + 1
    Next key
End Sub

Sub ShowRebillSelectionForm()
    ' ²áµî·î¥ì¥»¥×¥È¤Î°ìÍ÷¤ò¥æ¡¼¥¶¡¼¤ËÉ½¼¨¤·¡¢ÊÖÌáºÆÀÁµá¤¹¤ë¤â¤Î¤òÁªÂò¤·¤Æ¤â¤é¤¦
    Dim uf As Object, listData As Object
    Set listData = gOlderList
    If listData Is Nothing Or listData.Count = 0 Then Exit Sub
    ' ¥æ¡¼¥¶¡¼¥Õ¥©¡¼¥àºîÀ®¤ÈÉ½¼¨
    Set uf = CreateRebillSelectionForm(listData)
    Set gRebillForm = uf  ' ¥°¥í¡¼¥Ð¥ë»²¾ÈÊÝÂ¸
    gRebillForm.Show vbModal
    ' ¥Õ¥©¡¼¥àÊÄ¤¸¤¿¸å¡¢ÁªÂò·ë²Ì¤Ï gRebillData ¤È gLateData ¤Ë³ÊÇ¼ºÑ¤ß¡ÊProcessRebillSelection¤ÇÀßÄê¡Ë
End Sub

Function CreateRebillSelectionForm(listData As Object) As Object
    Dim uf As Object, listBox As Object, btnOK As Object
    Dim i As Long, rowData As Variant
    ' **UserForm ¤òÆ°Åª¤ËºîÀ®**
    Set uf = VBA.UserForms.Add()  ' ¿·µ¬UserForm
    uf.Caption = "ÊÖÌáºÆÀÁµá¤ÎÁªÂò"
    uf.Width = 400
    uf.Height = 500
    ' **ListBox¤òÄÉ²Ã**
    Set listBox = uf.Controls.Add("Forms.ListBox.1", "listBox", True)
    listBox.Left = 20
    listBox.Top = 20
    listBox.Width = 350
    listBox.Height = 350
    listBox.MultiSelect = 1  ' **Ê£¿ôÁªÂò²ÄÇ½**
    ' **¥ê¥¹¥È¤Ë¥Ç¡¼¥¿¤òÄÉ²Ã¡ÊÄ´ºÞÇ¯·î | ´µ¼Ô»áÌ¾ | °åÎÅµ¡´ØÌ¾ | ÅÀ¿ô¡Ë**
    For i = 0 To listData.Count - 1
        rowData = listData.Items()(i)
        listBox.AddItem rowData(1) & " | " & rowData(0) & " | " & rowData(2) & " | " & rowData(3)
    Next i
    ' **OK¥Ü¥¿¥ó¤òÄÉ²Ã**
    Set btnOK = uf.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    btnOK.Caption = "³ÎÄê"
    btnOK.Left = 150
    btnOK.Top = 400
    btnOK.Width = 100
    btnOK.Height = 30
    ' **¥Ü¥¿¥ó¥¯¥ê¥Ã¥¯»þ¤Î½èÍý¤òÀßÄê**
    btnOK.OnClick = "ProcessRebillSelection"
    Set CreateRebillSelectionForm = uf
End Function

Sub ProcessRebillSelection()
    ' ÊÖÌáºÆÀÁµáÁªÂò¥Õ¥©¡¼¥à¤ÎOK¥Ü¥¿¥ó½èÍý¡ÊÁªÂò¤µ¤ì¤¿¹àÌÜ¤òÊ¬Îà¡Ë
    Dim uf As Object, listBox As Object
    Dim i As Long
    ' Æ°Åª¥Õ¥©¡¼¥à¤ª¤è¤ÓListBox¤ò¼èÆÀ
    Set uf = gRebillForm
    Set listBox = uf.Controls("listBox")
    ' ·ë²ÌÍÑDictionary¤ò½é´ü²½
    Set gRebillData = CreateObject("Scripting.Dictionary")
    Set gLateData = CreateObject("Scripting.Dictionary")
    ' **ÁªÂò¾õÂÖ¤Ë±þ¤¸¤Æ¿¶¤êÊ¬¤±**
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) Then
            ' ÁªÂò¤µ¤ì¤¿¤â¤Î -> ÊÖÌáºÆÀÁµá
            gRebillData.Add gOlderList.Keys()(i), gOlderList.Items()(i)
        Else
            ' ÁªÂò¤µ¤ì¤Ê¤«¤Ã¤¿¤â¤Î -> ·îÃÙ¤ìÀÁµá
            gLateData.Add gOlderList.Keys()(i), gOlderList.Items()(i)
        End If
    Next i
    ' ¥Õ¥©¡¼¥à¤ò¥¢¥ó¥í¡¼¥É¤·¤ÆÊÄ¤¸¤ë
    Unload uf
    Set gRebillForm = Nothing
End Sub

Function AddUnclaimedRecords(payerType As String, targetYear As String, targetMonth As String) As Object
    Dim prevYear As String, prevMonth As String
    Dim prevFileName As String, prevFilePath As String
    Dim prevBook As Workbook, wsPrevDetails As Worksheet
    Dim startRow As Long, endRow As Long, row As Long
    ' Á°·î¤ò»»½Ð
    If CInt(targetMonth) = 1 Then
        prevYear = CStr(CInt(targetYear) - 1)
        prevMonth = "12"
    Else
        prevYear = targetYear
        prevMonth = CStr(CInt(targetMonth) - 1)
    End If
    ' Á°·î¤ÎÊó¹ð½ñ¥Õ¥¡¥¤¥ëÌ¾
    Dim prevYYMM As String
    prevYYMM = Format(CInt(prevYear) - 2018, "00") & Format(CInt(prevMonth), "00")
    prevFileName = "ÊÝ¸±ÀÁµá´ÉÍýÊó¹ð½ñ_R" & prevYYMM & ".xlsm"
    prevFilePath = GetSavePath() & "\" & prevFileName
    If Dir(prevFilePath) = "" Then
        ' ¥Õ¥¡¥¤¥ë¤¬Â¸ºß¤·¤Ê¤¤¾ì¹ç
        Set AddUnclaimedRecords = Nothing
        Exit Function
    End If
    ' Á°·î¥Õ¥¡¥¤¥ë¤ò³«¤¯¡ÊÆÉ¤ß¼è¤êÀìÍÑ¡Ë
    On Error Resume Next
    Set prevBook = Workbooks.Open(prevFilePath, ReadOnly:=True)
    On Error GoTo 0
    If prevBook Is Nothing Then
        Set AddUnclaimedRecords = Nothing
        Exit Function
    End If
    ' Á°·î¾ÜºÙ¥·¡¼¥È¤ò¼èÆÀ¡Ê¥·¡¼¥ÈÌ¾¤Ï¼ÒÊÝ/¹ñÊÝ¶¦ÄÌ¤Ç"B"¥·¡¼¥È¤È²¾Äê¡Ë
    Set wsPrevDetails = prevBook.Sheets(2)
    ' ÂÐ¾Ý¥«¥Æ¥´¥ê¤Î³«»Ï¹Ô¤ò¼èÆÀ
    Dim categoryLabel As String
    If payerType = "¼ÒÊÝ" Then
        categoryLabel = "¼ÒÊÝÌ¤ÀÁµá°·¤¤"
    Else
        categoryLabel = "¹ñÊÝÌ¤ÀÁµá°·¤¤"
    End If
    startRow = GetStartRow(wsPrevDetails, categoryLabel)
    If startRow = 0 Then
        ' ¥é¥Ù¥ë¤¬¸«¤Ä¤«¤é¤Ê¤¤¾ì¹ç¤Ï½ªÎ»
        prevBook.Close SaveChanges:=False
        Set AddUnclaimedRecords = Nothing
        Exit Function
    End If
    ' ³«»Ï¹Ô¤«¤é²¼Êý¸þ¤Ë¥Ç¡¼¥¿¤ò¼ý½¸
    Set gUnclaimedList = CreateObject("Scripting.Dictionary")
    endRow = startRow + 3  ' ´ðËÜÏÈ¤Ï4¹Ô
    ' ¥Ç¡¼¥¿¤¬ÄÉ²Ã¤µ¤ì¤Æ¤¤¤ë¾ì¹ç¡¢¶õÇò¹Ô¤¬½Ð¤ë¤Þ¤Ç±äÄ¹
    Do While wsPrevDetails.Cells(endRow, 4).Value <> "" Or wsPrevDetails.Cells(endRow + 1, 4).Value <> ""
        endRow = endRow + 1
        If endRow > wsPrevDetails.Rows.Count Then Exit Do
    Loop
    For row = startRow + 1 To endRow
        If wsPrevDetails.Cells(row, 4).Value <> "" Then   ' ´µ¼Ô»áÌ¾Îó¤¬¶õ¤Ç¤Ê¤±¤ì¤Ð¥Ç¡¼¥¿¤¢¤ê
            Dim prevRowData As Variant
            prevRowData = Array(wsPrevDetails.Cells(row, 4).Value, wsPrevDetails.Cells(row, 5).Value, wsPrevDetails.Cells(row, 6).Value, wsPrevDetails.Cells(row, 10).Value)
            gUnclaimedList.Add row, prevRowData
        End If
    Next row
    ' Á°·î¥Ö¥Ã¥¯¤òÊÄ¤¸¤ë
    prevBook.Close SaveChanges:=False
    ' ¥æ¡¼¥¶¡¼¤ËÁ°·îÌ¤ÀÁµá¤òÉ½¼¨¤·¡¢ÄÉ²Ã¤¹¤ë¤â¤Î¤òÁªÂò¤µ¤»¤ë
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
    ' ¥Õ¥©¡¼¥à¤¬ÊÄ¤¸¤é¤ì¤¿¸å¡¢gSelectedUnclaimed¤Ë·ë²Ì¤¬³ÊÇ¼¤µ¤ì¤ë
End Sub

Function CreateUnclaimedSelectionForm(listData As Object) As Object
    Dim uf As Object, listBox As Object, btnOK As Object
    Dim i As Long, rowData As Variant
    Set uf = VBA.UserForms.Add()
    uf.Caption = "Á°·î Ì¤ÀÁµá¥ì¥»¥×¥È¤ÎÄÉ²ÃÁªÂò"
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
    btnOK.Caption = "ÄÉ²Ã"
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
    circledNumbers = Array("­¡", "­¢", "­£", "­¤", "­¥", "­¦", "­§", "­¨", "­©", "­ª", "­«", "­¬")
    If month >= 1 And month <= 12 Then
        ConvertToCircledNumber = circledNumbers(month - 1)
    Else
        ConvertToCircledNumber = CStr(month)
    End If
End Function

Function ConvertToWesternDate(dispensingMonth As String) As String
    ' GYYMM·Á¼°¡ÊÏÂÎñ¡Ë¤òÀ¾ÎñÇ¯²¼2·å.·î·Á¼°¤ËÊÑ´¹
    Dim eraCode As String, yearPart As Integer, westernYear As Integer, monthPart As String
    eraCode = Left(dispensingMonth, 1)
    yearPart = CInt(Mid(dispensingMonth, 2, 2))
    monthPart = Right(dispensingMonth, 2)
    Select Case eraCode
        Case "5": westernYear = 2018 + yearPart   ' ÎáÏÂ (2019Ç¯=ÎáÏÂ1Ç¯)
        Case "4": westernYear = 1988 + yearPart   ' Ê¿À® (1989Ç¯=Ê¿À®1Ç¯) - ¢¨²áµî¥Ç¡¼¥¿ÂÐ±þ
        Case Else: westernYear = 2018 + yearPart  ' ¡Ê¥Ç¥Õ¥©¥ë¥È:ÎáÏÂ¤È¤·¤Æ·×»»¡Ë
    End Select
    ConvertToWesternDate = Right(CStr(westernYear), 2) & "." & monthPart
End Function

' **È¾Ç¯¤´¤È¤ÎÇä³Ý¥Ç¡¼¥¿Èæ³Ó¡¦¸íº¹Ê¬ÀÏµ¡Ç½**¡Ê²þÎÉÅÀ6¡Ë
Sub CompareHalfYearData()
    Dim inputYear As String, half As String
    inputYear = InputBox("Ê¬ÀÏ¤¹¤ëÇ¯¤òÆþÎÏ¤·¤Æ¤¯¤À¤µ¤¤¡ÊÀ¾Îñ¡Ë:", "È¾Ç¯Çä³ÝÈæ³Ó")
    If inputYear = "" Then Exit Sub
    half = InputBox("¾å´ü=1 ¤Þ¤¿¤Ï ²¼´ü=2 ¤òÆþÎÏ¤·¤Æ¤¯¤À¤µ¤¤:", "È¾Ç¯¶èÊ¬")
    If half = "" Then Exit Sub
    If half <> "1" And half <> "2" Then
        MsgBox "È¾´ü¶èÊ¬¤Ï1¤Þ¤¿¤Ï2¤ÇÆþÎÏ¤·¤Æ¤¯¤À¤µ¤¤¡£", vbExclamation, "ÆþÎÏ¥¨¥é¡¼"
        Exit Sub
    End If
    Dim startMonth As Integer, endMonth As Integer
    If half = "1" Then
        startMonth = 1: endMonth = 6
    Else
        startMonth = 7: endMonth = 12
    End If
    Dim analysisWb As Workbook
    Set analysisWb = ThisWorkbook  ' ·ë²Ì½ÐÎÏÀè¤ò¥Þ¥¯¥í¥Ö¥Ã¥¯¤ËÀßÄê
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
    outSheet.Range("A1:E1").Value = Array("·î", "Æü¼¡·×¾åÅÀ¿ô", "ÀÁµá³ÎÄêÅÀ¿ô", "¿¶¹þ³Û(±ß)", "ÅÀ¿ôº¹°Û")
    Dim m As Integer, rowIndex As Integer
    rowIndex = 2
    For m = startMonth To endMonth
        Dim yy As String, mm As String, fileCode As String
        yy = Format(CInt(inputYear) - 2018, "00")
        mm = Format(m, "00")
        fileCode = "R" & yy & mm
        Dim reportName As String
        reportName = "ÊÝ¸±ÀÁµá´ÉÍýÊó¹ð½ñ_" & fileCode & ".xlsm"
        Dim reportPath As String
        reportPath = GetSavePath() & "\" & reportName
        If Dir(reportPath) <> "" Then
            Dim repWb As Workbook
            Set repWb = Workbooks.Open(reportPath, ReadOnly:=True)
            Dim wsA As Worksheet, wsCSV As Worksheet, wsCSV2 As Worksheet
            Set wsA = repWb.Sheets(1)   ' Æü¼¡¥Ç¡¼¥¿¥·¡¼¥È
            ' Æü¼¡¥Ç¡¼¥¿¤«¤éÁíÅÀ¿ô¤ò¼èÆÀ¡ÊÎã¤È¤·¤Æ¥·¡¼¥ÈA¤Î¥»¥ëJ¡û¤Ê¤É¤Ë½¸·×¤¬¤¢¤ë¤È²¾Äê¡Ë
            Dim dailyTotal As Long
            dailyTotal = 0
            On Error Resume Next
            dailyTotal = CLng(wsA.Range("J50").Value) ' ¢¨Å¬ÀÚ¤Ê¥»¥ë»²¾È¤ËÍ×½¤Àµ
            On Error GoTo 0
            ' ÀÁµá³ÎÄêÅÀ¿ô¡Êfixf¥Ç¡¼¥¿¤ÎÁíÅÀ¿ô¡Ë¼èÆÀ¡Ê¥·¡¼¥ÈA¤Ë·×¤Þ¤¿¤Ï¥·¡¼¥ÈB¤ËÁí¹ç·×ÅÀ¡©²¾¤ËJ50¤È¤¹¤ë¡Ë
            Dim billedTotal As Long
            billedTotal = 0
            On Error Resume Next
            billedTotal = CLng(wsA.Range("J50").Value)
            On Error GoTo 0
            ' ¿¶¹þ³ÛÌÀºÙ¤«¤é¹ç·×¶â³Û¼èÆÀ¡ÊCSV¥·¡¼¥ÈÌ¾¤Ëfmei´Þ¤àÁÛÄê¡Ë
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
            ' ÅÀ¿ôº¹°Û·×»»
            Dim pointDiff As Long
            pointDiff = dailyTotal - billedTotal
            ' ·ë²Ì¤ò½ÐÎÏ
            outSheet.Cells(rowIndex, 1).Value = inputYear & "Ç¯" & m & "·î"
            outSheet.Cells(rowIndex, 2).Value = dailyTotal
            outSheet.Cells(rowIndex, 3).Value = billedTotal
            outSheet.Cells(rowIndex, 4).Value = payAmount
            outSheet.Cells(rowIndex, 5).Value = pointDiff
            rowIndex = rowIndex + 1
        Else
            ' ¥Õ¥¡¥¤¥ë¤¬¤Ê¤¤¾ì¹ç¤Ï¶õ¹Ô¤Þ¤¿¤Ï0½ÐÎÏ
            outSheet.Cells(rowIndex, 1).Value = inputYear & "Ç¯" & m & "·î"
            outSheet.Cells(rowIndex, 2).Value = "N/A"
            outSheet.Cells(rowIndex, 3).Value = "N/A"
            outSheet.Cells(rowIndex, 4).Value = "N/A"
            outSheet.Cells(rowIndex, 5).Value = "N/A"
            rowIndex = rowIndex + 1
        End If
    Next m
    MsgBox inputYear & "Ç¯ " & IIf(half = "1", "¾å´ü", "²¼´ü") & " ¤ÎÇä³Ý¥Ç¡¼¥¿Èæ³Ó¤¬´°Î»¤·¤Þ¤·¤¿¡£" & vbCrLf & _
            "¥·¡¼¥È[" & outSheet.Name & "]¤Ë·ë²Ì¤ò½ÐÎÏ¤·¤Þ¤·¤¿¡£", vbInformation, "Ê¬ÀÏ´°Î»"
End Sub