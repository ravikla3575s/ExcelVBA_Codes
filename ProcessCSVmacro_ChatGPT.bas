Option Explicit

' グローバル変数（ユーザーフォーム管理用）
Dim gRebillForm As Object          ' 動的に作成した返戻再請求選択フォーム
Dim gUnclaimedForm As Object       ' 動的に作成した未請求レセプト選択フォーム
Dim gOlderList As Object           ' 過去レセプトデータ一覧（返戻再請求/月遅れ選択用）
Dim gUnclaimedList As Object       ' 前月未請求データ一覧（未請求レセプト選択用）
Dim gRebillData As Object          ' ユーザー選択結果：返戻再請求に分類するデータ
Dim gLateData As Object            ' ユーザー選択結果：月遅れ請求に分類するデータ
Dim gSelectedUnclaimed As Object   ' ユーザー選択結果：前月未請求から追加するデータ

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

    ' 1. CSVフォルダをユーザーに選択させる
    csvFolder = SelectCSVFolder()
    If csvFolder = "" Then Exit Sub

    ' 1.1 フォルダが空なら処理を中止
    If IsFolderEmpty(csvFolder) Then
        MsgBox "選択したフォルダにはCSVファイルがありません。処理を中止します。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 2. テンプレートパス・保存フォルダ取得
    templatePath = GetTemplatePath()
    savePath = GetSavePath()
    If templatePath = "" Or savePath = "" Then Exit Sub

    ' 3. ファイルシステムオブジェクトの作成
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 4. フォルダ内のすべての「fixf」ファイルを取得
    Set fixfFiles = FindAllFixfFiles(fso, csvFolder)

    ' 5. fixfファイルがない場合は通常のCSV処理に切り替え
    If fixfFiles Is Nothing Or fixfFiles.Count = 0 Then
        ProcessWithoutFixf fso, csvFolder, savePath, templatePath
        Exit Sub
    End If

    ' 6. 複数のfixfファイルを順番に処理
    For Each file In fixfFiles
        fixfFile = file.Path

        ' 7. 対象年月を取得
        targetYear = "": targetMonth = ""
        GetYearMonthFromFixf fixfFile, targetYear, targetMonth

        ' 対象年月が取得できなかった場合はスキップ
        If targetYear = "" Or targetMonth = "" Then
            MsgBox "ファイル " & fixfFile & " から診療年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If

        ' 8. 対象Excelファイルが既に存在するか確認（存在する場合はスキップ）
        Dim csvYYMM As String, reportName As String
        csvYYMM = Format(CInt(targetYear) - 2018, "00") & targetMonth
        reportName = "保険請求管理報告書_R" & csvYYMM & ".xlsx"
        If fso.FileExists(savePath & "\" & reportName) Then
            MsgBox "診療年月 " & targetYear & "年" & targetMonth & "月 の報告書は既に処理済みです。", vbInformation, "処理済み"
            GoTo NextFile
        End If

        ' 対象Excelファイルを取得（存在しなければテンプレートから新規作成）
        targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)
        If targetFile = "" Then
            MsgBox "診療年月 " & targetYear & "年" & targetMonth & "月 のExcelファイルを作成できませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If

        ' 9. Excelを開く
        On Error Resume Next
        Set newBook = Workbooks.Open(targetFile)
        On Error GoTo 0
        If newBook Is Nothing Then
            MsgBox "ファイル " & targetFile & " を開けませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If

        ' 10. fixfファイルの内容をシート1に取り込む
        ImportCSVData fixfFile, newBook.Sheets(1), "請求確定状況"

        ' 11. テンプレート情報を設定（シート1への既定転記はスキップ）
        SetTemplateInfo newBook, targetYear, targetMonth, True

        ' 12. フォルダ内の他のCSVファイルを順に処理（fmei→henr→zogn）
        ProcessAllCSVFiles fso, newBook, csvFolder

        ' 13. 保存してブックを閉じる
        newBook.Save
        newBook.Close
NextFile:
    Next file

    ' 14. 処理完了メッセージ
    MsgBox "すべての fixf ファイルの処理が完了しました！", vbInformation, "処理完了"
End Sub

Sub ProcessWithoutFixf(fso As Object, csvFolder As String, savePath As String, templatePath As String)
    Dim targetYear As String, targetMonth As String
    Dim targetFile As String, newBook As Workbook

    ' 対象年月をCSV内容から取得
    targetYear = "": targetMonth = ""
    GetYearMonthFromCSV fso, csvFolder, targetYear, targetMonth
    If targetYear = "" Or targetMonth = "" Then
        MsgBox "CSVファイルから診療年月を取得できませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 対象Excelファイルが既に存在する場合はスキップ
    Dim csvYYMM As String, reportName As String, fsoLocal As Object
    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    csvYYMM = Format(CInt(targetYear) - 2018, "00") & targetMonth
    reportName = "保険請求管理報告書_R" & csvYYMM & ".xlsx"
    If fsoLocal.FileExists(savePath & "\" & reportName) Then
        MsgBox "診療年月 " & targetYear & "年" & targetMonth & "月 の報告書は既に処理済みです。", vbInformation, "処理済み"
        Exit Sub
    End If

    ' 対象Excelファイルを取得（存在しなければ新規作成）
    targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)
    If targetFile = "" Then
        MsgBox "診療年月 " & targetYear & "年" & targetMonth & "月 のExcelファイルを作成できませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' ブックを開く
    On Error Resume Next
    Set newBook = Workbooks.Open(targetFile)
    On Error GoTo 0
    If newBook Is Nothing Then
        MsgBox "ファイル " & targetFile & " を開けませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' テンプレート情報を設定（通常通り設定）
    SetTemplateInfo newBook, targetYear, targetMonth, False

    ' CSVファイルを順次処理（fixfなしでも他のCSVを処理可能）
    ProcessAllCSVFiles fso, newBook, csvFolder

    ' 保存してブックを閉じる
    newBook.Save
    newBook.Close

    MsgBox "CSVファイルの処理が完了しました。", vbInformation, "処理完了"
End Sub

Sub SetTemplateInfo(newBook As Workbook, targetYear As String, targetMonth As String, Optional skipSheet1Info As Boolean = False)
    Dim wsTemplate As Worksheet, wsTemplate2 As Worksheet
    Dim receiptYear As Integer, receiptMonth As Integer
    Dim sendMonth As Integer, sendDate As String

    ' 西暦年と調剤月の計算
    receiptYear = CInt(targetYear)
    receiptMonth = CInt(targetMonth)

    ' 請求月（調剤月の翌月）の計算
    sendMonth = receiptMonth + 1
    If sendMonth = 13 Then sendMonth = 1
    sendDate = sendMonth & "月10日請求分"

    ' シート1(A), シート2(B)を取得
    Set wsTemplate = newBook.Sheets(1)
    Set wsTemplate2 = newBook.Sheets(2)

    ' シート名変更（シート1を "R{令和YY}.{M}", シート2を丸数字の月に変更）
    wsTemplate.Name = "R" & (receiptYear - 2018) & "." & receiptMonth
    wsTemplate2.Name = ConvertToCircledNumber(receiptMonth)

    ' 情報転記
    If Not skipSheet1Info Then
        wsTemplate.Range("G2").Value = targetYear & "年" & targetMonth & "月調剤分"
        wsTemplate.Range("I2").Value = sendDate
        wsTemplate.Range("J2").Value = ThisWorkbook.Sheets(1).Range("B1").Value  ' 請求機関（施設名等）
    End If
    wsTemplate2.Range("H1").Value = targetYear & "年" & targetMonth & "月調剤分"
    wsTemplate2.Range("J1").Value = sendDate
    wsTemplate2.Range("L1").Value = ThisWorkbook.Sheets(1).Range("B1").Value     ' 請求機関（施設名等）
End Sub

Sub ProcessAllCSVFiles(fso As Object, newBook As Workbook, csvFolder As String)
    Dim csvFile As Object
    Dim fileType As String
    Dim wsDetails As Worksheet
    Dim wsCSV As Worksheet
    Dim sheetName As String
    Dim sheetIndex As Integer

    ' シート2（詳細データ用）を取得
    Set wsDetails = newBook.Sheets(2)

    ' 1. 振込額明細書（fmei）の処理
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" And InStr(csvFile.Name, "fmei") > 0 Then
            fileType = "振込額明細書"
            sheetName = fso.GetBaseName(csvFile.Name)
            sheetName = GetUniqueSheetName(newBook, sheetName)
            sheetIndex = Application.WorksheetFunction.Min(3, newBook.Sheets.Count + 1)
            Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(sheetIndex - 1))
            wsCSV.Name = sheetName
            ImportCSVData csvFile.Path, wsCSV, fileType
            Call TransferBillingDetails(newBook, sheetName, csvFile.Name)
        End If
    Next csvFile

    ' 2. 返戻内訳書（henr）の処理
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" And InStr(csvFile.Name, "henr") > 0 Then
            fileType = "返戻内訳書"
            sheetName = fso.GetBaseName(csvFile.Name)
            sheetName = GetUniqueSheetName(newBook, sheetName)
            sheetIndex = Application.WorksheetFunction.Min(3, newBook.Sheets.Count + 1)
            Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(sheetIndex - 1))
            wsCSV.Name = sheetName
            ImportCSVData csvFile.Path, wsCSV, fileType
            Call TransferBillingDetails(newBook, sheetName, csvFile.Name)
        End If
    Next csvFile

    ' 3. 増減点連絡書（zogn）の処理
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" And InStr(csvFile.Name, "zogn") > 0 Then
            fileType = "増減点連絡書"
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

    ' シート設定（請求データシートと詳細シート）
    Set wsBilling = newBook.Sheets(1)
    Set wsDetails = newBook.Sheets(2)

    ' 診療年月（YYMM形式）を取得
    Dim csvYYMM As String
    csvYYMM = Right(CStr(wsBilling.Cells(2, 2).Value), 4)

    ' CSVファイル名から請求先区分を判別
    payerCode = Mid(sheetName, 7, 1)
    Select Case payerCode
        Case "1": payerType = "社保"
        Case "2": payerType = "国保"
        Case Else: payerType = "労災"
    End Select

    ' 開始行位置の辞書を作成（シート2の各カテゴリ見出し行を取得）
    Set startRowDict = CreateObject("Scripting.Dictionary")
    If payerType = "社保" Then
        startRowDict.Add "返戻再請求", GetStartRow(wsDetails, "社保返戻再請求")
        startRowDict.Add "月遅れ請求", GetStartRow(wsDetails, "社保月遅れ請求")
        startRowDict.Add "返戻・査定", GetStartRow(wsDetails, "社保返戻・査定")
        startRowDict.Add "未請求扱い", GetStartRow(wsDetails, "社保未請求扱い")
    ElseIf payerType = "国保" Then
        startRowDict.Add "返戻再請求", GetStartRow(wsDetails, "国保返戻再請求")
        startRowDict.Add "月遅れ請求", GetStartRow(wsDetails, "国保月遅れ請求")
        startRowDict.Add "返戻・査定", GetStartRow(wsDetails, "国保返戻・査定")
        startRowDict.Add "未請求扱い", GetStartRow(wsDetails, "国保未請求扱い")
    End If

    ' 各カテゴリ用のディクショナリを作成
    Set rebillDict = CreateObject("Scripting.Dictionary")    ' 返戻再請求
    Set lateDict = CreateObject("Scripting.Dictionary")      ' 月遅れ請求
    Set unpaidDict = CreateObject("Scripting.Dictionary")    ' 未請求扱い
    Set assessmentDict = CreateObject("Scripting.Dictionary") ' 返戻・査定

    ' 請求データシートの最終行を取得
    lastRowBilling = wsBilling.Cells(Rows.Count, "D").End(xlUp).Row

    ' fixfファイルがない場合（請求データがシート1にない場合）、各CSVから詳細データを転記
    If lastRowBilling < 2 Then
        If InStr(csvFileName, "henr") > 0 Then
            Set wsCSV = newBook.Sheets(sheetName)
            lastRowDetails = wsCSV.Cells(Rows.Count, 1).End(xlUp).Row
            For j = 2 To lastRowDetails
                dispensingMonth = CStr(wsCSV.Cells(j, 1).Value)
                If dispensingMonth <> "" Then
                    If Len(dispensingMonth) = 4 Then dispensingMonth = "5" & dispensingMonth
                    convertedMonth = ConvertToWesternDate(dispensingMonth)
                    rowData = Array(wsCSV.Cells(j, 2).Value, convertedMonth, wsCSV.Cells(j, 5).Value, wsCSV.Cells(j, 14).Value)  ' 受付番号, 診療月, 請求点数, 事由コード
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
                    rowData = Array(wsCSV.Cells(j, 2).Value, convertedMonth, wsCSV.Cells(j, 6).Value, wsCSV.Cells(j, 7).Value)  ' 受付番号, 調剤月, 増減点数, 事由
                    unpaidDict.Add CStr(wsCSV.Cells(j, 2).Value) & "_" & j, rowData
                End If
            Next j
        End If
    End If

    ' 請求データ（fixf）をディクショナリに格納（fixfファイルがある場合のみ該当）
    Dim dispGYM As String
    For i = 2 To lastRowBilling
        dispensingMonth = wsBilling.Cells(i, 2).Value       ' GYYMM形式の診療月
        convertedMonth = ConvertToWesternDate(dispensingMonth)
        rowData = Array(wsBilling.Cells(i, 4).Value, convertedMonth, wsBilling.Cells(i, 5).Value, wsBilling.Cells(i, 10).Value)
        ' 対象診療月（csvYYMM）と異なる場合のみ各カテゴリに追加
        If Right(dispensingMonth, 4) <> csvYYMM Then
            If InStr(csvFileName, "fixf") > 0 Then
                ' fixfエントリの場合、ユーザーに返戻再請求か月遅れ請求か選択させる
                If ShowRebillSelectionForm(rowData) Then
                    rebillDict.Add wsBilling.Cells(i, 1).Value, rowData   ' 返戻再請求
                Else
                    lateDict.Add wsBilling.Cells(i, 1).Value, rowData    ' 月遅れ請求
                End If
            ElseIf InStr(csvFileName, "zogn") > 0 Then
                unpaidDict.Add wsBilling.Cells(i, 1).Value, rowData      ' 未請求扱い
            ElseIf InStr(csvFileName, "henr") > 0 Then
                assessmentDict.Add wsBilling.Cells(i, 1).Value, rowData  ' 返戻・査定
            End If
        End If
    Next i

    ' 各カテゴリの追加行数を計算（各カテゴリ4行を超える分）
    a = 0: b = 0: c = 0
    If rebillDict.Count > 4 Then a = rebillDict.Count - 4
    If lateDict.Count > 4 Then b = lateDict.Count - 4
    If assessmentDict.Count > 4 Then c = assessmentDict.Count - 4

    ' 各カテゴリの開始行を調整
    Dim lateStartRow As Long, assessmentStartRow As Long, unpaidStartRow As Long
    lateStartRow = startRowDict("月遅れ請求") + 1 + a
    assessmentStartRow = startRowDict("返戻・査定") + 1 + a + b
    unpaidStartRow = startRowDict("未請求扱い") + 1 + a + b + c

    ' 必要に応じて行を挿入して枠を確保
    If a + b + c > 0 Then
        wsDetails.Rows(lateStartRow & ":" & lateStartRow + a).Insert Shift:=xlDown
        wsDetails.Rows(assessmentStartRow & ":" & assessmentStartRow + b).Insert Shift:=xlDown
        wsDetails.Rows(unpaidStartRow & ":" & unpaidStartRow + c).Insert Shift:=xlDown
    End If

    ' 各ディクショナリのデータをシート2に転記（ディクショナリが空の場合はスキップ）
    If rebillDict.Count > 0 Then
        j = startRowDict("返戻再請求")
        TransferData rebillDict, wsDetails, j, payerType
    End If
    If lateDict.Count > 0 Then
        j = startRowDict("月遅れ請求")
        TransferData lateDict, wsDetails, j, payerType
    End If
    If unpaidDict.Count > 0 Then
        j = startRowDict("未請求扱い")
        TransferData unpaidDict, wsDetails, j, payerType
    End If
    If assessmentDict.Count > 0 Then
        j = startRowDict("返戻・査定")
        TransferData assessmentDict, wsDetails, j, payerType
    End If

    MsgBox payerType & " のデータ転記が完了しました！", vbInformation, "処理完了"
End Sub

Function SelectCSVFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "CSVフォルダを選択してください"
        If .Show = -1 Then
            SelectCSVFolder = .SelectedItems(1) & "\"
        Else
            MsgBox "フォルダが選択されませんでした。処理を中止します。", vbExclamation, "確認"
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
        IsFolderEmpty = True   ' **フォルダにファイルがない場合 True**
    Else
        IsFolderEmpty = False
    End If
End Function

Function GetTemplatePath() As String
    ' テンプレートファイルのパスをシート1のセルB2から取得
    GetTemplatePath = ThisWorkbook.Sheets(1).Range("B2").Value & "\保険請求管理報告書テンプレート.xltm"
End Function

Function GetSavePath() As String
    ' 保存先フォルダのパスをシート1のセルB3から取得
    GetSavePath = ThisWorkbook.Sheets(1).Range("B3").Value
End Function

Function FindAllFixfFiles(fso As Object, csvFolder As String) As Collection
    Dim csvFile As Object
    Dim fixfFiles As New Collection
    ' **フォルダ内のすべてのファイルをチェック**
    For Each csvFile In fso.GetFolder(csvFolder).Files
        ' **拡張子が "csv" であり、名前に "fixf" を含む場合**
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" And InStr(LCase(csvFile.Name), "fixf") > 0 Then
            fixfFiles.Add csvFile  ' **fixfファイルをリストに追加**
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
    ' **fixfファイル名から年月コードを推定**
    code = ""
    '  他のCSVファイル名から GYYMM を取得（例: 振込額明細書など） 
    Dim csvFile As Object, folderPath As String
    folderPath = fso.GetFile(fixfFile).ParentFolder.Path
    For Each csvFile In fso.GetFolder(folderPath).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" Then
            If InStr(LCase(csvFile.Name), "fmei") > 0 Or InStr(LCase(csvFile.Name), "zogn") > 0 Or InStr(LCase(csvFile.Name), "henr") > 0 Then
                ' 名前末尾4桁が数字ならそれを年月コードとする
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
        targetYear = CStr(2018 + CInt(yrCode))    ' **和暦年コードを西暦年に変換**
        targetMonth = CStr(CInt(monCode))         ' **月コード（先頭0含む）を整数化**
    Else
        ' **fallback: fixfファイルから診療年月を抽出**（簡易）
        Dim ts As Object, lineText As String
        On Error Resume Next
        Set ts = fso.OpenTextFile(fixfFile, 1, False, -2)
        On Error GoTo 0
        If Not ts Is Nothing Then
            ' 先頭数行を読み診療年月を含む行を探す
            Dim i As Integer
            For i = 1 To 5
                If ts.AtEndOfStream Then Exit For
                lineText = ts.ReadLine
                If InStr(lineText, "G") > 0 And InStr(lineText, ",") = 0 Then
                    ' **例: "5XXXX" 形式の文字列を含む場合**
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
        ' **取得失敗時、ユーザーに入力を促す**
        If targetYear = "" Or targetMonth = "" Then
            MsgBox "診療年月を自動取得できませんでした。指定してください。", vbExclamation, "確認"
            targetYear = InputBox("西暦年を入力してください（例: 2023）:", "診療年")
            targetMonth = InputBox("月を入力してください（1~12）:", "診療月")
            If targetYear = "" Or targetMonth = "" Then
                ' ユーザー未入力の場合
                Exit Sub
            End If
        End If
    End If
End Sub

Function FindOrCreateReport(savePath As String, targetYear As String, targetMonth As String, templatePath As String) As String
    Dim fso As Object, existingFile As Object
    Dim fileName As String, filePath As String
    Dim csvYYMM As String
    csvYYMM = Format(CInt(targetYear) - 2018, "00") & Format(CInt(targetMonth), "00")  ' **和暦年+月コード**
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' **保存フォルダ内に既存の RYYMM ファイルがあるか確認**
    For Each existingFile In fso.GetFolder(savePath).Files
        If LCase(fso.GetExtensionName(existingFile.Name)) = "xlsm" Or LCase(fso.GetExtensionName(existingFile.Name)) = "xlsx" Then
            If InStr(existingFile.Name, "保険請求管理報告書_R" & csvYYMM) > 0 Then
                FindOrCreateReport = existingFile.Path  ' **既存ファイルのパスを返す**
                Exit Function
            End If
        End If
    Next existingFile
    ' **該当するファイルがなければ、新規作成**
    fileName = "保険請求管理報告書_R" & csvYYMM & ".xlsm"   ' **xlsm形式で保存**（改良点1）
    filePath = savePath & "\" & fileName
    ' **テンプレートを元に新規ブック作成**
    On Error Resume Next
    Dim tmplWb As Workbook
    Set tmplWb = Workbooks.Open(templatePath)   ' **テンプレートブックを開く**
    On Error GoTo 0
    If tmplWb Is Nothing Then
        MsgBox "テンプレートを開けませんでした: " & templatePath, vbCritical, "エラー"
        FindOrCreateReport = ""
        Exit Function
    End If
    On Error Resume Next
    tmplWb.SaveAs filePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled  ' **xlsm形式で保存**
    If Err.Number <> 0 Then
        MsgBox "ファイルを保存できませんでした: " & filePath, vbCritical, "エラー"
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

    ' 項目マッピングを取得
    Set colMap = GetColumnMapping(fileType)
    ' シートをクリアして項目名を1行目に設定
    ws.Cells.Clear
    j = 1
    For Each key In colMap.Keys
        ws.Cells(1, j).Value = colMap(key)
        j = j + 1
    Next key

    ' CSVファイルをUTF-8テキストとして開く
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(csvFile, 1, False, -2)  ' -2: UTF-8

    ' データ部分を転記
    i = 2
    isHeader = True
    Do While Not ts.AtEndOfStream
        lineText = ts.ReadLine
        dataArray = Split(lineText, ",")
        If isHeader Then
            ' 一行目（ヘッダー行）はスキップ
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
    MsgBox "CSV読込中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
    If Not ts Is Nothing Then ts.Close
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Function GetColumnMapping(fileType As String) As Object
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    Select Case fileType
        Case "振込額明細書"
            colMap.Add 2, "振込年月"
            colMap.Add 3, "振込金額"
            colMap.Add 4, "振込日"
            ' （必要な列を追加）
        Case "増減点連絡書"
            colMap.Add 2, "調剤年月"
            colMap.Add 4, "受付番号"
            colMap.Add 11, "区分"
            colMap.Add 14, "老人減免区分"
            colMap.Add 15, "氏名"
            colMap.Add 21, "増減点数（金額）"
            colMap.Add 22, "事由"
        Case "返戻内訳書"
            colMap.Add 2, "調剤年月(YYMM形式)"
            colMap.Add 3, "受付番号"
            colMap.Add 4, "保険者番号"
            colMap.Add 7, "氏名"
            colMap.Add 9, "請求点数"
            colMap.Add 10, "薬剤一部負担金"
            colMap.Add 12, "一部負担金額"
            colMap.Add 13, "患者負担金額（公費）"
            colMap.Add 14, "事由コード"
        Case Else
            ' その他（必要に応じて追加）
            colMap.Add 1, "項目1"
    End Select
    Set GetColumnMapping = colMap
End Function

Function GetStartRow(ws As Worksheet, category As String) As Long
    ' 詳細シートから指定カテゴリの行番号を取得
    Dim rng As Range
    Set rng = ws.Cells.Find(what:=category, LookAt:=xlWhole)
    If rng Is Nothing Then
        MsgBox "詳細シート上でカテゴリ """ & category & """ を見つけられませんでした。", vbExclamation, "エラー"
        GetStartRow = 0
    Else
        GetStartRow = rng.Row
    End If
End Function

Sub TransferData(dataDict As Object, ws As Worksheet, startRow As Long, payerType As String)
    Dim key As Variant, rowData As Variant
    Dim j As Long, payerColumn As Long
    ' **Dictionaryが空なら処理しない**
    If dataDict.Count = 0 Then Exit Sub
    ' **payerTypeに応じた転記列を決定**
    If payerType = "社保" Then
        payerColumn = 8   ' 社保はH列に請求先マーク
    ElseIf payerType = "国保" Then
        payerColumn = 9   ' 国保はI列に請求先マーク
    Else
        payerColumn = 8   ' （労災等は社保列に仮設定）
    End If
    j = startRow
    For Each key In dataDict.Keys
        rowData = dataDict(key)
        ws.Cells(j, 4).Value = rowData(0)    ' 患者氏名
        ws.Cells(j, 5).Value = rowData(1)    ' 調剤年月（西暦表記）
        ws.Cells(j, 6).Value = rowData(2)    ' 医療機関名
        ws.Cells(j, payerColumn).Value = payerType   ' 請求先区分（社保/国保）
        ws.Cells(j, payerColumn).Font.Bold = True    ' **強調表示**
        ws.Cells(j, 10).Value = rowData(3)   ' 請求点数
        j = j + 1
    Next key
End Sub

Sub ShowRebillSelectionForm()
    ' 過去月レセプトの一覧をユーザーに表示し、返戻再請求するものを選択してもらう
    Dim uf As Object, listData As Object
    Set listData = gOlderList
    If listData Is Nothing Or listData.Count = 0 Then Exit Sub
    ' ユーザーフォーム作成と表示
    Set uf = CreateRebillSelectionForm(listData)
    Set gRebillForm = uf  ' グローバル参照保存
    gRebillForm.Show vbModal
    ' フォーム閉じた後、選択結果は gRebillData と gLateData に格納済み（ProcessRebillSelectionで設定）
End Sub

Function CreateRebillSelectionForm(listData As Object) As Object
    Dim uf As Object, listBox As Object, btnOK As Object
    Dim i As Long, rowData As Variant
    ' **UserForm を動的に作成**
    Set uf = VBA.UserForms.Add()  ' 新規UserForm
    uf.Caption = "返戻再請求の選択"
    uf.Width = 400
    uf.Height = 500
    ' **ListBoxを追加**
    Set listBox = uf.Controls.Add("Forms.ListBox.1", "listBox", True)
    listBox.Left = 20
    listBox.Top = 20
    listBox.Width = 350
    listBox.Height = 350
    listBox.MultiSelect = 1  ' **複数選択可能**
    ' **リストにデータを追加（調剤年月 | 患者氏名 | 医療機関名 | 点数）**
    For i = 0 To listData.Count - 1
        rowData = listData.Items()(i)
        listBox.AddItem rowData(1) & " | " & rowData(0) & " | " & rowData(2) & " | " & rowData(3)
    Next i
    ' **OKボタンを追加**
    Set btnOK = uf.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    btnOK.Caption = "確定"
    btnOK.Left = 150
    btnOK.Top = 400
    btnOK.Width = 100
    btnOK.Height = 30
    ' **ボタンクリック時の処理を設定**
    btnOK.OnClick = "ProcessRebillSelection"
    Set CreateRebillSelectionForm = uf
End Function

Sub ProcessRebillSelection()
    ' 返戻再請求選択フォームのOKボタン処理（選択された項目を分類）
    Dim uf As Object, listBox As Object
    Dim i As Long
    ' 動的フォームおよびListBoxを取得
    Set uf = gRebillForm
    Set listBox = uf.Controls("listBox")
    ' 結果用Dictionaryを初期化
    Set gRebillData = CreateObject("Scripting.Dictionary")
    Set gLateData = CreateObject("Scripting.Dictionary")
    ' **選択状態に応じて振り分け**
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) Then
            ' 選択されたもの -> 返戻再請求
            gRebillData.Add gOlderList.Keys()(i), gOlderList.Items()(i)
        Else
            ' 選択されなかったもの -> 月遅れ請求
            gLateData.Add gOlderList.Keys()(i), gOlderList.Items()(i)
        End If
    Next i
    ' フォームをアンロードして閉じる
    Unload uf
    Set gRebillForm = Nothing
End Sub

Function AddUnclaimedRecords(payerType As String, targetYear As String, targetMonth As String) As Object
    Dim prevYear As String, prevMonth As String
    Dim prevFileName As String, prevFilePath As String
    Dim prevBook As Workbook, wsPrevDetails As Worksheet
    Dim startRow As Long, endRow As Long, row As Long
    ' 前月を算出
    If CInt(targetMonth) = 1 Then
        prevYear = CStr(CInt(targetYear) - 1)
        prevMonth = "12"
    Else
        prevYear = targetYear
        prevMonth = CStr(CInt(targetMonth) - 1)
    End If
    ' 前月の報告書ファイル名
    Dim prevYYMM As String
    prevYYMM = Format(CInt(prevYear) - 2018, "00") & Format(CInt(prevMonth), "00")
    prevFileName = "保険請求管理報告書_R" & prevYYMM & ".xlsm"
    prevFilePath = GetSavePath() & "\" & prevFileName
    If Dir(prevFilePath) = "" Then
        ' ファイルが存在しない場合
        Set AddUnclaimedRecords = Nothing
        Exit Function
    End If
    ' 前月ファイルを開く（読み取り専用）
    On Error Resume Next
    Set prevBook = Workbooks.Open(prevFilePath, ReadOnly:=True)
    On Error GoTo 0
    If prevBook Is Nothing Then
        Set AddUnclaimedRecords = Nothing
        Exit Function
    End If
    ' 前月詳細シートを取得（シート名は社保/国保共通で"B"シートと仮定）
    Set wsPrevDetails = prevBook.Sheets(2)
    ' 対象カテゴリの開始行を取得
    Dim categoryLabel As String
    If payerType = "社保" Then
        categoryLabel = "社保未請求扱い"
    Else
        categoryLabel = "国保未請求扱い"
    End If
    startRow = GetStartRow(wsPrevDetails, categoryLabel)
    If startRow = 0 Then
        ' ラベルが見つからない場合は終了
        prevBook.Close SaveChanges:=False
        Set AddUnclaimedRecords = Nothing
        Exit Function
    End If
    ' 開始行から下方向にデータを収集
    Set gUnclaimedList = CreateObject("Scripting.Dictionary")
    endRow = startRow + 3  ' 基本枠は4行
    ' データが追加されている場合、空白行が出るまで延長
    Do While wsPrevDetails.Cells(endRow, 4).Value <> "" Or wsPrevDetails.Cells(endRow + 1, 4).Value <> ""
        endRow = endRow + 1
        If endRow > wsPrevDetails.Rows.Count Then Exit Do
    Loop
    For row = startRow + 1 To endRow
        If wsPrevDetails.Cells(row, 4).Value <> "" Then   ' 患者氏名列が空でなければデータあり
            Dim prevRowData As Variant
            prevRowData = Array(wsPrevDetails.Cells(row, 4).Value, wsPrevDetails.Cells(row, 5).Value, wsPrevDetails.Cells(row, 6).Value, wsPrevDetails.Cells(row, 10).Value)
            gUnclaimedList.Add row, prevRowData
        End If
    Next row
    ' 前月ブックを閉じる
    prevBook.Close SaveChanges:=False
    ' ユーザーに前月未請求を表示し、追加するものを選択させる
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
    ' フォームが閉じられた後、gSelectedUnclaimedに結果が格納される
End Sub

Function CreateUnclaimedSelectionForm(listData As Object) As Object
    Dim uf As Object, listBox As Object, btnOK As Object
    Dim i As Long, rowData As Variant
    Set uf = VBA.UserForms.Add()
    uf.Caption = "前月 未請求レセプトの追加選択"
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
    btnOK.Caption = "追加"
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
    ' GYYMM形式（和暦）を西暦年下2桁.月形式に変換
    Dim eraCode As String, yearPart As Integer, westernYear As Integer, monthPart As String
    eraCode = Left(dispensingMonth, 1)
    yearPart = CInt(Mid(dispensingMonth, 2, 2))
    monthPart = Right(dispensingMonth, 2)
    Select Case eraCode
        Case "5": westernYear = 2018 + yearPart   ' 令和 (2019年=令和1年)
        Case "4": westernYear = 1988 + yearPart   ' 平成 (1989年=平成1年) - ※過去データ対応
        Case Else: westernYear = 2018 + yearPart  ' （デフォルト:令和として計算）
    End Select
    ConvertToWesternDate = Right(CStr(westernYear), 2) & "." & monthPart
End Function

' **半年ごとの売掛データ比較・誤差分析機能**（改良点6）
Sub CompareHalfYearData()
    Dim inputYear As String, half As String
    inputYear = InputBox("分析する年を入力してください（西暦）:", "半年売掛比較")
    If inputYear = "" Then Exit Sub
    half = InputBox("上期=1 または 下期=2 を入力してください:", "半年区分")
    If half = "" Then Exit Sub
    If half <> "1" And half <> "2" Then
        MsgBox "半期区分は1または2で入力してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    Dim startMonth As Integer, endMonth As Integer
    If half = "1" Then
        startMonth = 1: endMonth = 6
    Else
        startMonth = 7: endMonth = 12
    End If
    Dim analysisWb As Workbook
    Set analysisWb = ThisWorkbook  ' 結果出力先をマクロブックに設定
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
    outSheet.Range("A1:E1").Value = Array("月", "日次計上点数", "請求確定点数", "振込額(円)", "点数差異")
    Dim m As Integer, rowIndex As Integer
    rowIndex = 2
    For m = startMonth To endMonth
        Dim yy As String, mm As String, fileCode As String
        yy = Format(CInt(inputYear) - 2018, "00")
        mm = Format(m, "00")
        fileCode = "R" & yy & mm
        Dim reportName As String
        reportName = "保険請求管理報告書_" & fileCode & ".xlsm"
        Dim reportPath As String
        reportPath = GetSavePath() & "\" & reportName
        If Dir(reportPath) <> "" Then
            Dim repWb As Workbook
            Set repWb = Workbooks.Open(reportPath, ReadOnly:=True)
            Dim wsA As Worksheet, wsCSV As Worksheet, wsCSV2 As Worksheet
            Set wsA = repWb.Sheets(1)   ' 日次データシート
            ' 日次データから総点数を取得（例としてシートAのセルJ○などに集計があると仮定）
            Dim dailyTotal As Long
            dailyTotal = 0
            On Error Resume Next
            dailyTotal = CLng(wsA.Range("J50").Value) ' ※適切なセル参照に要修正
            On Error GoTo 0
            ' 請求確定点数（fixfデータの総点数）取得（シートAに計またはシートBに総合計点？仮にJ50とする）
            Dim billedTotal As Long
            billedTotal = 0
            On Error Resume Next
            billedTotal = CLng(wsA.Range("J50").Value)
            On Error GoTo 0
            ' 振込額明細から合計金額取得（CSVシート名にfmei含む想定）
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
            ' 点数差異計算
            Dim pointDiff As Long
            pointDiff = dailyTotal - billedTotal
            ' 結果を出力
            outSheet.Cells(rowIndex, 1).Value = inputYear & "年" & m & "月"
            outSheet.Cells(rowIndex, 2).Value = dailyTotal
            outSheet.Cells(rowIndex, 3).Value = billedTotal
            outSheet.Cells(rowIndex, 4).Value = payAmount
            outSheet.Cells(rowIndex, 5).Value = pointDiff
            rowIndex = rowIndex + 1
        Else
            ' ファイルがない場合は空行または0出力
            outSheet.Cells(rowIndex, 1).Value = inputYear & "年" & m & "月"
            outSheet.Cells(rowIndex, 2).Value = "N/A"
            outSheet.Cells(rowIndex, 3).Value = "N/A"
            outSheet.Cells(rowIndex, 4).Value = "N/A"
            outSheet.Cells(rowIndex, 5).Value = "N/A"
            rowIndex = rowIndex + 1
        End If
    Next m
    MsgBox inputYear & "年 " & IIf(half = "1", "上期", "下期") & " の売掛データ比較が完了しました。" & vbCrLf & _
            "シート[" & outSheet.Name & "]に結果を出力しました。", vbInformation, "分析完了"
End Sub