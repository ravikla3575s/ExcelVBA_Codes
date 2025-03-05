Option Explicit

' グローバル変数（ユーザーフォーム管理用）
Public gRebillForm As Object          ' 動的に作成した返戻再請求選択フォーム
Public gUnclaimedForm As Object       ' 動的に作成した未請求レセプト選択フォーム
Public gOlderList As Object           ' 過去レセプトデータ一覧（返戻再請求/月遅れ選択用）
Public gUnclaimedList As Object       ' 前月未請求データ一覧（未請求レセプト選択用）
Public gRebillData As Object          ' ユーザー選択結果：返戻再請求に分類するデータ
Public gLateData As Object            ' ユーザー選択結果：月遅れ請求に分類するデータ
Public gSelectedUnclaimed As Object   ' ユーザー選択結果：前月未請求から追加するデータ

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

    ' 6. 各fixfファイルに対して処理を実行
    For Each fixfFile In fixfFiles
        ' 6.1 対象年月を取得（ファイル名や内容から推定）
        targetYear = "": targetMonth = ""
        GetYearMonthFromFixf fixfFile, targetYear, targetMonth
        If targetYear = "" Or targetMonth = "" Then
            MsgBox "fixfファイルから診療年月を取得できませんでした。", vbExclamation, "エラー"
            Exit Sub
        End If

        ' 6.2 対象Excelファイルを準備（既存報告書があれば開く、なければテンプレートから新規作成）
        targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)
        If targetFile = "" Then Exit Sub
        Set newBook = Workbooks.Open(targetFile)

        ' 6.3 帳票ブックに基本情報を設定
        SetTemplateInfo newBook, targetYear, targetMonth

        ' 6.4 CSVファイルを順次読み込み
        ProcessAllCSVFiles fso, newBook, csvFolder

        ' 6.5 帳票ブックを保存して閉じる
        newBook.Save
        newBook.Close False
        MsgBox "[" & fixfFile.Name & "] のデータ転記完了", vbInformation, "完了"
    Next fixfFile

    MsgBox "全てのfixfファイル処理完了！", vbInformation, "完了"
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
        MsgBox "既に対象年月の報告書ファイルが存在します: " & reportName, vbInformation, "処理スキップ"
        Exit Sub
    End If

    ' 新規報告書ブックを作成
    targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)
    If targetFile = "" Then Exit Sub
    Set newBook = Workbooks.Open(targetFile)
    SetTemplateInfo newBook, targetYear, targetMonth
    ProcessAllCSVFiles fso, newBook, csvFolder
    newBook.Save
    newBook.Close False
    MsgBox "CSVデータの転記が完了しました。", vbInformation, "完了"
End Sub

Sub SetTemplateInfo(newBook As Workbook, targetYear As String, targetMonth As String, Optional skipSheet1Info As Boolean = False)
    Dim wsTemplate As Worksheet, wsTemplate2 As Worksheet
    Dim receiptYear As Integer, receiptMonth As Integer
    Dim sendMonth As Integer, sendDate As String

    ' シート取得（テンプレートはシート1=帳票A, シート2=帳票B想定）
    Set wsTemplate = newBook.Sheets(1)
    Set wsTemplate2 = newBook.Sheets(2)
    ' シート名変更（シート1を "R{令和YY}.{M}", シート2を丸数字の月に変更）
    receiptYear = CInt(targetYear)
    receiptMonth = CInt(targetMonth)
    wsTemplate.Name = "R" & Format(receiptYear - 2018, "0") & "." & Format(receiptMonth, "0")
    wsTemplate2.Name = ConvertToCircledNumber(receiptMonth)

    ' 情報転記（シート1のヘッダー情報設定）
    If Not skipSheet1Info Then
        wsTemplate.Range("G2").Value = targetYear & "年" & Format(receiptMonth, "00") & "月調剤分"
        sendMonth = receiptMonth + 1
        sendDate = Format(sendMonth, "00") & "月10日"
        wsTemplate.Range("I2").Value = "提出日: " & targetYear & "年" & sendDate
    End If
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

    ' 1. 各CSVの処理
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If InStr(csvFile.Name, "fmei") > 0 Then
            ' 振込額明細書CSV
            Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(newBook.Sheets.Count))
            wsCSV.Name = "振込額明細書"
            ImportCSVData csvFile.Path, wsCSV, "振込額明細書"
        ElseIf InStr(csvFile.Name, "zogn") > 0 Then
            ' 増減点連絡書CSV
            Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(newBook.Sheets.Count))
            wsCSV.Name = "増減点連絡書"
            ImportCSVData csvFile.Path, wsCSV, "増減点連絡書"
        ElseIf InStr(csvFile.Name, "henr") > 0 Then
            ' 返戻内訳書CSV
            Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(newBook.Sheets.Count))
            wsCSV.Name = "返戻内訳書"
            ImportCSVData csvFile.Path, wsCSV, "返戻内訳書"
        ElseIf InStr(csvFile.Name, "fixf") > 0 Then
            ' 請求確定CSV（fixf）
            ImportCSVData csvFile.Path, newBook.Sheets(1), "fixf"
            sheetName = csvFile.Name
        End If
    Next csvFile

    ' 2. 返戻再請求・月遅れ請求データの詳細転記
    If sheetName <> "" Then
        Call TransferBillingDetails(newBook, sheetName, csvFile.Name)
    End If
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
    If wsBilling.Range("G2").Value <> "" Then
        csvYYMM = Right(CStr(wsBilling.Range("G2").Value), 4)
        ' G2には "YYYY年MM月調剤分" 形式で入っているため、数値部分を抽出
        If InStr(csvYYMM, "年") > 0 Or InStr(csvYYMM, "月") > 0 Then
            csvYYMM = Replace(Replace(csvYYMM, "年", ""), "月", "")
        End If
    Else
        csvYYMM = ""
    End If

    ' CSVファイル名から請求先区分を判別
    payerCode = Mid(sheetName, 7, 1)
    Select Case payerCode
        Case "1": payerType = "社保"
        Case "2": payerType = "国保"
        Case Else: payerType = "労災"
    End Select

    ' 開始行位置の辞書を作成（シート2の各カテゴリ見出し行を取得）
    Set startRowDict = CreateObject("Scripting.Dictionary")
    startRowDict.Add "返戻再請求", GetStartRow(wsDetails, payerType & "返戻再請求")
    startRowDict.Add "月遅れ請求", GetStartRow(wsDetails, payerType & "月遅れ請求")
    startRowDict.Add "未請求扱い", GetStartRow(wsDetails, payerType & "未請求扱い")
    startRowDict.Add "返戻・査定", GetStartRow(wsDetails, payerType & "返戻・査定")

    ' Mainシート上の全行を走査し、過去月データを収集
    lastRowBilling = wsBilling.Cells(wsBilling.Rows.Count, 1).End(xlUp).Row
    Set gOlderList = Nothing
    For i = 2 To lastRowBilling
        dispensingMonth = Trim(CStr(wsBilling.Cells(i, 3).Value))
        If dispensingMonth = "" Then GoTo ContinueLoop
        If IsNumeric(Left(dispensingMonth, 1)) Then
            ' Eraコードがない場合は令和(5)を補完
            dispensingMonth = "5" & dispensingMonth
        End If
        convertedMonth = ConvertToWesternDate(dispensingMonth)
        ' 判定：過去月かどうか
        Dim itemYYMM As Integer
        itemYYMM = 0
        If convertedMonth <> "" Then
            itemYYMM = CInt(Replace(convertedMonth, ".", ""))
        End If
        If itemYYMM <> 0 And csvYYMM <> "" Then
            If itemYYMM < CInt(csvYYMM) Then
                ' 過去月データをリスト収集
                rowData = Array(wsBilling.Cells(i, 2).Value, convertedMonth, wsBilling.Cells(i, 5).Value, wsBilling.Cells(i, 14).Value)
                If gOlderList Is Nothing Then Set gOlderList = CreateObject("Scripting.Dictionary")
                Dim keyStr As String
                keyStr = CStr(wsBilling.Cells(i, 2).Value) & "_" & CStr(i)
                If Not gOlderList.Exists(keyStr) Then gOlderList.Add keyStr, rowData
            End If
        End If
ContinueLoop:
    Next i

    ' ユーザーフォーム表示による返戻再請求選択
    If Not gOlderList Is Nothing And gOlderList.Count > 0 Then
        ShowRebillSelectionForm
    End If

    ' 選択結果（または空）のディクショナリを準備
    Set rebillDict = CreateObject("Scripting.Dictionary")
    Set lateDict = CreateObject("Scripting.Dictionary")
    If Not gRebillData Is Nothing Then Set rebillDict = gRebillData
    If Not gLateData Is Nothing Then Set lateDict = gLateData
    ' フォーム未使用や全未選択の場合、全件を月遅れ扱いにする
    If Not gOlderList Is Nothing And gOlderList.Count > 0 And rebillDict.Count = 0 Then
        Dim k As Variant
        For Each k In gOlderList.Keys
            If Not lateDict.Exists(k) Then lateDict.Add k, gOlderList(k)
        Next k
    End If

    ' 返戻内訳書シート（査定データ）の収集
    Set assessmentDict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set wsCSV = newBook.Sheets("返戻内訳書")
    On Error GoTo 0
    If Not wsCSV Is Nothing Then
        lastRowDetails = wsCSV.Cells(wsCSV.Rows.Count, 2).End(xlUp).Row
        For j = 2 To lastRowDetails
            If Trim(wsCSV.Cells(j, 2).Value) <> "" Then
                dispensingMonth = Trim(CStr(wsCSV.Cells(j, 2).Value))
                If IsNumeric(Left(dispensingMonth, 1)) Then
                    dispensingMonth = "5" & dispensingMonth  ' Eraコード補完
                End If
                convertedMonth = ConvertToWesternDate(dispensingMonth)
                rowData = Array(wsCSV.Cells(j, 3).Value, convertedMonth, wsCSV.Cells(j, 9).Value, wsCSV.Cells(j, 14).Value)
                assessmentDict.Add CStr(wsCSV.Cells(j, 3).Value) & "_" & j, rowData
            End If
        Next j
    End If

    ' 前月未請求データの追加選択
    Dim addDict As Object
    Set addDict = AddUnclaimedRecords(payerType, Mid(targetYear, Len(targetYear) - 1), Format(CInt(targetMonth), "0"))
    If Not addDict Is Nothing Then
        For Each k In addDict.Keys
            If Not lateDict.Exists(k) Then lateDict.Add k, addDict(k)
        Next k
    End If

    ' 各カテゴリの追加行数を計算（各カテゴリ4行を超える分）
    a = 0: b = 0: c = 0
    If rebillDict.Count > 4 Then a = rebillDict.Count - 4
    If lateDict.Count > 4 Then b = lateDict.Count - 4
    If assessmentDict.Count > 4 Then c = assessmentDict.Count - 4

    ' シートに行挿入（各カテゴリ枠の拡張）
    If a > 0 Then wsDetails.Rows(startRowDict("返戻再請求") + 4).Resize(a).Insert Shift:=xlDown
    If b > 0 Then wsDetails.Rows(startRowDict("月遅れ請求") + 4).Resize(b).Insert Shift:=xlDown
    If c > 0 Then wsDetails.Rows(startRowDict("返戻・査定") + 4).Resize(c).Insert Shift:=xlDown

    ' データ転記（各カテゴリセクションに書き込み）
    TransferData rebillDict, wsDetails, startRowDict("返戻再請求") + 1, payerType
    TransferData lateDict, wsDetails, startRowDict("月遅れ請求") + 1, payerType
    TransferData assessmentDict, wsDetails, startRowDict("返戻・査定") + 1, payerType
    ' 未請求扱いカテゴリは必要時のみ（基本未使用）
    If Not unpaidDict Is Nothing And unpaidDict.Count > 0 Then
        TransferData unpaidDict, wsDetails, startRowDict("未請求扱い") + 1, payerType
    End If
End Sub

Function SelectCSVFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "CSVフォルダを選択してください"
        .AllowMultiSelect = False
        If .Show = -1 Then
            SelectCSVFolder = .SelectedItems(1)
        Else
            SelectCSVFolder = ""
        End If
    End With
End Function

Function IsFolderEmpty(folderPath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.GetFolder(folderPath).Files.Count = 0 And fso.GetFolder(folderPath).SubFolders.Count = 0 Then
        IsFolderEmpty = True
    Else
        IsFolderEmpty = False
    End If
End Function

Function GetTemplatePath() As String
    ' テンプレートパス取得（ここでは固定値 or 別シートから取得する想定）
    On Error Resume Next
    GetTemplatePath = ThisWorkbook.Sheets("Settings").Range("B1").Value
    On Error GoTo 0
End Function

Function GetSavePath() As String
    ' 保存先フォルダパス取得（ここでは固定値 or 別シートから取得する想定）
    On Error Resume Next
    GetSavePath = ThisWorkbook.Sheets("Settings").Range("B2").Value
    On Error GoTo 0
End Function

Function FindAllFixfFiles(fso As Object, csvFolder As String) As Collection
    Dim fixfFiles As New Collection
    Dim file As Object
    For Each file In fso.GetFolder(csvFolder).Files
        If InStr(file.Name, "fixf") > 0 Then
            fixfFiles.Add file
        End If
    Next file
    If fixfFiles.Count = 0 Then
        Set FindAllFixfFiles = Nothing
    Else
        Set FindAllFixfFiles = fixfFiles
    End If
End Function

Sub GetYearMonthFromFixf(fixfFile As String, ByRef targetYear As String, ByRef targetMonth As String)
    Dim fso As Object, fileName As String, nameNoExt As String
    Dim code As String, yrCode As String, monCode As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = fso.GetFileName(fixfFile)
    nameNoExt = fso.GetBaseName(fileName)
    code = ""
    ' ファイル名から GYYMM コード抽出（末尾4桁が GYYMM想定）
    If Len(nameNoExt) >= 4 Then
        If IsNumeric(Right(nameNoExt, 4)) Then
            code = Right(nameNoExt, 4)
        End If
    End If
    If code <> "" Then
        yrCode = Left(code, 2)
        monCode = Right(code, 2)
        targetYear = CStr(2018 + CInt(yrCode))    ' 和暦年コードを西暦年に変換
        targetMonth = CStr(CInt(monCode))         ' 月コードを整数化
    Else
        ' ファイル名に無い場合、ファイル内容先頭から診療年月推定
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
                    ' 例: "5XXXX" 形式の文字列を含む場合
                    Dim matchStr As Variant
                    matchStr = Replace(lineText, """", "")
                    If Len(matchStr) >= 4 And IsNumeric(matchStr) Then
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
        ' 取得失敗時、ユーザーに入力を促す
        If targetYear = "" Or targetMonth = "" Then
            MsgBox "診療年月を自動取得できませんでした。指定してください。", vbExclamation, "確認"
            targetYear = InputBox("西暦年を入力してください（例: 2023）:", "診療年")
            targetMonth = InputBox("月を入力してください（1~12）:", "診療月")
            If targetYear = "" Or targetMonth = "" Then
                Exit Sub
            End If
        End If
    End If
End Sub

Function GetYearMonthFromCSV(fso As Object, csvFolder As String, ByRef targetYear As String, ByRef targetMonth As String)
    ' フォルダ内の任意のCSVから診療年月を推定
    Dim file As Object, fileName As String
    targetYear = "": targetMonth = ""
    For Each file In fso.GetFolder(csvFolder).Files
        fileName = LCase(file.Name)
        If InStr(fileName, "fmei") > 0 Or InStr(fileName, "zogn") > 0 Or InStr(fileName, "henr") > 0 Then
            ' ファイル名から推定（例: "RYYMM" 部分を含む場合）
            If InStr(fileName, "_r") > 0 Then
                Dim pos As Long
                pos = InStr(fileName, "_r") + 2
                If Len(fileName) >= pos + 3 Then
                    targetYear = CStr(2018 + CInt(Mid(fileName, pos, 2)))
                    targetMonth = CStr(CInt(Mid(fileName, pos + 2, 2)))
                    Exit Function
                End If
            End If
        End If
    Next file
    ' 推定失敗時、ユーザー入力要求
    If targetYear = "" Or targetMonth = "" Then
        targetYear = InputBox("西暦年を入力してください（例: 2023）:", "診療年")
        targetMonth = InputBox("月を入力してください（1~12）:", "診療月")
    End If
End Function

Function FindOrCreateReport(savePath As String, targetYear As String, targetMonth As String, templatePath As String) As String
    Dim fso As Object, existingFile As Object
    Dim fileName As String, filePath As String
    Dim csvYYMM As String
    csvYYMM = Format(CInt(targetYear) - 2018, "00") & Format(CInt(targetMonth), "00")  ' 和暦年+月コード
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' 保存フォルダ内に既存のRYYMMファイルがあるか確認
    For Each existingFile In fso.GetFolder(savePath).Files
        If LCase(fso.GetExtensionName(existingFile.Name)) = "xlsm" Or LCase(fso.GetExtensionName(existingFile.Name)) = "xlsx" Then
            If InStr(existingFile.Name, "保険請求管理報告書_R" & csvYYMM) > 0 Then
                FindOrCreateReport = existingFile.Path  ' 既存ファイルのパスを返す
                Exit Function
            End If
        End If
    Next existingFile
    ' 新規作成
    fileName = "保険請求管理報告書_R" & csvYYMM & ".xlsm"
    filePath = savePath & "\" & fileName
    ' テンプレートを元に新規ブック作成
    On Error Resume Next
    Dim tmplWb As Workbook
    Set tmplWb = Workbooks.Open(templatePath)
    On Error GoTo 0
    If tmplWb Is Nothing Then
        MsgBox "テンプレートを開けませんでした: " & templatePath, vbCritical, "エラー"
        FindOrCreateReport = ""
        Exit Function
    End If
    On Error Resume Next
    tmplWb.SaveAs filePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
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
    ' Dictionaryが空なら処理しない
    If dataDict.Count = 0 Then Exit Sub
    ' payerTypeに応じた転記列を決定
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
        ws.Cells(j, payerColumn).Font.Bold = True
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
    Set gRebillForm = uf
    gRebillForm.Show vbModal
    ' フォーム終了後、選択結果は gRebillData と gLateData に格納済み
End Sub

Function CreateRebillSelectionForm(listData As Object) As Object
    Dim uf As Object, listBox As Object, btnOK As Object
    Dim i As Long, rowData As Variant
    ' UserForm を動的に作成
    Set uf = VBA.UserForms.Add()
    uf.Caption = "返戻再請求の選択"
    uf.Width = 400
    uf.Height = 500
    ' ListBoxを追加
    Set listBox = uf.Controls.Add("Forms.ListBox.1", "listBox", True)
    listBox.Left = 20
    listBox.Top = 20
    listBox.Width = 350
    listBox.Height = 350
    listBox.MultiSelect = 1  ' 複数選択可能
    ' リストにデータを追加（調剤年月 | 患者氏名 | 医療機関名 | 点数）
    For i = 0 To listData.Count - 1
        rowData = listData.Items()(i)
        listBox.AddItem rowData(1) & " | " & rowData(0) & " | " & rowData(2) & " | " & rowData(3)
    Next i
    ' OKボタンを追加
    Set btnOK = uf.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    btnOK.Caption = "確定"
    btnOK.Left = 150
    btnOK.Top = 400
    btnOK.Width = 100
    btnOK.Height = 30
    ' ボタンクリック時の処理を設定
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
    ' 選択状態に応じて振り分け
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) Then
            ' 選択されたもの -> 返戻再請求
            gRebillData.Add gOlderList.Keys()(i), gOlderList.Items()(i)
        Else
            ' 選択されなかったもの -> 月遅れ請求
            gLateData.Add gOlderList.Keys()(i), gOlderList.Items()(i)
        End If
    Next i
    ' フォームを閉じる
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
        ' ファイルなし
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
    ' 前月詳細シートを取得（シート2想定）
    Set wsPrevDetails = prevBook.Sheets(2)
    ' 対象カテゴリ開始行を取得
    Dim categoryLabel As String
    If payerType = "社保" Then
        categoryLabel = "社保未請求扱い"
    Else
        categoryLabel = "国保未請求扱い"
    End If
    startRow = GetStartRow(wsPrevDetails, categoryLabel)
    If startRow = 0 Then
        prevBook.Close False
        Set AddUnclaimedRecords = Nothing
        Exit Function
    End If
    ' データ収集（開始行以降、空白行まで）
    Set gUnclaimedList = CreateObject("Scripting.Dictionary")
    endRow = startRow + 3  ' 基本枠4行
    Do While wsPrevDetails.Cells(endRow, 4).Value <> "" Or wsPrevDetails.Cells(endRow + 1, 4).Value <> ""
        endRow = endRow + 1
        If endRow > wsPrevDetails.Rows.Count Then Exit Do
    Loop
    For row = startRow + 1 To endRow
        If wsPrevDetails.Cells(row, 4).Value <> "" Then
            Dim prevRowData As Variant
            prevRowData = Array(wsPrevDetails.Cells(row, 4).Value, wsPrevDetails.Cells(row, 5).Value, wsPrevDetails.Cells(row, 6).Value, wsPrevDetails.Cells(row, 10).Value)
            gUnclaimedList.Add row, prevRowData
        End If
    Next row
    prevBook.Close False
    ' ユーザーフォームで追加選択
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
    ' フォーム終了後、gSelectedUnclaimed に結果格納済み
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
    circledNumbers = Array("①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫")
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
        Case "4": westernYear = 1988 + yearPart   ' 平成 (1989年=平成1年)
        Case Else: westernYear = 2018 + yearPart  ' （デフォルト:令和として計算）
    End Select
    ConvertToWesternDate = Right(CStr(westernYear), 2) & "." & monthPart
End Function

' 半年ごとの売掛データ比較・誤差分析機能
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
    Set analysisWb = ThisWorkbook
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
            Dim wsA As Worksheet, wsCSV As Worksheet
            Set wsA = repWb.Sheets(1)
            Dim dailyTotal As Long
            dailyTotal = 0
            On Error Resume Next
            dailyTotal = CLng(wsA.Range("J50").Value)
            On Error GoTo 0
            Dim billedTotal As Long
            billedTotal = 0
            On Error Resume Next
            billedTotal = CLng(wsA.Range("J50").Value)
            On Error GoTo 0
            Dim payAmount As Long
            payAmount = 0
            For Each wsCSV In repWb.Worksheets
                If InStr(wsCSV.Name, "振込額明細書") > 0 Then
                    On Error Resume Next
                    payAmount = CLng(wsCSV.Cells(wsCSV.Rows.Count, 3).End(xlUp).Value)
                    On Error GoTo 0
                    Exit For
                End If
            Next wsCSV
            repWb.Close False
            Dim pointDiff As Long
            pointDiff = dailyTotal - billedTotal
            outSheet.Cells(rowIndex, 1).Value = inputYear & "年" & m & "月"
            outSheet.Cells(rowIndex, 2).Value = dailyTotal
            outSheet.Cells(rowIndex, 3).Value = billedTotal
            outSheet.Cells(rowIndex, 4).Value = payAmount
            outSheet.Cells(rowIndex, 5).Value = pointDiff
            rowIndex = rowIndex + 1
        Else
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