Option Explicit

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
    Dim fixfFiles As Collection
    Dim file As Object
    
    ' 1 CSVフォルダをユーザーに選択させる
    csvFolder = SelectCSVFolder()
    If csvFolder = "" Then Exit Sub

    ' 1.1 フォルダが空なら処理を中止
    If IsFolderEmpty(csvFolder) Then
        MsgBox "選択したフォルダにはCSVファイルがありません。処理を中止します。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 2 テンプレートパス・保存フォルダ取得
    templatePath = GetTemplatePath()
    savePath = GetSavePath()
    If templatePath = "" Or savePath = "" Then Exit Sub

    ' 3 ファイルシステムオブジェクトの作成
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 4 フォルダ内のすべての `fixf` ファイルを取得
    Set fixfFiles = FindAllFixfFiles(fso, csvFolder)
    
    ' 5 `fixfFiles` が空の場合、通常のCSV処理に切り替え
    If fixfFiles Is Nothing Or fixfFiles.Count = 0 Then
        ProcessWithoutFixf fso, csvFolder, savePath, templatePath
        Exit Sub
    End If

    ' 6 複数の `fixf` ファイルを順番に処理
    For Each file In fixfFiles
        fixfFile = file.Path
        
        ' 7 対象年月を取得
        targetYear = ""
        targetMonth = ""
        GetYearMonthFromFixf fixfFile, targetYear, targetMonth
        
        ' **対象年月が取得できなかった場合はスキップ**
        If targetYear = "" Or targetMonth = "" Then
            MsgBox "ファイル " & fixfFile & " から診療年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If

        ' 8 既存の報告書（RYYMMファイル）がある場合はスキップ
        Dim eraYearCode As String, csvYYMM As String, fileName As String, filePath As String
        eraYearCode = Format(CInt(targetYear) - 2018, "00")  ' 和暦年（令和）を2桁コード化
        csvYYMM = eraYearCode & targetMonth
        fileName = "保険請求管理報告書_R" & csvYYMM & ".xlsx"
        filePath = savePath & "\" & fileName
        If fso.FileExists(filePath) Then
            MsgBox "報告書 " & fileName & " は既に存在するため、処理をスキップします。", vbInformation, "スキップ"
            GoTo NextFile
        End If

        ' 9 対象Excelファイルを取得（既存がなければ新規作成）
        targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)
        
        ' **ファイルを作成・取得できなかった場合はスキップ**
        If targetFile = "" Then
            MsgBox "診療年月 " & targetYear & "年" & targetMonth & "月 のExcelファイルを作成できませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If
        
        ' 10 Excelを開く
        On Error Resume Next
        Set newBook = Workbooks.Open(targetFile)
        On Error GoTo 0
        If newBook Is Nothing Then
            MsgBox "ファイル " & targetFile & " を開けませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If

        ' 11 テンプレート情報を設定
        SetTemplateInfo newBook, targetYear, targetMonth

        ' 12 フォルダ内の各種CSVファイルを順に処理（fixf → 振込額明細書 → 返戻内訳書 → 増減点連絡書）
        ProcessAllCSVFiles fso, newBook, csvFolder, targetYear, targetMonth

        ' 13 保存して閉じる
        newBook.Save
        newBook.Close False

NextFile:
    Next file

    ' 14 処理完了メッセージ
    MsgBox "すべての `fixf` ファイルの処理が完了しました！", vbInformation, "完了"
End Sub

Sub ProcessWithoutFixf(fso As Object, csvFolder As String, savePath As String, templatePath As String)
    Dim targetYear As String
    Dim targetMonth As String
    Dim targetFile As String
    Dim newBook As Workbook

    ' 1 fixfがない場合、最初のCSVから診療年月を取得
    targetYear = ""
    targetMonth = ""
    GetYearMonthFromCSV fso, csvFolder, targetYear, targetMonth

    ' **診療年月が取得できなかった場合は処理中止**
    If targetYear = "" Or targetMonth = "" Then
        MsgBox "CSVファイルから診療年月を取得できませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 2 対象Excelファイルを取得（既存がなければ新規作成）
    targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)
    If targetFile = "" Then
        MsgBox "診療年月 " & targetYear & "年" & targetMonth & "月 のExcelファイルを作成できませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 3 Excelを開く
    On Error Resume Next
    Set newBook = Workbooks.Open(targetFile)
    On Error GoTo 0
    If newBook Is Nothing Then
        MsgBox "ファイル " & targetFile & " を開けませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 4 テンプレート情報を設定
    SetTemplateInfo newBook, targetYear, targetMonth

    ' 5 CSVファイルを順次処理（振込額明細書 → 返戻内訳書 → 増減点連絡書）
    ProcessAllCSVFiles fso, newBook, csvFolder, targetYear, targetMonth

    ' 6 保存して閉じる
    newBook.Save
    newBook.Close False

    ' 7 完了メッセージ
    MsgBox "CSVファイルの処理が完了しました！", vbInformation, "完了"
End Sub

Function FindOrCreateReport(savePath As String, targetYear As String, targetMonth As String, templatePath As String) As String
    Dim fso As Object
    Dim filePath As String
    Dim fileName As String
    Dim existingFile As Object
    Dim newWb As Workbook
    Dim csvYYMM As String

    ' **診療年月を "YYMM" 形式に変換（令和）**
    csvYYMM = Format(CInt(targetYear) - 2018, "00") & targetMonth

    ' FileSystemObject の作成
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' **保存フォルダ内に既存の報告書 (`RYYMM.xlsx`) があるか検索**
    For Each existingFile In fso.GetFolder(savePath).Files
        If LCase(fso.GetExtensionName(existingFile.Name)) = "xlsx" Then
            If Right(fso.GetBaseName(existingFile.Name), 4) = csvYYMM Then
                ' **同じ診療年月のファイルが見つかった場合、そのパスを返す**
                FindOrCreateReport = existingFile.Path
                Exit Function
            End If
        End If
    Next existingFile

    ' **該当ファイルがなければ、新規作成**
    fileName = "保険請求管理報告書_R" & csvYYMM & ".xlsx"
    filePath = savePath & "\" & fileName

    ' **テンプレートを開き、新規ファイルとして保存**
    Set newWb = Workbooks.Open(templatePath)
    newWb.SaveAs filePath, FileFormat:=xlOpenXMLWorkbook  ' マクロ無しExcel形式で保存
    newWb.Close SaveChanges:=True

    ' パスを返す
    FindOrCreateReport = filePath
End Function

Sub SetTemplateInfo(newBook As Workbook, targetYear As String, targetMonth As String)
    Dim wsA As Worksheet, wsB As Worksheet
    Dim receiptYear As Integer, receiptMonth As Integer
    Dim sendMonth As Integer, sendDate As String

    ' **西暦年と調剤月の数値取得**
    receiptYear = CInt(targetYear)
    receiptMonth = CInt(targetMonth)

    ' **請求月（翌月）の算出**
    sendMonth = receiptMonth + 1
    If sendMonth = 13 Then sendMonth = 1
    sendDate = sendMonth & "月10日請求分"

    ' **テンプレートシートA, Bを取得**
    Set wsA = newBook.Sheets("A")
    Set wsB = newBook.Sheets("B")

    ' **シート名変更**（例：「R7.2」や「②」など）
    wsA.Name = "R" & (receiptYear - 2018) & "." & receiptMonth
    wsB.Name = ConvertToCircledNumber(receiptMonth)

    ' **帳票内の情報を設定**
    wsA.Range("G2").Value = targetYear & "年" & targetMonth & "月調剤分"
    wsA.Range("I2").Value = sendDate
    wsA.Range("J2").Value = ThisWorkbook.Sheets(1).Range("B1").Value  ' 発行者情報など
    wsB.Range("H1").Value = targetYear & "年" & targetMonth & "月調剤分"
    wsB.Range("J1").Value = sendDate
    wsB.Range("L1").Value = ThisWorkbook.Sheets(1).Range("B1").Value
End Sub

Sub ProcessAllCSVFiles(fso As Object, newBook As Workbook, csvFolder As String, targetYear As String, targetMonth As String)
    Dim csvFile As Object
    Dim fileType As String
    Dim wsDetails As Worksheet, wsCSV As Worksheet
    Dim sheetName As String
    Dim sheetIndex As Integer

    ' 詳細データシート（2番目）を取得
    Set wsDetails = newBook.Sheets(2)

    ' 処理対象年月の和暦コードを生成（GYYMM形式の文字列）
    Dim eraCode As String, eraYear As Integer, GYYMM As String
    If CInt(targetYear) >= 2019 Then
        eraCode = "5"  ' 令和
        eraYear = CInt(targetYear) - 2018
    ElseIf CInt(targetYear) >= 1989 Then
        eraCode = "4"  ' 平成
        eraYear = CInt(targetYear) - 1988
    ElseIf CInt(targetYear) >= 1926 Then
        eraCode = "3"  ' 昭和
        eraYear = CInt(targetYear) - 1925
    ElseIf CInt(targetYear) >= 1912 Then
        eraCode = "2"  ' 大正
        eraYear = CInt(targetYear) - 1911
    Else
        eraCode = "1"  ' 明治（または不明）
        eraYear = CInt(targetYear) - 1867
    End If
    GYYMM = eraCode & Format(eraYear, "00") & targetMonth

    ' 種類別のコレクションを用意
    Dim fmeiFilesColl As New Collection
    Dim henrFilesColl As New Collection
    Dim zognFilesColl As New Collection

    ' フォルダ内のCSVファイルを分類してコレクションに格納
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" Then
            If InStr(LCase(csvFile.Name), "fmei") > 0 And Right(fso.GetBaseName(csvFile.Name), Len(GYYMM)) = GYYMM Then
                fmeiFilesColl.Add csvFile
            ElseIf InStr(LCase(csvFile.Name), "henr") > 0 And Right(fso.GetBaseName(csvFile.Name), Len(GYYMM)) = GYYMM Then
                henrFilesColl.Add csvFile
            ElseIf InStr(LCase(csvFile.Name), "zogn") > 0 And Right(fso.GetBaseName(csvFile.Name), Len(GYYMM)) = GYYMM Then
                zognFilesColl.Add csvFile
            End If
        End If
    Next csvFile

    ' 1) 振込額明細書（fmei）CSVを処理
    fileType = "振込額明細書"
    For Each csvFile In fmeiFilesColl
        sheetName = fso.GetBaseName(csvFile.Name)
        sheetName = GetUniqueSheetName(newBook, sheetName)
        sheetIndex = Application.WorksheetFunction.Min(3, newBook.Sheets.Count + 1)
        Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(sheetIndex - 1))
        wsCSV.Name = sheetName

        ImportCSVData csvFile.Path, wsCSV, fileType
        TransferBillingDetails newBook, csvFile.Name
    Next csvFile

    ' 2) 返戻内訳書（henr）CSVを処理
    fileType = "返戻内訳書"
    For Each csvFile In henrFilesColl
        sheetName = fso.GetBaseName(csvFile.Name)
        sheetName = GetUniqueSheetName(newBook, sheetName)
        sheetIndex = Application.WorksheetFunction.Min(3, newBook.Sheets.Count + 1)
        Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(sheetIndex - 1))
        wsCSV.Name = sheetName

        ImportCSVData csvFile.Path, wsCSV, fileType
        TransferBillingDetails newBook, csvFile.Name
    Next csvFile

    ' 3) 増減点連絡書（zogn）CSVを処理
    fileType = "増減点連絡書"
    For Each csvFile In zognFilesColl
        sheetName = fso.GetBaseName(csvFile.Name)
        sheetName = GetUniqueSheetName(newBook, sheetName)
        sheetIndex = Application.WorksheetFunction.Min(3, newBook.Sheets.Count + 1)
        Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(sheetIndex - 1))
        wsCSV.Name = sheetName

        ImportCSVData csvFile.Path, wsCSV, fileType
        TransferBillingDetails newBook, csvFile.Name
    Next csvFile
End Sub

Function SelectCSVFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "CSVフォルダを選択してください"
        If .Show = -1 Then
            SelectCSVFolder = .SelectedItems(1) & "\"
        Else
            MsgBox "フォルダが選択されませんでした。処理を中止します。", vbExclamation, "エラー"
            SelectCSVFolder = ""
        End If
    End With
End Function

Function IsFolderEmpty(folderPath As String) As Boolean
    Dim fso As Object
    Dim folder As Object

    ' FileSystemObject を作成
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' フォルダが存在しない場合は True
    If Not fso.FolderExists(folderPath) Then
        IsFolderEmpty = True
        Exit Function
    End If
    ' フォルダ内のファイル数を確認
    Set folder = fso.GetFolder(folderPath)
    If folder.Files.Count = 0 Then
        IsFolderEmpty = True   ' ファイルがない場合
    Else
        IsFolderEmpty = False  ' ファイルが存在する場合
    End If
End Function

Function GetTemplatePath() As String
    ' 設定シートのB2セルからテンプレートパスを取得
    GetTemplatePath = ThisWorkbook.Sheets(1).Range("B2").Value & "\保険請求管理報告書テンプレート20250222.xltm"
End Function

Function GetSavePath() As String
    ' 設定シートのB3セルから保存先フォルダを取得
    GetSavePath = ThisWorkbook.Sheets(1).Range("B3").Value
End Function

Function FindAllFixfFiles(fso As Object, csvFolder As String) As Collection
    Dim csvFile As Object
    Dim fixfFiles As New Collection

    ' フォルダ内のすべてのファイルをチェック
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" And InStr(LCase(csvFile.Name), "fixf") > 0 Then
            fixfFiles.Add csvFile  ' fixfファイルをコレクションに追加
        End If
    Next csvFile

    Set FindAllFixfFiles = fixfFiles
End Function

Sub GetYearMonthFromFixf(fixfFile As String, ByRef targetYear As String, ByRef targetMonth As String)
    Dim fileName As String
    Dim datePart As String
    Dim yearPart As String, monthPart As String

    ' 1 ファイル名から日時部分を取得（例: "RTfixf...20250228150730.csv" の "20250228150730"）
    fileName = Mid(fixfFile, InStrRev(fixfFile, "\") + 1)
    datePart = Mid(fileName, 18, 14)   ' ファイル名の18文字目から14桁が日時部分

    ' 2 年月日に分解
    yearPart = Left(datePart, 4)      ' "2025"
    monthPart = Mid(datePart, 5, 2)   ' "02"
    ' ※日付や時刻部分（dayPart, hourPart 等）は使用しないため取得省略

    ' 3 戻り値に設定
    targetYear = yearPart    ' 西暦年
    targetMonth = monthPart  ' 月（2桁）
End Sub

Sub GetYearMonthFromCSV(fso As Object, csvFolder As String, ByRef targetYear As String, ByRef targetMonth As String)
    Dim csvFile As Object
    Dim ts As Object
    Dim lineText As String
    Dim era As String
    Dim yearPart As String, monthPart As String
    Dim westernYear As Integer

    ' フォルダ内の最初のCSVから GYYMM を取得
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(Right(csvFile.Name, 4)) = ".csv" Then
            ' CSVテキストを開く（読み取り専用, UTF-8）
            Set ts = fso.OpenTextFile(csvFile.Path, 1, False, -2)
            Do While Not ts.AtEndOfStream
                lineText = ts.ReadLine
                If Len(lineText) >= 5 Then
                    era = Left(lineText, 1)         ' 和暦元号コード (1:明治, 2:大正, 3:昭和, 4:平成, 5:令和)
                    yearPart = Mid(lineText, 2, 2)   ' 元号年（2桁）
                    monthPart = Right(lineText, 2)   ' 月（2桁）
                    ' 和暦コードを西暦年に変換
                    Select Case era
                        Case "1": westernYear = 1867 + CInt(yearPart)  ' 明治
                        Case "2": westernYear = 1911 + CInt(yearPart)  ' 大正
                        Case "3": westernYear = 1925 + CInt(yearPart)  ' 昭和
                        Case "4": westernYear = 1988 + CInt(yearPart)  ' 平成
                        Case "5": westernYear = 2018 + CInt(yearPart)  ' 令和
                        Case Else: westernYear = 2000 + CInt(yearPart)  ' その他（仮）
                    End Select
                    targetYear = CStr(westernYear)
                    targetMonth = monthPart
                    Exit Do
                End If
            Loop
            ts.Close
            If targetYear <> "" And targetMonth <> "" Then
                Exit For
            End If
        End If
    Next csvFile
End Sub

Sub ImportCSVData(csvFile As String, ws As Worksheet, fileType As String)
    Dim colMap As Object
    Dim fso As Object
    Dim ts As Object
    Dim lineText As String
    Dim dataArray As Variant
    Dim i As Long, j As Long, key
    Dim isHeader As Boolean

    On Error GoTo ErrorHandler

    ' 画面更新と自動計算を停止
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' CSV項目のマッピングを取得
    Set colMap = GetColumnMapping(fileType)

    ' シートをクリア
    ws.Cells.Clear

    ' 1行目に項目名を書き込む
    j = 1
    For Each key In colMap.Keys
        ws.Cells(1, j).Value = colMap(key)
        j = j + 1
    Next key

    ' CSVファイルを読み込んでデータ転記
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(csvFile, 1, False, -2)  ' UTF-8として開く
    
    i = 2
    isHeader = True
    Do While Not ts.AtEndOfStream
        lineText = ts.ReadLine
        dataArray = Split(lineText, ",")
        If isHeader Then
            ' 最初の行（ヘッダー）はスキップ
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

    ' 列幅を自動調整
    ws.Cells.EntireColumn.AutoFit

    ' 画面更新と自動計算を再開
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
    If Not ts Is Nothing Then ts.Close
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub TransferBillingDetails(newBook As Workbook, csvFileName As String)
    Dim wsBilling As Worksheet, wsDetails As Worksheet
    Dim lastRowBilling As Long
    Dim i As Long
    Dim dispensingMonth As String, convertedMonth As String
    Dim payerCode As String, payerType As String
    Dim receiptNo As String
    Dim startRowDict As Object
    Dim rebillDict As Object, lateDict As Object, unpaidDict As Object, assessmentDict As Object
    Dim category As String
    Dim startRow As Long
    Dim dataDict As Object
    Dim rowData As Variant
    Dim a As Long, b As Long, c As Long
    Dim csvYYMM As String

    ' シート設定
    Set wsBilling = newBook.Sheets(1)  ' 請求データシート（メイン）
    Set wsDetails = newBook.Sheets(2)  ' 詳細データ用シート

    ' **診療年月 (csvYYMM) を取得（GYYMMの下2桁YYとMM）**
    csvYYMM = ""
    If wsBilling.Cells(2, 2).Value <> "" Then
        csvYYMM = Right(wsBilling.Cells(2, 2).Value, 4)
    End If

    ' 請求先種別の判定（ファイル名の7文字目で判断）
    Dim baseName As String
    If InStr(csvFileName, ".") > 0 Then
        baseName = Left(csvFileName, InStrRev(csvFileName, ".") - 1)
    Else
        baseName = csvFileName
    End If
    If Len(baseName) >= 7 Then
        payerCode = Mid(baseName, 7, 1)
    Else
        payerCode = ""
    End If
    Select Case payerCode
        Case "1": payerType = "社保"
        Case "2": payerType = "国保"
        Case Else: payerType = "労災"
    End Select

    ' **開始行管理用 Dictionary 作成**
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

    ' **区分ごとの Dictionary を作成**
    Set rebillDict = CreateObject("Scripting.Dictionary")     ' 返戻再請求
    Set lateDict = CreateObject("Scripting.Dictionary")       ' 月遅れ請求
    Set unpaidDict = CreateObject("Scripting.Dictionary")     ' 未請求扱い
    Set assessmentDict = CreateObject("Scripting.Dictionary") ' 返戻・査定

    ' 請求データ最終行を取得（D列を基準に下方向検索）
    lastRowBilling = wsBilling.Cells(wsBilling.Rows.Count, "D").End(xlUp).Row

    ' **請求データを走査して Dictionary に振り分け**
    For i = 2 To lastRowBilling
        dispensingMonth = wsBilling.Cells(i, 2).Value  ' GYYMM形式の診療月
        convertedMonth = ConvertToWesternDate(dispensingMonth)
        ' 対象診療月（csvYYMM）と異なる過去月データのみ対象
        If csvYYMM <> "" And Right(dispensingMonth, 4) <> csvYYMM Then
            rowData = Array(wsBilling.Cells(i, 4).Value, convertedMonth, wsBilling.Cells(i, 5).Value, wsBilling.Cells(i, 10).Value)
            If InStr(LCase(csvFileName), "fixf") > 0 Then
                ' fixfファイルの場合：過去月データはデフォルトで「月遅れ請求」に分類
                lateDict.Add wsBilling.Cells(i, 1).Value, rowData
            ElseIf InStr(LCase(csvFileName), "zogn") > 0 Then
                ' 増減点連絡書：未請求扱い
                unpaidDict.Add wsBilling.Cells(i, 1).Value, rowData
            ElseIf InStr(LCase(csvFileName), "henr") > 0 Then
                ' 返戻内訳書：返戻・査定
                assessmentDict.Add wsBilling.Cells(i, 1).Value, rowData
            End If
        End If
    Next i

    ' **各カテゴリの超過分件数を算出（初期枠4件を超えた数）**
    a = 0: b = 0: c = 0
    If rebillDict.Count > 4 Then a = rebillDict.Count - 4
    If lateDict.Count > 4 Then b = lateDict.Count - 4
    If assessmentDict.Count > 4 Then c = assessmentDict.Count - 4

    ' **各カテゴリの転記開始行を調整**
    Dim lateStartRow As Long, assessmentStartRow As Long, unpaidStartRow As Long
    lateStartRow = startRowDict("月遅れ請求") + 1 + a
    assessmentStartRow = startRowDict("返戻・査定") + 1 + a + b
    unpaidStartRow = startRowDict("未請求扱い") + 1 + a + b + c

    ' **必要な行を挿入**
    If a + b + c > 0 Then
        If a > 0 Then wsDetails.Rows(startRowDict("月遅れ請求") + 1 & ":" & startRowDict("月遅れ請求") + a).Insert Shift:=xlDown
        If b > 0 Then wsDetails.Rows(startRowDict("返戻・査定") + 1 & ":" & startRowDict("返戻・査定") + b).Insert Shift:=xlDown
        If c > 0 Then wsDetails.Rows(startRowDict("未請求扱い") + 1 & ":" & startRowDict("未請求扱い") + c).Insert Shift:=xlDown
    End If

    ' **各 Dictionary のデータを詳細シートに転記**
    If rebillDict.Count > 0 Then
        startRow = startRowDict("返戻再請求")
        TransferData rebillDict, wsDetails, startRow, payerType
    End If
    If lateDict.Count > 0 Then
        startRow = startRowDict("月遅れ請求")
        TransferData lateDict, wsDetails, startRow, payerType
    End If
    If unpaidDict.Count > 0 Then
        startRow = startRowDict("未請求扱い")
        TransferData unpaidDict, wsDetails, startRow, payerType
    End If
    If assessmentDict.Count > 0 Then
        startRow = startRowDict("返戻・査定")
        TransferData assessmentDict, wsDetails, startRow, payerType
    End If

    MsgBox payerType & " のデータ転記が完了しました！", vbInformation, "処理完了"
End Sub

Sub TransferData(dataDict As Object, ws As Worksheet, startRow As Long, payerType As String)
    Dim key As Variant, rowData As Variant
    Dim r As Long
    Dim payerColumn As Long

    ' Dictionaryが空の場合は何もしない
    If dataDict.Count = 0 Then Exit Sub

    ' payerTypeに応じた転記列を決定（社保=H列, 国保=I列）
    If payerType = "社保" Then
        payerColumn = 8   ' H列
    ElseIf payerType = "国保" Then
        payerColumn = 9   ' I列
    Else
        Exit Sub  ' 労災の場合は転記不要
    End If

    r = startRow
    For Each key In dataDict.Keys
        rowData = dataDict(key)
        ws.Cells(r, 4).Value = rowData(0)           ; ws.Cells(r, 4).Value = rowData(0)  ' 患者氏名
        ws.Cells(r, 5).Value = rowData(1)           ; ws.Cells(r, 5).Value = rowData(1)  ' 調剤年月（YY.MM）
        ws.Cells(r, 6).Value = rowData(2)           ; ws.Cells(r, 6).Value = rowData(2)  ' 医療機関名
        ws.Cells(r, payerColumn).Value = payerType  ; ws.Cells(r, payerColumn).Value = payerType  ' 請求先種別（社保/国保）
        ws.Cells(r, payerColumn).Font.Bold = True   ; ws.Cells(r, payerColumn).Font.Bold = True   ' 強調表示
        ws.Cells(r, 10).Value = rowData(3)          ; ws.Cells(r, 10).Value = rowData(3)  ' 請求点数
        r = r + 1
    Next key
End Sub

Function GetColumnMapping(fileType As String) As Object
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")

    Select Case fileType
        Case "振込額明細書"
            colMap.Add 2, "診療（調剤）年月"
            colMap.Add 5, "受付番号"
            colMap.Add 14, "氏名"
            colMap.Add 16, "生年月日"
            colMap.Add 22, "医療保険＿療養の給付＿請求点数"
            colMap.Add 23, "医療保険＿療養の給付＿決定点数"
            colMap.Add 24, "医療保険＿療養の給付＿一部負担金"
            colMap.Add 25, "医療保険＿療養の給付＿金額"
            ' 公費（第1～第5）のデータ
            Dim k As Integer
            For k = 1 To 5
                colMap.Add 33 + (k - 1) * 10, "第" & k & "公費_請求点数"
                colMap.Add 34 + (k - 1) * 10, "第" & k & "公費_決定点数"
                colMap.Add 35 + (k - 1) * 10, "第" & k & "公費_患者負担金"
                colMap.Add 36 + (k - 1) * 10, "第" & k & "公費_金額"
            Next k
            colMap.Add 82, "算定額合計"

        Case "請求確定状況"
            colMap.Add 4, "診療（調剤）年月"
            colMap.Add 5, "氏名"
            colMap.Add 7, "生年月日"
            colMap.Add 9, "医療機関名称"
            colMap.Add 13, "総合計点数"
            For k = 1 To 4
                colMap.Add 16 + (k - 1) * 3, "第" & k & "公費_請求点数"
            Next k
            colMap.Add 30, "請求確定状況"
            colMap.Add 31, "エラー区分"

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
    End Select

    Set GetColumnMapping = colMap
End Function

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
    Dim era As String, yearNum As Integer, westernYear As Integer, monthPart As String
    If Len(dispensingMonth) < 5 Then
        ConvertToWesternDate = ""
        Exit Function
    End If
    era = Left(dispensingMonth, 1)        ' 元号コード
    yearNum = CInt(Mid(dispensingMonth, 2, 2))
    monthPart = Right(dispensingMonth, 2)
    Select Case era
        Case "5": westernYear = 2018 + yearNum   ' 令和（2019年=令和元年）
        Case "4": westernYear = 1988 + yearNum   ' 平成（1989年=平成元年）
        Case "3": westernYear = 1925 + yearNum   ' 昭和（1926年=昭和元年）
        Case "2": westernYear = 1911 + yearNum   ' 大正（1912年=大正元年）
        Case "1": westernYear = 1867 + yearNum   ' 明治（1868年=明治元年）
        Case Else: westernYear = 2000 + yearNum  ' その他（仮定）
    End Select
    ConvertToWesternDate = Right(westernYear, 2) & "." & monthPart  ' "YY.MM"形式
End Function

Function GetUniqueSheetName(wb As Workbook, baseName As String) As String
    Dim newName As String
    Dim counter As Integer
    Dim ws As Worksheet
    Dim exists As Boolean

    newName = baseName
    counter = 1
    Do
        exists = False
        For Each ws In wb.Sheets
            If LCase(ws.Name) = LCase(newName) Then
                exists = True
                Exit For
            End If
        Next ws
        If exists Then
            newName = baseName & "_" & counter
            counter = counter + 1
        End If
    Loop While exists

    GetUniqueSheetName = newName
End Function

Function GetStartRow(ws As Worksheet, categoryName As String) As Long
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(what:=categoryName, LookAt:=xlWhole)
    If Not foundCell Is Nothing Then
        GetStartRow = foundCell.Row
    Else
        GetStartRow = 0
    End If
End Function