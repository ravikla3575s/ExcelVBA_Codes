Option Explicit

' CSV一括処理マクロ: 請求CSV(`fixf`)と各種明細CSV(`fmei`, `henr`, `zogn`)を読み込み、報告書Excelを作成・更新

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

    ' 1. CSVフォルダをユーザーに選択させる
    csvFolder = SelectCSVFolder()
    If csvFolder = "" Then Exit Sub

    ' 1.1 フォルダが空なら処理を中止
    If IsFolderEmpty(csvFolder) Then
        MsgBox "選択したフォルダにはCSVファイルがありません。処理を中止します。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 2. テンプレートパス・保存先フォルダを取得
    templatePath = GetTemplatePath()
    savePath = GetSavePath()
    If templatePath = "" Or savePath = "" Then Exit Sub

    ' 3. FileSystemObjectの用意
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 4. フォルダ内のすべての`fixf`ファイルを取得
    Set fixfFiles = FindAllFixfFiles(fso, csvFolder)

    ' 5. `fixf`ファイルがない場合、通常のCSV処理(ProcessWithoutFixf)に切り替え
    If fixfFiles Is Nothing Or fixfFiles.Count = 0 Then
        ProcessWithoutFixf fso, csvFolder, savePath, templatePath
        Exit Sub
    End If

    ' 6. 複数の`fixf`ファイルを順に処理
    For Each file In fixfFiles
        fixfFile = file.Path

        ' 7. `fixf`ファイル名から対象の診療年・月を取得
        targetYear = ""
        targetMonth = ""
        GetYearMonthFromFixf(fixfFile, targetYear, targetMonth)

        ' **診療年月が取得できなかった場合はスキップ**
        If targetYear = "" Or targetMonth = "" Then
            MsgBox "ファイル " & fixfFile & " から診療年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If

        ' 8. 既存の報告書ファイル(RYYMM形式)がある場合はスキップ
        Dim eraYearCode As String, csvYYMM As String, fileName As String, filePath As String
        eraYearCode = Format(CInt(targetYear) - 2018, "00")  ' 和暦年（令和）のコード (例:2025→07)
        csvYYMM = eraYearCode & targetMonth                  ' RYYMM形式の文字列作成 (例: R0702)
        fileName = "保険請求管理報告書_R" & csvYYMM & ".xlsx"
        filePath = savePath & "\" & fileName
        If fso.FileExists(filePath) Then
            MsgBox "報告書 " & fileName & " は既に存在するため、処理をスキップします。", vbInformation, "スキップ"
            GoTo NextFile
        End If

        ' 9. 対象Excelファイルを取得（既存がなければテンプレートから新規作成）
        targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)

        ' **ファイルを作成・取得できなかった場合はスキップ**
        If targetFile = "" Then
            MsgBox "診療年月 " & targetYear & "年" & targetMonth & "月 のExcelファイルを作成できませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If

        ' 10. 報告書Excelを開く
        On Error Resume Next
        Set newBook = Workbooks.Open(targetFile)
        On Error GoTo 0
        If newBook Is Nothing Then
            MsgBox "ファイル " & targetFile & " を開けませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If

        ' 11. テンプレート情報を設定（シート名やタイトル情報の更新）
        SetTemplateInfo newBook, targetYear, targetMonth

        ' 11.5 `fixf` CSVをメインシートにインポート（請求確定状況データ）
        ImportCSVData fixfFile, newBook.Sheets(1), "請求確定状況"

        ' 12. フォルダ内の各種CSVファイルを順次処理（振込額明細書→返戻内訳書→増減点連絡書）
        ProcessAllCSVFiles fso, newBook, csvFolder, targetYear, targetMonth

        ' 13. 保存してブックを閉じる
        newBook.Save
        newBook.Close False

NextFile:
    Next file

    ' 14. 全`fixf`ファイルの処理完了メッセージ
    MsgBox "すべての請求データの処理が完了しました！", vbInformation, "処理完了"
End Sub

Sub ProcessWithoutFixf(fso As Object, csvFolder As String, savePath As String, templatePath As String)
    Dim targetYear As String
    Dim targetMonth As String
    Dim targetFile As String
    Dim newBook As Workbook

    ' 1. `fixf`がない場合、フォルダ内最初のCSVから診療年月を取得（振込明細などから推定）
    targetYear = ""
    targetMonth = ""
    GetYearMonthFromCSV(fso, csvFolder, targetYear, targetMonth)

    ' **診療年月が取得できなかった場合は処理中止**
    If targetYear = "" Or targetMonth = "" Then
        MsgBox "CSVファイルから診療年月を取得できませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 2. 対象Excelファイルを取得（既存がなければ新規作成）
    targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)
    If targetFile = "" Then
        MsgBox "診療年月 " & targetYear & "年" & targetMonth & "月 のExcelファイルを作成できませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 3. 報告書Excelを開く
    On Error Resume Next
    Set newBook = Workbooks.Open(targetFile)
    On Error GoTo 0
    If newBook Is Nothing Then
        MsgBox "ファイル " & targetFile & " を開けませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 4. テンプレート情報を設定
    SetTemplateInfo newBook, targetYear, targetMonth

    ' 5. CSVファイルを順次処理（振込額明細書→返戻内訳書→増減点連絡書）
    ProcessAllCSVFiles fso, newBook, csvFolder, targetYear, targetMonth

    ' 6. 保存してブックを閉じる
    newBook.Save
    newBook.Close False

    ' 7. 完了メッセージ
    MsgBox "CSVファイルの処理が完了しました！", vbInformation, "完了"
End Sub

Function FindOrCreateReport(savePath As String, targetYear As String, targetMonth As String, templatePath As String) As String
    Dim fso As Object
    Dim filePath As String
    Dim fileName As String
    Dim existingFile As Object
    Dim newWb As Workbook
    Dim csvYYMM As String

    ' **診療年月をRYYMM形式に変換（令和年コード＋月）**
    Dim eraCode As String, eraYear As Integer
    If CInt(targetYear) >= 2019 Then
        eraCode = "5"                          ' 令和
        eraYear = CInt(targetYear) - 2018
    ElseIf CInt(targetYear) >= 1989 Then
        eraCode = "4"                          ' 平成
        eraYear = CInt(targetYear) - 1988
    ElseIf CInt(targetYear) >= 1926 Then
        eraCode = "3"                          ' 昭和
        eraYear = CInt(targetYear) - 1925
    ElseIf CInt(targetYear) >= 1912 Then
        eraCode = "2"                          ' 大正
        eraYear = CInt(targetYear) - 1911
    Else
        eraCode = "1"                          ' 明治（または不明）
        eraYear = CInt(targetYear) - 1867
    End If
    csvYYMM = Format(eraYear, "00") & targetMonth  ' 和暦年を2桁にして月を連結

    ' FileSystemObjectを作成
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

    ' テンプレートを開き、新規ファイルとして保存
    Set newWb = Workbooks.Open(templatePath)
    newWb.SaveAs filePath, FileFormat:=xlOpenXMLWorkbook  ' マクロ無しExcel形式で保存
    newWb.Close SaveChanges:=True

    ' 作成したファイルのパスを返す
    FindOrCreateReport = filePath
End Function

Sub SetTemplateInfo(newBook As Workbook, targetYear As String, targetMonth As String)
    Dim wsA As Worksheet, wsB As Worksheet
    Dim receiptYear As Integer, receiptMonth As Integer
    Dim sendMonth As Integer, sendDate As String

    ' **西暦年と調剤月を数値化**
    receiptYear = CInt(targetYear)
    receiptMonth = CInt(targetMonth)

    ' **請求月（翌月）の日付文字列を作成** （例: 調剤2月→請求月3月10日請求分）
    sendMonth = receiptMonth + 1
    If sendMonth = 13 Then sendMonth = 1   ' 12月の場合は翌年1月
    sendDate = sendMonth & "月10日請求分"

    ' **テンプレートシートA, Bを取得（名前「A」「B」でテンプレ内に存在）**
    Set wsA = newBook.Sheets("A")
    Set wsB = newBook.Sheets("B")

    ' **シート名を診療年月に合わせて変更**（例: wsA→"R7.2", wsB→"②"）
    wsA.Name = "R" & (receiptYear - 2018) & "." & receiptMonth    ' 令和year.month形式
    wsB.Name = ConvertToCircledNumber(receiptMonth)              ' 丸数字（調剤月）シート名

    ' **帳票内のタイトル情報を設定**
    wsA.Range("G2").Value = targetYear & "年" & targetMonth & "月調剤分"
    wsA.Range("I2").Value = sendDate
    wsA.Range("J2").Value = ThisWorkbook.Sheets(1).Range("B1").Value  ' 発行者情報（設定シートB1）
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

    ' 詳細データシート（2番目のシート）を取得
    Set wsDetails = newBook.Sheets(2)

    ' 処理対象年月の和暦GYYMMコードを生成
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
    GYYMM = eraCode & Format(eraYear, "00") & targetMonth  ' 例: 2025年02月→令和7年=07, GYYMM="50702"

    ' 種類別のCSVファイルコレクションを準備
    Dim fmeiFilesColl As New Collection
    Dim henrFilesColl As New Collection
    Dim zognFilesColl As New Collection

    ' フォルダ内のCSVファイルを種類ごとに振り分け
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" Then
            If InStr(LCase(csvFile.Name), "fmei") > 0 And Right(fso.GetBaseName(csvFile.Name), Len(GYYMM)) = GYYMM Then
                fmeiFilesColl.Add csvFile    ' 振込額明細書
            ElseIf InStr(LCase(csvFile.Name), "henr") > 0 And Right(fso.GetBaseName(csvFile.Name), Len(GYYMM)) = GYYMM Then
                henrFilesColl.Add csvFile    ' 返戻内訳書
            ElseIf InStr(LCase(csvFile.Name), "zogn") > 0 And Right(fso.GetBaseName(csvFile.Name), Len(GYYMM)) = GYYMM Then
                zognFilesColl.Add csvFile    ' 増減点連絡書
            End If
        End If
    Next csvFile

    ' 1) 振込額明細書（fmei）CSVを処理
    fileType = "振込額明細書"
    For Each csvFile In fmeiFilesColl
        sheetName = fso.GetBaseName(csvFile.Name)
        sheetName = GetUniqueSheetName(newBook, sheetName)   ' 一意なシート名に調整
        sheetIndex = Application.WorksheetFunction.Min(3, newBook.Sheets.Count + 1)
        Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(sheetIndex - 1))
        wsCSV.Name = sheetName

        ImportCSVData csvFile.Path, wsCSV, fileType
        TransferBillingDetails newBook, csvFile.Name    ' 入金明細反映（過去データの分類転記）
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
        TransferBillingDetails newBook, csvFile.Name    ' 返戻データ反映
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
        TransferBillingDetails newBook, csvFile.Name    ' 減点査定データ反映
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
    Dim fso As Object, folder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        IsFolderEmpty = True
        Exit Function
    End If
    Set folder = fso.GetFolder(folderPath)
    If folder.Files.Count = 0 Then
        IsFolderEmpty = True   ' ファイルが一つもない
    Else
        IsFolderEmpty = False  ' ファイルが存在する
    End If
End Function

Function GetTemplatePath() As String
    ' 設定シート（ThisWorkbook Sheets(1)）のB2セルからテンプレート格納パスを取得
    GetTemplatePath = ThisWorkbook.Sheets(1).Range("B2").Value & "\保険請求管理報告書テンプレート.xltm"
End Function

Function GetSavePath() As String
    ' 設定シートのB3セルから保存先フォルダパスを取得
    GetSavePath = ThisWorkbook.Sheets(1).Range("B3").Value
End Function

Function FindAllFixfFiles(fso As Object, csvFolder As String) As Collection
    Dim csvFile As Object
    Dim fixfFiles As New Collection

    ' フォルダ内の全ファイルから`fixf`を含むCSVを収集
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" And InStr(LCase(csvFile.Name), "fixf") > 0 Then
            fixfFiles.Add csvFile  ' `fixf`ファイルをコレクションに追加
        End If
    Next csvFile

    Set FindAllFixfFiles = fixfFiles
End Function

Sub GetYearMonthFromFixf(fixfFile As String, ByRef targetYear As String, ByRef targetMonth As String)
    Dim fileName As String
    Dim datePart As String
    Dim yearPart As String, monthPart As String

    ' 1. ファイル名から日時部分を取得（例: "RTfixf_..._20250228150730.csv" の "20250228150730"部分）
    fileName = Mid(fixfFile, InStrRev(fixfFile, "\") + 1)
    datePart = Mid(fileName, 18, 14)   ' ファイル名の18文字目から14桁がタイムスタンプ部分(YYYYMMDDhhmmss)

    ' 2. 年月日に分解
    yearPart = Left(datePart, 4)      ' "2025"
    monthPart = Mid(datePart, 5, 2)   ' "02"
    ' ※日・時刻部分は使用しない

    ' 3. 戻り値（診療年と月）をセット
    targetYear = yearPart    ' 西暦年 (例: "2025")
    targetMonth = monthPart  ' 月（2桁, 例: "02")
End Sub

Sub GetYearMonthFromCSV(fso As Object, csvFolder As String, ByRef targetYear As String, ByRef targetMonth As String)
    Dim csvFile As Object, ts As Object
    Dim lineText As String
    Dim era As String, yearPart As String, monthPart As String
    Dim westernYear As Integer

    ' フォルダ内の最初のCSVファイルからGYYMM形式の年月コードを取得する
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(Right(csvFile.Name, 4)) = ".csv" Then
            ' テキストファイルとして開く（読み取り専用, UTF-8）
            Set ts = fso.OpenTextFile(csvFile.Path, 1, False, -2)
            Do While Not ts.AtEndOfStream
                lineText = ts.ReadLine
                If Len(lineText) >= 5 Then
                    era = Left(lineText, 1)         ' 和暦元号コード (1:明治,2:大正,3:昭和,4:平成,5:令和)
                    yearPart = Mid(lineText, 2, 2)   ' 元号年（2桁）
                    monthPart = Right(lineText, 2)   ' 月（2桁）
                    ' 元号コードを西暦年に変換
                    Select Case era
                        Case "5": westernYear = 2018 + CInt(yearPart)  ' 令和 (2019年=令和元年)
                        Case "4": westernYear = 1988 + CInt(yearPart)  ' 平成 (1989年=平成元年)
                        Case "3": westernYear = 1925 + CInt(yearPart)  ' 昭和 (1926年=昭和元年)
                        Case "2": westernYear = 1911 + CInt(yearPart)  ' 大正 (1912年=大正元年)
                        Case "1": westernYear = 1867 + CInt(yearPart)  ' 明治 (1868年=明治元年)
                        Case Else: westernYear = 2000 + CInt(yearPart) ' その他（仮定）
                    End Select
                    targetYear = CStr(westernYear)
                    targetMonth = monthPart
                    Exit Do
                End If
            Loop
            ts.Close
            If targetYear <> "" And targetMonth <> "" Then Exit For
        End If
    Next csvFile
End Sub

Sub ImportCSVData(csvFile As String, ws As Worksheet, fileType As String)
    Dim colMap As Object
    Dim fso As Object, ts As Object
    Dim lineText As String
    Dim dataArray As Variant
    Dim i As Long, j As Long, key
    Dim isHeader As Boolean

    On Error GoTo ErrorHandler

    ' 画面更新と自動計算を一時停止
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' CSV項目のマッピング定義を取得
    Set colMap = GetColumnMapping(fileType)

    ' シートをクリアしてヘッダ行作成
    ws.Cells.Clear
    j = 1
    For Each key In colMap.Keys
        ws.Cells(1, j).Value = colMap(key)
        j = j + 1
    Next key

    ' CSVファイルを開き、内容を順次読み込んでシートに転記
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(csvFile, 1, False, -2)  ' UTF-8として開く
    i = 2
    isHeader = True
    Do While Not ts.AtEndOfStream
        lineText = ts.ReadLine
        dataArray = Split(lineText, ",")
        If isHeader Then
            ' 最初の行（ヘッダー行）は読み飛ばす
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
    MsgBox "CSVデータ読込中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
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
    Dim rowData As Variant
    Dim a As Long, b As Long, c As Long
    Dim csvYYMM As String

    ' シート設定： 請求データシート（1枚目）と詳細データシート（2枚目）
    Set wsBilling = newBook.Sheets(1)
    Set wsDetails = newBook.Sheets(2)

    ' **診療年月コード(csvYYMM)**を取得（請求データシートB2セルの下4桁YYMMを使用）
    csvYYMM = ""
    If wsBilling.Cells(2, 2).Value <> "" Then
        csvYYMM = Right(wsBilling.Cells(2, 2).Value, 4)
    End If

    ' 請求先区分の判定（ファイル名の7文字目で判断: "1"社保, "2"国保, その他は労災等）
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

    ' **開始行位置管理用Dictionaryの作成**（詳細シート上の各カテゴリ開始行を取得）
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

    ' **カテゴリ別Dictionaryの準備**
    Set rebillDict = CreateObject("Scripting.Dictionary")     ' 返戻再請求（過去返戻分の再請求）
    Set lateDict = CreateObject("Scripting.Dictionary")       ' 月遅れ請求（当月請求に含めた過去月分）
    Set unpaidDict = CreateObject("Scripting.Dictionary")     ' 未請求扱い（請求漏れ等未提出）
    Set assessmentDict = CreateObject("Scripting.Dictionary") ' 返戻・査定（返戻・減点により未収となった分）

    ' 請求データ（メインシート）の最終行を取得（D列に基準値があるものとする）
    lastRowBilling = wsBilling.Cells(wsBilling.Rows.Count, "D").End(xlUp).Row

    ' **請求データを走査し、過去月分を各Dictionaryに振り分け**
    For i = 2 To lastRowBilling
        dispensingMonth = wsBilling.Cells(i, 2).Value           ' GYYMM形式の調剤年月（例: "50701"）
        convertedMonth = ConvertToWesternDate(dispensingMonth)  ' "YY.MM"形式に変換（例: "07.01"）

        ' 対象診療月(csvYYMM)と異なる場合＝過去月のレセプトデータのみ処理対象
        If csvYYMM <> "" And Right(dispensingMonth, 4) <> csvYYMM Then
            ' 転記用データ（患者氏名, 調剤年月(YY.MM), 医療機関名, 請求点数）を配列に格納
            rowData = Array(wsBilling.Cells(i, 4).Value, convertedMonth, wsBilling.Cells(i, 5).Value, wsBilling.Cells(i, 10).Value)
            ' ファイル種別ごとに分類
            If InStr(LCase(csvFileName), "fixf") > 0 Then
                ' `fixf`ファイルでは過去月レセプトをすべて「月遅れ請求」として登録
                lateDict.Add wsBilling.Cells(i, 1).Value, rowData
            ElseIf InStr(LCase(csvFileName), "fmei") > 0 Then
                ' 入金明細では過去月レセプトを「返戻再請求」として登録（前月返戻分が当月入金されたケース）
                rebillDict.Add wsBilling.Cells(i, 1).Value, rowData
            ElseIf InStr(LCase(csvFileName), "zogn") > 0 Then
                ' 増減点連絡書では過去月レセプトを「未請求扱い」として登録（請求除外または未処理分）
                unpaidDict.Add wsBilling.Cells(i, 1).Value, rowData
            ElseIf InStr(LCase(csvFileName), "henr") > 0 Then
                ' 返戻内訳書では過去月レセプトを「返戻・査定」として登録（過去の査定残データ）
                assessmentDict.Add wsBilling.Cells(i, 1).Value, rowData
            End If
        End If
    Next i

    ' **各カテゴリの件数超過分を算出**（初期枠4件を超えた分だけ挿入用行数とする）
    a = 0: b = 0: c = 0
    If rebillDict.Count > 4 Then a = rebillDict.Count - 4
    If lateDict.Count > 4 Then b = lateDict.Count - 4
    If assessmentDict.Count > 4 Then c = assessmentDict.Count - 4
    ' 未請求扱い(unpaidDict)は必要行数算出対象外（将来の請求候補として残すため常時4行枠）

    ' **必要な追加行を挿入**（再請求・月遅れ・返戻査定の各セクション）
    If a + b + c > 0 Then
        If a > 0 Then wsDetails.Rows(startRowDict("月遅れ請求") + 1 & ":" & startRowDict("月遅れ請求") + a).Insert Shift:=xlDown
        If b > 0 Then wsDetails.Rows(startRowDict("返戻・査定") + 1 & ":" & startRowDict("返戻・査定") + b).Insert Shift:=xlDown
        If c > 0 Then wsDetails.Rows(startRowDict("未請求扱い") + 1 & ":" & startRowDict("未請求扱い") + c).Insert Shift:=xlDown
    End If

    ' **各Dictionaryのデータを詳細シートに転記**
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

    ' 完了メッセージ（処理区分ごと）
    MsgBox payerType & " のデータ転記が完了しました！", vbInformation, "処理完了"
End Sub

Sub TransferData(dataDict As Object, ws As Worksheet, startRow As Long, payerType As String)
    Dim key As Variant, rowData As Variant
    Dim r As Long
    Dim payerColumn As Long

    ' Dictionaryが空の場合は何もしない
    If dataDict.Count = 0 Then Exit Sub

    ' payerTypeに応じて転記する列を決定（社保→H列(8), 国保→I列(9)）
    If payerType = "社保" Then
        payerColumn = 8   ' H列
    ElseIf payerType = "国保" Then
        payerColumn = 9   ' I列
    Else
        Exit Sub  ' 労災等は詳細シート対象外
    End If

    r = startRow
    For Each key In dataDict.Keys
        rowData = dataDict(key)
        ws.Cells(r, 4).Value = rowData(0)            ' 患者氏名
        ws.Cells(r, 5).Value = rowData(1)            ' 調剤年月（YY.MM形式）
        ws.Cells(r, 6).Value = rowData(2)            ' 医療機関名
        ws.Cells(r, payerColumn).Value = payerType   ' 請求先種別（社保/国保）をセット
        ws.Cells(r, payerColumn).Font.Bold = True    ' 種別を強調表示
        ws.Cells(r, 10).Value = rowData(3)           ' 請求点数
        r = r + 1
    Next key
End Sub

Function GetColumnMapping(fileType As String) As Object
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")

    Select Case fileType
        Case "振込額明細書"    ' 支払基金からの振込額明細CSV
            colMap.Add 2, "診療（調剤）年月"
            colMap.Add 5, "受付番号"
            colMap.Add 14, "氏名"
            colMap.Add 16, "生年月日"
            colMap.Add 22, "医療保険_請求点数"
            colMap.Add 23, "医療保険_決定点数"
            colMap.Add 24, "医療保険_一部負担金"
            colMap.Add 25, "医療保険_金額"
            ' 第1～第5公費のデータ列（各10列間隔で請求点数・決定点数・患者負担金・金額）
            Dim k As Integer
            For k = 1 To 5
                colMap.Add 33 + (k - 1) * 10, "第" & k & "公費_請求点数"
                colMap.Add 34 + (k - 1) * 10, "第" & k & "公費_決定点数"
                colMap.Add 35 + (k - 1) * 10, "第" & k & "公費_患者負担金"
                colMap.Add 36 + (k - 1) * 10, "第" & k & "公費_金額"
            Next k
            colMap.Add 82, "算定額合計"

        Case "請求確定状況"    ' 請求確定CSV（`fixf`データ）
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

        Case "増減点連絡書"    ' 減点・増点連絡書CSV（査定減点等）
            colMap.Add 2, "調剤年月"
            colMap.Add 4, "受付番号"
            colMap.Add 11, "区分"
            colMap.Add 14, "老人減免区分"
            colMap.Add 15, "氏名"
            colMap.Add 21, "増減点数（金額）"
            colMap.Add 22, "事由"

        Case "返戻内訳書"      ' 返戻内訳書CSV（返戻レセプト詳細）
            colMap.Add 2, "調剤年月(YYMM)"
            colMap.Add 3, "受付番号"
            colMap.Add 4, "保険者番号"
            colMap.Add 7, "氏名"
            colMap.Add 9, "請求点数"
            colMap.Add 10, "薬剤一部負担金"
            colMap.Add 12, "一部負担金額"
            colMap.Add 13, "公費負担金額"
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
        Case "5": westernYear = 2018 + yearNum   ' 令和 (2019年=令和元年)
        Case "4": westernYear = 1988 + yearNum   ' 平成 (1989年=平成元年)
        Case "3": westernYear = 1925 + yearNum   ' 昭和 (1926年=昭和元年)
        Case "2": westernYear = 1911 + yearNum   ' 大正 (1912年=大正元年)
        Case "1": westernYear = 1867 + yearNum   ' 明治 (1868年=明治元年)
        Case Else: westernYear = 2000 + yearNum  ' その他（仮定）
    End Select
    ConvertToWesternDate = Right(westernYear, 2) & "." & monthPart  ' "YY.MM"形式の文字列を返す
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
    ' 指定カテゴリー名に完全一致するセルの行番号を取得
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(what:=categoryName, LookAt:=xlWhole)
    If Not foundCell Is Nothing Then
        GetStartRow = foundCell.Row
    Else
        GetStartRow = 0
    End If
End Function

Sub InvestigateHalfYearDiscrepancy()
    Dim yearStr As String, halfStr As String
    Dim yearNum As Integer, half As Integer
    Dim startMonth As Integer, endMonth As Integer
    Dim fso As Object, folderPath As String
    Dim month As Integer
    Dim fileName As String, filePath As String
    Dim wb As Workbook, wsMain As Worksheet, wsDep As Worksheet
    Dim totalClaim As Long, totalDecided As Long
    Dim eraCode As String, eraYear As Integer, eraYY As String
    Dim resultMsg As String

    ' ユーザーに対象年と半期区分を入力させる
    yearStr = InputBox("調査する年（西暦）を入力してください:", "半期請求誤差調査")
    If yearStr = "" Then Exit Sub
    halfStr = InputBox("上期(1) または 下期(2) を指定してください:", "半期請求誤差調査")
    If halfStr = "" Then Exit Sub
    If Not IsNumeric(yearStr) Or Not IsNumeric(halfStr) Then
        MsgBox "入力が正しくありません。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    yearNum = CInt(yearStr)
    half = CInt(halfStr)
    If half <> 1 And half <> 2 Then
        MsgBox "半期の指定が不正です。1（上期）または2（下期）を指定してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    ' 半期の開始月と終了月を設定（上期:1-6月, 下期:7-12月）
    If half = 1 Then
        startMonth = 1: endMonth = 6
    Else
        startMonth = 7: endMonth = 12
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = GetSavePath()
    If folderPath = "" Then
        MsgBox "保存フォルダが設定されていません。", vbExclamation, "エラー"
        Exit Sub
    End If

    resultMsg = "【" & yearNum & "年 " & IIf(half = 1, "上期", "下期") & " 請求誤差調査】" & vbCrLf

    ' 指定期間の各月について、請求点数と決定点数の差異を集計
    For month = startMonth To endMonth
        ' 対象ファイル名を構築（和暦年コードを算出）
        If yearNum >= 2019 Then
            eraCode = "5"  ' 令和
            eraYear = yearNum - 2018
        ElseIf yearNum >= 1989 Then
            eraCode = "4"  ' 平成
            eraYear = yearNum - 1988
        ElseIf yearNum >= 1926 Then
            eraCode = "3"  ' 昭和
            eraYear = yearNum - 1925
        ElseIf yearNum >= 1912 Then
            eraCode = "2"  ' 大正
            eraYear = yearNum - 1911
        Else
            eraCode = "1"  ' 明治
            eraYear = yearNum - 1867
        End If
        eraYY = Format(eraYear, "00")  ' 和暦年2桁
        fileName = "保険請求管理報告書_R" & eraYY & Format(month, "00") & ".xlsx"
        filePath = folderPath & "\" & fileName

        If fso.FileExists(filePath) Then
            Set wb = Workbooks.Open(filePath, ReadOnly:=True)
            Set wsMain = wb.Sheets(1)
            wsMain.Activate  ' メインシート（請求確定状況データ）

            ' 当月請求の総請求点数合計を算出（メインシートの「総合計点数」列合計）
            Dim totalPointsClaim As Long, totalPointsDecided As Long
            totalPointsClaim = 0: totalPointsDecided = 0
            Dim hdrCell As Range, colClaim As Long, colDec As Long

            Set hdrCell = wsMain.Rows(1).Find("総合計点数", LookAt:=xlWhole)
            If Not hdrCell Is Nothing Then
                colClaim = hdrCell.Column
                Dim lastRow As Long
                lastRow = wsMain.Cells(wsMain.Rows.Count, colClaim).End(xlUp).Row
                If lastRow >= 2 Then
                    totalPointsClaim = Application.WorksheetFunction.Sum(wsMain.Range(wsMain.Cells(2, colClaim), wsMain.Cells(lastRow, colClaim)))
                End If
            End If

            ' 当月の決定点数合計を算出（振込額明細シートの各決定点数列の合計）
            ' 振込額明細シートを特定（ヘッダ行に"決定点数"を含むシートを探す）
            Set wsDep = Nothing
            Dim ws As Worksheet, foundHdr As Range
            For Each ws In wb.Sheets
                Set foundHdr = ws.Rows(1).Find("決定点数", LookAt:=xlPart)
                If Not foundHdr Is Nothing Then
                    If LCase(ws.Name) <> LCase(wsMain.Name) And LCase(ws.Name) <> LCase(wb.Sheets(2).Name) Then
                        Set wsDep = ws
                        Exit For
                    End If
                End If
            Next ws
            If Not wsDep Is Nothing Then
                ' 決定点数列（社保および公費）を合計
                Dim col As Long
                For col = 1 To wsDep.UsedRange.Columns.Count
                    If InStr(wsDep.Cells(1, col).Value, "決定点数") > 0 Then
                        Dim lastRowDep As Long
                        lastRowDep = wsDep.Cells(wsDep.Rows.Count, col).End(xlUp).Row
                        If lastRowDep >= 2 Then
                            totalPointsDecided = totalPointsDecided + Application.WorksheetFunction.Sum(wsDep.Range(wsDep.Cells(2, col), wsDep.Cells(lastRowDep, col)))
                        End If
                    End If
                Next col
            End If

            wb.Close SaveChanges:=False

            ' 差異算出
            Dim diffPoints As Long
            diffPoints = totalPointsClaim - totalPointsDecided
            If diffPoints <> 0 Then
                resultMsg = resultMsg & "・" & yearNum & "年" & month & "月: 請求点数=" & totalPointsClaim & " , 決定点数=" & totalPointsDecided & " （差異 " & diffPoints & " 点）" & vbCrLf
            End If
        Else
            ' ファイルがない場合
            resultMsg = resultMsg & "・" & yearNum & "年" & month & "月: 報告書未作成" & vbCrLf
        End If
    Next month

    MsgBox resultMsg, vbInformation, "半期ごとの請求誤差調査結果"
End Sub