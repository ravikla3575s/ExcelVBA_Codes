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
    
    ' 【1】CSVフォルダをユーザーに選択させる
    csvFolder = SelectCSVFolder()
    If csvFolder = "" Then Exit Sub

    ' 【1.1】フォルダが空なら処理を中止
    If IsFolderEmpty(csvFolder) Then
        MsgBox "選択したフォルダにはCSVファイルがありません。処理を中止します。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 【2】テンプレートパス・保存フォルダ取得
    templatePath = GetTemplatePath()
    savePath = GetSavePath()
    If templatePath = "" Or savePath = "" Then Exit Sub

    ' 【3】ファイルシステムオブジェクトの作成
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 【4】フォルダ内のすべての `fixf` ファイルを取得
    Set fixfFiles = FindAllFixfFiles(fso, csvFolder)
    
    ' 【5】`fixfFiles` が `Nothing` の場合、または `fixfFiles.Count = 0` の場合、通常のCSV処理に切り替え
    If fixfFiles Is Nothing Or fixfFiles.Count = 0 Then
        Call ProcessWithoutFixf(fso, csvFolder, savePath, templatePath)
        Exit Sub
    End If

    ' 【6】複数の `fixf` を順番に処理
    For Each file In fixfFiles
        fixfFile = file.Path
        
        ' 【7】対象年月を取得（変数を初期化）
        targetYear = ""
        targetMonth = ""
        Call GetYearMonthFromFixf(fixfFile, targetYear, targetMonth)
        
        ' **対象年月が取得できなかった場合はスキップ**
        If targetYear = "" Or targetMonth = "" Then
            MsgBox "ファイル " & fixfFile & " から診療年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If

        ' 【8】対象Excelファイルを取得 or 作成
        targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)
        
        ' **ファイルが取得できなかった場合はスキップ**
        If targetFile = "" Then
            MsgBox "診療年月 " & targetYear & "年" & targetMonth & "月 のExcelファイルを作成できませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If
        
        ' 【9】Excelを開く
        On Error Resume Next
        Set newBook = Workbooks.Open(targetFile)
        If newBook Is Nothing Then
            MsgBox "ファイル " & targetFile & " を開けませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If
        On Error GoTo 0

        ' 【10】テンプレート情報を設定
        SetTemplateInfo newBook, targetYear, targetMonth
        
        ' 【11】フォルダ内のすべてのCSVを処理
        ProcessAllCSVFiles fso, newBook, csvFolder
        
        ' 【12】保存して閉じる
        newBook.Save
        newBook.Close

NextFile:
    Next file

    ' 【13】処理完了メッセージ
    MsgBox "すべての `fixf` ファイルを処理しました！", vbInformation
End Sub

Function SelectCSVFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "CSVフォルダを選択してください"
        If .Show = -1 Then
            SelectCSVFolder = .SelectedItems(1) & "\"
        Else
            MsgBox "フォルダが選択されませんでした。処理を中止します。", vbExclamation
            SelectCSVFolder = ""
        End If
    End With
End Function

Function IsFolderEmpty(folderPath As String) As Boolean
    Dim fso As Object
    Dim folder As Object
    
    ' **FileSystemObject を作成**
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' **フォルダが存在しない場合は True を返す**
    If Not fso.FolderExists(folderPath) Then
        IsFolderEmpty = True
        Exit Function
    End If

    ' **フォルダ内のファイル数をチェック**
    Set folder = fso.GetFolder(folderPath)
    If folder.Files.Count = 0 Then
        IsFolderEmpty = True ' **ファイルがない場合は True**
    Else
        IsFolderEmpty = False ' **ファイルがある場合は False**
    End If
End Function

Function GetTemplatePath() As String
    GetTemplatePath = ThisWorkbook.Sheets(1).Range("B2").Value & "\保険請求管理報告書テンプレート20250222.xltm"
End Function

Function GetSavePath() As String
    GetSavePath = ThisWorkbook.Sheets(1).Range("B3").Value
End Function

Function FindAllFixfFiles(fso As Object, csvFolder As String) As Collection
    Dim csvFile As Object
    Dim fixfFiles As New Collection

    ' **フォルダ内のすべてのファイルをループ**
    For Each csvFile In fso.GetFolder(csvFolder).Files
        ' **拡張子が "csv" であり、名前に "fixf" を含む場合**
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" And InStr(LCase(csvFile.Name), "fixf") > 0 Then
            fixfFiles.Add csvFile ' **fixf ファイルをコレクションに追加**
        End If
    Next csvFile

    ' **コレクションを返す**
    Set FindAllFixfFiles = fixfFiles
End Function

Sub GetYearMonthFromFixf(fixfFile As String, ByRef targetYear As String, ByRef targetMonth As String)
    Dim fileName As String
    Dim datePart As String
    Dim yearPart As String, monthPart As String, dayPart As String
    Dim hourPart As String, minPart As String, secPart As String
    
    ' 【1】ファイル名を取得（フォルダパスを除く）
    fileName = Mid(fixfFile, InStrRev(fixfFile, "\") + 1)

    ' 【2】fixfファイルの日付部分を取得（後半部分）
    ' 例: RTfixf1014123456720250228150730.csv → "20250228150730"
    datePart = Mid(fileName, 18, 14)

    ' 【3】年月日を分解
    yearPart = Left(datePart, 4)    ' "2025"
    monthPart = Mid(datePart, 5, 2) ' "02"
    dayPart = Mid(datePart, 7, 2)   ' "28"
    hourPart = Mid(datePart, 9, 2)  ' "15"
    minPart = Mid(datePart, 11, 2)  ' "07"
    secPart = Mid(datePart, 13, 2)  ' "30"

    ' 【4】取得した年と月を戻り値に設定
    targetYear = yearPart  ' 西暦のまま
    targetMonth = monthPart ' 2桁の月

    ' **確認用ログ（必要なら表示）**
    ' MsgBox "診療年月取得: " & targetYear & "年 " & targetMonth & "月 " & dayPart & "日 " & hourPart & ":" & minPart & ":" & secPart, vbInformation, "確認"
End Sub

Sub GetYearMonthFromCSV(fso As Object, csvFolder As String, ByRef targetYear As String, ByRef targetMonth As String)
    Dim csvFile As Object
    Dim ts As Object
    Dim lineText As String
    Dim yearPart As String, monthPart As String
    Dim era As String, westernYear As Integer
    
    ' フォルダ内の最初の CSV を探す
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(Right(csvFile.Name, 4)) = ".csv" Then
            ' CSV を開く
            Set ts = fso.OpenTextFile(csvFile.Path, 1, False, -2) ' UTF-8対応

            ' 1行目または2行目から GYYMM を取得（例: "50406" = 令和4年6月）
            Do While Not ts.AtEndOfStream
                lineText = ts.ReadLine
                If Len(lineText) >= 5 Then
                    era = Left(lineText, 1) ' 和暦の元号 (1: 明治, 2: 大正, 3: 昭和, 4: 平成, 5: 令和)
                    yearPart = Mid(lineText, 2, 2) ' 2桁の年
                    monthPart = Right(lineText, 2) ' 2桁の月

                    ' 和暦を西暦に変換
                    Select Case era
                        Case "1": westernYear = 1867 + CInt(yearPart) ' 明治
                        Case "2": westernYear = 1911 + CInt(yearPart) ' 大正
                        Case "3": westernYear = 1925 + CInt(yearPart) ' 昭和
                        Case "4": westernYear = 1988 + CInt(yearPart) ' 平成
                        Case "5": westernYear = 2018 + CInt(yearPart) ' 令和
                        Case Else: westernYear = 2000 ' 不明な場合は適当なデフォルト値
                    End Select

                    ' 取得した年月をセット
                    targetYear = CStr(westernYear)
                    targetMonth = monthPart
                    Exit Sub
                End If
            Loop

            ' ファイルを閉じる
            ts.Close
        End If
    Next csvFile
End Sub

Function FindOrCreateReport(savePath As String, targetYear As String, targetMonth As String, templatePath As String) As String
    Dim fso As Object
    Dim filePath As String
    Dim fileName As String
    Dim existingFile As Object
    Dim newWb As Workbook
    Dim csvYYMM As String
    
    ' **診療年月を "YYMM" 形式に変換**
    csvYYMM = Cstr(Format(CInt(targetYear) - 2018, "00")) & targetMonth

    ' **FileSystemObject の作成**
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' **フォルダ内の報告書 (`RYYMM.xlsx`) を検索**
    For Each existingFile In fso.GetFolder(savePath).Files
        If LCase(fso.GetExtensionName(existingFile.Name)) = "xlsx" Then
            If Right(fso.GetBaseName(existingFile.Name), 4) = csvYYMM Then
                ' **診療年月が一致するファイルが見つかった場合、そのパスを返す**
                FindOrCreateReport = existingFile.Path
                Exit Function
            End If
        End If
    Next existingFile

    ' **該当するファイルがなければ、新規作成**
    fileName = "保険請求管理報告書_R" & csvYYMM & ".xlsx"
    filePath = savePath & "\" & fileName

    ' **テンプレートを元に新規作成**
    Set newWb = Workbooks.Open(templatePath)
    newWb.SaveAs filePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    
    ' **ブックを閉じてパスを返す**
    newWb.Close SaveChanges:=True
    FindOrCreateReport = filePath
End Function

Sub SetTemplateInfo(newBook As Workbook, targetYear As String, targetMonth As String)
    Dim wsTemplate As Worksheet, wsTemplate2 As Worksheet
    Dim receiptYear As Integer, receiptMonth As Integer
    Dim sendMonth As Integer, sendDate As String

    ' **西暦年と調剤月の計算**
    receiptYear = CInt(targetYear)
    receiptMonth = CInt(targetMonth)

    ' **請求月の計算**
    sendMonth = receiptMonth + 1
    If sendMonth = 13 Then sendMonth = 1
    sendDate = sendMonth & "月10日請求分"

    ' **シートA, Bを取得**
    Set wsTemplate = newBook.Sheets("A")
    Set wsTemplate2 = newBook.Sheets("B")

    ' **シート名変更**
    wsTemplate.Name = "R" & (receiptYear - 2018) & "." & receiptMonth
    wsTemplate2.Name = ConvertToCircledNumber(receiptMonth)

    ' **情報転記**
    wsTemplate.Range("G2").Value = targetYear & "年" & targetMonth & "月調剤分"
    wsTemplate.Range("I2").Value = sendDate
    wsTemplate.Range("J2").Value = ThisWorkbook.Sheets(1).Range("B1").Value
    wsTemplate2.Range("H1").Value = targetYear & "年" & targetMonth & "月調剤分"
    wsTemplate2.Range("J1").Value = sendDate
    wsTemplate2.Range("L1").Value = ThisWorkbook.Sheets(1).Range("B1").Value
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

    ' CSVファイルをループ処理
    For Each csvFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFile.Name)) = "csv" Then
            fileType = ""

            ' **CSVの種類を判別**
            Select Case True
                Case InStr(csvFile.Name, "fmei") > 0
                    fileType = "振込額明細書"
                Case InStr(csvFile.Name, "zogn") > 0
                    fileType = "増減点連絡書"
                Case InStr(csvFile.Name, "henr") > 0
                    fileType = "返戻内訳書"
                Case Else
                    GoTo NextFile
            End Select

            ' **シート名にファイル名（拡張子なし）を設定**
            sheetName = fso.GetBaseName(csvFile.Name)

            ' **シートがすでに存在する場合、"_1", "_2" を付けて回避**
            sheetName = GetUniqueSheetName(newBook, sheetName)

            ' **シートを3番目（Sheets(3)の位置）に追加**
            sheetIndex = Application.WorksheetFunction.Min(3, newBook.Sheets.Count + 1)
            Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(sheetIndex - 1))
            wsCSV.Name = sheetName

            ' **CSVデータを転記**
            ImportCSVData(csvFile.Path, wsCSV, fileType)

            ' **請求確定状況の詳細データをシート2に転記**
            TransferBillingDetails(newBook, csvFile.Name)

NextFile:
        End If
    Next csvFile
End Sub

Function ConvertToCircledNumber(month As Integer) As String
    Dim circledNumbers As Variant
    circledNumbers = Array("①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫")

    ' **1～12月の範囲内なら変換、範囲外ならそのまま返す**
    If month >= 1 And month <= 12 Then
        ConvertToCircledNumber = circledNumbers(month - 1)
    Else
        ConvertToCircledNumber = CStr(month) ' 予期しない値ならそのまま
    End If
End Function

Function GetUniqueSheetName(wb As Workbook, baseName As String) As String
    Dim newName As String
    Dim counter As Integer
    Dim ws As Worksheet
    Dim exists As Boolean

    newName = baseName
    counter = 1

    ' **同じ名前のシートが存在するか確認**
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

' CSVデータを転記
Sub ImportCSVData(csvFile As String, ws As Worksheet, fileType As String)
    Dim colMap As Object
    Dim fso As Object
    Dim ts As Object
    Dim lineText As String
    Dim dataArray As Variant
    Dim i As Long, j As Long, key
    Dim isHeader As Boolean

    ' エラーハンドリング
    On Error GoTo ErrorHandler

    ' 画面更新・計算を一時停止
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 項目マッピングを取得
    Set colMap = GetColumnMapping(fileType)

    ' シートをクリア
    ws.Cells.Clear

    ' 1行目に項目名を転記
    j = 1
    For Each key In colMap.Keys
        ws.Cells(1, j).Value = colMap(key)
        j = j + 1
    Next key

    ' CSVデータを読み込んで転記
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(csvFile, 1, False, -2) ' UTF-8対応（-2）

    ' データを転記
    i = 2
    isHeader = True ' 最初の行はヘッダー行
    Do While Not ts.AtEndOfStream
        lineText = ts.ReadLine
        dataArray = Split(lineText, ",")

        ' ヘッダー行をスキップ
        If isHeader Then
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

    ' 画面更新・計算を再開
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    Exit Sub

' エラーハンドリング
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    If Not ts Is Nothing Then ts.Close
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub TransferBillingDetails(newBook As Workbook, sheetName As String, csvFileName As String)
    Dim wsBilling As Worksheet, wsDetails As Worksheet
    Dim lastRowBilling As Long, lastRowDetails As Long
    Dim i As Long, j As Long
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
    Set wsBilling = newBook.Sheets(1) ' **請求データが格納されているシート**
    Set wsDetails = newBook.Sheets(2) ' **詳細用シート**

    ' **診療年月 (csvYYMM) を取得**
    csvYYMM = Right(wsBilling.Cells(2, 2).Value, 4) ' GYYMM の YYMM を取得

    ' CSVデータの請求先分類
    payerCode = Mid(sheetName, 7, 1)
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
    Set rebillDict = CreateObject("Scripting.Dictionary")   ' 返戻再請求
    Set lateDict = CreateObject("Scripting.Dictionary")     ' 月遅れ請求
    Set unpaidDict = CreateObject("Scripting.Dictionary")   ' 未請求扱い
    Set assessmentDict = CreateObject("Scripting.Dictionary") ' 返戻・査定

    lastRowBilling = wsBilling.Cells(Rows.Count, "D").End(xlUp).Row

    ' **請求データを Dictionary に格納**
    For i = 2 To lastRowBilling
        dispensingMonth = wsBilling.Cells(i, 2).Value ' **GYYMM形式の診療月**
        convertedMonth = ConvertToWesternDate(dispensingMonth)
        rowData = Array(wsBilling.Cells(i, 4).Value, convertedMonth, wsBilling.Cells(i, 5).Value, wsBilling.Cells(i, 10).Value)

        ' **現在処理中の診療月（csvYYMM）と異なる場合のみ追加**
        If Right(dispensingMonth, 4) <> csvYYMM Then
            ' **CSVの種類で振り分け**
            If InStr(csvFileName, "fixf") > 0 Then
                ' fixf → ユーザーに選択させる
                If ShowRebillSelectionForm(rowData) Then
                    rebillDict.Add wsBilling.Cells(i, 1).Value, rowData ' **返戻再請求**
                Else
                    lateDict.Add wsBilling.Cells(i, 1).Value, rowData ' **月遅れ請求**
                End If
            ElseIf InStr(csvFileName, "zogn") > 0 Then
                unpaidDict.Add wsBilling.Cells(i, 1).Value, rowData ' **未請求扱い**
            ElseIf InStr(csvFileName, "henr") > 0 Then
                assessmentDict.Add wsBilling.Cells(i, 1).Value, rowData ' **返戻・査定**
            End If
        End If
    Next i

    ' **各カテゴリの追加行数を計算**
    a = 0: b = 0: c = 0
    If rebillDict.Count > 4 Then a = rebillDict.Count - 4
    If lateDict.Count > 4 Then b = lateDict.Count - 4
    If assessmentDict.Count > 4 Then c = assessmentDict.Count - 4

    ' **各転記開始行の調整**
    Dim lateStartRow As Long, assessmentStartRow As Long, unpaidStartRow As Long

    lateStartRow = startRowDict("月遅れ請求") + 1 + a
    assessmentStartRow = startRowDict("返戻・査定") + 1 + a + b
    unpaidStartRow = startRowDict("未請求扱い") + 1 + a + b + c

    ' **行を追加**
    If a + b + c > 0 Then
        wsDetails.Rows(lateStartRow & ":" & lateStartRow + a).Insert Shift:=xlDown
        wsDetails.Rows(assessmentStartRow & ":" & assessmentStartRow + b).Insert Shift:=xlDown
        wsDetails.Rows(unpaidStartRow & ":" & unpaidStartRow + c).Insert Shift:=xlDown
    End If

    ' **各 Dictionary の転記処理（空ならスキップ）**
    If rebillDict.Count > 0 Then
        j = startRowDict("返戻再請求")
        Call TransferData(rebillDict, wsDetails, j, payerType)
    End If

    If lateDict.Count > 0 Then
        j = startRowDict("月遅れ請求")
        Call TransferData(lateDict, wsDetails, j, payerType)
    End If

    If unpaidDict.Count > 0 Then
        j = startRowDict("未請求扱い")
        Call TransferData(unpaidDict, wsDetails, j, payerType)
    End If

    If assessmentDict.Count > 0 Then
        j = startRowDict("返戻・査定")
        Call TransferData(assessmentDict, wsDetails, j, payerType)
    End If

    MsgBox payerType & " のデータ転記が完了しました！", vbInformation, "処理完了"
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
            
            ' 公費データ（第一〜第五）
            Dim i As Integer
            For i = 1 To 5
                colMap.Add 33 + (i - 1) * 10, "第" & i & "公費_請求点数"
                colMap.Add 34 + (i - 1) * 10, "第" & i & "公費_決定点数"
                colMap.Add 35 + (i - 1) * 10, "第" & i & "公費_患者負担金"
                colMap.Add 36 + (i - 1) * 10, "第" & i & "公費_金額"
            Next i

            colMap.Add 82, "算定額合計"

        Case "請求確定状況"
            colMap.Add 4, "診療（調剤）年月"
            colMap.Add 5, "氏名"
            colMap.Add 7, "生年月日"
            colMap.Add 9, "医療機関名称"
            colMap.Add 13, "総合計点数"

            ' 公費データ（第一〜第四）
            For i = 1 To 4
                colMap.Add 16 + (i - 1) * 3, "第" & i & "公費_請求点数"
            Next i

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

Function ConvertToWesternDate(dispensingMonth As String) As String
    Dim era As String, yearPart As Integer, westernYear As Integer, monthPart As String
    
    ' GYYMM 形式から元号と年月を取得
    era = Left(dispensingMonth, 1) ' 例: "5"（令和）
    yearPart = Mid(dispensingMonth, 2, 2) ' 例: "06"
    monthPart = Right(dispensingMonth, 2) ' 例: "06"

    ' 和暦を西暦に変換
    Select Case era
        Case "5": westernYear = 2018 + yearPart ' 令和（2019年開始）
        ' 他の元号（明治/大正/昭和/平成）は未対応
    End Select

    ' 変換結果（YY.MM）
    ConvertToWesternDate = Right(westernYear, 2) & "." & monthPart
End Function

Sub ShowRebillSelectionForm(newBook As Workbook)
    Dim wsBilling As Worksheet
    Dim lastRow As Long, i As Long
    Dim userForm As Object
    Dim listData As Object
    Dim rowData As Variant
    
    ' メインシート取得
    Set wsBilling = newBook.Sheets(1)
    lastRow = wsBilling.Cells(Rows.Count, "D").End(xlUp).Row

    ' Dictionary でリストを管理
    Set listData = CreateObject("Scripting.Dictionary")

    ' 現在の請求月取得
    Dim currentBillingMonth As String
    currentBillingMonth = wsBilling.Cells(2, 2).Value ' GYYMM

    ' 該当調剤月以外のデータをリスト化
    For i = 2 To lastRow
        If wsBilling.Cells(i, 2).Value <> currentBillingMonth Then
            rowData = Array(wsBilling.Cells(i, 2).Value, wsBilling.Cells(i, 4).Value, wsBilling.Cells(i, 5).Value, wsBilling.Cells(i, 6).Value)
            listData.Add i, rowData
        End If
    Next i

    ' リストにデータがあればフォーム表示
    If listData.Count > 0 Then
        Set userForm = CreateRebillSelectionForm(listData)
        userForm.Show
    Else
        MsgBox "該当するデータはありません。", vbInformation, "確認"
    End If
End Sub

Sub TransferData(dataDict As Object, ws As Worksheet, startRow As Long, payerType As String)
    Dim key As Variant, rowData As Variant
    Dim j As Long
    Dim payerColumn As Long

    ' **Dictionary が空なら処理しない**
    If dataDict.Count = 0 Then Exit Sub

    ' **payerType に応じた転記列を決定**
    If payerType = "社保" Then
        payerColumn = 8 ' 社保の請求先は H列（8列目）
    ElseIf payerType = "国保" Then
        payerColumn = 9 ' 国保の請求先は I列（9列目）
    Else
        Exit Sub ' 労災の場合は処理しない
    End If

    j = startRow
    For Each key In dataDict.Keys
        rowData = dataDict(key)
        ws.Cells(j, 4).Value = rowData(0) ' 患者氏名
        ws.Cells(j, 5).Value = rowData(1) ' 調剤年月
        ws.Cells(j, 6).Value = rowData(2) ' 医療機関名
        ws.Cells(j, payerColumn).Value = payerType ' 請求先（社保 or 国保）
        ws.Cells(j, payerColumn).Font.Bold = True ' **強調**
        ws.Cells(j, 10).Value = rowData(3) ' 請求点数
        j = j + 1
    Next key
End Sub

Sub CreateRebillSelectionForm(listData As Object)
    Dim uf As Object
    Dim listBox As Object
    Dim btnOK As Object
    Dim i As Long
    Dim rowData As Variant

    ' **UserForm を作成**
    Set uf = CreateObject("Forms.UserForm")
    uf.Caption = "返戻再請求の選択"
    uf.Width = 400
    uf.Height = 500

    ' **ListBox を追加**
    Set listBox = uf.Controls.Add("Forms.ListBox.1", "listBox", True)
    listBox.Left = 20
    listBox.Top = 20
    listBox.Width = 350
    listBox.Height = 350
    listBox.MultiSelect = 1 ' **複数選択可能**

    ' **リストデータ追加**
    For i = 0 To listData.Count - 1
        rowData = listData.Items()(i)
        listBox.AddItem rowData(0) & " | " & rowData(1) & " | " & rowData(2) & " | " & rowData(3)
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

    ' **フォームを表示**
    uf.Show
End Sub

Sub ProcessRebillSelection()
    Dim wsDetails As Worksheet
    Dim listBox As Object
    Dim selectedData As Object
    Dim i As Long, rowIndex As Long
    Dim startRowDict As Object
    Dim category As String
    Dim insertRows As Long
    Dim uf As Object

    ' 転記用ワークシート取得
    Set wsDetails = ThisWorkbook.Sheets(2) ' 詳細用シート
    Set uf = UserForm1 ' UserForm1 を明示的に取得
    Set listBox = uf.Controls("listBox")

    ' 選択データを格納
    Set selectedData = CreateObject("Scripting.Dictionary")

    ' 開始行管理用 Dictionary 作成
    Set startRowDict = CreateObject("Scripting.Dictionary")
    startRowDict.Add "社保返戻再請求", GetStartRow(wsDetails, "社保返戻再請求")
    startRowDict.Add "国保返戻再請求", GetStartRow(wsDetails, "国保返戻再請求")
    startRowDict.Add "社保月遅れ請求", GetStartRow(wsDetails, "社保月遅れ請求")
    startRowDict.Add "国保月遅れ請求", GetStartRow(wsDetails, "国保月遅れ請求")

    ' **選択された項目を取得**
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) Then
            selectedData.Add i, listBox.List(i)
        End If
    Next i

    ' **5行以上ある場合は、行を追加**
    insertRows = selectedData.Count - 4
    If insertRows > 0 Then
        wsDetails.Rows(startRowDict("社保返戻再請求") + 1 & ":" & startRowDict("社保返戻再請求") + insertRows).Insert Shift:=xlDown
    End If

    ' **データ転記**
    category = "社保返戻再請求"
    rowIndex = startRowDict(category)
    
    For Each i In selectedData.Keys
        wsDetails.Cells(rowIndex, 5).Value = selectedData(i) ' 調剤年月
        wsDetails.Cells(rowIndex, 6).Value = selectedData(i) ' 患者氏名
        wsDetails.Cells(rowIndex, 7).Value = selectedData(i) ' 医療機関名
        wsDetails.Cells(rowIndex, 10).Value = selectedData(i) ' 請求点数
        rowIndex = rowIndex + 1
    Next i

    ' UserForm を閉じる
    Unload uf

    MsgBox "転記が完了しました！", vbInformation, "処理完了"
End Sub