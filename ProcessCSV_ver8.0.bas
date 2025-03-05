Option Explicit

' **請求CSV一括処理マクロ:** 
' 指定フォルダ内の請求確定CSV(`fixf`)および各種明細CSV(`fmei`:振込額明細, `henr`:返戻内訳, `zogn`:増減点連絡)を読み込み、
' 月次の「保険請求管理報告書」Excelを作成・更新します。
' 処理後、報告書Excel（名称: 保険請求管理報告書_RYYMM.xlsx）が指定フォルダに出力されます。

Sub ProcessCSV()
    Dim csvFolder As String            ' CSVフォルダパス
    Dim fso As Object                  ' FileSystemObject
    Dim targetYear As String, targetMonth As String  ' 処理対象の診療年・月（西暦）
    Dim savePath As String            ' 報告書保存先フォルダ
    Dim templatePath As String        ' 報告書テンプレートファイル(.xltm)パス
    Dim reportWb As Workbook          ' 報告書Excelブック
    Dim fixfFiles As Collection, fmeiFiles As Collection, henrFiles As Collection, zognFiles As Collection
    Dim fileObj As Object
    Dim reportFilePath As String, reportFileName As String
    Dim eraLetter As String, yearCode As String, monthCode As String
    Dim fixfFile As String, sheetName As String
    Dim i As Long

    ' ... (前略: フォルダ選択やテンプレートパス取得などの処理) ...

    ' 1. CSVフォルダをユーザーに選択させる
    csvFolder = SelectCSVFolder()
    If csvFolder = "" Then Exit Sub  ' ユーザーがキャンセルした場合

    ' 1.1 フォルダが空なら処理を中止
    If IsFolderEmpty(csvFolder) Then
        MsgBox "選択したフォルダにはCSVファイルがありません。処理を中止します。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 2. テンプレートパス・保存先フォルダを取得
    templatePath = GetTemplatePath()    ' 設定シートのB2セル（テンプレート格納先）
    savePath = GetSavePath()           ' 設定シートのB3セル（保存先フォルダ）
    If templatePath = "" Or savePath = "" Then Exit Sub  ' 必須パスが取得できなければ中止

    ' 3. FileSystemObjectの用意
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 4. フォルダ内の全CSVファイルを種類別に収集（fixf, fmei, henr, zogn）
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fixfFiles = New Collection: Set fmeiFiles = New Collection
    Set henrFiles = New Collection: Set zognFiles = New Collection
    For Each fileObj In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(fileObj.Name)) = "csv" Then
            If InStr(LCase(fileObj.Name), "fixf") > 0 Then
                fixfFiles.Add fileObj
            ElseIf InStr(LCase(fileObj.Name), "fmei") > 0 Then
                fmeiFiles.Add fileObj
            ElseIf InStr(LCase(fileObj.Name), "henr") > 0 Then
                henrFiles.Add fileObj
            ElseIf InStr(LCase(fileObj.Name), "zogn") > 0 Then
                zognFiles.Add fileObj
            End If
        End If
    Next fileObj

    ' 4.1 対象CSVファイルが一つもない場合、処理を中止
    If fixfFiles.Count = 0 And fmeiFiles.Count = 0 And henrFiles.Count = 0 And zognFiles.Count = 0 Then
        MsgBox "選択したフォルダには処理対象のCSVファイルがありません。処理を中止します。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 5. 収集したCSVファイルを調剤年月が古い順にソート
    If fixfFiles.Count > 1 Then
        For i = 1 To fixfFiles.Count - 1
            For j = i + 1 To fixfFiles.Count
                Dim year1 As String, month1 As String
                Dim year2 As String, month2 As String
                year1 = "": month1 = ""
                year2 = "": month2 = ""
                GetYearMonthFromFixf fixfFiles(i).Path, year1, month1
                GetYearMonthFromFixf fixfFiles(j).Path, year2, month2
                If year1 <> "" And month1 <> "" And year2 <> "" And month2 <> "" Then
                    If (year1 & month1) > (year2 & month2) Then
                        Set tempObj = fixfFiles(i)
                        Set fixfFiles(i) = fixfFiles(j)
                        Set fixfFiles(j) = tempObj
                    End If
                End If
            Next j
        Next i
    End If
    If fmeiFiles.Count > 1 Then
        For i = 1 To fmeiFiles.Count - 1
            For j = i + 1 To fmeiFiles.Count
                Dim code1 As String, code2 As String
                code1 = Right(fso.GetBaseName(fmeiFiles(i).Name), 5)
                code2 = Right(fso.GetBaseName(fmeiFiles(j).Name), 5)
                If Len(code1) = 5 And Len(code2) = 5 And IsNumeric(code1) And IsNumeric(code2) Then
                    If CDbl(code1) > CDbl(code2) Then
                        Set tempObj = fmeiFiles(i)
                        Set fmeiFiles(i) = fmeiFiles(j)
                        Set fmeiFiles(j) = tempObj
                    End If
                End If
            Next j
        Next i
    End If
    If henrFiles.Count > 1 Then
        For i = 1 To henrFiles.Count - 1
            For j = i + 1 To henrFiles.Count
                code1 = Right(fso.GetBaseName(henrFiles(i).Name), 5)
                code2 = Right(fso.GetBaseName(henrFiles(j).Name), 5)
                If Len(code1) = 5 And Len(code2) = 5 And IsNumeric(code1) And IsNumeric(code2) Then
                    If CDbl(code1) > CDbl(code2) Then
                        Set tempObj = henrFiles(i)
                        Set henrFiles(i) = henrFiles(j)
                        Set henrFiles(j) = tempObj
                    End If
                End If
            Next j
        Next i
    End If
    If zognFiles.Count > 1 Then
        For i = 1 To zognFiles.Count - 1
            For j = i + 1 To zognFiles.Count
                code1 = Right(fso.GetBaseName(zognFiles(i).Name), 5)
                code2 = Right(fso.GetBaseName(zognFiles(j).Name), 5)
                If Len(code1) = 5 And Len(code2) = 5 And IsNumeric(code1) And IsNumeric(code2) Then
                    If CDbl(code1) > CDbl(code2) Then
                        Set tempObj = zognFiles(i)
                        Set zognFiles(i) = zognFiles(j)
                        Set zognFiles(j) = tempObj
                    End If
                End If
            Next j
        Next i
    End If

    ' 6. 請求確定CSV（fixf）の処理
    For Each fileObj In fixfFiles
        Dim treatmentYear as integer , treatmentMonth as integer
        treatmentYear = "":treatmentMonth = ""
        fixfFile = fileObj.Path
        targetYear = "": targetMonth = ""
        GetYearMonthFromFixf fixfFile, targetYear, targetMonth  ' fixfファイルから診療年月を取得
        If targetYear = "" Or targetMonth = "" Then
            MsgBox "ファイル " & fixfFile & " から診療年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextFixf
        End If
        ' 請求年月の1ヶ月前を診療年月とする（月を1減算）
        treatmentMonth = targetMonth - 1
        If treatmentMonth = 0 Then
        treatmentYear = targetYear - 1
        treatmentMonth = 12
        End If
        ' 出力報告書ファイル名（診療年月RYYMM形式）を決定
        If CInt(treatmentYear) >= 2019 Then
            eraLetter = "R": yearCode = Format(CInt(treatmentYear) - 2018, "00")   ' 令和
        ElseIf CInt(treatmentYear) >= 1989 Then
            eraLetter = "H": yearCode = Format(CInt(treatmentYear) - 1988, "00")   ' 平成
        ElseIf CInt(treatmentYear) >= 1926 Then
            eraLetter = "S": yearCode = Format(CInt(treatmentYear) - 1925, "00")   ' 昭和
        ElseIf CInt(treatmentYear) >= 1912 Then
            eraLetter = "T": yearCode = Format(CInt(treatmentYear) - 1911, "00")   ' 大正
        Else
            eraLetter = "M": yearCode = Format(CInt(treatmentYear) - 1867, "00")   ' 明治
        End If
        reportFileName = "保険請求管理報告書_" & eraLetter & yearCode & Format(CInt(treatmentMonth), "00") & ".xlsm"
        reportFilePath = savePath & "\" & reportFileName
        ' **既存の報告書ファイルがある場合、重複処理を避けてスキップ**
        If fso.FileExists(reportFilePath) Then
            GoTo NextFixf
        End If
        ' 報告書ブックを取得（存在しなければテンプレートから新規作成）
        If Not fso.FileExists(reportFilePath) Then
            Dim newWb As Workbook
            Set newWb = Workbooks.Add(templatePath)
            If newWb Is Nothing Then
                MsgBox "テンプレートファイルを開けませんでした: " & templatePath, vbExclamation, "エラー"
                GoTo NextFixf
            End If
            Application.DisplayAlerts = False
            newWb.SaveAs Filename:=reportFilePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            Application.DisplayAlerts = True
            newWb.Close False
        End If
        On Error Resume Next
        Set reportWb = Workbooks.Open(reportFilePath)
        On Error GoTo 0
        If reportWb Is Nothing Then
            MsgBox "ファイル " & reportFilePath & " を開けませんでした。", vbExclamation, "エラー"
            GoTo NextFixf
        End If
        ' テンプレート情報を設定（診療年月等を更新）
        SetTemplateInfo reportWb, targetYear, targetMonth

        ' fixf CSVデータをシート3～6に転記（シートがなければ作成）
        Dim neededSheets As Integer: neededSheets = 6
        If reportWb.Sheets.Count < neededSheets Then
            For i = reportWb.Sheets.Count + 1 To neededSheets
                reportWb.Sheets.Add After:=reportWb.Sheets(reportWb.Sheets.Count)
            Next i
        End If
        ' 対象シート3～6をクリア
        For i = 3 To 6
            If i <= reportWb.Sheets.Count Then reportWb.Sheets(i).Cells.Clear
        Next i
        ' fixf CSVファイルを読み込み、カテゴリ別にデータ行を振り分け
        Dim ts As Object, lineText As String, lines() As String
        Set ts = fso.OpenTextFile(fixfFile, 1, False, -2)  ' UTF-8で読み込み
        lineText = ts.ReadAll: ts.Close
        lines = Split(lineText, vbCrLf)
        Dim colMap As Object: Set colMap = GetColumnMapping("請求確定状況")
        Dim dataLinesCat1 As New Collection, dataLinesCat2 As New Collection
        ' 最初の行はヘッダ行とみなす
        Dim headerSkipped As Boolean: headerSkipped = False
        For i = LBound(lines) To UBound(lines)
            If Trim(lines(i)) = "" Then Continue For
            If Not headerSkipped Then
                headerSkipped = True
                Continue For  ' ヘッダ行をスキップ
            End If
            If Left(lines(i), 2) = "1," Then
                dataLinesCat1.Add lines(i)
            ElseIf Left(lines(i), 2) = "2," Then
                dataLinesCat2.Add lines(i)
            Else
                ' 想定外の行は無視
            End If
        Next i
        ' データ出力のヘルパーサブルーチン（指定シートにヘッダ＋指定行集合を書き込む）
        Dim key As Variant
        Sub WriteDataToSheet(ws As Worksheet, dataCol As Collection)
            ws.Cells.Clear
            ' ヘッダ行を書き込み
            Dim j As Long: j = 1
            For Each key In colMap.Keys
                ws.Cells(1, j).Value = colMap(key)
                j = j + 1
            Next key
            ' データ行を書き込み
            Dim rowIndex As Long: rowIndex = 2
            Dim arr As Variant
            For j = 1 To dataCol.Count
                arr = Split(dataCol(j), ",")
                Dim k As Long: k = 1
                For Each key In colMap.Keys
                    If key - 1 <= UBound(arr) Then
                        ws.Cells(rowIndex, k).Value = Trim(arr(key - 1))
                    End If
                    k = k + 1
                Next key
                rowIndex = rowIndex + 1
            Next j
        End Sub
        ' 1ページあたり最大行数（必要に応じて調整）
        Dim maxLinesPerSheet As Long: maxLinesPerSheet = 40

        ' カテゴリ1（社保）データの転記
        If dataLinesCat1.Count > 0 Then
            If dataLinesCat1.Count <= maxLinesPerSheet Then
                WriteDataToSheet reportWb.Sheets(3), dataLinesCat1
            Else
                ' 1ページに収まらない場合、Sheet3とSheet4に分割
                Dim tmpCol As New Collection
                For i = 1 To maxLinesPerSheet
                    tmpCol.Add dataLinesCat1(i)
                Next i
                WriteDataToSheet reportWb.Sheets(3), tmpCol
                tmpCol.Remove 1 To tmpCol.Count  ' コレクションをクリア
                For i = maxLinesPerSheet + 1 To dataLinesCat1.Count
                    tmpCol.Add dataLinesCat1(i)
                Next i
                WriteDataToSheet reportWb.Sheets(4), tmpCol
            End If
        End If
        ' カテゴリ2（国保）データの転記
        If dataLinesCat2.Count > 0 Then
            Dim startSheet As Integer
            startSheet = IIf(dataLinesCat1.Count > 0, 5, 3)  ' 社保ありならSheet5開始、なければSheet3開始
            If dataLinesCat2.Count <= maxLinesPerSheet Then
                WriteDataToSheet reportWb.Sheets(startSheet), dataLinesCat2
            Else
                ' 2ページに分割
                Dim tmpCol2 As New Collection
                For i = 1 To maxLinesPerSheet
                    tmpCol2.Add dataLinesCat2(i)
                Next i
                WriteDataToSheet reportWb.Sheets(startSheet), tmpCol2
                tmpCol2.Remove 1 To tmpCol2.Count
                For i = maxLinesPerSheet + 1 To dataLinesCat2.Count
                    tmpCol2.Add dataLinesCat2(i)
                Next i
                ' 次のシート（カテゴリ開始シートの+1）に続き出力
                WriteDataToSheet reportWb.Sheets(startSheet + 1), tmpCol2
            End If
        End If

        ' ブックを保存して閉じる
        reportWb.Save
        reportWb.Close False
NextFixf:
        ' 次のfixfファイルへ
    Next fileObj

    ' 7. 振込額明細CSV（fmei）の処理
    For Each fileObj In fmeiFiles
        targetYear = "": targetMonth = ""
        ' ファイル名から診療年月コードを取得
        Dim ymCode As String
        ymCode = GetTreatmentYearMonthFromFileName(fileObj.Name)
        If ymCode = "" Then
            MsgBox "ファイル " & fileObj.Name & " から診療年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextFmei
        End If
        ' 診療年月（西暦）を算出
        eraLetter = Left(ymCode, 1)
        yearCode = Mid(ymCode, 2, 2)
        monthCode = Right(ymCode, 2)
        Select Case eraLetter
            Case "R": targetYear = CStr(2018 + CInt(yearCode))   ' 令和
            Case "H": targetYear = CStr(1988 + CInt(yearCode))   ' 平成
            Case "S": targetYear = CStr(1925 + CInt(yearCode))   ' 昭和
            Case "T": targetYear = CStr(1911 + CInt(yearCode))   ' 大正
            Case "M": targetYear = CStr(1867 + CInt(yearCode))   ' 明治
            Case Else: targetYear = CStr(2000 + CInt(yearCode))  ' 仮（不明コード）
        End Select
        targetMonth = monthCode
        ' 報告書ファイル名を決定（存在しなければ作成）
        reportFileName = "保険請求管理報告書_" & ymCode & ".xlsm"
        reportFilePath = savePath & "\" & reportFileName
        If Not fso.FileExists(reportFilePath) Then
            Dim newWb2 As Workbook
            Set newWb2 = Workbooks.Add(templatePath)
            If newWb2 Is Nothing Then
                MsgBox "テンプレートファイルを開けませんでした: " & templatePath, vbExclamation, "エラー"
                GoTo NextFmei
            End If
            Application.DisplayAlerts = False
            newWb2.SaveAs Filename:=reportFilePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            Application.DisplayAlerts = True
            newWb2.Close False
        End If
        On Error Resume Next
        Set reportWb = Workbooks.Open(reportFilePath)
        On Error GoTo 0
        If reportWb Is Nothing Then
            MsgBox "ファイル " & reportFilePath & " を開けませんでした。", vbExclamation, "エラー"
            GoTo NextFmei
        End If
        ' 既に同じファイル名のシートが存在する場合はスキップ
        sheetName = fso.GetBaseName(fileObj.Name)
        Dim ws As Worksheet, sheetExists As Boolean: sheetExists = False
        For Each ws In reportWb.Sheets
            If InStr(ws.Name, sheetName) > 0 Then
                sheetExists = True
                Exit For
            End If
        Next ws
        If sheetExists Then
            reportWb.Close False
            GoTo NextFmei
        End If
        ' テンプレート情報を設定
        SetTemplateInfo reportWb, targetYear, targetMonth
        ' CSVデータを新規シートにインポート
        Dim insertIndex As Long
        insertIndex = Application.WorksheetFunction.Min(3, reportWb.Sheets.Count + 1)
        Dim wsCSV As Worksheet
        Set wsCSV = reportWb.Sheets.Add(After:=reportWb.Sheets(insertIndex - 1))
        wsCSV.Name = sheetName
        ImportCSVData fileObj.Path, wsCSV, "振込額明細書"
        TransferBillingDetails reportWb, fileObj.Name  ' 振込額明細の詳細転記処理

        reportWb.Save
        reportWb.Close False
NextFmei:
        ' 次のfmeiファイルへ
    Next fileObj

    ' 8. 返戻内訳CSV（henr）の処理
    For Each fileObj In henrFiles
        targetYear = "": targetMonth = ""
        ymCode = GetTreatmentYearMonthFromFileName(fileObj.Name)
        If ymCode = "" Then
            MsgBox "ファイル " & fileObj.Name & " から診療年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextHenr
        End If
        eraLetter = Left(ymCode, 1)
        yearCode = Mid(ymCode, 2, 2)
        monthCode = Right(ymCode, 2)
        Select Case eraLetter
            Case "R": targetYear = CStr(2018 + CInt(yearCode))
            Case "H": targetYear = CStr(1988 + CInt(yearCode))
            Case "S": targetYear = CStr(1925 + CInt(yearCode))
            Case "T": targetYear = CStr(1911 + CInt(yearCode))
            Case "M": targetYear = CStr(1867 + CInt(yearCode))
            Case Else: targetYear = CStr(2000 + CInt(yearCode))
        End Select
        targetMonth = monthCode
        reportFileName = "保険請求管理報告書_" & ymCode & ".xlsm"
        reportFilePath = savePath & "\" & reportFileName
        If Not fso.FileExists(reportFilePath) Then
            Dim newWb3 As Workbook
            Set newWb3 = Workbooks.Add(templatePath)
            If newWb3 Is Nothing Then
                MsgBox "テンプレートファイルを開けませんでした: " & templatePath, vbExclamation, "エラー"
                GoTo NextHenr
            End If
            Application.DisplayAlerts = False
            newWb3.SaveAs Filename:=reportFilePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            Application.DisplayAlerts = True
            newWb3.Close False
        End If
        On Error Resume Next
        Set reportWb = Workbooks.Open(reportFilePath)
        On Error GoTo 0
        If reportWb Is Nothing Then
            MsgBox "ファイル " & reportFilePath & " を開けませんでした。", vbExclamation, "エラー"
            GoTo NextHenr
        End If
        sheetName = fso.GetBaseName(fileObj.Name)
        sheetExists = False
        For Each ws In reportWb.Sheets
            If InStr(ws.Name, sheetName) > 0 Then
                sheetExists = True: Exit For
            End If
        Next ws
        If sheetExists Then
            reportWb.Close False
            GoTo NextHenr
        End If
        SetTemplateInfo reportWb, targetYear, targetMonth
        insertIndex = Application.WorksheetFunction.Min(3, reportWb.Sheets.Count + 1)
        Set wsCSV = reportWb.Sheets.Add(After:=reportWb.Sheets(insertIndex - 1))
        wsCSV.Name = sheetName
        ImportCSVData fileObj.Path, wsCSV, "返戻内訳書"
        TransferBillingDetails reportWb, fileObj.Name  ' 返戻内訳の詳細転記処理

        reportWb.Save
        reportWb.Close False
NextHenr:
        ' 次のhenrファイルへ
    Next fileObj

    ' 9. 増減点連絡CSV（zogn）の処理
    For Each fileObj In zognFiles
        targetYear = "": targetMonth = ""
        ymCode = GetTreatmentYearMonthFromFileName(fileObj.Name)
        If ymCode = "" Then
            MsgBox "ファイル " & fileObj.Name & " から診療年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextZogn
        End If
        eraLetter = Left(ymCode, 1)
        yearCode = Mid(ymCode, 2, 2)
        monthCode = Right(ymCode, 2)
        Select Case eraLetter
            Case "R": targetYear = CStr(2018 + CInt(yearCode))
            Case "H": targetYear = CStr(1988 + CInt(yearCode))
            Case "S": targetYear = CStr(1925 + CInt(yearCode))
            Case "T": targetYear = CStr(1911 + CInt(yearCode))
            Case "M": targetYear = CStr(1867 + CInt(yearCode))
            Case Else: targetYear = CStr(2000 + CInt(yearCode))
        End Select
        targetMonth = monthCode
        reportFileName = "保険請求管理報告書_" & ymCode & ".xlsm"
        reportFilePath = savePath & "\" & reportFileName
        If Not fso.FileExists(reportFilePath) Then
            Dim newWb4 As Workbook
            Set newWb4 = Workbooks.Add(templatePath)
            If newWb4 Is Nothing Then
                MsgBox "テンプレートファイルを開けませんでした: " & templatePath, vbExclamation, "エラー"
                GoTo NextZogn
            End If
            Application.DisplayAlerts = False
            newWb4.SaveAs Filename:=reportFilePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            Application.DisplayAlerts = True
            newWb4.Close False
        End If
        On Error Resume Next
        Set reportWb = Workbooks.Open(reportFilePath)
        On Error GoTo 0
        If reportWb Is Nothing Then
            MsgBox "ファイル " & reportFilePath & " を開けませんでした。", vbExclamation, "エラー"
            GoTo NextZogn
        End If
        sheetName = fso.GetBaseName(fileObj.Name)
        sheetExists = False
        For Each ws In reportWb.Sheets
            If InStr(ws.Name, sheetName) > 0 Then
                sheetExists = True: Exit For
            End If
        Next ws
        If sheetExists Then
            reportWb.Close False
            GoTo NextZogn
        End If
        SetTemplateInfo reportWb, targetYear, targetMonth
        insertIndex = Application.WorksheetFunction.Min(3, reportWb.Sheets.Count + 1)
        Set wsCSV = reportWb.Sheets.Add(After:=reportWb.Sheets(insertIndex - 1))
        wsCSV.Name = sheetName
        ImportCSVData fileObj.Path, wsCSV, "増減点連絡書"
        TransferBillingDetails reportWb, fileObj.Name  ' 増減点連絡の詳細転記処理

        reportWb.Save
        reportWb.Close False
NextZogn:
        ' 次のzognファイルへ
    Next fileObj

    ' 10. 完了メッセージ
    MsgBox "CSVファイルの処理が完了しました！", vbInformation, "完了"

End Sub

Sub ProcessWithoutFixf(fso As Object, csvFolder As String, savePath As String, templatePath As String)
    Dim targetYear As String, targetMonth As String
    Dim reportFile As String
    Dim reportWb As Workbook

    ' 1. フォルダ内の最初のCSVから診療年月を推定（fmei等の先頭行GYYMMコードを利用）
    targetYear = ""
    targetMonth = ""
    ' fixfファイルがなく報告書ファイルも存在しない場合、fmeiファイル名から診療年月を推定
    Dim fmeiFile As Object
    For Each fmeiFile In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(fmeiFile.Name)) = "csv" And InStr(LCase(fmeiFile.Name), "fmei") > 0 Then
            ' 最初に見つかったfmeiファイルを使用
            Dim codePart As String, eraCode As String, yearCode As String, monthCode As String
            Dim westernYear As Integer
            codePart = Right(fso.GetBaseName(fmeiFile.Name), 5)
            If Len(codePart) = 5 And IsNumeric(codePart) Then
                eraCode = Left(codePart, 1)
                yearCode = Mid(codePart, 2, 2)
                monthCode = Right(codePart, 2)
                Select Case eraCode
                    Case "5": westernYear = 2018 + CInt(yearCode)   ' 令和
                    Case "4": westernYear = 1988 + CInt(yearCode)   ' 平成
                    Case "3": westernYear = 1925 + CInt(yearCode)   ' 昭和
                    Case "2": westernYear = 1911 + CInt(yearCode)   ' 大正
                    Case "1": westernYear = 1867 + CInt(yearCode)   ' 明治
                    Case Else: westernYear = 2000 + CInt(yearCode)  ' 仮定
                End Select
                targetYear = CStr(westernYear)
                targetMonth = monthCode
            End If
            Exit For
        End If
    Next fmeiFile
    ' ファイル名から取得できなかった場合、CSV内容から診療年月を取得
    If targetYear = "" Or targetMonth = "" Then
        GetYearMonthFromCSV fso, csvFolder, targetYear, targetMonth 
    End If
    ' **診療年月が取得できなかった場合は処理中止**
    If targetYear = "" Or targetMonth = "" Then
        MsgBox "CSVファイルから診療年月を取得できませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 2. 報告書Excelファイルを取得（既存がなければ新規作成）
    reportFile = FindOrCreateReport savePath, targetYear, targetMonth, templatePath 
    If reportFile = "" Then
        MsgBox "診療年月 " & targetYear & "年" & targetMonth & "月 のExcelファイルを作成できませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 3. 報告書Excelを開く
    On Error Resume Next
    Set reportWb = Workbooks.Open(reportFile)
    On Error GoTo 0
    If reportWb Is Nothing Then
        MsgBox "ファイル " & reportFile & " を開けませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 4. テンプレート情報を設定（タイトル等）
    SetTemplateInfo reportWb, targetYear, targetMonth

    ' 5. CSVファイルを種類別に処理（fixfなしなので、振込明細・返戻・増減点のみ）
    ProcessAllCSVFiles fso, reportWb, csvFolder, targetYear, targetMonth

    ' 6. 保存してブックを閉じる
    reportWb.Save
    reportWb.Close False

    ' 7. 完了メッセージ
    MsgBox "CSVファイルの処理が完了しました！", vbInformation, "完了"
End Sub

Sub ProcessAllCSVFiles(fso As Object, reportWb As Workbook, csvFolder As String, targetYear As String, targetMonth As String)
    Dim eraCode As String, eraYear As Integer
    Dim GYYMM As String          ' 和暦元号コード付の対象年月 (例:50702)
    Dim csvFileObj As Object
    ' 受け取りCSVの種類別コレクションを用意
    Dim fmeiFiles As New Collection, henrFiles As New Collection, zognFiles As New Collection

    ' 対象年月を和暦GYYMM形式に変換（例: 2025年02月→令和7年=07 ⇒ "50702"）
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
        eraCode = "1"  ' 明治（想定外の場合）
        eraYear = CInt(targetYear) - 1867
    End If
    GYYMM = eraCode & Format(eraYear, "00") & targetMonth   ' 例: "50702"

    ' フォルダ内の全CSVファイルを走査し、ファイル名により種類別に振り分け
    For Each csvFileObj In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(csvFileObj.Name)) = "csv" Then
            Dim baseName As String
            baseName = fso.GetBaseName(csvFileObj.Name)
            ' ファイル名に種別キーワードを含み、かつ末尾のYYMMコードが対象年月かチェック
            If InStr(LCase(baseName), "fmei") > 0 And Right(baseName, Len(GYYMM)) = GYYMM Then
                fmeiFiles.Add csvFileObj    ' 振込額明細書CSVを収集
            ElseIf InStr(LCase(baseName), "henr") > 0 And Right(baseName, Len(GYYMM)) = GYYMM Then
                henrFiles.Add csvFileObj    ' 返戻内訳書CSVを収集
            ElseIf InStr(LCase(baseName), "zogn") > 0 And Right(baseName, Len(GYYMM)) = GYYMM Then
                zognFiles.Add csvFileObj    ' 増減点連絡書CSVを収集
            End If
        End If
    Next csvFileObj

    ' 1) 振込額明細書（fmei）CSVの処理
    ProcessFmeiFiles fso, reportWb, fmeiFiles

    ' 2) 返戻内訳書（henr）CSVの処理
    ProcessHenrFiles fso, reportWb, henrFiles

    ' 3) 増減点連絡書（zogn）CSVの処理
    ProcessZognFiles fso, reportWb, zognFiles
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

Sub ProcessFmeiFiles(fso As Object, reportWb As Workbook, fmeiFiles As Collection)
    Dim csvFileObj As Object, wsCSV As Worksheet
    Dim sheetName As String, insertIndex As Integer

    For Each csvFileObj In fmeiFiles
        ' 新しいシートを追加し、一意なシート名を設定（既存重複回避）
        sheetName = fso.GetBaseName(csvFileObj.Name)
        sheetName = GetUniqueSheetName(reportWb, sheetName)
        insertIndex = Application.WorksheetFunction.Min(3, reportWb.Sheets.Count + 1)
        Set wsCSV = reportWb.Sheets.Add(After:=reportWb.Sheets(insertIndex - 1))
        wsCSV.Name = sheetName

        ' CSVデータをインポートし転記（列マッピングは"振込額明細書"定義を使用）
        ImportCSVData csvFileObj.Path, wsCSV, "振込額明細書"
        ' 当該データの詳細分類転記（過去月入金＝返戻再請求の検出等）
        TransferBillingDetails reportWb, csvFileObj.Name
    Next csvFileObj
End Sub

Sub ProcessHenrFiles(fso As Object, reportWb As Workbook, henrFiles As Collection)
    Dim csvFileObj As Object, wsCSV As Worksheet
    Dim sheetName As String, insertIndex As Integer

    For Each csvFileObj In henrFiles
        sheetName = fso.GetBaseName(csvFileObj.Name)
        sheetName = GetUniqueSheetName(reportWb, sheetName)
        insertIndex = Application.WorksheetFunction.Min(3, reportWb.Sheets.Count + 1)
        Set wsCSV = reportWb.Sheets.Add(After:=reportWb.Sheets(insertIndex - 1))
        wsCSV.Name = sheetName

        ImportCSVData csvFileObj.Path, wsCSV, "返戻内訳書"
        ' 返戻データ（過去未収＝返戻・査定）の詳細シート反映
        TransferBillingDetails reportWb, csvFileObj.Name
    Next csvFileObj
End Sub

Sub ProcessZognFiles(fso As Object, reportWb As Workbook, zognFiles As Collection)
    Dim csvFileObj As Object, wsCSV As Worksheet
    Dim sheetName As String, insertIndex As Integer

    For Each csvFileObj In zognFiles
        sheetName = fso.GetBaseName(csvFileObj.Name)
        sheetName = GetUniqueSheetName(reportWb, sheetName)
        insertIndex = Application.WorksheetFunction.Min(3, reportWb.Sheets.Count + 1)
        Set wsCSV = reportWb.Sheets.Add(After:=reportWb.Sheets(insertIndex - 1))
        wsCSV.Name = sheetName

        ImportCSVData csvFileObj.Path, wsCSV, "増減点連絡書"
        ' 減点（未請求扱い）データの詳細シート反映
        TransferBillingDetails reportWb, csvFileObj.Name
    Next csvFileObj
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
    Dim fso As Object, folderObj As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        IsFolderEmpty = True
        Exit Function
    End If
    Set folderObj = fso.GetFolder(folderPath)
    If folderObj.Files.Count = 0 Then
        IsFolderEmpty = True   ' ファイルが一つもない
    Else
        IsFolderEmpty = False  ' ファイルが存在する
    End If
End Function

Function GetTemplatePath() As String
    GetTemplatePath = ThisWorkbook.Sheets(1).Range("B2").Value & "\保険請求管理報告書テンプレート20250222.xltm"
End Function

Function GetSavePath() As String
    GetSavePath = ThisWorkbook.Sheets(1).Range("B3").Value
End Function

Function FindAllFixfFiles(fso As Object, csvFolder As String) As Collection
    Dim fileObj As Object
    Dim fixfFiles As New Collection
    ' フォルダ内の全ファイルから名前に`fixf`を含むCSVを収集
    For Each fileObj In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(fileObj.Name)) = "csv" And InStr(LCase(fileObj.Name), "fixf") > 0 Then
            fixfFiles.Add fileObj
        End If
    Next fileObj
    Set FindAllFixfFiles = fixfFiles
End Function

Function FindOrCreateReport(savePath As String, targetYear As String, targetMonth As String, templatePath As String) As String
    Dim fso As Object, newWb As Workbook
    Dim reportPath As String, reportName As String
    Dim csvYYMM As String, eraCode As String, eraYear As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' RYYMM形式のファイル名を生成
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
        eraCode = "1"  ' 明治
        eraYear = CInt(targetYear) - 1867
    End If
    csvYYMM = eraCode & Format(eraYear, "00") & targetMonth  ' RYYMM文字列
    reportName = "保険請求管理報告書_" & csvYYMM & ".xlsm"
    reportPath = savePath & "\" & reportName
    If Not fso.FileExists(reportPath) Then
        On Error Resume Next
        Set newWb = Workbooks.Add(templatePath)
        On Error GoTo 0
        If newWb Is Nothing Then
            FindOrCreateReport = ""
            Exit Function
        End If
        Application.DisplayAlerts = False
        newWb.SaveAs Filename:=reportPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Application.DisplayAlerts = True
        newWb.Close False
    End If
    If fso.FileExists(reportPath) Then
        FindOrCreateReport = reportPath
    Else
        FindOrCreateReport = ""  ' 作成に失敗した場合
    End If
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

Sub GetYearMonthFromFixf(fixfFilePath As String, ByRef targetYear As String, ByRef targetMonth As String)
    Dim fileName As String, datePart As String
    Dim yearStr As String, monthStr As String
    ' ファイルパスからファイル名部分を取得
    fileName = Mid(fixfFilePath, InStrRev(fixfFilePath, "\") + 1)
    ' ファイル名中のタイムスタンプ部分(YYYYMMDDhhmmss)を抽出 (例: "..._20250228150730.csv"から"20250228150730")
    datePart = Mid(fileName, 18, 14)
    If Len(datePart) < 8 Then Exit Sub
    ' 年月に分解
    yearStr = Left(datePart, 4)    ' "2025"
    monthStr = Mid(datePart, 5, 2) ' "02"
    ' 結果を戻り値にセット
    targetYear = yearStr
    targetMonth = monthStr
End Sub

Sub GetYearMonthFromCSV(fso As Object, csvFolder As String, ByRef targetYear As String, ByRef targetMonth As String)
    Dim fileObj As Object, ts As Object
    Dim lineText As String
    Dim eraCode As String, yearCode As String, monthCode As String
    Dim westernYear As Integer

    ' フォルダ内のCSVファイルから先頭行のGYYMMコードを取得（対象ファイル以外はスキップ）
    For Each fileObj In fso.GetFolder(csvFolder).Files
        If LCase(fso.GetExtensionName(fileObj.Name)) = "csv" _
           And (InStr(LCase(fileObj.Name), "fixf") > 0 _
                Or InStr(LCase(fileObj.Name), "fmei") > 0 _
                Or InStr(LCase(fileObj.Name), "henr") > 0 _
                Or InStr(LCase(fileObj.Name), "zogn") > 0) Then
            Set ts = fso.OpenTextFile(fileObj.Path, 1, False, -2)  ' テキストストリーム (読み取り専用, UTF-8)
            Do While Not ts.AtEndOfStream
                lineText = ts.ReadLine
                If Len(lineText) >= 5 Then
                    eraCode = Left(lineText, 1)        ' 元号コード (1:明治,2:大正,3:昭和,4:平成,5:令和)
                    yearCode = Mid(lineText, 2, 2)     ' 元号年（2桁）
                    monthCode = Right(lineText, 2)     ' 月（2桁）
                    ' 元号コード＋年 を西暦年に変換
                    Select Case eraCode
                        Case "5": westernYear = 2018 + CInt(yearCode)   ' 令和 (2019=令和元年)
                        Case "4": westernYear = 1988 + CInt(yearCode)   ' 平成 (1989=平成元年)
                        Case "3": westernYear = 1925 + CInt(yearCode)   ' 昭和 (1926=昭和元年)
                        Case "2": westernYear = 1911 + CInt(yearCode)   ' 大正 (1912=大正元年)
                        Case "1": westernYear = 1867 + CInt(yearCode)   ' 明治 (1868=明治元年)
                        Case Else: westernYear = 2000 + CInt(yearCode)  ' 仮定
                    End Select
                    targetYear = CStr(westernYear)
                    targetMonth = monthCode
                    Exit Do   ' 必要な情報取得できたのでループ終了
                End If
            Loop
            ts.Close
            If targetYear <> "" And targetMonth <> "" Then Exit For
        End If
    Next fileObj
End Sub

Sub ImportCSVData(csvFilePath As String, ws As Worksheet, fileType As String)
    Dim fso As Object, ts As Object
    Dim colMap As Object          ' 列マッピング定義（Dictionary）
    Dim lineText As String
    Dim dataArray As Variant
    Dim i As Long, j As Long, key

    On Error GoTo ImportError
    ' 画面更新と計算を一時停止（パフォーマンス向上）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 1. CSV項目のマッピング定義を取得（fileTypeに応じた列マッピング辞書）
    Set colMap = GetColumnMapping(fileType)

    ' 2. 対象シートをクリアし、ヘッダー行を作成
    ws.Cells.Clear
    j = 1
    For Each key In colMap.Keys
        ws.Cells(1, j).Value = colMap(key)  ' マッピング定義の値＝ヘッダ名
        j = j + 1
    Next key

    ' 3. CSVファイルを開いて読み込み、データ部をシートに転記
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(csvFilePath, 1, False, -2)  ' UTF-8でテキストストリーム開く
    i = 2
    Dim isHeader As Boolean: isHeader = True
    Do While Not ts.AtEndOfStream
        lineText = ts.ReadLine
        dataArray = Split(lineText, ",")
        If isHeader Then
            ' 最初の行はCSVヘッダ行とみなし読み飛ばす
            isHeader = False
        Else
            j = 1
            For Each key In colMap.Keys
                ' CSV列index=(key-1)に対応する値をセット（範囲外チェック）
                If key - 1 <= UBound(dataArray) Then
                    ws.Cells(i, j).Value = Trim(dataArray(key - 1))
                End If
                j = j + 1
            Next key
            i = i + 1
        End If
    Loop
    ts.Close

    ' 4. 読み込んだデータの列幅を自動調整
    ws.Cells.EntireColumn.AutoFit

    ' 5. 自動計算と画面更新を再開
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ImportError:
    MsgBox "CSVデータ読込中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
    If Not ts Is Nothing Then ts.Close
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub TransferBillingDetails(reportWb As Workbook, csvFileName As String)
    Dim wsMain As Worksheet, wsDetails As Worksheet
    Dim lastRowMain As Long, i As Long
    Dim dispensingCode As String, dispensingYM As String
    Dim payerCode As String, payerType As String
    Dim receiptNo As String
    Dim startRowDict As Object                  ' 各カテゴリ開始行(Dictionary)
    Dim rebillDict As Object, lateDict As Object, unpaidDict As Object, assessmentDict As Object
    Dim category As String, startRow As Long
    Dim rowData As Variant
    Dim a As Long, b As Long, c As Long         ' 追加行数算出用

    ' 1. シートオブジェクト取得
    Set wsMain = reportWb.Sheets(1)    ' メインシート（請求確定状況データ）
    Set wsDetails = reportWb.Sheets(2) ' 詳細データシート

    ' 2. 処理対象の調剤年月コード(csvYYMM)を取得（メインシートB2セルの下4桁がRYYMM）
    Dim csvYYMM As String: csvYYMM = ""
    If wsMain.Cells(2, 2).Value <> "" Then
        csvYYMM = Right(wsMain.Cells(2, 2).Value, 4)
    End If

    ' 3. 請求先区分の判定（CSVファイル名の7文字目: "1"社保, "2"国保, その他=労災等）
    Dim baseName As String
    baseName = csvFileName
    If InStr(baseName, ".") > 0 Then baseName = Left(baseName, InStrRev(baseName, ".") - 1)
    If Len(baseName) >= 7 Then
        payerCode = Mid(baseName, 7, 1)
    Else
        payerCode = ""
    End If
    Select Case payerCode
        Case "1": payerType = "社保"
        Case "2": payerType = "国保"
        Case Else: payerType = "労災"   ' 想定外のものは労災等その他扱い
    End Select

    ' 4. 詳細シート上の各カテゴリ開始行を取得してDictionaryに格納
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
    Else
        ' 労災等は詳細シート対象外（処理不要）
        Exit Sub
    End If

    ' 5. 過去データ分類用のDictionaryを準備
    Set rebillDict = CreateObject("Scripting.Dictionary")     ' 返戻再請求（過去返戻分で当月入金）
    Set lateDict = CreateObject("Scripting.Dictionary")       ' 月遅れ請求（今回請求に含めた過去月分）
    Set unpaidDict = CreateObject("Scripting.Dictionary")     ' 未請求扱い（請求漏れ・除外分）
    Set assessmentDict = CreateObject("Scripting.Dictionary") ' 返戻・査定（返戻・減点で未収）

    ' 6. メインシート（請求データ）の最終行を取得（D列に値がある最後の行）
    lastRowMain = wsMain.Cells(wsMain.Rows.Count, "D").End(xlUp).Row

    ' 7. メインシートの各レコードを走査し、当月ではないデータを各カテゴリに振り分け
    For i = 2 To lastRowMain
        dispensingCode = wsMain.Cells(i, 2).Value            ' 元号付き調剤年月 (例: "50701")
        dispensingYM = ConvertToWesternDate(dispensingCode)   ' YY.MM形式に変換 (例: "07.01")
        If csvYYMM <> "" And Right(dispensingCode, 4) <> csvYYMM Then
            ' ※対象診療月(csvYYMM)と異なる＝過去月レセプト
            ' 転記用データ配列（患者氏名, 調剤年月(YY.MM), 医療機関名, 請求点数）を用意
            rowData = Array(wsMain.Cells(i, 4).Value, dispensingYM, wsMain.Cells(i, 5).Value, wsMain.Cells(i, 10).Value)
            ' ファイル種別ごとに過去月データのカテゴリ振り分け
            If InStr(LCase(csvFileName), "fixf") > 0 Then
                ' `fixf`（請求確定）では過去月レセプトはすべて「月遅れ請求」に分類
                lateDict(wsMain.Cells(i, 1).Value) = rowData
            ElseIf InStr(LCase(csvFileName), "fmei") > 0 Then
                ' 振込明細では過去月レセプトを「返戻再請求」として分類（前月返戻→当月入金）
                rebillDict(wsMain.Cells(i, 1).Value) = rowData
            ElseIf InStr(LCase(csvFileName), "zogn") > 0 Then
                ' 増減点連絡書では過去月レセプトを「未請求扱い」に分類（請求除外/未処理）
                unpaidDict(wsMain.Cells(i, 1).Value) = rowData
            ElseIf InStr(LCase(csvFileName), "henr") > 0 Then
                ' 返戻内訳書では過去月レセプトを「返戻・査定」に分類（査定等で未収）
                assessmentDict(wsMain.Cells(i, 1).Value) = rowData
            End If
        End If
    Next i

    ' 8. 各カテゴリの件数超過分を算出（初期枠4件を超えた分）
    a = 0: b = 0: c = 0
    If rebillDict.Count > 4 Then a = rebillDict.Count - 4
    If lateDict.Count > 4 Then b = lateDict.Count - 4
    If assessmentDict.Count > 4 Then c = assessmentDict.Count - 4
    ' ※未請求扱い(unpaidDict)は今後の請求候補として枠固定（超過行挿入しない）

    ' 9. 必要な追加行を各カテゴリセクションに挿入
    If a + b + c > 0 Then
        If a > 0 Then wsDetails.Rows(startRowDict("月遅れ請求") + 1 & ":" & startRowDict("月遅れ請求") + a).Insert Shift:=xlDown
        If b > 0 Then wsDetails.Rows(startRowDict("返戻・査定") + 1 & ":" & startRowDict("返戻・査定") + b).Insert Shift:=xlDown
        If c > 0 Then wsDetails.Rows(startRowDict("未請求扱い") + 1 & ":" & startRowDict("未請求扱い") + c).Insert Shift:=xlDown
    End If

    ' 10. 各Dictionaryのデータを詳細シートに順次転記
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

    ' 11. 完了メッセージ（処理区分ごとに表示）
    MsgBox payerType & " のデータ転記が完了しました！", vbInformation, "処理完了"
End Sub

Sub TransferData(dataDict As Object, ws As Worksheet, startRow As Long, payerType As String)
    If dataDict.Count = 0 Then Exit Sub

    Dim key As Variant, rowData As Variant
    Dim r As Long: r = startRow
    Dim payerCol As Long

    ' 社保はH列(8), 国保はI列(9)に種別を記載
    If payerType = "社保" Then
        payerCol = 8
    ElseIf payerType = "国保" Then
        payerCol = 9
    Else
        Exit Sub  ' その他（労災等）は対象外
    End If

    ' 各レコードを書き込み
    For Each key In dataDict.Keys
        rowData = dataDict(key)
        ws.Cells(r, 4).Value = rowData(0)          ' 患者氏名
        ws.Cells(r, 5).Value = rowData(1)          ' 調剤年月 (YY.MM形式)
        ws.Cells(r, 6).Value = rowData(2)          ' 医療機関名
        ws.Cells(r, payerCol).Value = payerType    ' 請求先種別 (社保/国保)
        ws.Cells(r, payerCol).Font.Bold = True     ' 強調表示
        ws.Cells(r, 10).Value = rowData(3)         ' 請求点数
        r = r + 1
    Next key
End Sub

Function GetColumnMapping(fileType As String) As Object
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    Dim k As Integer

    Select Case fileType
        Case "振込額明細書"
            ' 支払基金からの振込額明細CSV列 → シート列見出し の対応
            colMap.Add 2, "診療（調剤）年月"
            colMap.Add 5, "受付番号"
            colMap.Add 14, "氏名"
            colMap.Add 16, "生年月日"
            colMap.Add 22, "医療保険_請求点数"
            colMap.Add 23, "医療保険_決定点数"
            colMap.Add 24, "医療保険_一部負担金"
            colMap.Add 25, "医療保険_金額"
            ' 第1～第5公費 分の列（各10列間隔: 請求点数・決定点数・患者負担金・金額）
            For k = 1 To 5
                colMap.Add 33 + (k - 1) * 10, "第" & k & "公費_請求点数"
                colMap.Add 34 + (k - 1) * 10, "第" & k & "公費_決定点数"
                colMap.Add 35 + (k - 1) * 10, "第" & k & "公費_患者負担金"
                colMap.Add 36 + (k - 1) * 10, "第" & k & "公費_金額"
            Next k
            colMap.Add 82, "算定額合計"

        Case "請求確定状況"
            ' 請求確定CSV（fixfデータ）の列対応
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
            colMap.Add 21, "増減点数(金額)"
            colMap.Add 22, "事由"

        Case "返戻内訳書"
            colMap.Add 2, "調剤年月(YYMM)"
            colMap.Add 3, "受付番号"
            colMap.Add 4, "保険者番号"
            colMap.Add 7, "氏名"
            colMap.Add 9, "請求点数"
            colMap.Add 10, "薬剤一部負担金"
            colMap.Add 12, "一部負担金額"
            colMap.Add 13, "公費負担金額"
            colMap.Add 14, "事由コード"

        Case Else
            ' その他データ種別（必要に応じ追加）
    End Select

    Set GetColumnMapping = colMap
End Function

Function ConvertToWesternDate(dispensingCode As String) As String
    Dim eraCode As String, yearNum As Integer, westernYear As Integer, monthPart As String
    If Len(dispensingCode) < 5 Then
        ConvertToWesternDate = ""
        Exit Function
    End If
    eraCode = Left(dispensingCode, 1)                ' 元号コード
    yearNum = CInt(Mid(dispensingCode, 2, 2))        ' 元号年2桁
    monthPart = Right(dispensingCode, 2)             ' 月2桁
    Select Case eraCode
        Case "5": westernYear = 2018 + yearNum   ' 令和 (2019=元年)
        Case "4": westernYear = 1988 + yearNum   ' 平成 (1989=元年)
        Case "3": westernYear = 1925 + yearNum   ' 昭和 (1926=元年)
        Case "2": westernYear = 1911 + yearNum   ' 大正 (1912=元年)
        Case "1": westernYear = 1867 + yearNum   ' 明治 (1868=元年)
        Case Else: westernYear = 2000 + yearNum  ' 仮置き
    End Select
    ' WesternYearの下2桁と月を組み合わせ "YY.MM"形式文字列を返す
    ConvertToWesternDate = Right(CStr(westernYear), 2) & "." & monthPart
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

Function GetUniqueSheetName(wb As Workbook, baseName As String) As String
    Dim newName As String, counter As Integer
    Dim ws As Worksheet, exists As Boolean
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

' --- （参考）半期ごとの請求誤差調査 ---
Sub InvestigateHalfYearDiscrepancy()
    ' ユーザー入力の年・半期区分について、保存済み報告書ファイルを集計し
    ' 請求点数と決定点数の差異を一覧表示する。
    Dim yearStr As String, halfStr As String
    Dim yearNum As Integer, half As Integer
    Dim startMonth As Integer, endMonth As Integer
    Dim fso As Object, folderPath As String
    Dim m As Integer
    Dim fileName As String, filePath As String
    Dim wb As Workbook, wsMain As Worksheet, wsDep As Worksheet
    Dim totalPointsClaim As Long, totalPointsDecided As Long
    Dim eraCode As String, eraYear As Integer, eraYY As String
    Dim resultMsg As String

    ' 1. 対象年と半期を入力させる
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

    ' 2. 半期の開始月・終了月を設定
    If half = 1 Then
        startMonth = 1: endMonth = 6   ' 上期: 1～6月
    Else
        startMonth = 7: endMonth = 12  ' 下期: 7～12月
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = GetSavePath()
    If folderPath = "" Then
        MsgBox "保存フォルダが設定されていません。", vbExclamation, "エラー"
        Exit Sub
    End If

    resultMsg = yearNum & "年 " & IIf(half = 1, "上期", "下期") & " 請求誤差調査結果:" & vbCrLf

    ' 3. 指定期間各月の報告書ファイルを順次開き、請求点数と決定点数を集計
    For m = startMonth To endMonth
        ' ファイル名（RYYMM形式）を構築
        If yearNum >= 2019 Then
            eraCode = "5": eraYear = yearNum - 2018   ' 令和
        ElseIf yearNum >= 1989 Then
            eraCode = "4": eraYear = yearNum - 1988   ' 平成
        ElseIf yearNum >= 1926 Then
            eraCode = "3": eraYear = yearNum - 1925   ' 昭和
        ElseIf yearNum >= 1912 Then
            eraCode = "2": eraYear = yearNum - 1911   ' 大正
        Else
            eraCode = "1": eraYear = yearNum - 1867   ' 明治
        End If
        eraYY = Format(eraYear, "00")
        fileName = "保険請求管理報告書_" & eraYY & Format(m, "00") & ".xlsm"
        filePath = folderPath & "\" & fileName

        If fso.FileExists(filePath) Then
            ' 報告書Excelを開いて集計
            Set wb = Workbooks.Open(filePath, ReadOnly:=True)
            Set wsMain = wb.Sheets(1)  ' メインシート
            totalPointsClaim = 0: totalPointsDecided = 0

            ' メインシート「総合計点数」列合計を算出（請求点数合計）
            Dim hdrCell As Range, colClaim As Long
            Set hdrCell = wsMain.Rows(1).Find("総合計点数", LookAt:=xlWhole)
            If Not hdrCell Is Nothing Then
                colClaim = hdrCell.Column
                Dim lastRow As Long
                lastRow = wsMain.Cells(wsMain.Rows.Count, colClaim).End(xlUp).Row
                If lastRow >= 2 Then
                    totalPointsClaim = Application.WorksheetFunction.Sum(wsMain.Range(wsMain.Cells(2, colClaim), wsMain.Cells(lastRow, colClaim)))
                End If
            End If

            ' 振込額明細シート上の「決定点数」列合計を算出（実際の支払点数合計）
            Set wsDep = Nothing
            Dim ws As Worksheet, foundHdr As Range
            For Each ws In wb.Sheets
                Set foundHdr = ws.Rows(1).Find("決定点数", LookAt:=xlPart)
                If Not foundHdr Is Nothing Then
                    ' ヘッダに"決定点数"を含むシート（メインシートおよび詳細シートは除く）を振込額明細シートとみなす
                    If LCase(ws.Name) <> LCase(wsMain.Name) And LCase(ws.Name) <> LCase(wb.Sheets(2).Name) Then
                        Set wsDep = ws
                        Exit For
                    End If
                End If
            Next ws
            If Not wsDep Is Nothing Then
                ' 決定点数列（複数列: 社保・各公費）を順次合計
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

            ' 差異を算出しメッセージ文字列に追加
            Dim diffPoints As Long
            diffPoints = totalPointsClaim - totalPointsDecided
            If diffPoints <> 0 Then
                resultMsg = resultMsg & "・" & yearNum & "年" & m & "月: 請求=" & totalPointsClaim & " , 決定=" & totalPointsDecided & " （差異 " & diffPoints & "点）" & vbCrLf
            End If
        Else
            resultMsg = resultMsg & "・" & yearNum & "年" & m & "月: 報告書未作成" & vbCrLf
        End If
    Next m

    ' 4. 集計結果を表示
    MsgBox resultMsg, vbInformation, "半期ごとの請求誤差調査結果"
End Sub

Function ConvertEraCodeToLetter(eraCode As String) As String
    Select Case eraCode
        Case "1": ConvertEraCodeToLetter = "M"
        Case "2": ConvertEraCodeToLetter = "T"
        Case "3": ConvertEraCodeToLetter = "S"
        Case "4": ConvertEraCodeToLetter = "H"
        Case "5": ConvertEraCodeToLetter = "R"
        Case Else: ConvertEraCodeToLetter = "E"
    End Select
End Function

Option Explicit

' (追加) 請求年月（GYYMM形式）から診療年月（RYYMM形式）を取得する関数
Function GetTreatmentYearMonthFromFileName(fileName As String) As String
    Dim baseName As String, codePart As String
    Dim eraCodeNum As String, eraYearCode As String, monthCode As String
    Dim westernYear As Integer, westernMonth As Integer
    ' ファイル名から拡張子を除いた部分を取得
    baseName = fileName
    If InStr(fileName, ".") > 0 Then baseName = Left(fileName, InStrRev(fileName, ".") - 1)
    ' ファイル名中のGYYMMコード部分を抽出（末尾の5文字が数字の場合に取得）
    codePart = Right(baseName, 5)
    If Not codePart Like "*#####" Then
        ' 末尾5文字が数字でない場合、"TEST"挿入などを考慮して数字部分を検索
        Dim i As Long
        For i = Len(baseName) To 1 Step -1
            If Mid(baseName, i, 5) Like "#####" Then
                codePart = Mid(baseName, i, 5)
                Exit For
            End If
        Next i
    End If
    If Len(codePart) <> 5 Or Not IsNumeric(codePart) Then
        GetTreatmentYearMonthFromFileName = ""  ' 変換失敗時は空文字
        Exit Function
    End If
    ' 請求年月（GYYMM）のコードから西暦年・月を取得
    eraCodeNum = Left(codePart, 1)          ' 元号コード（数字）
    eraYearCode = Mid(codePart, 2, 2)       ' 元号年（2桁）
    monthCode = Right(codePart, 2)          ' 月（2桁）
    Select Case eraCodeNum
        Case "5": westernYear = 2018 + CInt(eraYearCode)   ' 令和(2019=令和元年)
        Case "4": westernYear = 1988 + CInt(eraYearCode)   ' 平成(1989=平成元年)
        Case "3": westernYear = 1925 + CInt(eraYearCode)   ' 昭和(1926=昭和元年)
        Case "2": westernYear = 1911 + CInt(eraYearCode)   ' 大正(1912=大正元年)
        Case "1": westernYear = 1867 + CInt(eraYearCode)   ' 明治(1868=明治元年)
        Case Else: westernYear = 2000 + CInt(eraYearCode)  ' 不明なコード（仮）
    End Select
    westernMonth = CInt(monthCode)
    ' 請求年月の1ヶ月前を診療年月とする（月を1減算）
    westernMonth = westernMonth - 1
    If westernMonth = 0 Then
        westernYear = westernYear - 1
        westernMonth = 12
    End If
    ' 診療年月を元号コード(アルファベット)付きのRYYMM形式文字列に変換
    Dim newEraCode As String, newEraYear As Integer, newEraYearCode As String, eraLetter As String
    If westernYear >= 2019 Then
        newEraCode = "5": newEraYear = westernYear - 2018   ' 令和
    ElseIf westernYear >= 1989 Then
        newEraCode = "4": newEraYear = westernYear - 1988   ' 平成
    ElseIf westernYear >= 1926 Then
        newEraCode = "3": newEraYear = westernYear - 1925   ' 昭和
    ElseIf westernYear >= 1912 Then
        newEraCode = "2": newEraYear = westernYear - 1911   ' 大正
    Else
        newEraCode = "1": newEraYear = westernYear - 1867   ' 明治
    End If
    newEraYearCode = Format(newEraYear, "00")
    eraLetter = ConvertEraCodeToLetter(newEraCode)
    GetTreatmentYearMonthFromFileName = eraLetter & newEraYearCode & Format(westernMonth, "00")
End Function

' (追加) GYYMM形式に基づき報告書ファイル名を生成する関数
Function GetReportFileName(fileName As String) As String
    Dim Rcode As String
    Rcode = GetTreatmentYearMonthFromFileName(fileName)
    If Rcode = "" Then
        GetReportFileName = ""
    Else
        GetReportFileName = "保険請求管理報告書_" & Rcode & ".xlsm"
    End If
End Function