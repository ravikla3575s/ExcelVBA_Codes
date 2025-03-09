' **請求CSV一括処理マクロ:** 
' 指定フォルダ内の請求確定CSV(`fixf`)および各種明細CSV(`fmei`:振込額明細, `henr`:返戻内訳, `zogn`:増減点連絡)を読み込み、
' 月次の「保険請求管理報告書」Excelを作成・更新します。
' 処理後、報告書Excel（名称: 保険請求管理報告書_RYYMM.xlsx）が指定フォルダに出力されます。

Sub ProcessCSV()
    Dim csv_folder As String            ' CSVフォルダパス
    Dim file_system As Object          ' FileSystemObject
    Dim invoice_year As String, invoice_month As String  ' 処理対象の請求年・月（西暦）
    Dim dispensing_year As Integer, dispensing_month As Integer
    Dim save_path As String            ' 報告書保存先フォルダ
    Dim template_path As String        ' 報告書テンプレートファイル(.xltm)パス
    Dim report_workbook As Workbook    ' 報告書Excelブック
    Dim fixf_files As Collection, fmei_files As Collection, henr_files As Collection, zogn_files As Collection
    Dim csv_file As Object
    Dim report_file_path As String, report_file_name As String
    Dim era_letter As String, year_code As String, month_code As String
    Dim fixf_file_path As String, sheet_name As String
    Dim i As Long

    ' ... (前略: フォルダ選択やテンプレートパス取得などの処理) ...

    ' 1. CSVフォルダをユーザーに選択させる
    csv_folder = SelectCsvFolder()
    If csv_folder = "" Then Exit Sub  ' ユーザーがキャンセルした場合

    ' 1.1 フォルダが空なら処理を中止
    If IsFolderEmpty(csv_folder) Then
        MsgBox "選択したフォルダにはCSVファイルがありません。処理を中止します。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 2. テンプレートパス・保存先フォルダを取得
    template_path = GetTemplate_path()    ' 設定シートのB2セル（テンプレート格納先）
    save_path = GetSave_path()           ' 設定シートのB3セル（保存先フォルダ）
    If template_path = "" Or save_path = "" Then Exit Sub  ' 必須パスが取得できなければ中止

    ' 3. FileSystemObjectの用意
    Set file_system = CreateObject("Scripting.FileSystemObject")

    ' 4. フォルダ内の全CSVファイルを種類別に収集（fixf, fmei, henr, zogn）
    Set file_system = CreateObject("Scripting.FileSystemObject")
    Set fixf_files = New Collection: Set fmei_files = New Collection
    Set henr_files = New Collection: Set zogn_files = New Collection
    For Each csv_file In file_system.GetFolder(csv_folder).Files
        If LCase(file_system.GetExtensionName(csv_file.Name)) = "csv" Then
            If InStr(LCase(csv_file.Name), "fixf") > 0 Then
                fixf_files.Add csv_file
            ElseIf InStr(LCase(csv_file.Name), "fmei") > 0 Then
                fmei_files.Add csv_file
            ElseIf InStr(LCase(csv_file.Name), "henr") > 0 Then
                henr_files.Add csv_file
            ElseIf InStr(LCase(csv_file.Name), "zogn") > 0 Then
                zogn_files.Add csv_file
            End If
        End If
    Next csv_file

    ' 4.1 対象CSVファイルが一つもない場合、処理を中止
    If fixf_files.Count = 0 And fmei_files.Count = 0 And henr_files.Count = 0 And zogn_files.Count = 0 Then
        MsgBox "選択したフォルダには処理対象のCSVファイルがありません。処理を中止します。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 5. 収集したCSVファイルを調剤年月が古い順にソート
    If fixf_files.Count > 1 Then
        SortCsvFilesByDate fixf_files, "fixf"
    End If
    If fmei_files.Count > 1 Then
        SortCsvFilesByDate fmei_files, "fmei"
    End If
    If henr_files.Count > 1 Then
        SortCsvFilesByDate henr_files, "henr"
    End If
    If zogn_files.Count > 1 Then
        SortCsvFilesByDate zogn_files, "zogn"
    End If

    ' 6. 請求確定CSV（fixf）の処理
    For Each csv_file In fixf_files
        fixf_file_path = csv_file.Path
        invoice_year = "": invoice_month = ""
        GetYearMonthFromFixf fixf_file_path, invoice_year, invoice_month  ' fixfファイルから請求年月を取得
        If invoice_year = "" Or invoice_month = "" Then
            MsgBox "ファイル " & fixf_file_path & " から調剤年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextFixf
        End If
        ' ファイルから得た請求年月から調剤年月を取得
        ExchangeToDispensingYearMonth invoice_year, invoice_month, dispensing_year, dispensing_month
        targetFile = FindOrCreateReport save_path, dispensing_year, dispensing_month, template_path
        ' テンプレート情報を設定（調剤年月等を更新）
        SetTemplateInfo report_workbook, invoice_year, invoice_month

        ' fixf CSVデータをシート3～6に転記（シートがなければ作成）
        Dim neededSheets As Integer: neededSheets = 6
        If report_workbook.Sheets.Count < neededSheets Then
            For i = report_workbook.Sheets.Count + 1 To neededSheets
                report_workbook.Sheets.Add After:=report_workbook.Sheets(report_workbook.Sheets.Count)
            Next i
        End If
        ' 対象シート3～6をクリア
        For i = 3 To 6
            If i <= report_workbook.Sheets.Count Then report_workbook.Sheets(i).Cells.Clear
        Next i
        ' fixf CSVファイルを読み込み、カテゴリ別にデータ行を振り分け
        Dim ts As Object, lineText As String, lines() As String
        Set ts = file_system.OpenTextFile(fixf_file_path, 1, False, -2)  ' UTF-8で読み込み
        lineText = ts.ReadAll: ts.Close
        lines = Split(lineText, vbCrLf)
        Dim column_map As Object: Set column_map = GetColumnMapping("請求確定状況")
        Dim data_lines_cat1 As New Collection, data_lines_cat2 As New Collection
        ' 最初の行はヘッダ行とみなす
        Dim header_skipped As Boolean: header_skipped = False
        For i = LBound(lines) To UBound(lines)
            If Trim(lines(i)) = "" Then Continue For
            If Not header_skipped Then
                header_skipped = True
                Continue For  ' ヘッダ行をスキップ
            End If
            If Left(lines(i), 2) = "1," Then
                data_lines_cat1.Add lines(i)
            ElseIf Left(lines(i), 2) = "2," Then
                data_lines_cat2.Add lines(i)
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
            For Each key In column_map.Keys
                ws.Cells(1, j).Value = column_map(key)
                j = j + 1
            Next key
            ' データ行を書き込み
            Dim rowIndex As Long: rowIndex = 2
            Dim arr As Variant
            For j = 1 To dataCol.Count
                arr = Split(dataCol(j), ",")
                Dim k As Long: k = 1
                For Each key In column_map.Keys
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
        If data_lines_cat1.Count > 0 Then
            If data_lines_cat1.Count <= maxLinesPerSheet Then
                WriteDataToSheet report_workbook.Sheets(3), data_lines_cat1
            Else
                ' 1ページに収まらない場合、Sheet3とSheet4に分割
                Dim temp_collection_1 As New Collection
                For i = 1 To maxLinesPerSheet
                    temp_collection_1.Add data_lines_cat1(i)
                Next i
                WriteDataToSheet report_workbook.Sheets(3), temp_collection_1
                temp_collection_1.Remove 1 To temp_collection_1.Count  ' コレクションをクリア
                For i = maxLinesPerSheet + 1 To data_lines_cat1.Count
                    temp_collection_1.Add data_lines_cat1(i)
                Next i
                WriteDataToSheet report_workbook.Sheets(4), temp_collection_1
            End If
        End If
        ' カテゴリ2（国保）データの転記
        If data_lines_cat2.Count > 0 Then
            Dim start_sheet As Integer
            start_sheet = IIf(data_lines_cat1.Count > 0, 5, 3)  ' 社保ありならSheet5開始、なければSheet3開始
            If data_lines_cat2.Count <= maxLinesPerSheet Then
                WriteDataToSheet report_workbook.Sheets(start_sheet), data_lines_cat2
            Else
                ' 2ページに分割
                Dim temp_collection_2 As New Collection
                For i = 1 To maxLinesPerSheet
                    temp_collection_2.Add data_lines_cat2(i)
                Next i
                WriteDataToSheet report_workbook.Sheets(start_sheet), temp_collection_2
                temp_collection_2.Remove 1 To temp_collection_2.Count
                For i = maxLinesPerSheet + 1 To data_lines_cat2.Count
                    temp_collection_2.Add data_lines_cat2(i)
                Next i
                ' 次のシート（カテゴリ開始シートの+1）に続き出力
                WriteDataToSheet report_workbook.Sheets(start_sheet + 1), temp_collection_2
            End If
        End If

        ' ブックを保存して閉じる
        report_workbook.Save
        report_workbook.Close False
NextFixf:
        ' 次のfixfファイルへ
    Next csv_file

    ' 7. 振込額明細CSV（fmei）の処理
    For Each csv_file In fmei_files
        invoice_year = "": invoice_month = ""
        ' ファイル名から調剤年月コードを取得
        Dim ym_code As String
        ym_code = GetDispensingYearMonthFromFileName(csv_file.Name)
        If ym_code = "" Then
            MsgBox "ファイル " & csv_file.Name & " から調剤年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextFmei
        End If
        ' 調剤年月（西暦）を算出
        era_letter = Left(ym_code, 1)
        year_code = Mid(ym_code, 2, 2)
        month_code = Right(ym_code, 2)
        Select Case era_letter
            Case "R": invoice_year = CStr(2018 + CInt(year_code))   ' 令和
            Case "H": invoice_year = CStr(1988 + CInt(year_code))   ' 平成
            Case "S": invoice_year = CStr(1925 + CInt(year_code))   ' 昭和
            Case "T": invoice_year = CStr(1911 + CInt(year_code))   ' 大正
            Case "M": invoice_year = CStr(1867 + CInt(year_code))   ' 明治
            Case Else: invoice_year = CStr(2000 + CInt(year_code))  ' 仮（不明コード）
        End Select
        invoice_month = month_code
        ' 報告書ファイル名を決定（存在しなければ作成）
        report_file_name = "保険請求管理報告書_" & ym_code & ".xlsm"
        report_file_path = save_path & "\" & report_file_name
        If Not file_system.FileExists(report_file_path) Then
            Dim new_wb2 As Workbook
            Set new_wb2 = Workbooks.Add(template_path)
            If new_wb2 Is Nothing Then
                MsgBox "テンプレートファイルを開けませんでした: " & template_path, vbExclamation, "エラー"
                GoTo NextFmei
            End If
            Application.DisplayAlerts = False
            new_wb2.SaveAs Filename:=report_file_path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            Application.DisplayAlerts = True
            new_wb2.Close False
        End If
        On Error Resume Next
        Set report_workbook = Workbooks.Open(report_file_path)
        On Error GoTo 0
        If report_workbook Is Nothing Then
            MsgBox "ファイル " & report_file_path & " を開けませんでした。", vbExclamation, "エラー"
            GoTo NextFmei
        End If
        ' 既に同じファイル名のシートが存在する場合はスキップ
        sheet_name = file_system.GetBaseName(csv_file.Name)
        Dim sheet As Worksheet, sheetExists As Boolean: sheetExists = False
        For Each sheet In report_workbook.Sheets
            If InStr(sheet.Name, sheet_name) > 0 Then
                sheetExists = True
                Exit For
            End If
        Next sheet
        If sheetExists Then
            report_workbook.Close False
            GoTo NextFmei
        End If
        ' テンプレート情報を設定
        SetTemplateInfo report_workbook, invoice_year, invoice_month
        ' CSVデータを新規シートにインポート
        Dim insertIndex As Long
        insertIndex = Application.WorksheetFunction.Min(3, report_workbook.Sheets.Count + 1)
        Dim csv_sheet As Worksheet
        Set csv_sheet = report_workbook.Sheets.Add(After:=report_workbook.Sheets(insertIndex - 1))
        csv_sheet.Name = sheet_name
        ImportCsvData csv_file.Path, csv_sheet, "振込額明細書"
        TransferBillingDetails report_workbook, csv_file.Name  ' 振込額明細の詳細転記処理

        report_workbook.Save
        report_workbook.Close False
NextFmei:
        ' 次のfmeiファイルへ
    Next csv_file

    ' 8. 返戻内訳CSV（henr）の処理
    For Each csv_file In henr_files
        invoice_year = "": invoice_month = ""
        ym_code = GetDispensingYearMonthFromFileName(csv_file.Name)
        If ym_code = "" Then
            MsgBox "ファイル " & csv_file.Name & " から調剤年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextHenr
        End If
        era_letter = Left(ym_code, 1)
        year_code = Mid(ym_code, 2, 2)
        month_code = Right(ym_code, 2)
        Select Case era_letter
            Case "R": invoice_year = CStr(2018 + CInt(year_code))
            Case "H": invoice_year = CStr(1988 + CInt(year_code))
            Case "S": invoice_year = CStr(1925 + CInt(year_code))
            Case "T": invoice_year = CStr(1911 + CInt(year_code))
            Case "M": invoice_year = CStr(1867 + CInt(year_code))
            Case Else: invoice_year = CStr(2000 + CInt(year_code))
        End Select
        invoice_month = month_code
        report_file_name = "保険請求管理報告書_" & ym_code & ".xlsm"
        report_file_path = save_path & "\" & report_file_name
        If Not file_system.FileExists(report_file_path) Then
            Dim new_wb3 As Workbook
            Set new_wb3 = Workbooks.Add(template_path)
            If new_wb3 Is Nothing Then
                MsgBox "テンプレートファイルを開けませんでした: " & template_path, vbExclamation, "エラー"
                GoTo NextHenr
            End If
            Application.DisplayAlerts = False
            new_wb3.SaveAs Filename:=report_file_path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            Application.DisplayAlerts = True
            new_wb3.Close False
        End If
        On Error Resume Next
        Set report_workbook = Workbooks.Open(report_file_path)
        On Error GoTo 0
        If report_workbook Is Nothing Then
            MsgBox "ファイル " & report_file_path & " を開けませんでした。", vbExclamation, "エラー"
            GoTo NextHenr
        End If
        sheet_name = file_system.GetBaseName(csv_file.Name)
        sheetExists = False
        For Each sheet In report_workbook.Sheets
            If InStr(sheet.Name, sheet_name) > 0 Then
                sheetExists = True: Exit For
            End If
        Next sheet
        If sheetExists Then
            report_workbook.Close False
            GoTo NextHenr
        End If
        SetTemplateInfo report_workbook, invoice_year, invoice_month
        insertIndex = Application.WorksheetFunction.Min(3, report_workbook.Sheets.Count + 1)
        Set csv_sheet = report_workbook.Sheets.Add(After:=report_workbook.Sheets(insertIndex - 1))
        csv_sheet.Name = sheet_name
        ImportCsvData csv_file.Path, csv_sheet, "返戻内訳書"
        TransferBillingDetails report_workbook, csv_file.Name  ' 返戻内訳の詳細転記処理

        report_workbook.Save
        report_workbook.Close False
NextHenr:
        ' 次のhenrファイルへ
    Next csv_file

    ' 9. 増減点連絡CSV（zogn）の処理
    For Each csv_file In zogn_files
        invoice_year = "": invoice_month = ""
        ym_code = GetDispensingYearMonthFromFileName(csv_file.Name)
        If ym_code = "" Then
            MsgBox "ファイル " & csv_file.Name & " から調剤年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextZogn
        End If
        era_letter = Left(ym_code, 1)
        year_code = Mid(ym_code, 2, 2)
        month_code = Right(ym_code, 2)
        Select Case era_letter
            Case "R": invoice_year = CStr(2018 + CInt(year_code))
            Case "H": invoice_year = CStr(1988 + CInt(year_code))
            Case "S": invoice_year = CStr(1925 + CInt(year_code))
            Case "T": invoice_year = CStr(1911 + CInt(year_code))
            Case "M": invoice_year = CStr(1867 + CInt(year_code))
            Case Else: invoice_year = CStr(2000 + CInt(year_code))
        End Select
        invoice_month = month_code
        report_file_name = "保険請求管理報告書_" & ym_code & ".xlsm"
        report_file_path = save_path & "\" & report_file_name
        If Not file_system.FileExists(report_file_path) Then
            Dim new_wb4 As Workbook
            Set new_wb4 = Workbooks.Add(template_path)
            If new_wb4 Is Nothing Then
                MsgBox "テンプレートファイルを開けませんでした: " & template_path, vbExclamation, "エラー"
                GoTo NextZogn
            End If
            Application.DisplayAlerts = False
            new_wb4.SaveAs Filename:=report_file_path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            Application.DisplayAlerts = True
            new_wb4.Close False
        End If
        On Error Resume Next
        Set report_workbook = Workbooks.Open(report_file_path)
        On Error GoTo 0
        If report_workbook Is Nothing Then
            MsgBox "ファイル " & report_file_path & " を開けませんでした。", vbExclamation, "エラー"
            GoTo NextZogn
        End If
        sheet_name = file_system.GetBaseName(csv_file.Name)
        sheetExists = False
        For Each sheet In report_workbook.Sheets
            If InStr(sheet.Name, sheet_name) > 0 Then
                sheetExists = True: Exit For
            End If
        Next sheet
        If sheetExists Then
            report_workbook.Close False
            GoTo NextZogn
        End If
        SetTemplateInfo report_workbook, invoice_year, invoice_month
        insertIndex = Application.WorksheetFunction.Min(3, report_workbook.Sheets.Count + 1)
        Set csv_sheet = report_workbook.Sheets.Add(After:=report_workbook.Sheets(insertIndex - 1))
        csv_sheet.Name = sheet_name
        ImportCsvData csv_file.Path, csv_sheet, "増減点連絡書"
        TransferBillingDetails report_workbook, csv_file.Name  ' 増減点連絡の詳細転記処理

        report_workbook.Save
        report_workbook.Close False
NextZogn:
        ' 次のzognファイルへ
    Next csv_file

    ' 10. 完了メッセージ
    MsgBox "CSVファイルの処理が完了しました！", vbInformation, "完了"

End Sub

' 調剤年月順にCSVファイルコレクションをソートする関数
Private Sub SortCsvFilesByDate(file_list As Collection, file_type As String)
    Dim i As Long, j As Long
    Dim temp As Object
    For i = 1 To file_list.Count - 1
        For j = i + 1 To file_list.Count
            Dim key1 As String, key2 As String
            ' ファイル種別に応じて比較キーを取得
            If file_type = "fixf" Then
                ' fixfファイルは名前中のタイムスタンプから年+月を抽出
                Dim yy1 As String, mm1 As String, yy2 As String, mm2 As String
                yy1 = "": mm1 = "": yy2 = "": mm2 = ""
                GetYearMonthFromFixf file_list(i).Path, yy1, mm1
                GetYearMonthFromFixf file_list(j).Path, yy2, mm2
                If yy1 <> "" And mm1 <> "" And yy2 <> "" And mm2 <> "" Then
                    key1 = yy1 & mm1
                    key2 = yy2 & mm2
                End If
            Else
                ' その他のファイルはファイル名末尾の5桁コード（GYYMM形式）を使用
                key1 = Right(file_list(i).Name, 5)
                key2 = Right(file_list(j).Name, 5)
            End If
            ' 取得したキーが有効な数値なら比較して並べ替え
            If IsNumeric(key1) And IsNumeric(key2) Then
                If CDbl(key1) > CDbl(key2) Then
                    Set temp = file_list(i)
                    Set file_list(i) = file_list(j)
                    Set file_list(j) = temp
                End If
            End If
        Next j
    Next i
End Sub

' 指定年月の報告書Excelを取得（存在しなければテンプレートから作成）
Private Function FindOrCreateReport(save_path As String, dispensing_year As String, dispensing_month As String, template_path As String) As String
    ' 元号コード付きのファイル名を生成（例：「保険請求管理報告書_R0504.xlsm」）
    Dim era_letter As String, year_code As String, month_code As String
    Dim yr As Integer: yr = CInt(dispending_year)
    Dim mon As Integer: mon = CInt(dispending_month)
    ' 年月から元号コードを決定
    Select Case yr
        Case Is >= 2019: era_letter = "R": year_code = Format(yr - 2018, "00")  ' 2019年以降: 令和
        Case Is >= 1989: era_letter = "H": year_code = Format(yr - 1988, "00")  ' 1989年以降: 平成
        Case Is >= 1926: era_letter = "S": year_code = Format(yr - 1925, "00")  ' 1926年以降: 昭和
        Case Is >= 1912: era_letter = "T": year_code = Format(yr - 1911, "00")  ' 1912年以降: 大正
        Case Else:      era_letter = "M": year_code = Format(yr - 1867, "00")  ' その他: 明治など
    End Select
    month_code = Format(mon, "00")
    Dim report_fileName As String
    report_fileName = "保険請求管理報告書_" & era_letter & year_code & month_code & ".xlsm"
    Dim report_path As String
    report_path = save_path & "\" & report_fileName
    ' ファイル存在チェックと作成処理
    If Dir(report_path) = "" Then  ' ファイル未存在の場合
        Dim new_wb As Workbook
        On Error Resume Next
        Set new_wb = Workbooks.Add(template_path)  ' テンプレートから新規ブック作成
        On Error GoTo 0
        If new_wb Is Nothing Then
            MsgBox "テンプレートファイルを開けませんでした: " & template_path, vbExclamation, "エラー"
            FindOrCreateReport = ""
            Exit Function
        End If
        Application.DisplayAlerts = False
        new_wb.SaveAs Filename:=report_path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Application.DisplayAlerts = True
        new_wb.Close False
    End If
    ' 作成または既存のファイルパスを戻り値とする
    FindOrCreateReport = report_path
End Function

Private Function ExchangeToDispensingYearMonth(invoice_year, invoice_month, dispensing_year, dispensing_month)
    ' 請求年月の1ヶ月前を調剤年月とする（月を1減算）
    dispensing_month = invoice_month - 1
    If dispensing_month = 0 Then
        dispensing_year = invoice_year - 1
        dispensing_month = 12
    End If
End Function