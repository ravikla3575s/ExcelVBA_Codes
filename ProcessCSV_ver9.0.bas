Option Explicit

' **請求CSV一括処理マクロ:** 
' 指定フォルダ内の請求確定CSV(`fixf`)および各種明細CSV(`fmei`:振込額明細, `henr`:返戻内訳, `zogn`:増減点連絡)を読み込み、
' 月次の「保険請求管理報告書」Excelを作成・更新します。
' 処理後、報告書Excel（名称: 保険請求管理報告書_RYYMM.xlsx）が指定フォルダに出力されます。

Sub ProcessCsv()
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
    template_path = GetTemplatePath()    ' 設定シートのB2セル（テンプレート格納先）
    save_path = GetSavePath()           ' 設定シートのB3セル（保存先フォルダ）
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
        For i = 1 To fixf_files.Count - 1
            For j = i + 1 To fixf_files.Count
                Dim year_1 As String, month_1 As String
                Dim year_2 As String, month_2 As String
                year_1 = "": month_1 = ""
                year_2 = "": month_2 = ""
                GetYearMonthFromFixf fixf_files(i).Path, year_1, month_1
                GetYearMonthFromFixf fixf_files(j).Path, year_2, month_2
                If year_1 <> "" And month_1 <> "" And year_2 <> "" And month_2 <> "" Then
                    If (year_1 & month_1) > (year_2 & month_2) Then
                        Set temp_obj = fixf_files(i)
                        Set fixf_files(i) = fixf_files(j)
                        Set fixf_files(j) = temp_obj
                    End If
                End If
            Next j
        Next i
    End If
    If fmei_files.Count > 1 Then
        For i = 1 To fmei_files.Count - 1
            For j = i + 1 To fmei_files.Count
                Dim code_1 As String, code_2 As String
                code_1 = Right(file_system.GetBaseName(fmei_files(i).Name), 5)
                code_2 = Right(file_system.GetBaseName(fmei_files(j).Name), 5)
                If Len(code_1) = 5 And Len(code_2) = 5 And IsNumeric(code_1) And IsNumeric(code_2) Then
                    If CDbl(code_1) > CDbl(code_2) Then
                        Set temp_obj = fmei_files(i)
                        Set fmei_files(i) = fmei_files(j)
                        Set fmei_files(j) = temp_obj
                    End If
                End If
            Next j
        Next i
    End If
    If henr_files.Count > 1 Then
        For i = 1 To henr_files.Count - 1
            For j = i + 1 To henr_files.Count
                code_1 = Right(file_system.GetBaseName(henr_files(i).Name), 5)
                code_2 = Right(file_system.GetBaseName(henr_files(j).Name), 5)
                If Len(code_1) = 5 And Len(code_2) = 5 And IsNumeric(code_1) And IsNumeric(code_2) Then
                    If CDbl(code_1) > CDbl(code_2) Then
                        Set temp_obj = henr_files(i)
                        Set henr_files(i) = henr_files(j)
                        Set henr_files(j) = temp_obj
                    End If
                End If
            Next j
        Next i
    End If
    If zogn_files.Count > 1 Then
        For i = 1 To zogn_files.Count - 1
            For j = i + 1 To zogn_files.Count
                code_1 = Right(file_system.GetBaseName(zogn_files(i).Name), 5)
                code_2 = Right(file_system.GetBaseName(zogn_files(j).Name), 5)
                If Len(code_1) = 5 And Len(code_2) = 5 And IsNumeric(code_1) And IsNumeric(code_2) Then
                    If CDbl(code_1) > CDbl(code_2) Then
                        Set temp_obj = zogn_files(i)
                        Set zogn_files(i) = zogn_files(j)
                        Set zogn_files(j) = temp_obj
                    End If
                End If
            Next j
        Next i
    End If

    ' 6. 請求確定CSV（fixf）の処理
    For Each csv_file In fixf_files
        Dim treatmentYear_shadow As Integer, treatmentMonth_shadow As Integer  ' ※一時的に使用（スコープ内で再宣言）
        treatmentYear_shadow = "": treatmentMonth_shadow = ""
        fixf_file_path = csv_file.Path
        invoice_year = "": invoice_month = ""
        GetYearMonthFromFixf fixf_file_path, invoice_year, invoice_month  ' fixfファイルから調剤年月を取得
        If invoice_year = "" Or invoice_month = "" Then
            MsgBox "ファイル " & fixf_file_path & " から調剤年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextFixf
        End If
        ' 請求年月の1ヶ月前を調剤年月とする（月を1減算）
        treatmentMonth_shadow = invoice_month - 1
        If treatmentMonth_shadow = 0 Then
            treatmentYear_shadow = invoice_year - 1
            treatmentMonth_shadow = 12
        End If
        ' 出力報告書ファイル名（調剤年月RYYMM形式）を決定
        If CInt(treatmentYear_shadow) >= 2019 Then
            era_letter = "R": year_code = Format(CInt(treatmentYear_shadow) - 2018, "00")  ' 令和
        ElseIf CInt(treatmentYear_shadow) >= 1989 Then
            era_letter = "H": year_code = Format(CInt(treatmentYear_shadow) - 1988, "00")  ' 平成
        ElseIf CInt(treatmentYear_shadow) >= 1926 Then
            era_letter = "S": year_code = Format(CInt(treatmentYear_shadow) - 1925, "00")  ' 昭和
        ElseIf CInt(treatmentYear_shadow) >= 1912 Then
            era_letter = "T": year_code = Format(CInt(treatmentYear_shadow) - 1911, "00")  ' 大正
        Else
            era_letter = "M": year_code = Format(CInt(treatmentYear_shadow) - 1867, "00")  ' 明治
        End If
        report_file_name = "保険請求管理報告書_" & era_letter & year_code & Format(CInt(treatmentMonth_shadow), "00") & ".xlsm"
        report_file_path = save_path & "\" & report_file_name
        ' **既存の報告書ファイルがある場合、重複処理を避けてスキップ**
        If file_system.FileExists(report_file_path) Then
            GoTo NextFixf
        End If
        ' 報告書ブックを取得（存在しなければテンプレートから新規作成）
        If Not file_system.FileExists(report_file_path) Then
            Dim newWb As Workbook
            Set newWb = Workbooks.Add(template_path)
            If newWb Is Nothing Then
                MsgBox "テンプレートファイルを開けませんでした: " & template_path, vbExclamation, "エラー"
                GoTo NextFixf
            End If
            Application.DisplayAlerts = False
            newWb.SaveAs Filename:=report_file_path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            Application.DisplayAlerts = True
            newWb.Close False
        End If
        On Error Resume Next
        Set report_workbook = Workbooks.Open(report_file_path)
        On Error GoTo 0
        If report_workbook Is Nothing Then
            MsgBox "ファイル " & report_file_path & " を開けませんでした。", vbExclamation, "エラー"
            GoTo NextFixf
        End If
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
            Dim newWb2 As Workbook
            Set newWb2 = Workbooks.Add(template_path)
            If newWb2 Is Nothing Then
                MsgBox "テンプレートファイルを開けませんでした: " & template_path, vbExclamation, "エラー"
                GoTo NextFmei
            End If
            Application.DisplayAlerts = False
            newWb2.SaveAs Filename:=report_file_path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            Application.DisplayAlerts = True
            newWb2.Close False
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
            Dim newWb3 As Workbook
            Set newWb3 = Workbooks.Add(template_path)
            If newWb3 Is Nothing Then
                MsgBox "テンプレートファイルを開けませんでした: " & template_path, vbExclamation, "エラー"
                GoTo NextHenr
            End If
            Application.DisplayAlerts = False
            newWb3.SaveAs Filename:=report_file_path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            Application.DisplayAlerts = True
            newWb3.Close False
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
            Dim newWb4 As Workbook
            Set newWb4 = Workbooks.Add(template_path)
            If newWb4 Is Nothing Then
                MsgBox "テンプレートファイルを開けませんでした: " & template_path, vbExclamation, "エラー"
                GoTo NextZogn
            End If
            Application.DisplayAlerts = False
            newWb4.SaveAs Filename:=report_file_path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            Application.DisplayAlerts = True
            newWb4.Close False
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

Sub ProcessWithoutFixf(file_system As Object, csv_folder As String, save_path As String, template_path As String)
    Dim invoice_year As String, invoice_month As String
    Dim reportFile As String
    Dim report_workbook As Workbook

    ' 1. フォルダ内の最初のCSVから調剤年月を推定（fmei等の先頭行GYYMMコードを利用）
    invoice_year = ""
    invoice_month = ""
    ' fixfファイルがなく報告書ファイルも存在しない場合、fmeiファイル名から調剤年月を推定
    Dim fmeiFile As Object
    For Each fmeiFile In file_system.GetFolder(csv_folder).Files
        If LCase(file_system.GetExtensionName(fmeiFile.Name)) = "csv" And InStr(LCase(fmeiFile.Name), "fmei") > 0 Then
            ' 最初に見つかったfmeiファイルを使用
            Dim codePart As String, era_code As String, year_code As String, month_code As String
            Dim westernYear As Integer
            codePart = Right(file_system.GetBaseName(fmeiFile.Name), 5)
            If Len(codePart) = 5 And IsNumeric(codePart) Then
                era_code = Left(codePart, 1)
                year_code = Mid(codePart, 2, 2)
                month_code = Right(codePart, 2)
                Select Case era_code
                    Case "5": westernYear = 2018 + CInt(year_code)   ' 令和
                    Case "4": westernYear = 1988 + CInt(year_code)   ' 平成
                    Case "3": westernYear = 1925 + CInt(year_code)   ' 昭和
                    Case "2": westernYear = 1911 + CInt(year_code)   ' 大正
                    Case "1": westernYear = 1867 + CInt(year_code)   ' 明治
                    Case Else: westernYear = 2000 + CInt(year_code)  ' 仮定
                End Select
                invoice_year = CStr(westernYear)
                invoice_month = month_code
            End If
            Exit For
        End If
    Next fmeiFile
    ' ファイル名から取得できなかった場合、CSV内容から調剤年月を取得
    If invoice_year = "" Or invoice_month = "" Then
        GetYearMonthFromCSV file_system, csv_folder, invoice_year, invoice_month
    End If
    ' **調剤年月が取得できなかった場合は処理中止**
    If invoice_year = "" Or invoice_month = "" Then
        MsgBox "CSVファイルから調剤年月を取得できませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 2. 報告書Excelファイルを取得（既存がなければ新規作成）
    reportFile = FindOrCreateReport(save_path, invoice_year, invoice_month, template_path)
    If reportFile = "" Then
        MsgBox "調剤年月 " & invoice_year & "年" & invoice_month & "月 のExcelファイルを作成できませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 3. 報告書Excelを開く
    On Error Resume Next
    Set report_workbook = Workbooks.Open(reportFile)
    On Error GoTo 0
    If report_workbook Is Nothing Then
        MsgBox "ファイル " & reportFile & " を開けませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 4. テンプレート情報を設定（タイトル等）
    SetTemplateInfo report_workbook, invoice_year, invoice_month

    ' 5. CSVファイルを種類別に処理（fixfなしなので、振込明細・返戻・増減点のみ）
    ProcessAllCsvFiles file_system, report_workbook, csv_folder, invoice_year, invoice_month

    ' 6. 保存してブックを閉じる
    report_workbook.Save
    report_workbook.Close False

    ' 7. 完了メッセージ
    MsgBox "CSVファイルの処理が完了しました！", vbInformation, "完了"
End Sub

Sub ProcessAllCsvFiles(file_system As Object, report_workbook As Workbook, csv_folder As String, invoice_year As String, invoice_month As String)
    Dim era_code As String, eraYear As Integer
    Dim GYYMM As String          ' 和暦元号コード付の対象年月 (例:50702)
    Dim csvFileObj As Object
    ' 受け取りCSVの種類別コレクションを用意
    Dim fmei_files As New Collection, henr_files As New Collection, zogn_files As New Collection

    ' 対象年月を和暦GYYMM形式に変換（例: 2025年02月→令和7年=07 ⇒ "50702"）
    If CInt(invoice_year) >= 2019 Then
        era_code = "5"  ' 令和
        eraYear = CInt(invoice_year) - 2018
    ElseIf CInt(invoice_year) >= 1989 Then
        era_code = "4"  ' 平成
        eraYear = CInt(invoice_year) - 1988
    ElseIf CInt(invoice_year) >= 1926 Then
        era_code = "3"  ' 昭和
        eraYear = CInt(invoice_year) - 1925
    ElseIf CInt(invoice_year) >= 1912 Then
        era_code = "2"  ' 大正
        eraYear = CInt(invoice_year) - 1911
    Else
        era_code = "1"  ' 明治（想定外の場合）
        eraYear = CInt(invoice_year) - 1867
    End If
    GYYMM = era_code & Format(eraYear, "00") & invoice_month   ' 例: "50702"

    ' フォルダ内の全CSVファイルを走査し、ファイル名により種類別に振り分け
    For Each csvFileObj In file_system.GetFolder(csv_folder).Files
        If LCase(file_system.GetExtensionName(csvFileObj.Name)) = "csv" Then
            Dim base_name As String
            base_name = file_system.GetBaseName(csvFileObj.Name)
            ' ファイル名に種別キーワードを含み、かつ末尾のYYMMコードが対象年月かチェック
            If InStr(LCase(base_name), "fmei") > 0 And Right(base_name, Len(GYYMM)) = GYYMM Then
                fmei_files.Add csvFileObj    ' 振込額明細書CSVを収集
            ElseIf InStr(LCase(base_name), "henr") > 0 And Right(base_name, Len(GYYMM)) = GYYMM Then
                henr_files.Add csvFileObj    ' 返戻内訳書CSVを収集
            ElseIf InStr(LCase(base_name), "zogn") > 0 And Right(base_name, Len(GYYMM)) = GYYMM Then
                zogn_files.Add csvFileObj    ' 増減点連絡書CSVを収集
            End If
        End If
    Next csvFileObj

    ' 1) 振込額明細書（fmei）CSVの処理
    ProcessFmeiFiles file_system, report_workbook, fmei_files

    ' 2) 返戻内訳書（henr）CSVの処理
    ProcessHenrFiles file_system, report_workbook, henr_files

    ' 3) 増減点連絡書（zogn）CSVの処理
    ProcessZognFiles file_system, report_workbook, zogn_files
End Sub

Function ConvertToCircledNumber(month As Integer) As String
    Dim circledNumbers As Variant
    circledNumbers = Array("①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫")

    ' **1～12月の範囲内なら変換、範囲外ならそのまま返す**
    If month >= 1 And month <= 12 Then
        ConvertToCircledNumber = circledNumbers(month - 1)
    Else
        ConvertToCircledNumber = CStr(month)  ' 予期しない値ならそのまま
    End If
End Function

Sub ProcessFmeiFiles(file_system As Object, report_workbook As Workbook, fmei_files As Collection)
    Dim csvFileObj As Object, csv_sheet As Worksheet
    Dim sheet_name As String, insertIndex As Integer

    For Each csvFileObj In fmei_files
        ' 新しいシートを追加し、一意なシート名を設定（既存重複回避）
        sheet_name = file_system.GetBaseName(csvFileObj.Name)
        sheet_name = GetUniqueSheetName(report_workbook, sheet_name)
        insertIndex = Application.WorksheetFunction.Min(3, report_workbook.Sheets.Count + 1)
        Set csv_sheet = report_workbook.Sheets.Add(After:=report_workbook.Sheets(insertIndex - 1))
        csv_sheet.Name = sheet_name

        ' CSVデータをインポートし転記（列マッピングは"振込額明細書"定義を使用）
        ImportCsvData csvFileObj.Path, csv_sheet, "振込額明細書"
        ' 当該データの詳細分類転記（過去月入金＝返戻再請求の検出等）
        TransferBillingDetails report_workbook, csvFileObj.Name
    Next csvFileObj
End Sub

Sub ProcessHenrFiles(file_system As Object, report_workbook As Workbook, henr_files As Collection)
    Dim csvFileObj As Object, csv_sheet As Worksheet
    Dim sheet_name As String, insertIndex As Integer

    For Each csvFileObj In henr_files
        sheet_name = file_system.GetBaseName(csvFileObj.Name)
        sheet_name = GetUniqueSheetName(report_workbook, sheet_name)
        insertIndex = Application.WorksheetFunction.Min(3, report_workbook.Sheets.Count + 1)
        Set csv_sheet = report_workbook.Sheets.Add(After:=report_workbook.Sheets(insertIndex - 1))
        csv_sheet.Name = sheet_name

        ImportCsvData csvFileObj.Path, csv_sheet, "返戻内訳書"
        ' 返戻データ（過去未収＝返戻・査定）の詳細シート反映
        TransferBillingDetails report_workbook, csvFileObj.Name
    Next csvFileObj
End Sub

Sub ProcessZognFiles(file_system As Object, report_workbook As Workbook, zogn_files As Collection)
    Dim csvFileObj As Object, csv_sheet As Worksheet
    Dim sheet_name As String, insertIndex As Integer

    For Each csvFileObj In zogn_files
        sheet_name = file_system.GetBaseName(csvFileObj.Name)
        sheet_name = GetUniqueSheetName(report_workbook, sheet_name)
        insertIndex = Application.WorksheetFunction.Min(3, report_workbook.Sheets.Count + 1)
        Set csv_sheet = report_workbook.Sheets.Add(After:=report_workbook.Sheets(insertIndex - 1))
        csv_sheet.Name = sheet_name

        ImportCsvData csvFileObj.Path, csv_sheet, "増減点連絡書"
        ' 減点（未請求扱い）データの詳細シート反映
        TransferBillingDetails report_workbook, csvFileObj.Name
    Next csvFileObj
End Sub

Function SelectCsvFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "CSVフォルダを選択してください"
        If .Show = -1 Then
            SelectCsvFolder = .SelectedItems(1) & "\"
        Else
            MsgBox "フォルダが選択されませんでした。処理を中止します。", vbExclamation, "エラー"
            SelectCsvFolder = ""
        End If
    End With
End Function

Function IsFolderEmpty(folderPath As String) As Boolean
    Dim file_system_local As Object, folderObj As Object
    Set file_system_local = CreateObject("Scripting.FileSystemObject")
    If Not file_system_local.FolderExists(folderPath) Then
        IsFolderEmpty = True
        Exit Function
    End If
    Set folderObj = file_system_local.GetFolder(folderPath)
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

Function FindAllFixfFiles(file_system As Object, csv_folder As String) As Collection
    Dim csv_file As Object
    Dim fixf_files_local As New Collection
    ' フォルダ内の全ファイルから名前に`fixf`を含むCSVを収集
    For Each csv_file In file_system.GetFolder(csv_folder).Files
        If LCase(file_system.GetExtensionName(csv_file.Name)) = "csv" And InStr(LCase(csv_file.Name), "fixf") > 0 Then
            fixf_files_local.Add csv_file
        End If
    Next csv_file
    Set FindAllFixfFiles = fixf_files_local
End Function

Function FindOrCreateReport(save_path As String, invoice_year As String, invoice_month As String, template_path As String) As String
    Dim file_system_local As Object, newWb As Workbook
    Dim reportPath As String, reportName As String
    Dim csvYYMM As String, era_code As String, eraYear As Integer
    Set file_system_local = CreateObject("Scripting.FileSystemObject")
    ' RYYMM形式のファイル名を生成
    If CInt(invoice_year) >= 2019 Then
        era_code = "5"  ' 令和
        eraYear = CInt(invoice_year) - 2018
    ElseIf CInt(invoice_year) >= 1989 Then
        era_code = "4"  ' 平成
        eraYear = CInt(invoice_year) - 1988
    ElseIf CInt(invoice_year) >= 1926 Then
        era_code = "3"  ' 昭和
        eraYear = CInt(invoice_year) - 1925
    ElseIf CInt(invoice_year) >= 1912 Then
        era_code = "2"  ' 大正
        eraYear = CInt(invoice_year) - 1911
    Else
        era_code = "1"  ' 明治
        eraYear = CInt(invoice_year) - 1867
    End If
    csvYYMM = era_code & Format(eraYear, "00") & invoice_month  ' RYYMM文字列
    reportName = "保険請求管理報告書_" & csvYYMM & ".xlsm"
    reportPath = save_path & "\" & reportName
    If Not file_system_local.FileExists(reportPath) Then
        On Error Resume Next
        Set newWb = Workbooks.Add(template_path)
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
    If file_system_local.FileExists(reportPath) Then
        FindOrCreateReport = reportPath
    Else
        FindOrCreateReport = ""  ' 作成に失敗した場合
    End If
End Function

Sub SetTemplateInfo(newBook As Workbook, invoice_year As String, invoice_month As String)
    Dim sheet_a As Worksheet, sheet_b As Worksheet
    Dim receiptYear As Integer, receiptMonth As Integer
    Dim treatmentMonth_shadow As Integer, treatmentYear_shadow As Integer, send_date As String

    ' **西暦年と調剤月の計算**
    receiptYear = CInt(invoice_year)
    receiptMonth = CInt(invoice_month)

    ' **請求月の計算**
    treatmentMonth_shadow = receiptMonth - 1
    If treatmentMonth_shadow <= 0 Then
        treatmentYear_shadow = receiptYear - 1
        treatmentMonth_shadow = 12
    End If
    send_date = receiptMonth & "月10日請求分"

    ' **シートA, Bを取得**
    Set sheet_a = newBook.Sheets("A")
    Set sheet_b = newBook.Sheets("B")

    ' **シート名変更**
    sheet_a.Name = "R" & (treatmentYear_shadow - 2018) & "." & treatmentMonth_shadow
    sheet_b.Name = ConvertToCircledNumber(treatmentMonth_shadow)

    ' **情報転記**
    sheet_a.Range("G2").Value = treatmentYear_shadow & "年" & treatmentMonth_shadow & "月調剤分"
    sheet_a.Range("I2").Value = send_date
    sheet_a.Range("J2").Value = ThisWorkbook.Sheets("設定").Range("B1").Value
    sheet_b.Range("H1").Value = treatmentYear_shadow & "年" & treatmentMonth_shadow & "月調剤分"
    sheet_b.Range("J1").Value = send_date
    sheet_b.Range("L1").Value = ThisWorkbook.Sheets("設定").Range("B1").Value
End Sub

Sub GetYearMonthFromFixf(fixfFilePath As String, ByRef invoice_year As String, ByRef invoice_month As String)
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
    invoice_year = yearStr
    invoice_month = monthStr
End Sub

Sub GetYearMonthFromCSV(file_system As Object, csv_folder As String, ByRef invoice_year As String, ByRef invoice_month As String)
    Dim csv_file As Object, ts As Object
    Dim lineText As String
    Dim era_code As String, year_code As String, month_code As String
    Dim westernYear As Integer

    ' フォルダ内のCSVファイルから先頭行のGYYMMコードを取得（対象ファイル以外はスキップ）
    For Each csv_file In file_system.GetFolder(csv_folder).Files
        If LCase(file_system.GetExtensionName(csv_file.Name)) = "csv" _
           And (InStr(LCase(csv_file.Name), "fixf") > 0 _
                Or InStr(LCase(csv_file.Name), "fmei") > 0 _
                Or InStr(LCase(csv_file.Name), "henr") > 0 _
                Or InStr(LCase(csv_file.Name), "zogn") > 0) Then
            Set ts = file_system.OpenTextFile(csv_file.Path, 1, False, -2)  ' テキストストリーム (読み取り専用, UTF-8)
            Do While Not ts.AtEndOfStream
                lineText = ts.ReadLine
                If Len(lineText) >= 5 Then
                    era_code = Left(lineText, 1)        ' 元号コード (1:明治,2:大正,3:昭和,4:平成,5:令和)
                    year_code = Mid(lineText, 2, 2)     ' 元号年（2桁）
                    month_code = Right(lineText, 2)     ' 月（2桁）
                    ' 元号コード＋年を西暦年に変換
                    Select Case era_code
                        Case "5": westernYear = 2018 + CInt(year_code)   ' 令和 (2019=令和元年)
                        Case "4": westernYear = 1988 + CInt(year_code)   ' 平成 (1989=平成元年)
                        Case "3": westernYear = 1925 + CInt(year_code)   ' 昭和 (1926=昭和元年)
                        Case "2": westernYear = 1911 + CInt(year_code)   ' 大正 (1912=大正元年)
                        Case "1": westernYear = 1867 + CInt(year_code)   ' 明治 (1868=明治元年)
                        Case Else: westernYear = 2000 + CInt(year_code)  ' 仮定
                    End Select
                    invoice_year = CStr(westernYear)
                    invoice_month = month_code
                    Exit Do   ' 必要な情報取得できたのでループ終了
                End If
            Loop
            ts.Close
            If invoice_year <> "" And invoice_month <> "" Then Exit For
        End If
    Next csv_file
End Sub

Sub ImportCsvData(csvFilePath As String, ws As Worksheet, fileType As String)
    Dim file_system_local As Object, ts As Object
    Dim column_map As Object          ' 列マッピング定義（Dictionary）
    Dim lineText As String
    Dim dataArray As Variant
    Dim i As Long, j As Long, key

    On Error GoTo ImportError
    ' 画面更新と計算を一時停止（パフォーマンス向上）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 1. CSV項目のマッピング定義を取得（fileTypeに応じた列マッピング辞書）
    Set column_map = GetColumnMapping(fileType)

    ' 2. 対象シートをクリアし、ヘッダー行を作成
    ws.Cells.Clear
    j = 1
    For Each key In column_map.Keys
        ws.Cells(1, j).Value = column_map(key)  ' マッピング定義の値＝ヘッダ名
        j = j + 1
    Next key

    ' 3. CSVファイルを開いて読み込み、データ部をシートに転記
    Set file_system_local = CreateObject("Scripting.FileSystemObject")
    Set ts = file_system_local.OpenTextFile(csvFilePath, 1, False, -2)  ' UTF-8でテキストストリーム開く
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
            For Each key In column_map.Keys
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

Sub TransferBillingDetails(report_workbook As Workbook, csvFileName As String)
    Dim main_sheet As Worksheet, details_sheet As Worksheet
    Dim last_row_main As Long, i As Long
    Dim dispensingCode As String, dispensingYM As String
    Dim payer_code As String, payer_type As String
    Dim receipt_no As String
    Dim start_row_dict As Object                  ' 各カテゴリ開始行(Dictionary)
    Dim rebill_dict As Object, late_dict As Object, unpaid_dict As Object, assessment_dict As Object
    Dim category As String, start_row As Long
    Dim row_data As Variant
    Dim rebill_extra_count As Long, late_extra_count As Long, assessment_extra_count As Long         ' 追加行数算出用

    ' 1. シートオブジェクト取得
    Set main_sheet = report_workbook.Sheets(1)    ' メインシート（請求確定状況データ）
    Set details_sheet = report_workbook.Sheets(2) ' 詳細データシート

    ' 2. 処理対象の調剤年月コード(csvYYMM)を取得（メインシートB2セルの下4桁がRYYMM）
    Dim csvYYMM As String: csvYYMM = ""
    If main_sheet.Cells(2, 2).Value <> "" Then
        csvYYMM = Right(main_sheet.Cells(2, 2).Value, 4)
    End If

    ' 3. 請求先区分の判定（CSVファイル名の7文字目: "1"社保, "2"国保, その他=労災等）
    Dim base_name As String
    base_name = csvFileName
    If InStr(base_name, ".") > 0 Then base_name = Left(base_name, InStrRev(base_name, ".") - 1)
    If Len(base_name) >= 7 Then
        payer_code = Mid(base_name, 7, 1)
    Else
        payer_code = ""
    End If
    Select Case payer_code
        Case "1": payer_type = "社保"
        Case "2": payer_type = "国保"
        Case Else: payer_type = "労災"   ' 想定外のものは労災等その他扱い
    End Select

    ' 4. 詳細シート上の各カテゴリ開始行を取得してDictionaryに格納
    Set start_row_dict = CreateObject("Scripting.Dictionary")
    If payer_type = "社保" Then
        start_row_dict.Add "返戻再請求", GetStartRow(details_sheet, "社保返戻再請求")
        start_row_dict.Add "月遅れ請求", GetStartRow(details_sheet, "社保月遅れ請求")
        start_row_dict.Add "返戻・査定", GetStartRow(details_sheet, "社保返戻・査定")
        start_row_dict.Add "未請求扱い", GetStartRow(details_sheet, "社保未請求扱い")
    ElseIf payer_type = "国保" Then
        start_row_dict.Add "返戻再請求", GetStartRow(details_sheet, "国保返戻再請求")
        start_row_dict.Add "月遅れ請求", GetStartRow(details_sheet, "国保月遅れ請求")
        start_row_dict.Add "返戻・査定", GetStartRow(details_sheet, "国保返戻・査定")
        start_row_dict.Add "未請求扱い", GetStartRow(details_sheet, "国保未請求扱い")
    Else
        ' 労災等は詳細シート対象外（処理不要）
        Exit Sub
    End If

    ' 5. 過去データ分類用のDictionaryを準備
    Set rebill_dict = CreateObject("Scripting.Dictionary")     ' 返戻再請求（過去返戻分で当月入金）
    Set late_dict = CreateObject("Scripting.Dictionary")       ' 月遅れ請求（今回請求に含めた過去月分）
    Set unpaid_dict = CreateObject("Scripting.Dictionary")     ' 未請求扱い（請求漏れ・除外分）
    Set assessment_dict = CreateObject("Scripting.Dictionary") ' 返戻・査定（返戻・減点で未収）

    ' 6. メインシート（請求データ）の最終行を取得（D列に値がある最後の行）
    last_row_main = main_sheet.Cells(main_sheet.Rows.Count, "D").End(xlUp).Row

    ' 7. メインシートの各レコードを走査し、当月ではないデータを各カテゴリに振り分け
    For i = 2 To last_row_main
        dispensingCode = main_sheet.Cells(i, 2).Value            ' 元号付き調剤年月 (例: "50701")
        dispensingYM = ConvertToWesternDate(dispensingCode)      ' YY.MM形式に変換 (例: "07.01")
        If csvYYMM <> "" And Right(dispensingCode, 4) <> csvYYMM Then
            ' ※対象請求月(csvYYMM)と異なる＝過去月レセプト
            ' 転記用データ配列（患者氏名, 調剤年月(YY.MM), 医療機関名, 請求点数）を用意
            row_data = Array(main_sheet.Cells(i, 4).Value, dispensingYM, main_sheet.Cells(i, 5).Value, main_sheet.Cells(i, 10).Value)
            ' ファイル種別ごとに過去月データのカテゴリ振り分け
            If InStr(LCase(csvFileName), "fixf") > 0 Then
                ' `fixf`（請求確定）では過去月レセプトはすべて「月遅れ請求」に分類
                late_dict(main_sheet.Cells(i, 1).Value) = row_data
            ElseIf InStr(LCase(csvFileName), "fmei") > 0 Then
                ' 振込明細では過去月レセプトを「返戻再請求」として分類（前月返戻→当月入金）
                rebill_dict(main_sheet.Cells(i, 1).Value) = row_data
            ElseIf InStr(LCase(csvFileName), "zogn") > 0 Then
                ' 増減点連絡書では過去月レセプトを「未請求扱い」に分類（請求除外/未処理）
                unpaid_dict(main_sheet.Cells(i, 1).Value) = row_data
            ElseIf InStr(LCase(csvFileName), "henr") > 0 Then
                ' 返戻内訳書では過去月レセプトを「返戻・査定」に分類（査定等で未収）
                assessment_dict(main_sheet.Cells(i, 1).Value) = row_data
            End If
        End If
    Next i

    ' 8. 各カテゴリの件数超過分を算出（初期枠4件を超えた分）
    rebill_extra_count = 0: late_extra_count = 0: assessment_extra_count = 0
    If rebill_dict.Count > 4 Then rebill_extra_count = rebill_dict.Count - 4
    If late_dict.Count > 4 Then late_extra_count = late_dict.Count - 4
    If assessment_dict.Count > 4 Then assessment_extra_count = assessment_dict.Count - 4
    ' ※未請求扱い(unpaid_dict)は今後の請求候補として枠固定（超過行挿入しない）

    ' 9. 必要な追加行を各カテゴリセクションに挿入
    If rebill_extra_count + late_extra_count + assessment_extra_count > 0 Then
        If rebill_extra_count > 0 Then details_sheet.Rows(start_row_dict("月遅れ請求") + 1 & ":" & start_row_dict("月遅れ請求") + rebill_extra_count).Insert Shift:=xlDown
        If late_extra_count > 0 Then details_sheet.Rows(start_row_dict("返戻・査定") + 1 & ":" & start_row_dict("返戻・査定") + late_extra_count).Insert Shift:=xlDown
        If assessment_extra_count > 0 Then details_sheet.Rows(start_row_dict("未請求扱い") + 1 & ":" & start_row_dict("未請求扱い") + assessment_extra_count).Insert Shift:=xlDown
    End If

    ' 10. 各Dictionaryのデータを詳細シートに順次転記
    If rebill_dict.Count > 0 Then
        start_row = start_row_dict("返戻再請求")
        TransferData rebill_dict, details_sheet, start_row, payer_type
    End If
    If late_dict.Count > 0 Then
        start_row = start_row_dict("月遅れ請求")
        TransferData late_dict, details_sheet, start_row, payer_type
    End If
    If unpaid_dict.Count > 0 Then
        start_row = start_row_dict("未請求扱い")
        TransferData unpaid_dict, details_sheet, start_row, payer_type
    End If
    If assessment_dict.Count > 0 Then
        start_row = start_row_dict("返戻・査定")
        TransferData assessment_dict, details_sheet, start_row, payer_type
    End If

    ' 11. 完了メッセージ（処理区分ごとに表示）
    MsgBox payer_type & " のデータ転記が完了しました！", vbInformation, "処理完了"
End Sub

Sub TransferData(dataDict As Object, ws As Worksheet, start_row As Long, payer_type As String)
    If dataDict.Count = 0 Then Exit Sub

    Dim key As Variant, row_data As Variant
    Dim r As Long: r = start_row
    Dim payer_col As Long

    ' 社保はH列(8), 国保はI列(9)に種別を記載
    If payer_type = "社保" Then
        payer_col = 8
    ElseIf payer_type = "国保" Then
        payer_col = 9
    Else
        Exit Sub  ' その他（労災等）は対象外
    End If

    ' 各レコードを書き込み
    For Each key In dataDict.Keys
        row_data = dataDict(key)
        ws.Cells(r, 4).Value = row_data(0)          ' 患者氏名
        ws.Cells(r, 5).Value = row_data(1)          ' 調剤年月 (YY.MM形式)
        ws.Cells(r, 6).Value = row_data(2)          ' 医療機関名
        ws.Cells(r, payer_col).Value = payer_type   ' 請求先種別 (社保/国保)
        ws.Cells(r, payer_col).Font.Bold = True     ' 強調表示
        ws.Cells(r, 10).Value = row_data(3)         ' 請求点数
        r = r + 1
    Next key
End Sub

Function GetColumnMapping(fileType As String) As Object
    Dim column_map As Object
    Set column_map = CreateObject("Scripting.Dictionary")
    Dim k As Integer

    Select Case fileType
        Case "振込額明細書"
            ' 支払基金からの振込額明細CSV列 → シート列見出し の対応
            column_map.Add 2, "診療（調剤）年月"
            column_map.Add 5, "受付番号"
            column_map.Add 14, "氏名"
            column_map.Add 16, "生年月日"
            column_map.Add 22, "医療保険_請求点数"
            column_map.Add 23, "医療保険_決定点数"
            column_map.Add 24, "医療保険_一部負担金"
            column_map.Add 25, "医療保険_金額"
            ' 第1～第5公費 分の列（各10列間隔: 請求点数・決定点数・患者負担金・金額）
            For k = 1 To 5
                column_map.Add 33 + (k - 1) * 10, "第" & k & "公費_請求点数"
                column_map.Add 34 + (k - 1) * 10, "第" & k & "公費_決定点数"
                column_map.Add 35 + (k - 1) * 10, "第" & k & "公費_患者負担金"
                column_map.Add 36 + (k - 1) * 10, "第" & k & "公費_金額"
            Next k
            column_map.Add 82, "算定額合計"

        Case "請求確定状況"
            ' 請求確定CSV（fixfデータ）の列対応
            column_map.Add 4, "診療（調剤）年月"
            column_map.Add 5, "氏名"
            column_map.Add 7, "生年月日"
            column_map.Add 9, "医療機関名称"
            column_map.Add 13, "総合計点数"
            For k = 1 To 4
                column_map.Add 16 + (k - 1) * 3, "第" & k & "公費_請求点数"
            Next k
            column_map.Add 30, "請求確定状況"
            column_map.Add 31, "エラー区分"

        Case "増減点連絡書"
            column_map.Add 2, "調剤年月"
            column_map.Add 4, "受付番号"
            column_map.Add 11, "区分"
            column_map.Add 14, "老人減免区分"
            column_map.Add 15, "氏名"
            column_map.Add 21, "増減点数(金額)"
            column_map.Add 22, "事由"

        Case "返戻内訳書"
            column_map.Add 2, "調剤年月(YYMM)"
            column_map.Add 3, "受付番号"
            column_map.Add 4, "保険者番号"
            column_map.Add 7, "氏名"
            column_map.Add 9, "請求点数"
            column_map.Add 10, "薬剤一部負担金"
            column_map.Add 12, "一部負担金額"
            column_map.Add 13, "公費負担金額"
            column_map.Add 14, "事由コード"

        Case Else
            ' その他データ種別（必要に応じ追加）
    End Select

    Set GetColumnMapping = column_map
End Function

Function ConvertToWesternDate(dispensingCode As String) As String
    Dim era_code As String, yearNum As Integer, westernYear As Integer, monthPart As String
    If Len(dispensingCode) < 5 Then
        ConvertToWesternDate = ""
        Exit Function
    End If
    era_code = Left(dispensingCode, 1)                ' 元号コード
    yearNum = CInt(Mid(dispensingCode, 2, 2))         ' 元号年2桁
    monthPart = Right(dispensingCode, 2)             ' 月2桁
    Select Case era_code
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

Function GetUniqueSheetName(wb As Workbook, base_name As String) As String
    Dim newName As String, counter As Integer
    Dim sheet As Worksheet, exists As Boolean
    newName = base_name
    counter = 1
    Do
        exists = False
        For Each sheet In wb.Sheets
            If LCase(sheet.Name) = LCase(newName) Then
                exists = True
                Exit For
            End If
        Next sheet
        If exists Then
            newName = base_name & "_" & counter
            counter = counter + 1
        End If
    Loop While exists
    GetUniqueSheetName = newName
End Function

' --- （参考）半期ごとの請求誤差調査 ---
Sub InvestigateHalfYearDiscrepancy()
    ' ユーザー入力の年（西暦）と半期区分について、保存済み報告書ファイルを集計し請求点数と決定点数の差異を一覧表示する。
    Dim year_str As String, half_str As String
    Dim year_num As Integer, half_term As Integer
    Dim start_month As Integer, end_month As Integer
    Dim file_system_local As Object, folder_path As String
    Dim m As Integer
    Dim file_name As String, file_path As String
    Dim wb As Workbook, wsMain As Worksheet, deposit_sheet As Worksheet
    Dim total_points_claim As Long, total_points_decided As Long
    Dim era_code As String, era_year As Integer, era_year_code As String
    Dim result_msg As String

    ' 1. 対象年と半期を入力させる
    year_str = InputBox("調査する年（西暦）を入力してください:", "半期請求誤差調査")
    If year_str = "" Then Exit Sub
    half_str = InputBox("上期(1) または 下期(2) を指定してください:", "半期請求誤差調査")
    If half_str = "" Then Exit Sub
    If Not IsNumeric(year_str) Or Not IsNumeric(half_str) Then
        MsgBox "入力が正しくありません。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    year_num = CInt(year_str)
    half_term = CInt(half_str)
    If half_term <> 1 And half_term <> 2 Then
        MsgBox "半期の指定が不正です。1（上期）または2（下期）を指定してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    ' 2. 半期の開始月・終了月を設定
    If half_term = 1 Then
        start_month = 1: end_month = 6   ' 上期: 1～6月
    Else
        start_month = 7: end_month = 12  ' 下期: 7～12月
    End If

    Set file_system_local = CreateObject("Scripting.FileSystemObject")
    folder_path = GetSavePath()
    If folder_path = "" Then
        MsgBox "保存フォルダが設定されていません。", vbExclamation, "エラー"
        Exit Sub
    End If

    result_msg = year_num & "年 " & IIf(half_term = 1, "上期", "下期") & " 請求誤差調査結果:" & vbCrLf

    ' 3. 指定期間各月の報告書ファイルを順次開き、請求点数と決定点数を集計
    For m = start_month To end_month
        ' ファイル名（RYYMM形式）を構築
        If year_num >= 2019 Then
            era_code = "5": era_year = year_num - 2018   ' 令和
        ElseIf year_num >= 1989 Then
            era_code = "4": era_year = year_num - 1988   ' 平成
        ElseIf year_num >= 1926 Then
            era_code = "3": era_year = year_num - 1925   ' 昭和
        ElseIf year_num >= 1912 Then
            era_code = "2": era_year = year_num - 1911   ' 大正
        Else
            era_code = "1": era_year = year_num - 1867   ' 明治
        End If
        era_year_code = Format(era_year, "00")
        file_name = "保険請求管理報告書_" & era_year_code & Format(m, "00") & ".xlsm"
        file_path = folder_path & "\" & file_name

        If file_system_local.FileExists(file_path) Then
            ' 報告書Excelを開いて集計
            Set wb = Workbooks.Open(file_path, ReadOnly:=True)
            Set wsMain = wb.Sheets(1)  ' メインシート
            total_points_claim = 0: total_points_decided = 0

            ' メインシート「総合計点数」列合計を算出（請求点数合計）
            Dim header_cell As Range, colClaim As Long
            Set header_cell = wsMain.Rows(1).Find("総合計点数", LookAt:=xlWhole)
            If Not header_cell Is Nothing Then
                colClaim = header_cell.Column
                Dim last_row As Long
                last_row = wsMain.Cells(wsMain.Rows.Count, colClaim).End(xlUp).Row
                If last_row >= 2 Then
                    total_points_claim = Application.WorksheetFunction.Sum(wsMain.Range(wsMain.Cells(2, colClaim), wsMain.Cells(last_row, colClaim)))
                End If
            End If

            ' 振込額明細シート上の「決定点数」列合計を算出（実際の支払点数合計）
            Set deposit_sheet = Nothing
            Dim sheet As Worksheet, found_header As Range
            For Each sheet In wb.Sheets
                Set found_header = sheet.Rows(1).Find("決定点数", LookAt:=xlPart)
                If Not found_header Is Nothing Then
                    ' ヘッダに"決定点数"を含むシート（メインシートおよび詳細シートは除く）を振込額明細シートとみなす
                    If LCase(sheet.Name) <> LCase(wsMain.Name) And LCase(sheet.Name) <> LCase(wb.Sheets(2).Name) Then
                        Set deposit_sheet = sheet
                        Exit For
                    End If
                End If
            Next sheet
            If Not deposit_sheet Is Nothing Then
                ' 決定点数列（複数列: 社保・各公費）を順次合計
                Dim col As Long
                For col = 1 To deposit_sheet.UsedRange.Columns.Count
                    If InStr(deposit_sheet.Cells(1, col).Value, "決定点数") > 0 Then
                        Dim last_row_dep As Long
                        last_row_dep = deposit_sheet.Cells(deposit_sheet.Rows.Count, col).End(xlUp).Row
                        If last_row_dep >= 2 Then
                            total_points_decided = total_points_decided + Application.WorksheetFunction.Sum(deposit_sheet.Range(deposit_sheet.Cells(2, col), deposit_sheet.Cells(last_row_dep, col)))
                        End If
                    End If
                Next col
            End If

            wb.Close SaveChanges:=False

            ' 差異を算出しメッセージ文字列に追加
            Dim diff_points As Long
            diff_points = total_points_claim - total_points_decided
            If diff_points <> 0 Then
                result_msg = result_msg & "・" & year_num & "年" & m & "月: 請求=" & total_points_claim & " , 決定=" & total_points_decided & " （差異 " & diff_points & "点）" & vbCrLf
            End If
        Else
            result_msg = result_msg & "・" & year_num & "年" & m & "月: 報告書未作成" & vbCrLf
        End If
    Next m

    ' 4. 集計結果を表示
    MsgBox result_msg, vbInformation, "半期ごとの請求誤差調査結果"
End Sub

Function ConvertEraCodeToLetter(era_code_num As String) As String
    Select Case era_code_num
        Case "1": ConvertEraCodeToLetter = "M"
        Case "2": ConvertEraCodeToLetter = "T"
        Case "3": ConvertEraCodeToLetter = "S"
        Case "4": ConvertEraCodeToLetter = "H"
        Case "5": ConvertEraCodeToLetter = "R"
        Case Else: ConvertEraCodeToLetter = "E"
    End Select
End Function

' (追加) 請求年月（GYYMM形式）から調剤年月（RYYMM形式）を取得する関数
Function GetDispensingYearMonthFromFileName(fileName As String) As String
    Dim base_name As String, codePart As String
    Dim era_code_num As String, era_year_code As String, month_code As String
    Dim westernYear As Integer, western_month As Integer
    ' ファイル名から拡張子を除いた部分を取得
    base_name = fileName
    If InStr(fileName, ".") > 0 Then base_name = Left(fileName, InStrRev(fileName, ".") - 1)
    ' ファイル名中のGYYMMコード部分を抽出（末尾の5文字が数字の場合に取得）
    codePart = Right(base_name, 5)
    If Not codePart Like "*#####" Then
        ' 末尾5文字が数字でない場合、"TEST"挿入などを考慮して数字部分を検索
        Dim i As Long
        For i = Len(base_name) To 1 Step -1
            If Mid(base_name, i, 5) Like "#####" Then
                codePart = Mid(base_name, i, 5)
                Exit For
            End If
        Next i
    End If
    If Len(codePart) <> 5 Or Not IsNumeric(codePart) Then
        GetDispensingYearMonthFromFileName = ""  ' 変換失敗時は空文字
        Exit Function
    End If
    ' 請求年月（GYYMM）のコードから西暦年・月を取得
    era_code_num = Left(codePart, 1)          ' 元号コード（数字）
    era_year_code = Mid(codePart, 2, 2)       ' 元号年（2桁）
    month_code = Right(codePart, 2)          ' 月（2桁）
    Select Case era_code_num
        Case "5": westernYear = 2018 + CInt(era_year_code)   ' 令和(2019=令和元年)
        Case "4": westernYear = 1988 + CInt(era_year_code)   ' 平成(1989=平成元年)
        Case "3": westernYear = 1925 + CInt(era_year_code)   ' 昭和(1926=昭和元年)
        Case "2": westernYear = 1911 + CInt(era_year_code)   ' 大正(1912=大正元年)
        Case "1": westernYear = 1867 + CInt(era_year_code)   ' 明治(1868=明治元年)
        Case Else: westernYear = 2000 + CInt(era_year_code)  ' 不明なコード（仮）
    End Select
    western_month = CInt(month_code)
    ' 請求年月の1ヶ月前を調剤年月とする（月を1減算）
    western_month = western_month - 1
    If western_month = 0 Then
        westernYear = westernYear - 1
        western_month = 12
    End If
    ' 調剤年月を元号コード(アルファベット)付きのRYYMM形式文字列に変換
    Dim new_era_code As String, new_era_year As Integer, new_era_year_code As String, era_letter As String
    If westernYear >= 2019 Then
        new_era_code = "5": new_era_year = westernYear - 2018   ' 令和
    ElseIf westernYear >= 1989 Then
        new_era_code = "4": new_era_year = westernYear - 1988   ' 平成
    ElseIf westernYear >= 1926 Then
        new_era_code = "3": new_era_year = westernYear - 1925   ' 昭和
    ElseIf westernYear >= 1912 Then
        new_era_code = "2": new_era_year = westernYear - 1911   ' 大正
    Else
        new_era_code = "1": new_era_year = westernYear - 1867   ' 明治
    End If
    new_era_year_code = Format(new_era_year, "00")
    era_letter = ConvertEraCodeToLetter(new_era_code)
    GetDispensingYearMonthFromFileName = era_letter & new_era_year_code & Format(western_month, "00")
End Function

' (追加) GYYMM形式に基づき報告書ファイル名を生成する関数
Function GetReportFileName(fileName As String) As String
    Dim report_code As String
    report_code = GetDispensingYearMonthFromFileName(fileName)
    If report_code = "" Then
        GetReportFileName = ""
    Else
        GetReportFileName = "保険請求管理報告書_" & report_code & ".xlsm"
    End If
End Function