Option Explicit

' 定数定義
Const MAX_LINES_PER_SHEET As Long = 40
Const REQUIRED_SHEETS_COUNT As Integer = 6
Const BASE_DETAIL_ROWS As Integer = 4
Const TEMPLATE_FILE_NAME As String = "保険請求管理報告書テンプレート20250222.xltm"

Sub ProcessCsv()
    Dim csv_folder As String            ' CSVフォルダパス
    Dim file_system As Object          ' FileSystemObject
    Dim target_year As String, target_month As String  ' 処理対象の診療年・月（西暦）
    Dim save_path As String            ' 報告書保存先フォルダ
    Dim template_path As String        ' 報告書テンプレートファイル(.xltm)パス
    Dim fixf_files As New Collection, fmei_files As New Collection
    Dim henr_files As New Collection, zogn_files As New Collection
    Dim file_obj As Object

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
    For Each file_obj In file_system.GetFolder(csv_folder).Files
        If LCase(file_system.GetExtensionName(file_obj.Name)) = "csv" Then
            If InStr(LCase(file_obj.Name), "fixf") > 0 Then
                fixf_files.Add file_obj
            ElseIf InStr(LCase(file_obj.Name), "fmei") > 0 Then
                fmei_files.Add file_obj
            ElseIf InStr(LCase(file_obj.Name), "henr") > 0 Then
                henr_files.Add file_obj
            ElseIf InStr(LCase(file_obj.Name), "zogn") > 0 Then
                zogn_files.Add file_obj
            End If
        End If
    Next file_obj

    ' 4.1 対象CSVファイルが一つもない場合、処理を中止
    If fixf_files.Count = 0 And fmei_files.Count = 0 And henr_files.Count = 0 And zogn_files.Count = 0 Then
        MsgBox "選択したフォルダには処理対象のCSVファイルがありません。処理を中止します。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 5. fixfファイルがない場合は通常のCSV処理に切り替え
    If fixf_files.Count = 0 Then
        ProcessWithoutFixf file_system, csv_folder, save_path, template_path
        Exit Sub
    End If

    ' 6. 請求確定CSV（fixf）の処理
    ProcessFixfFiles file_system, fixf_files, save_path, template_path

    ' 7. 各種明細CSV（fmei, henr, zogn）の処理
    ProcessCsvFilesByType file_system, fmei_files, "振込額明細書"
    ProcessCsvFilesByType file_system, henr_files, "返戻内訳書"
    ProcessCsvFilesByType file_system, zogn_files, "増減点連絡書"

    ' 8. 完了メッセージ
    MsgBox "CSVファイルの処理が完了しました！", vbInformation, "完了"
End Sub

Sub ProcessWithoutFixf(file_system As Object, csv_folder As String, save_path As String, template_path As String)
    Dim target_year As String, target_month As String
    target_year = "": target_month = ""
    ' CSV内容から診療年月を取得
    GetYearMonthFromCsv file_system, csv_folder, target_year, target_month
    If target_year = "" Or target_month = "" Then
        MsgBox "CSVファイルから診療年月を取得できませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 出力報告書ファイル名を決定（診療年月RYYMM形式）
    Dim era_letter As String, era_year_val As Integer, year_code As String
    If CInt(target_year) >= 2019 Then
        era_letter = "R": era_year_val = CInt(target_year) - 2018  ' 令和
    ElseIf CInt(target_year) >= 1989 Then
        era_letter = "H": era_year_val = CInt(target_year) - 1988  ' 平成
    ElseIf CInt(target_year) >= 1926 Then
        era_letter = "S": era_year_val = CInt(target_year) - 1925  ' 昭和
    ElseIf CInt(target_year) >= 1912 Then
        era_letter = "T": era_year_val = CInt(target_year) - 1911  ' 大正
    Else
        era_letter = "M": era_year_val = CInt(target_year) - 1867  ' 明治
    End If
    year_code = Format(era_year_val, "00")
    Dim report_file_name As String, report_file_path As String
    report_file_name = "保険請求管理報告書_" & era_letter & year_code & Format(CInt(target_month), "00") & ".xlsm"
    report_file_path = save_path & "\" & report_file_name

    ' 既に報告書ファイルが存在する場合はスキップ
    If file_system.FileExists(report_file_path) Then
        MsgBox "既に対象年月の報告書ファイルが存在します: " & report_file_name, vbInformation, "処理スキップ"
        Exit Sub
    End If

    ' 報告書ブックを新規作成
    Dim report_wb As Workbook
    Set report_wb = Workbooks.Add(template_path)
    If report_wb Is Nothing Then
        MsgBox "診療年月 " & target_year & "年" & target_month & "月 のExcelファイルを作成できませんでした。", vbExclamation, "エラー"
        Exit Sub
    End If
    Application.DisplayAlerts = False
    report_wb.SaveAs Filename:=report_file_path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True

    ' テンプレート情報を設定
    SetTemplateInfo report_wb, target_year, target_month

    ' フォルダ内の全CSVファイルを種類別に処理（fixfなしのため振込明細・返戻・増減点のみ）
    Dim csv_file_obj As Object
    Dim base_name As String
    Dim fmei_files As New Collection, henr_files As New Collection, zogn_files As New Collection
    Dim era_code As String, era_year_code As String, gyy_mm As String
    ' 和暦コード付き年月(GYYMM)を生成
    If CInt(target_year) >= 2019 Then
        era_code = "5": era_year_code = Format(CInt(target_year) - 2018, "00")
    ElseIf CInt(target_year) >= 1989 Then
        era_code = "4": era_year_code = Format(CInt(target_year) - 1988, "00")
    ElseIf CInt(target_year) >= 1926 Then
        era_code = "3": era_year_code = Format(CInt(target_year) - 1925, "00")
    ElseIf CInt(target_year) >= 1912 Then
        era_code = "2": era_year_code = Format(CInt(target_year) - 1911, "00")
    Else
        era_code = "1": era_year_code = Format(CInt(target_year) - 1867, "00")
    End If
    gyy_mm = era_code & era_year_code & Format(CInt(target_month), "00")

    Dim sheet_exists As Boolean, ws As Worksheet, ws_new As Worksheet, insert_idx As Long

    ' 種類別のCSVを収集（対象年月のみ）
    For Each csv_file_obj In file_system.GetFolder(csv_folder).Files
        If LCase(file_system.GetExtensionName(csv_file_obj.Name)) = "csv" Then
            base_name = file_system.GetBaseName(csv_file_obj.Name)
            If InStr(LCase(base_name), "fmei") > 0 And Right(base_name, Len(gyy_mm)) = gyy_mm Then
                fmei_files.Add csv_file_obj
            ElseIf InStr(LCase(base_name), "henr") > 0 And Right(base_name, Len(gyy_mm)) = gyy_mm Then
                henr_files.Add csv_file_obj
            ElseIf InStr(LCase(base_name), "zogn") > 0 And Right(base_name, Len(gyy_mm)) = gyy_mm Then
                zogn_files.Add csv_file_obj
            End If
        End If
    Next csv_file_obj

    ' 振込額明細書（fmei）CSVの処理
    For Each csv_file_obj In fmei_files
        base_name = file_system.GetBaseName(csv_file_obj.Name)
        sheet_exists = False
        For Each ws In report_wb.Sheets
            If LCase(ws.Name) = LCase(base_name) Then
                sheet_exists = True
                Exit For
            End If
        Next ws
        If sheet_exists = False Then
            insert_idx = Application.WorksheetFunction.Min(3, report_wb.Sheets.Count + 1)
            Set ws_new = report_wb.Sheets.Add(After:=report_wb.Sheets(insert_idx - 1))
            ws_new.Name = base_name
            ImportCsvData csv_file_obj.Path, ws_new, "振込額明細書"
            TransferBillingDetails report_wb, csv_file_obj.Name
        End If
    Next csv_file_obj

    ' 返戻内訳書（henr）CSVの処理
    For Each csv_file_obj In henr_files
        base_name = file_system.GetBaseName(csv_file_obj.Name)
        sheet_exists = False
        For Each ws In report_wb.Sheets
            If LCase(ws.Name) = LCase(base_name) Then
                sheet_exists = True
                Exit For
            End If
        Next ws
        If sheet_exists = False Then
            insert_idx = Application.WorksheetFunction.Min(3, report_wb.Sheets.Count + 1)
            Set ws_new = report_wb.Sheets.Add(After:=report_wb.Sheets(insert_idx - 1))
            ws_new.Name = base_name
            ImportCsvData csv_file_obj.Path, ws_new, "返戻内訳書"
            TransferBillingDetails report_wb, csv_file_obj.Name
        End If
    Next csv_file_obj

    ' 増減点連絡書（zogn）CSVの処理
    For Each csv_file_obj In zogn_files
        base_name = file_system.GetBaseName(csv_file_obj.Name)
        sheet_exists = False
        For Each ws In report_wb.Sheets
            If LCase(ws.Name) = LCase(base_name) Then
                sheet_exists = True
                Exit For
            End If
        Next ws
        If sheet_exists = False Then
            insert_idx = Application.WorksheetFunction.Min(3, report_wb.Sheets.Count + 1)
            Set ws_new = report_wb.Sheets.Add(After:=report_wb.Sheets(insert_idx - 1))
            ws_new.Name = base_name
            ImportCsvData csv_file_obj.Path, ws_new, "増減点連絡書"
            TransferBillingDetails report_wb, csv_file_obj.Name
        End If
    Next csv_file_obj

    ' 保存してブックを閉じる
    report_wb.Save
    report_wb.Close False

    ' 完了メッセージ
    MsgBox "CSVファイルの処理が完了しました！", vbInformation, "完了"
End Sub

Sub ProcessFixfFiles(file_system As Object, fixf_files As Collection, save_path As String, template_path As String)
    Dim file_obj As Object
    Dim target_year As String, target_month As String
    Dim treatment_year As Integer, treatment_month As Integer
    Dim report_file_path As String, report_file_name As String
    Dim era_letter As String, year_code As String, month_code As String
    Dim report_wb As Workbook

    ' 収集したfixfファイルを調剤年月が古い順にソート
    SortFixfFiles file_system, fixf_files

    ' fixfファイルごとに処理
    For Each file_obj In fixf_files
        target_year = "": target_month = ""
        treatment_year = 0: treatment_month = 0
        ' fixfファイルから診療年月を取得
        GetYearMonthFromFixf file_obj.Path, target_year, target_month
        If target_year = "" Or target_month = "" Then
            MsgBox "ファイル " & file_obj.Name & " から診療年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextFixf
        End If
        ' 請求年月の1ヶ月前を診療年月とする（月を1減算）
        treatment_month = CInt(target_month) - 1
        If treatment_month <= 0 Then
            treatment_year = CInt(target_year) - 1
            treatment_month = 12
        Else
            treatment_year = CInt(target_year)
        End If
        ' 出力報告書ファイル名（診療年月RYYMM形式）を決定
        If treatment_year >= 2019 Then
            era_letter = "R": year_code = Format(treatment_year - 2018, "00")   ' 令和
        ElseIf treatment_year >= 1989 Then
            era_letter = "H": year_code = Format(treatment_year - 1988, "00")   ' 平成
        ElseIf treatment_year >= 1926 Then
            era_letter = "S": year_code = Format(treatment_year - 1925, "00")   ' 昭和
        ElseIf treatment_year >= 1912 Then
            era_letter = "T": year_code = Format(treatment_year - 1911, "00")   ' 大正
        Else
            era_letter = "M": year_code = Format(treatment_year - 1867, "00")   ' 明治
        End If
        month_code = Format(treatment_month, "00")
        report_file_name = "保険請求管理報告書_" & era_letter & year_code & month_code & ".xlsm"
        report_file_path = save_path & "\" & report_file_name
        ' 既存の報告書ファイルがある場合、そのfixfはスキップ（重複処理回避）
        If file_system.FileExists(report_file_path) Then
            GoTo NextFixf
        End If
        ' 報告書ブックを取得（存在しなければテンプレートから新規作成）
        If Not file_system.FileExists(report_file_path) Then
            Dim new_wb As Workbook
            Set new_wb = Workbooks.Add(template_path)
            If new_wb Is Nothing Then
                MsgBox "テンプレートファイルを開けませんでした: " & template_path, vbExclamation, "エラー"
                GoTo NextFixf
            End If
            Application.DisplayAlerts = False
            new_wb.SaveAs Filename:=report_file_path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            Application.DisplayAlerts = True
            new_wb.Close False
        End If
        On Error Resume Next
        Set report_wb = Workbooks.Open(report_file_path)
        On Error GoTo 0
        If report_wb Is Nothing Then
            MsgBox "ファイル " & report_file_path & " を開けませんでした。", vbExclamation, "エラー"
            GoTo NextFixf
        End If
        ' テンプレート情報を設定（診療年月等を更新）
        SetTemplateInfo report_wb, target_year, target_month

        ' 必要なシートを確保（Sheet1, Sheet2, およびカテゴリ用Sheet3～Sheet6）
        If report_wb.Sheets.Count < REQUIRED_SHEETS_COUNT Then
            Dim idx As Long
            For idx = report_wb.Sheets.Count + 1 To REQUIRED_SHEETS_COUNT
                report_wb.Sheets.Add After:=report_wb.Sheets(report_wb.Sheets.Count)
            Next idx
        End If
        ' 対象シート3～6をクリア
        Dim sh As Long
        For sh = 3 To 6
            If sh <= report_wb.Sheets.Count Then
                report_wb.Sheets(sh).Cells.Clear
            End If
        Next sh

        ' fixf CSVファイルを読み込み、カテゴリ別にデータ行を振り分け
        Dim text_stream As Object, line_text As String, lines_arr() As String
        Set text_stream = file_system.OpenTextFile(file_obj.Path, 1, False, -2)  ' UTF-8で読み込み
        line_text = text_stream.ReadAll
        text_stream.Close
        lines_arr = Split(line_text, vbCrLf)
        Dim column_map As Object: Set column_map = GetColumnMapping("請求確定状況")
        Dim data_lines_cat1 As New Collection, data_lines_cat2 As New Collection
        Dim header_skipped As Boolean: header_skipped = False
        Dim lineIdx As Long
        For lineIdx = LBound(lines_arr) To UBound(lines_arr)
            If Trim(lines_arr(lineIdx)) = "" Then GoTo ContinueLine
            If Not header_skipped Then
                header_skipped = True
                GoTo ContinueLine  ' ヘッダ行をスキップ
            End If
            If Left(lines_arr(lineIdx), 2) = "1," Then
                data_lines_cat1.Add lines_arr(lineIdx)
            ElseIf Left(lines_arr(lineIdx), 2) = "2," Then
                data_lines_cat2.Add lines_arr(lineIdx)
            Else
                ' 想定外の行は無視
            End If
ContinueLine:
        Next lineIdx

        ' カテゴリ1（社保）のデータ転記
        If data_lines_cat1.Count > 0 Then
            If data_lines_cat1.Count <= MAX_LINES_PER_SHEET Then
                WriteDataToSheet report_wb.Sheets(3), data_lines_cat1, column_map
            Else
                Dim tmp_col As Collection
                Set tmp_col = New Collection
                Dim j As Long
                For j = 1 To MAX_LINES_PER_SHEET
                    tmp_col.Add data_lines_cat1(j)
                Next j
                WriteDataToSheet report_wb.Sheets(3), tmp_col, column_map
                Set tmp_col = New Collection
                For j = MAX_LINES_PER_SHEET + 1 To data_lines_cat1.Count
                    tmp_col.Add data_lines_cat1(j)
                Next j
                WriteDataToSheet report_wb.Sheets(4), tmp_col, column_map
            End If
        End If

        ' カテゴリ2（国保）のデータ転記
        If data_lines_cat2.Count > 0 Then
            Dim start_sheet As Integer
            start_sheet = IIf(data_lines_cat1.Count > 0, 5, 3)  ' 社保データありならSheet5開始、なければSheet3開始
            If data_lines_cat2.Count <= MAX_LINES_PER_SHEET Then
                WriteDataToSheet report_wb.Sheets(start_sheet), data_lines_cat2, column_map
            Else
                Dim tmp_col2 As Collection
                Set tmp_col2 = New Collection
                For j = 1 To MAX_LINES_PER_SHEET
                    tmp_col2.Add data_lines_cat2(j)
                Next j
                WriteDataToSheet report_wb.Sheets(start_sheet), tmp_col2, column_map
                Set tmp_col2 = New Collection
                For j = MAX_LINES_PER_SHEET + 1 To data_lines_cat2.Count
                    tmp_col2.Add data_lines_cat2(j)
                Next j
                WriteDataToSheet report_wb.Sheets(start_sheet + 1), tmp_col2, column_map
            End If
        End If

        ' ブックを保存して閉じる
        report_wb.Save
        report_wb.Close False
NextFixf:
        ' 次のfixfファイルへ
    Next file_obj
End Sub

Sub ProcessCsvFilesByType(file_system As Object, csv_files As Collection, file_type_name As String)
    Dim file_obj As Object
    Dim save_path_local As String, template_path_local As String
    save_path_local = GetSavePath()
    template_path_local = GetTemplatePath()
    For Each file_obj In csv_files
        Dim report_file_name As String, report_file_path As String
        Dim base_name As String, sheet_name As String
        Dim report_wb As Workbook
        Dim sheet_exists As Boolean

        ' 出力報告書ファイル名（RYYMM形式）を生成
        report_file_name = GetReportFileName(file_obj.Name)
        If report_file_name = "" Then
            MsgBox "ファイル " & file_obj.Name & " から診療年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If
        report_file_path = save_path_local & "\" & report_file_name

        ' 報告書ファイルを取得（存在しなければテンプレートから作成）
        If Not file_system.FileExists(report_file_path) Then
            Dim new_book As Workbook
            Set new_book = Workbooks.Add(template_path_local)
            If new_book Is Nothing Then
                MsgBox "テンプレートファイルを開けませんでした: " & template_path_local, vbExclamation, "エラー"
                GoTo NextFile
            End If
            Application.DisplayAlerts = False
            new_book.SaveAs Filename:=report_file_path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            Application.DisplayAlerts = True
            new_book.Close False
        End If

        On Error Resume Next
        Set report_wb = Workbooks.Open(report_file_path)
        On Error GoTo 0
        If report_wb Is Nothing Then
            MsgBox "ファイル " & report_file_path & " を開けませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If

        ' 既に同名のシートが存在する場合はスキップ
        base_name = file_system.GetBaseName(file_obj.Name)
        sheet_exists = False
        Dim ws As Worksheet
        For Each ws In report_wb.Sheets
            If LCase(ws.Name) = LCase(base_name) Then
                sheet_exists = True
                Exit For
            End If
        Next ws
        If sheet_exists Then
            report_wb.Close False
            GoTo NextFile
        End If

        ' テンプレート情報を設定
        Dim year_val As String, month_val As String
        ' report_file_nameは "保険請求管理報告書_RYYMM.xlsm" 形式なので、西暦年と月を取得
        Dim era_letter As String, era_year_num As Integer, era_year_code As String
        era_letter = Mid(report_file_name, 11, 1)  ' 例: "R"
        era_year_code = Mid(report_file_name, 12, 2) ' 例: "07"
        month_val = Mid(report_file_name, 14, 2)     ' 例: "02"
        Select Case era_letter
            Case "R": year_val = CStr(2018 + CInt(era_year_code))
            Case "H": year_val = CStr(1988 + CInt(era_year_code))
            Case "S": year_val = CStr(1925 + CInt(era_year_code))
            Case "T": year_val = CStr(1911 + CInt(era_year_code))
            Case "M": year_val = CStr(1867 + CInt(era_year_code))
            Case Else: year_val = CStr(2000 + CInt(era_year_code))
        End Select
        SetTemplateInfo report_wb, year_val, month_val

        ' CSVデータをインポートして新規シートに転記
        sheet_name = base_name
        Dim insert_index As Long
        insert_index = Application.WorksheetFunction.Min(3, report_wb.Sheets.Count + 1)
        Dim ws_csv As Worksheet
        Set ws_csv = report_wb.Sheets.Add(After:=report_wb.Sheets(insert_index - 1))
        ws_csv.Name = sheet_name
        ImportCsvData file_obj.Path, ws_csv, file_type_name

        ' 詳細データを詳細シートに反映
        TransferBillingDetails report_wb, file_obj.Name

        ' 保存してブックを閉じる
        report_wb.Save
        report_wb.Close False
NextFile:
        ' 次のCSVファイルへ
    Next file_obj
End Sub

Sub WriteDataToSheet(ws As Worksheet, data_col As Collection, column_map As Object)
    ' 指定ワークシートをクリアしてヘッダ行を書き込み
    ws.Cells.Clear
    Dim key As Variant
    Dim col_index As Long: col_index = 1
    For Each key In column_map.Keys
        ws.Cells(1, col_index).Value = column_map(key)
        col_index = col_index + 1
    Next key
    ' データ行を書き込み
    Dim row_index As Long: row_index = 2
    Dim arr As Variant
    Dim item_idx As Long
    For item_idx = 1 To data_col.Count
        arr = Split(data_col(item_idx), ",")
        Dim k As Long: k = 1
        For Each key In column_map.Keys
            If key - 1 <= UBound(arr) Then
                ws.Cells(row_index, k).Value = Trim(arr(key - 1))
            End If
            k = k + 1
        Next key
        row_index = row_index + 1
    Next item_idx
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

Function IsFolderEmpty(folder_path As String) As Boolean
    Dim fso_local As Object, folder_obj As Object
    Set fso_local = CreateObject("Scripting.FileSystemObject")
    If Not fso_local.FolderExists(folder_path) Then
        IsFolderEmpty = True
        Exit Function
    End If
    Set folder_obj = fso_local.GetFolder(folder_path)
    If folder_obj.Files.Count = 0 Then
        IsFolderEmpty = True   ' ファイルが一つもない
    Else
        IsFolderEmpty = False  ' ファイルが存在する
    End If
End Function

Function GetTemplatePath() As String
    GetTemplatePath = ThisWorkbook.Sheets(1).Range("B2").Value & "\" & TEMPLATE_FILE_NAME
End Function

Function GetSavePath() As String
    GetSavePath = ThisWorkbook.Sheets(1).Range("B3").Value
End Function

Function FindAllFixfFiles(file_system As Object, csv_folder As String) As Collection
    Dim f_obj As Object
    Dim fixf_coll As New Collection
    For Each f_obj In file_system.GetFolder(csv_folder).Files
        If LCase(file_system.GetExtensionName(f_obj.Name)) = "csv" And InStr(LCase(f_obj.Name), "fixf") > 0 Then
            fixf_coll.Add f_obj
        End If
    Next f_obj
    Set FindAllFixfFiles = fixf_coll
End Function

Function FindOrCreateReport(save_path As String, target_year As String, target_month As String, template_path As String) As String
    Dim fso_local As Object
    Dim new_wb As Workbook
    Dim report_path As String, report_name As String
    Dim csv_yymm As String, era_code As String, era_year_val As Integer
    Set fso_local = CreateObject("Scripting.FileSystemObject")
    ' RYYMM形式のファイル名を生成
    If CInt(target_year) >= 2019 Then
        era_code = "5"
        era_year_val = CInt(target_year) - 2018
    ElseIf CInt(target_year) >= 1989 Then
        era_code = "4"
        era_year_val = CInt(target_year) - 1988
    ElseIf CInt(target_year) >= 1926 Then
        era_code = "3"
        era_year_val = CInt(target_year) - 1925
    ElseIf CInt(target_year) >= 1912 Then
        era_code = "2"
        era_year_val = CInt(target_year) - 1911
    Else
        era_code = "1"
        era_year_val = CInt(target_year) - 1867
    End If
    csv_yymm = era_code & Format(era_year_val, "00") & target_month  ' RYYMM文字列
    report_name = "保険請求管理報告書_" & csv_yymm & ".xlsm"
    report_path = save_path & "\" & report_name
    If Not fso_local.FileExists(report_path) Then
        On Error Resume Next
        Set new_wb = Workbooks.Add(template_path)
        On Error GoTo 0
        If new_wb Is Nothing Then
            FindOrCreateReport = ""
            Exit Function
        End If
        Application.DisplayAlerts = False
        new_wb.SaveAs Filename:=report_path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Application.DisplayAlerts = True
        new_wb.Close False
    End If
    If fso_local.FileExists(report_path) Then
        FindOrCreateReport = report_path
    Else
        FindOrCreateReport = ""
    End If
End Function

Sub SetTemplateInfo(report_book As Workbook, target_year As String, target_month As String)
    Dim ws_main As Worksheet, ws_sub As Worksheet
    Dim receipt_year As Integer, receipt_month As Integer
    Dim treatment_year As Integer, treatment_month As Integer
    Dim send_date As String

    ' 西暦年と調剤月の計算
    receipt_year = CInt(target_year)
    receipt_month = CInt(target_month)

    ' 請求月の計算（請求月 = 調剤月の翌月）
    treatment_month = receipt_month - 1
    If treatment_month <= 0 Then
        treatment_year = receipt_year - 1
        treatment_month = 12
    Else
        treatment_year = receipt_year
    End If

    send_date = receipt_month & "月10日請求分"

    ' テンプレートシート（シート1: A, シート2: B）を取得
    Set ws_main = report_book.Sheets(1)
    Set ws_sub = report_book.Sheets(2)

    ' シート名変更（シート1を "R{令和YY}.{M}"形式, シート2を丸数字月に）
    ws_main.Name = "R" & (receipt_year - 2018) & "." & receipt_month
    ws_sub.Name = ConvertToCircledNumber(receipt_month)

    ' 情報転記（ヘッダ部）
    ws_main.Range("G2").Value = treatment_year & "年" & treatment_month & "月調剤分"
    ws_main.Range("I2").Value = send_date
    ws_main.Range("J2").Value = ThisWorkbook.Sheets("設定").Range("B1").Value
    ws_sub.Range("H1").Value = treatment_year & "年" & treatment_month & "月調剤分"
    ws_sub.Range("J1").Value = send_date
    ws_sub.Range("L1").Value = ThisWorkbook.Sheets("設定").Range("B1").Value
End Sub

Sub GetYearMonthFromFixf(fixf_file_path As String, ByRef target_year As String, ByRef target_month As String)
    Dim file_name As String, date_part As String
    Dim year_str As String, month_str As String
    file_name = Mid(fixf_file_path, InStrRev(fixf_file_path, "\") + 1)
    date_part = Mid(file_name, 18, 14)
    If Len(date_part) < 8 Then Exit Sub
    year_str = Left(date_part, 4)
    month_str = Mid(date_part, 5, 2)
    target_year = year_str
    target_month = month_str
End Sub

Sub GetYearMonthFromCsv(file_system As Object, csv_folder As String, ByRef target_year As String, ByRef target_month As String)
    Dim f_obj As Object, text_stream As Object
    Dim line_text As String
    Dim era_code As String, year_code As String, month_code As String
    Dim western_year As Integer

    For Each f_obj In file_system.GetFolder(csv_folder).Files
        If LCase(file_system.GetExtensionName(f_obj.Name)) = "csv" _
           And (InStr(LCase(f_obj.Name), "fixf") > 0 _
                Or InStr(LCase(f_obj.Name), "fmei") > 0 _
                Or InStr(LCase(f_obj.Name), "henr") > 0 _
                Or InStr(LCase(f_obj.Name), "zogn") > 0) Then
            Set text_stream = file_system.OpenTextFile(f_obj.Path, 1, False, -2)
            Do While Not text_stream.AtEndOfStream
                line_text = text_stream.ReadLine
                If Len(line_text) >= 5 Then
                    era_code = Left(line_text, 1)
                    year_code = Mid(line_text, 2, 2)
                    month_code = Right(line_text, 2)
                    Select Case era_code
                        Case "5": western_year = 2018 + CInt(year_code)
                        Case "4": western_year = 1988 + CInt(year_code)
                        Case "3": western_year = 1925 + CInt(year_code)
                        Case "2": western_year = 1911 + CInt(year_code)
                        Case "1": western_year = 1867 + CInt(year_code)
                        Case Else: western_year = 2000 + CInt(year_code)
                    End Select
                    target_year = CStr(western_year)
                    target_month = month_code
                    Exit Do
                End If
            Loop
            text_stream.Close
            If target_year <> "" And target_month <> "" Then Exit For
        End If
    Next f_obj
End Sub

Sub ImportCsvData(csv_file_path As String, ws As Worksheet, file_type As String)
    Dim file_system_local As Object, text_stream As Object
    Dim column_map As Object
    Dim line_text As String
    Dim data_array As Variant
    Dim row_index As Long, col_index As Long
    Dim key As Variant

    On Error GoTo ImportError
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set file_system_local = CreateObject("Scripting.FileSystemObject")
    Set text_stream = file_system_local.OpenTextFile(csv_file_path, 1, False, -2)
    Set column_map = GetColumnMapping(file_type)

    ' ヘッダ行を作成
    ws.Cells.Clear
    col_index = 1
    For Each key In column_map.Keys
        ws.Cells(1, col_index).Value = column_map(key)
        col_index = col_index + 1
    Next key

    ' CSVファイルを読み込み、データ部分を転記
    row_index = 2
    Dim is_header As Boolean: is_header = True
    Do While Not text_stream.AtEndOfStream
        line_text = text_stream.ReadLine
        data_array = Split(line_text, ",")
        If is_header Then
            is_header = False
        Else
            col_index = 1
            For Each key In column_map.Keys
                If key - 1 <= UBound(data_array) Then
                    ws.Cells(row_index, col_index).Value = Trim(data_array(key - 1))
                End If
                col_index = col_index + 1
            Next key
            row_index = row_index + 1
        End If
    Loop
    text_stream.Close

    ws.Cells.EntireColumn.AutoFit

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
ImportError:
    MsgBox "CSVデータ読込中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
    If Not text_stream Is Nothing Then text_stream.Close
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub TransferBillingDetails(report_wb As Workbook, csv_file_name As String)
    Dim ws_main As Worksheet, ws_details As Worksheet
    Dim last_row_main As Long, row As Long
    Dim dispensing_code As String, dispensing_ym As String
    Dim payer_code As String, payer_type As String
    Dim start_row_dict As Object
    Dim rebill_dict As Object, late_dict As Object, unpaid_dict As Object, assessment_dict As Object
    Dim row_data As Variant
    Dim a As Long, b As Long, c As Long

    Set ws_main = report_wb.Sheets(1)   ' メインシート
    Set ws_details = report_wb.Sheets(2) ' 詳細データシート

    ' 調剤年月コード(csvYYMM)を取得（メインシートB2セルの下4桁）
    Dim csv_yymm As String: csv_yymm = ""
    If ws_main.Cells(2, 2).Value <> "" Then
        csv_yymm = Right(CStr(ws_main.Cells(2, 2).Value), 4)
        If InStr(csv_yymm, "年") > 0 Or InStr(csv_yymm, "月") > 0 Then
            csv_yymm = Replace(Replace(csv_yymm, "年", ""), "月", "")
        End If
    End If

    ' 請求先区分の判定（ファイル名の7文字目: "1"社保, "2"国保, その他=労災等）
    Dim base_name As String
    base_name = csv_file_name
    If InStr(base_name, ".") > 0 Then base_name = Left(base_name, InStrRev(base_name, ".") - 1)
    If Len(base_name) >= 7 Then
        payer_code = Mid(base_name, 7, 1)
    Else
        payer_code = ""
    End If
    Select Case payer_code
        Case "1": payer_type = "社保"
        Case "2": payer_type = "国保"
        Case Else: payer_type = "労災"   ' 想定外は労災等
    End Select

    ' 詳細シート上の各カテゴリ開始行を取得
    Set start_row_dict = CreateObject("Scripting.Dictionary")
    If payer_type = "社保" Then
        start_row_dict.Add "返戻再請求", GetStartRow(ws_details, "社保返戻再請求")
        start_row_dict.Add "月遅れ請求", GetStartRow(ws_details, "社保月遅れ請求")
        start_row_dict.Add "返戻・査定", GetStartRow(ws_details, "社保返戻・査定")
        start_row_dict.Add "未請求扱い", GetStartRow(ws_details, "社保未請求扱い")
    ElseIf payer_type = "国保" Then
        start_row_dict.Add "返戻再請求", GetStartRow(ws_details, "国保返戻再請求")
        start_row_dict.Add "月遅れ請求", GetStartRow(ws_details, "国保月遅れ請求")
        start_row_dict.Add "返戻・査定", GetStartRow(ws_details, "国保返戻・査定")
        start_row_dict.Add "未請求扱い", GetStartRow(ws_details, "国保未請求扱い")
    Else
        Exit Sub  ' 労災等は詳細シート対象外
    End If

    ' 各カテゴリ辞書の初期化
    Set rebill_dict = CreateObject("Scripting.Dictionary")
    Set late_dict = CreateObject("Scripting.Dictionary")
    Set unpaid_dict = CreateObject("Scripting.Dictionary")
    Set assessment_dict = CreateObject("Scripting.Dictionary")

    ' メインシートの最終行(D列の最終値がある行)を取得
    last_row_main = ws_main.Cells(ws_main.Rows.Count, "D").End(xlUp).Row

    ' メインシート各レコードを走査し、当月ではないデータをカテゴリ別に振り分け
    For row = 2 To last_row_main
        dispensing_code = ws_main.Cells(row, 2).Value
        dispensing_ym = ConvertToWesternDate(dispensing_code)
        If csv_yymm <> "" And Right(dispensing_code, 4) <> csv_yymm Then
            ' 過去月のレセプト
            row_data = Array(ws_main.Cells(row, 4).Value, dispensing_ym, ws_main.Cells(row, 5).Value, ws_main.Cells(row, 10).Value)
            ' ファイル種別ごとにカテゴリ振り分け
            If InStr(LCase(csv_file_name), "fixf") > 0 Then
                ' fixfでは過去月レセプトを「月遅れ請求」に分類
                late_dict(ws_main.Cells(row, 1).Value) = row_data
            ElseIf InStr(LCase(csv_file_name), "fmei") > 0 Then
                ' 振込明細では過去月レセプトを「返戻再請求」に分類
                rebill_dict(ws_main.Cells(row, 1).Value) = row_data
            ElseIf InStr(LCase(csv_file_name), "zogn") > 0 Then
                ' 増減点連絡書では過去月レセプトを「未請求扱い」に分類
                unpaid_dict(ws_main.Cells(row, 1).Value) = row_data
            ElseIf InStr(LCase(csv_file_name), "henr") > 0 Then
                ' 返戻内訳書では過去月レセプトを「返戻・査定」に分類
                assessment_dict(ws_main.Cells(row, 1).Value) = row_data
            End If
        End If
    Next row

    ' 各カテゴリの件数超過分を算出（初期枠4件を超えた分）
    a = 0: b = 0: c = 0
    If rebill_dict.Count > BASE_DETAIL_ROWS Then a = rebill_dict.Count - BASE_DETAIL_ROWS
    If late_dict.Count > BASE_DETAIL_ROWS Then b = late_dict.Count - BASE_DETAIL_ROWS
    If assessment_dict.Count > BASE_DETAIL_ROWS Then c = assessment_dict.Count - BASE_DETAIL_ROWS
    ' 未請求扱い(unpaid_dict)は固定枠（超過行は挿入しない）

    ' 必要な追加行を各カテゴリセクションに挿入
    If a > 0 Then ws_details.Rows(start_row_dict("返戻再請求") + 1 & ":" & start_row_dict("返戻再請求") + a).Insert Shift:=xlDown
    If b > 0 Then ws_details.Rows(start_row_dict("月遅れ請求") + 1 & ":" & start_row_dict("月遅れ請求") + b).Insert Shift:=xlDown
    If c > 0 Then ws_details.Rows(start_row_dict("返戻・査定") + 1 & ":" & start_row_dict("返戻・査定") + c).Insert Shift:=xlDown

    ' 各Dictionaryのデータを詳細シートに転記
    Dim start_row As Long
    If rebill_dict.Count > 0 Then
        start_row = start_row_dict("返戻再請求")
        TransferData rebill_dict, ws_details, start_row, payer_type
    End If
    If late_dict.Count > 0 Then
        start_row = start_row_dict("月遅れ請求")
        TransferData late_dict, ws_details, start_row, payer_type
    End If
    If unpaid_dict.Count > 0 Then
        start_row = start_row_dict("未請求扱い")
        TransferData unpaid_dict, ws_details, start_row, payer_type
    End If
    If assessment_dict.Count > 0 Then
        start_row = start_row_dict("返戻・査定")
        TransferData assessment_dict, ws_details, start_row, payer_type
    End If

    ' 完了メッセージ（処理区分ごとに表示）
    MsgBox payer_type & " のデータ転記が完了しました！", vbInformation, "処理完了"
End Sub

Function GetStartRow(ws As Worksheet, category_name As String) As Long
    Dim found_cell As Range
    Set found_cell = ws.Cells.Find(what:=category_name, LookAt:=xlWhole)
    If Not found_cell Is Nothing Then
        GetStartRow = found_cell.Row
    Else
        GetStartRow = 0
    End If
End Function

Function GetUniqueSheetName(workbook_obj As Workbook, base_name As String) As String
    Dim new_name As String, counter As Integer
    Dim ws As Worksheet, exists As Boolean
    new_name = base_name
    counter = 1
    Do
        exists = False
        For Each ws In workbook_obj.Sheets
            If LCase(ws.Name) = LCase(new_name) Then
                exists = True
                Exit For
            End If
        Next ws
        If exists Then
            new_name = base_name & "_" & counter
            counter = counter + 1
        End If
    Loop While exists
    GetUniqueSheetName = new_name
End Function

' 半期ごとの請求誤差調査マクロ
Sub InvestigateHalfYearDiscrepancy()
    Dim year_str As String, half_str As String
    Dim year_num As Integer, half_val As Integer
    Dim start_month As Integer, end_month As Integer
    Dim file_system As Object, folder_path As String
    Dim m As Integer
    Dim file_name As String, file_path As String
    Dim wb As Workbook, ws_main As Worksheet, ws_dep As Worksheet
    Dim total_points_claim As Long, total_points_decided As Long
    Dim era_code As String, era_year_val As Integer, era_yy As String, era_letter As String
    Dim result_msg As String

    ' 1. 対象年と半期を入力
    year_str = InputBox("調査する年（西暦）を入力してください:", "半期請求誤差調査")
    If year_str = "" Then Exit Sub
    half_str = InputBox("上期(1) または 下期(2) を指定してください:", "半期請求誤差調査")
    If half_str = "" Then Exit Sub
    If Not IsNumeric(year_str) Or Not IsNumeric(half_str) Then
        MsgBox "入力が正しくありません。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    year_num = CInt(year_str)
    half_val = CInt(half_str)
    If half_val <> 1 And half_val <> 2 Then
        MsgBox "半期の指定が不正です。1（上期）または2（下期）を指定してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    ' 2. 半期の開始月・終了月を設定
    If half_val = 1 Then
        start_month = 1: end_month = 6   ' 上期: 1～6月
    Else
        start_month = 7: end_month = 12  ' 下期: 7～12月
    End If

    Set file_system = CreateObject("Scripting.FileSystemObject")
    folder_path = GetSavePath()
    If folder_path = "" Then
        MsgBox "保存フォルダが設定されていません。", vbExclamation, "エラー"
        Exit Sub
    End If

    result_msg = year_num & "年 " & IIf(half_val = 1, "上期", "下期") & " 請求誤差調査結果:" & vbCrLf

    ' 3. 指定期間各月の報告書ファイルを順次開き、請求点数と決定点数を集計
    For m = start_month To end_month
        ' ファイル名（RYYMM形式）を構築
        If year_num >= 2019 Then
            era_code = "5": era_year_val = year_num - 2018   ' 令和
        ElseIf year_num >= 1989 Then
            era_code = "4": era_year_val = year_num - 1988   ' 平成
        ElseIf year_num >= 1926 Then
            era_code = "3": era_year_val = year_num - 1925   ' 昭和
        ElseIf year_num >= 1912 Then
            era_code = "2": era_year_val = year_num - 1911   ' 大正
        Else
            era_code = "1": era_year_val = year_num - 1867   ' 明治
        End If
        era_yy = Format(era_year_val, "00")
        era_letter = ConvertEraCodeToLetter(era_code)
        file_name = "保険請求管理報告書_" & era_letter & era_yy & Format(m, "00") & ".xlsm"
        file_path = folder_path & "\" & file_name

        If file_system.FileExists(file_path) Then
            Set wb = Workbooks.Open(file_path, ReadOnly:=True)
            Set ws_main = wb.Sheets(1)
            total_points_claim = 0: total_points_decided = 0
            ' メインシートの総合計点数列合計を算出（請求点数合計）
            Dim hdr_cell As Range, col_claim As Long
            Set hdr_cell = ws_main.Rows(1).Find("総合計点数", LookAt:=xlWhole)
            If Not hdr_cell Is Nothing Then
                col_claim = hdr_cell.Column
                Dim last_row As Long
                last_row = ws_main.Cells(ws_main.Rows.Count, col_claim).End(xlUp).Row
                If last_row >= 2 Then
                    total_points_claim = Application.WorksheetFunction.Sum(ws_main.Range(ws_main.Cells(2, col_claim), ws_main.Cells(last_row, col_claim)))
                End If
            End If
            ' 振込額明細シート上の決定点数列合計を算出（決定点数合計）
            Set ws_dep = Nothing
            Dim sheet_obj As Worksheet, found_hdr As Range
            For Each sheet_obj In wb.Sheets
                Set found_hdr = sheet_obj.Rows(1).Find("決定点数", LookAt:=xlPart)
                If Not found_hdr Is Nothing Then
                    If LCase(sheet_obj.Name) <> LCase(ws_main.Name) And LCase(sheet_obj.Name) <> LCase(wb.Sheets(2).Name) Then
                        Set ws_dep = sheet_obj
                        Exit For
                    End If
                End If
            Next sheet_obj
            If Not ws_dep Is Nothing Then
                Dim col_idx As Long
                For col_idx = 1 To ws_dep.UsedRange.Columns.Count
                    If InStr(ws_dep.Cells(1, col_idx).Value, "決定点数") > 0 Then
                        Dim last_row_dep As Long
                        last_row_dep = ws_dep.Cells(ws_dep.Rows.Count, col_idx).End(xlUp).Row
                        If last_row_dep >= 2 Then
                            total_points_decided = total_points_decided + Application.WorksheetFunction.Sum(ws_dep.Range(ws_dep.Cells(2, col_idx), ws_dep.Cells(last_row_dep, col_idx)))
                        End If
                    End If
                Next col_idx
            End If
            wb.Close SaveChanges:=False

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

Function ConvertToCircledNumber(month As Integer) As String
    Dim circled_numbers As Variant
    circled_numbers = Array("①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫")
    If month >= 1 And month <= 12 Then
        ConvertToCircledNumber = circled_numbers(month - 1)
    Else
        ConvertToCircledNumber = CStr(month)
    End If
End Function

Function ConvertToWesternDate(dispensing_code As String) As String
    Dim era_code As String, year_num As Integer, western_year As Integer, month_part As String
    If Len(dispensing_code) < 5 Then
        ConvertToWesternDate = ""
        Exit Function
    End If
    era_code = Left(dispensing_code, 1)
    year_num = CInt(Mid(dispensing_code, 2, 2))
    month_part = Right(dispensing_code, 2)
    Select Case era_code
        Case "5": western_year = 2018 + year_num   ' 令和
        Case "4": western_year = 1988 + year_num   ' 平成
        Case "3": western_year = 1925 + year_num   ' 昭和
        Case "2": western_year = 1911 + year_num   ' 大正
        Case "1": western_year = 1867 + year_num   ' 明治
        Case Else: western_year = 2000 + year_num
    End Select
    ConvertToWesternDate = Right(CStr(western_year), 2) & "." & month_part
End Function

Function GetColumnMapping(file_type As String) As Object
    Dim column_map As Object
    Set column_map = CreateObject("Scripting.Dictionary")
    Dim k As Integer

    Select Case file_type
        Case "振込額明細書"
            column_map.Add 2, "診療（調剤）年月"
            column_map.Add 5, "受付番号"
            column_map.Add 14, "氏名"
            column_map.Add 16, "生年月日"
            column_map.Add 22, "医療保険_請求点数"
            column_map.Add 23, "医療保険_決定点数"
            column_map.Add 24, "医療保険_一部負担金"
            column_map.Add 25, "医療保険_金額"
            ' 第1～第5公費（各10列間隔: 請求点数・決定点数・患者負担金・金額）
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
            ' その他のデータ種別があれば追加
    End Select

    Set GetColumnMapping = column_map
End Function

Function GetReportFileName(file_name As String) As String
    Dim report_code As String
    report_code = GetTreatmentYearMonthFromFileName(file_name)
    If report_code = "" Then
        GetReportFileName = ""
    Else
        GetReportFileName = "保険請求管理報告書_" & report_code & ".xlsm"
    End If
End Function

Function GetTreatmentYearMonthFromFileName(file_name As String) As String
    Dim base_name As String, code_part As String
    Dim era_code_num As String, era_year_code As String, month_code As String
    Dim western_year As Integer, western_month As Integer
    base_name = file_name
    If InStr(file_name, ".") > 0 Then base_name = Left(file_name, InStrRev(file_name, ".") - 1)
    code_part = Right(base_name, 5)
    If Not code_part Like "*#####" Then
        Dim i As Long
        For i = Len(base_name) To 1 Step -1
            If Mid(base_name, i, 5) Like "#####" Then
                code_part = Mid(base_name, i, 5)
                Exit For
            End If
        Next i
    End If
    If Len(code_part) <> 5 Or Not IsNumeric(code_part) Then
        GetTreatmentYearMonthFromFileName = ""
        Exit Function
    End If
    era_code_num = Left(code_part, 1)          ' 元号コード（数字）
    era_year_code = Mid(code_part, 2, 2)       ' 元号年（2桁）
    month_code = Right(code_part, 2)           ' 月（2桁）
    Select Case era_code_num
        Case "5": western_year = 2018 + CInt(era_year_code)   ' 令和(2019=令和元年)
        Case "4": western_year = 1988 + CInt(era_year_code)   ' 平成(1989=平成元年)
        Case "3": western_year = 1925 + CInt(era_year_code)   ' 昭和(1926=昭和元年)
        Case "2": western_year = 1911 + CInt(era_year_code)   ' 大正(1912=大正元年)
        Case "1": western_year = 1867 + CInt(era_year_code)   ' 明治(1868=明治元年)
        Case Else: western_year = 2000 + CInt(era_year_code)  ' 不明なコードの場合
    End Select
    western_month = CInt(month_code)
    western_month = western_month - 1
    If western_month = 0 Then
        western_year = western_year - 1
        western_month = 12
    End If
    Dim new_era_code As String, new_era_year As Integer, new_era_year_code As String, era_letter_code As String
    If western_year >= 2019 Then
        new_era_code = "5": new_era_year = western_year - 2018   ' 令和
    ElseIf western_year >= 1989 Then
        new_era_code = "4": new_era_year = western_year - 1988   ' 平成
    ElseIf western_year >= 1926 Then
        new_era_code = "3": new_era_year = western_year - 1925   ' 昭和
    ElseIf western_year >= 1912 Then
        new_era_code = "2": new_era_year = western_year - 1911   ' 大正
    Else
        new_era_code = "1": new_era_year = western_year - 1867   ' 明治
    End If
    new_era_year_code = Format(new_era_year, "00")
    era_letter_code = ConvertEraCodeToLetter(new_era_code)
    GetTreatmentYearMonthFromFileName = era_letter_code & new_era_year_code & Format(western_month, "00")
End Function

Function ConvertEraCodeToLetter(era_code As String) As String
    Select Case era_code
        Case "1": ConvertEraCodeToLetter = "M"
        Case "2": ConvertEraCodeToLetter = "T"
        Case "3": ConvertEraCodeToLetter = "S"
        Case "4": ConvertEraCodeToLetter = "H"
        Case "5": ConvertEraCodeToLetter = "R"
        Case Else: ConvertEraCodeToLetter = "E"
    End Select
End Function

Sub SortFixfFiles(file_system As Object, fixf_files As Collection)
    Dim i As Long, j As Long
    For i = 1 To fixf_files.Count - 1
        For j = i + 1 To fixf_files.Count
            Dim year1 As String, month1 As String
            Dim year2 As String, month2 As String
            year1 = "": month1 = ""
            year2 = "": month2 = ""
            GetYearMonthFromFixf fixf_files(i).Path, year1, month1
            GetYearMonthFromFixf fixf_files(j).Path, year2, month2
            If year1 <> "" And month1 <> "" And year2 <> "" And month2 <> "" Then
                If (year1 & month1) > (year2 & month2) Then
                    Dim temp_obj As Object
                    Set temp_obj = fixf_files(i)
                    Set fixf_files(i) = fixf_files(j)
                    Set fixf_files(j) = temp_obj
                End If
            End If
        Next j
    Next i
End Sub

Sub SortFilesBySuffixCode(file_coll As Collection, file_system As Object)
    Dim i As Long, j As Long
    For i = 1 To file_coll.Count - 1
        For j = i + 1 To file_coll.Count
            Dim code1 As String, code2 As String
            code1 = Right(file_system.GetBaseName(file_coll(i).Name), 5)
            code2 = Right(file_system.GetBaseName(file_coll(j).Name), 5)
            If Len(code1) = 5 And Len(code2) = 5 And IsNumeric(code1) And IsNumeric(code2) Then
                If CDbl(code1) > CDbl(code2) Then
                    Dim temp_obj As Object
                    Set temp_obj = file_coll(i)
                    Set file_coll(i) = file_coll(j)
                    Set file_coll(j) = temp_obj
                End If
            End If
        Next j
    Next i
End Sub