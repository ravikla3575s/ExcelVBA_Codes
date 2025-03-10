Option Explicit

' 定数定義
Const MAX_LINES_PER_SHEET As Long = 40
Const REQUIRED_SHEETS_COUNT As Integer = 6
Const BASE_DETAIL_ROWS As Integer = 4

' テンプレート・保存先パス
Dim template_path As String
Dim save_path As String

Sub CreateReportsFromCSV()
    On Error GoTo ErrorHandler
    
    ' パスの設定
    template_path = ThisWorkbook.Sheets("設定").Range("B2").Value & "\保険請求管理報告書テンプレート20250222.xltm"
    save_path = ThisWorkbook.Sheets("設定").Range("B3").Value
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim csv_folder As String            ' CSVフォルダパス
    Dim file_system As Object          ' FileSystemObject
    Dim billing_year As String, billing_month As String  ' 処理対象の診療年・月（西暦）
    Dim fixf_files As New Collection, fmei_files As New Collection
    Dim henr_files As New Collection, zogn_files As New Collection
    Dim file_obj As Object

    ' 1. CSVフォルダをユーザーに選択させる
    csv_folder = SelectCsvFolder()
    If csv_folder = "" Then Exit Sub  ' ユーザーがキャンセルした場合

    ' 2. フォルダが空なら処理を中止
    If IsFolderEmpty(csv_folder) Then
        MsgBox "選択したフォルダにはCSVファイルがありません。処理を中止します。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 3. テンプレートパス・保存先フォルダの存在確認
    If template_path = "" Or save_path = "" Then
        MsgBox "テンプレートパスまたは保存先フォルダが設定されていません。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 4. FileSystemObjectの用意
    Set file_system = CreateObject("Scripting.FileSystemObject")

    ' 4. フォルダ内の全CSVファイルを種類別に収集（fixf, fmei, henr, zogn）
    For Each file_obj In file_system.GetFolder(csv_folder).Files
        If LCase(file_system.GetExtensionName(file_obj.Name)) = "csv" Then
            If InStr(LCase(file_obj.Name), "fixf") > 0 Then
                fixf_files.Add file_obj
                Set fixf_files = SortFileCollection(fixf_files, file_system, "fixf")
            ElseIf InStr(LCase(file_obj.Name), "fmei") > 0 Then
                fmei_files.Add file_obj
                Set fmei_files = SortFileCollection(fmei_files, file_system, "fmei")
            ElseIf InStr(LCase(file_obj.Name), "henr") > 0 Then
                henr_files.Add file_obj
                Set henr_files = SortFileCollection(henr_files, file_system, "henr")
            ElseIf InStr(LCase(file_obj.Name), "zogn") > 0 Then
                zogn_files.Add file_obj
                Set zogn_files = SortFileCollection(zogn_files, file_system, "zogn")
            End If
        End If
    Next file_obj


    ' 6. 対象CSVファイルが一つもない場合、処理を中止
    If fixf_files.Count = 0 And fmei_files.Count = 0 And henr_files.Count = 0 And zogn_files.Count = 0 Then
        MsgBox "選択したフォルダには処理対象のCSVファイルがありません。処理を中止します。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 7. fixfファイルとfmeiファイルの有無による処理分岐
    If fixf_files.Count > 0 Then
        CreateReportFiles file_system, fixf_files, save_path, template_path
    End If
    If fmei_files.Count > 0 Then
        CreateReportFiles file_system, fmei_files, save_path, template_path
    End If

    ' 8. 各種明細CSV（fmei, henr, zogn）の処理
    ProcessCsvFilesByType file_system, fixf_files, "請求確定状況"
    ProcessCsvFilesByType file_system, fmei_files, "振込額明細書"
    ProcessCsvFilesByType file_system, henr_files, "返戻内訳書" 
    ProcessCsvFilesByType file_system, zogn_files, "増減点連絡書"
    
    ' 9. 完了メッセージ
    MsgBox "CSVファイルの処理が完了しました！", vbInformation, "完了"

    ' オブジェクトの解放処理を追加
    Set file_system = Nothing
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    MsgBox "メイン処理でエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "発生箇所: CreateReportsFromCSV", _
           vbCritical, "エラー"
    
    ' クリーンアップ処理
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    ' 開いているワークブックをクリーンアップ
    Call CleanupObjects(Workbooks)
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

Function CreateReportFiles(file_system As Object, files As Collection, save_path As String, template_path As String)
    On Error GoTo ErrorHandler
    
    ' テンプレートファイルの存在確認を追加
    If Not file_system.FileExists(template_path) Then
        MsgBox "テンプレートファイルが見つかりません。" & vbCrLf & _
               "パス: " & template_path & vbCrLf & _
               "発生箇所: CreateReportFiles", _
               vbCritical, "エラー"
        Exit Function
    End If
    
    ' 保存先フォルダの存在確認を追加
    If Not file_system.FolderExists(save_path) Then
        MsgBox "保存先フォルダが見つかりません。" & vbCrLf & _
               "パス: " & save_path & vbCrLf & _
               "発生箇所: CreateReportFiles", _
               vbCritical, "エラー"
        Exit Function
    End If
    
    Dim file As Object
    Dim billing_year As String, billing_month As String
    Dim era_letter As String, era_year_val As Integer, year_code As String
    
    For Each file In files
        On Error GoTo ErrorHandler
        
        ' CSVから年月を取得
        billing_year = "": billing_month = ""
        
        ' ファイル種類によって年月取得方法を変える
        If InStr(LCase(file.Name), "fixf") > 0 Then
            If Len(file.Name) < 25 Then
                MsgBox "FIXFファイルの形式が不正です。" & vbCrLf & _
                       "ファイル名: " & file.Name & vbCrLf & _
                       "必要な長さ: 25文字以上", _
                       vbExclamation, "CreateReportFiles - エラー"
                GoTo NextFile
            End If
            billing_year = Mid(file.Name, 18, 4)
            billing_month = Mid(file.Name, 22, 2)
            
        ElseIf InStr(LCase(file.Name), "fmei") > 0 Then
            If Len(file.Name) < 22 Then
                MsgBox "FMEIファイルの形式が不正です。" & vbCrLf & _
                       "ファイル名: " & file.Name & vbCrLf & _
                       "必要な長さ: 22文字以上", _
                       vbExclamation, "CreateReportFiles - エラー"
                GoTo NextFile
            End If
            Dim era_code As String
            era_code = Mid(file.Name, 18, 1)
            billing_month = Mid(file.Name, 21, 2)
            
            ' 元号コードを設定
            Select Case era_code
                Case "5"  ' 令和
                    era_letter = "R"
                    era_year_val = CInt(Mid(file.Name, 19, 2))
                    billing_year = CStr(2018 + era_year_val)
                Case "4"  ' 平成
                    era_letter = "H"
                    era_year_val = CInt(Mid(file.Name, 19, 2))
                    billing_year = CStr(1988 + era_year_val)
                Case "3"  ' 昭和
                    era_letter = "S"
                    era_year_val = CInt(Mid(file.Name, 19, 2))
                    billing_year = CStr(1925 + era_year_val)
                Case "2"  ' 大正
                    era_letter = "T"
                    era_year_val = CInt(Mid(file.Name, 19, 2))
                    billing_year = CStr(1911 + era_year_val)
                Case "1"  ' 明治
                    era_letter = "M"
                    era_year_val = CInt(Mid(file.Name, 19, 2))
                    billing_year = CStr(1867 + era_year_val)
            End Select
        End If
        
        If billing_year <> "" And billing_month <> "" Then
            year_code = Format(era_year_val, "00")
            
            ' 報告書ファイル名を生成
            Dim report_file_name As String, report_file_path As String
            report_file_name = GenerateReportFileNameFromBillingDate(CInt(billing_year), CInt(billing_month))
            If report_file_name = "" Then
                MsgBox "ファイル名の生成に失敗しました。", vbExclamation, "エラー"
                GoTo NextFile  ' エラー時は次のファイルへ
            End If
            report_file_path = save_path & "\" & report_file_name
            
            ' ファイルが存在しない場合のみ新規作成
            If Not file_system.FileExists(report_file_path) Then
                Dim report_wb As Workbook
                On Error GoTo ErrorHandler
                Set report_wb = Workbooks.Add(template_path)
                
                If Not report_wb Is Nothing Then
                    Application.DisplayAlerts = False
                    report_wb.SaveAs Filename:=report_file_path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
                    ' テンプレート情報を設定（シート名の変更も含む）
                    SetTemplateInfo report_wb, billing_year, billing_month
                    Application.DisplayAlerts = True
                End If
            End If
        End If
NextFile:
    Next file
    Exit Function

ErrorHandler:
    MsgBox "レポートファイル作成中にエラーが発生しました。" & vbCrLf & _
           "ファイル: " & IIf(Not file Is Nothing, file.Name, "不明") & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "発生箇所: CreateReportFiles", _
           vbCritical, "エラー"
    Resume NextFile
End Function

Function ProcessCsvFilesByType(file_system As Object, csv_files As Collection, file_type_name As String)
    On Error GoTo ErrorHandler
    
    Dim file_obj As Object
    For Each file_obj In csv_files
        Dim report_file_name As String, report_file_path As String
        Dim base_name As String, sheet_name As String
        Dim report_wb As Workbook
        Dim sheet_exists As Boolean
        Dim western_year As Integer, month As Integer
        Dim billing_year As String, billing_month As String
        Dim dispensing_year As String, dispensing_month As String

        ' CSVファイル名から請求年月を取得
        Dim year_str As String, month_str As String
        If Not GetYearMonthFromFile(file_obj.Path, file_type_name, year_str, month_str) Then
            MsgBox "ファイル " & file_obj.Name & " から診療年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If
        
        western_year = CInt(year_str)
        month = CInt(month_str)
        
        ' 請求年月を設定
        billing_year = CStr(western_year)
        billing_month = CStr(month)
        
        ' 報告書ファイル名を生成
        report_file_name = GenerateReportFileNameFromBillingDate(western_year, month)
        If report_file_name = "" Then
            MsgBox "ファイル名の生成に失敗しました。", vbExclamation, "エラー"
            GoTo NextFile
        End If
        
        ' 調剤年月を計算
        Call CalculateDispensingDate(western_year, month, dispensing_year, dispensing_month)
        
        report_file_path = save_path & "\" & report_file_name

        ' 追加すべきチェック
        If Not file_system.FileExists(report_file_path) Then
            MsgBox "ファイルが存在しません: " & report_file_path, vbExclamation
            GoTo NextFile
        End If

        ' ワークブックを開く処理を先に行う
        On Error GoTo ErrorHandler
        Set report_wb = Workbooks.Open(report_file_path, ReadOnly:=True)
        If report_wb Is Nothing Then
            MsgBox "ファイル " & report_file_path & " を開けませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If

        ' CSVデータをインポートして新規シートに転記
        base_name = file_system.GetBaseName(file_obj.Name)
        sheet_name = base_name
        Dim insert_index As Long
        insert_index = Application.WorksheetFunction.Min(3, report_wb.Sheets.Count + 1)
        Dim ws_csv As Worksheet
        Set ws_csv = report_wb.Sheets.Add(After:=report_wb.Sheets(insert_index - 1))
        ws_csv.Name = sheet_name
        ImportCsvData file_obj.Path, ws_csv, file_type_name

        ' 詳細データを詳細シートに反映
        Call TransferBillingDetails(report_wb, file_obj.Name, dispensing_year, dispensing_month)

        ' 保存してブックを閉じる
        report_wb.Save
        report_wb.Close False

        ' オブジェクトの解放処理を追加
        Set report_wb = Nothing
        Exit Function

ErrorHandler:
        If Not report_wb Is Nothing Then
            report_wb.Close SaveChanges:=False
            Set report_wb = Nothing
        End If
NextFile:
        ' 次のCSVファイルへ
    Next file_obj
End Function

' 報告書ファイル名を生成する関数
Function GenerateReportFileNameFromBillingDate(ByVal billing_year As Integer, ByVal billing_month As Integer) As String
    ' 入力値の検証を追加
    If billing_year < 1868 Or billing_year > 2100 Then
        MsgBox "無効な年が指定されました。" & vbCrLf & _
               "年: " & billing_year & vbCrLf & _
               "発生箇所: GenerateReportFileNameFromBillingDate", _
               vbExclamation, "エラー"
        GenerateReportFileNameFromBillingDate = ""
        Exit Function
    End If
    
    If billing_month < 1 Or billing_month > 12 Then
        MsgBox "無効な月が指定されました。" & vbCrLf & _
               "月: " & billing_month & vbCrLf & _
               "発生箇所: GenerateReportFileNameFromBillingDate", _
               vbExclamation, "エラー"
        GenerateReportFileNameFromBillingDate = ""
        Exit Function
    End If
    
    Dim dispensing_year As String, dispensing_month As String
    Dim era_info As Object
    
    ' 請求年月から調剤年月を計算
    If Not CalculateDispensingDate(billing_year, billing_month, dispensing_year, dispensing_month) Then
        MsgBox "調剤年月の計算に失敗しました。" & vbCrLf & _
               "請求年月: " & billing_year & "年" & billing_month & "月", _
               vbExclamation, "GenerateReportFileNameFromBillingDate - エラー"
        GenerateReportFileNameFromBillingDate = ""
        Exit Function
    End If
    
    ' 調剤年（西暦）から元号情報を取得
    Set era_info = ConvertEraYear(CInt(dispensing_year), True)
    If era_info("era") = "" Then
        MsgBox "元号への変換に失敗しました。" & vbCrLf & _
               "調剤年月: " & dispensing_year & "年" & dispensing_month & "月", _
               vbExclamation, "GenerateReportFileNameFromBillingDate - エラー"
        GenerateReportFileNameFromBillingDate = ""
        Exit Function
    End If
    
    GenerateReportFileNameFromBillingDate = "保険請求管理報告書_" & _
                            era_info("era") & _
                            Format(era_info("year"), "00") & "年" & _
                            Format(CInt(dispensing_month), "00") & "月調剤分.xlsm"
End Function

Function ConvertEraYearToWestern(ByVal era_code As String, ByVal era_year As String) As Integer
    Dim year_num As Integer
    year_num = CInt(era_year)
    
    Select Case era_code
        Case "5": ConvertEraYearToWestern = 2018 + year_num  ' 令和
        Case "4": ConvertEraYearToWestern = 1988 + year_num  ' 平成
        Case "3": ConvertEraYearToWestern = 1925 + year_num  ' 昭和
        Case "2": ConvertEraYearToWestern = 1911 + year_num  ' 大正
        Case "1": ConvertEraYearToWestern = 1867 + year_num  ' 明治
        Case Else: ConvertEraYearToWestern = 0
    End Select
End Function

Function CalculateDispensingDate(ByVal western_year As Integer, ByVal western_month As Integer, _
    ByRef dispensing_year As String, ByRef dispensing_month As String) As Boolean
    
    Dim temp_month As Integer
    temp_month = western_month - 1
    
    If temp_month < 1 Then
        temp_month = 12
        dispensing_year = CStr(western_year - 1)
    Else
        dispensing_year = CStr(western_year)
    End If
    
    dispensing_month = Format(temp_month, "00")  ' 常に2桁の形式で保存
    CalculateDispensingDate = True
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

Function SetTemplateInfo(report_book As Workbook, billing_year As String, billing_month As String) As Boolean
    Dim ws_main As Worksheet, ws_sub As Worksheet
    Dim billing_year_num As Integer, billing_month_num As Integer
    Dim dispensing_year As Integer, dispensing_month As Integer
    Dim send_date As String

    On Error GoTo ErrorHandler

    ' 西暦年と調剤月の計算
    billing_year_num = CInt(billing_year)
    billing_month_num = CInt(billing_month)

    ' 請求月の計算（請求月 = 調剤月の翌月）
    dispensing_month = billing_month_num - 1
    If dispensing_month < 0 Then
        dispensing_year = billing_year_num - 1  
        dispensing_month = 12
    Else
        dispensing_year = billing_year_num
    End If

    send_date = billing_month_num & "月10日請求分"

    ' テンプレートシート（シート1: A, シート2: B）を取得
    Set ws_main = report_book.Sheets(1)
    Set ws_sub = report_book.Sheets(2)

    ' シート名変更（シート1を "R{令和YY}.{M}"形式, シート2を丸数字月に）
    ws_main.Name = "R" & (dispensing_year - 2018) & "." & dispensing_month
    ws_sub.Name = ConvertToCircledNumber(dispensing_month)

    ' 情報転記（ヘッダ部）
    ws_main.Range("G2").Value = dispensing_year & "年" & dispensing_month & "月調剤分"
    ws_main.Range("I2").Value = send_date
    ws_main.Range("J2").Value = ThisWorkbook.Sheets("設定").Range("B1").Value
    ws_sub.Range("H1").Value = dispensing_year & "年" & dispensing_month & "月調剤分"
    ws_sub.Range("J1").Value = send_date
    ws_sub.Range("L1").Value = ThisWorkbook.Sheets("設定").Range("B1").Value

    SetTemplateInfo = True
    Exit Function

ErrorHandler:
    SetTemplateInfo = False
End Function

Function ConvertToCircledNumber(month As Integer) As String
    Dim circled_numbers As Variant
    circled_numbers = Array("①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫")
    If month >= 1 And month <= 12 Then
        ConvertToCircledNumber = circled_numbers(month - 1)
    Else
        ConvertToCircledNumber = CStr(month)
    End If
End Function

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
    Set text_stream = Nothing
    Set file_system_local = Nothing
    Set column_map = Nothing
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

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

Sub TransferBillingDetails(report_wb As Workbook, csv_file_name As String, dispensing_year As String, dispensing_month As String)
    Dim ws_main As Worksheet, ws_details As Worksheet
    Dim csv_yymm As String
    Dim payer_type As String
    Dim start_row_dict As Object
    Dim rebill_dict As Object, late_dict As Object, unpaid_dict As Object, assessment_dict As Object
    Dim era_year As Integer
    
    ' 西暦から和暦年を計算
    Call GetEraInfo(CInt(dispensing_year), era_code, era_year)
    
    ' ワークシートの設定
    Set ws_main = report_wb.Sheets("R" & era_year & "." & dispensing_month)   ' メインシート
    Set ws_details = report_wb.Sheets(ConvertToCircledNumber(CInt(dispensing_month))) ' 詳細データシート

    ' 調剤年月と請求先区分の取得
    csv_yymm = GetDispenseYearMonth(ws_main)
    payer_type = GetPayerType(csv_file_name)
    If payer_type = "労災" Then Exit Sub ' 労災等は詳細シート対象外

    ' 詳細シート上の各カテゴリ開始行を取得
    Set start_row_dict = GetCategoryStartRows(ws_details, payer_type)

    ' データの分類と辞書の作成
    Set rebill_dict = CreateObject("Scripting.Dictionary")
    Set late_dict = CreateObject("Scripting.Dictionary")
    Set unpaid_dict = CreateObject("Scripting.Dictionary")
    Set assessment_dict = CreateObject("Scripting.Dictionary")
    
    ' メインシートのデータを分類
    Call ClassifyMainSheetData(ws_main, csv_yymm, csv_file_name, rebill_dict, late_dict, unpaid_dict, assessment_dict)

    ' 行の追加処理
    Call InsertAdditionalRows(ws_details, start_row_dict, rebill_dict.Count, late_dict.Count, assessment_dict.Count)

    ' データの転記
    Call WriteDataToDetails(ws_details, start_row_dict, rebill_dict, late_dict, unpaid_dict, assessment_dict, payer_type)

    ' 完了メッセージ
    MsgBox payer_type & " のデータ転記が完了しました！", vbInformation, "処理完了"
End Sub

Private Function GetDispenseYearMonth(ws As Worksheet) As String
    GetDispenseYearMonth = ""
    If ws.Cells(2, 2).Value <> "" Then
        GetDispenseYearMonth = Right(CStr(ws.Cells(2, 2).Value), 4)
        If InStr(GetDispenseYearMonth, "年") > 0 Or InStr(GetDispenseYearMonth, "月") > 0 Then
            GetDispenseYearMonth = Replace(Replace(GetDispenseYearMonth, "年", ""), "月", "")
        End If
    End If
End Function

Private Function GetPayerType(csv_file_name As String) As String
    Dim base_name As String, payer_code As String
    
    base_name = csv_file_name
    If InStr(base_name, ".") > 0 Then base_name = Left(base_name, InStrRev(base_name, ".") - 1)
    
    If Len(base_name) >= 7 Then
        payer_code = Mid(base_name, 7, 1)
    Else
        payer_code = ""
    End If
    
    Select Case payer_code
        Case "1": GetPayerType = "社保"
        Case "2": GetPayerType = "国保"
        Case Else: GetPayerType = "労災"
    End Select
End Function

Private Function GetCategoryStartRows(ws As Worksheet, payer_type As String) As Object
    Dim start_row_dict As Object
    Set start_row_dict = CreateObject("Scripting.Dictionary")
    
    If payer_type = "社保" Then
        start_row_dict.Add "返戻再請求", GetStartRow(ws, "社保返戻再請求")
        start_row_dict.Add "月遅れ請求", GetStartRow(ws, "社保月遅れ請求")
        start_row_dict.Add "返戻・査定", GetStartRow(ws, "社保返戻・査定")
        start_row_dict.Add "未請求扱い", GetStartRow(ws, "社保未請求扱い")
    ElseIf payer_type = "国保" Then
        start_row_dict.Add "返戻再請求", GetStartRow(ws, "国保返戻再請求")
        start_row_dict.Add "月遅れ請求", GetStartRow(ws, "国保月遅れ請求")
        start_row_dict.Add "返戻・査定", GetStartRow(ws, "国保返戻・査定")
        start_row_dict.Add "未請求扱い", GetStartRow(ws, "国保未請求扱い")
    End If
    
    Set GetCategoryStartRows = start_row_dict
End Function

Private Sub ClassifyMainSheetData(ws As Worksheet, csv_yymm As String, csv_file_name As String, _
    ByRef rebill_dict As Object, ByRef late_dict As Object, ByRef unpaid_dict As Object, ByRef assessment_dict As Object)
    
    Dim last_row As Long, row As Long
    Dim dispensing_code As String, dispensing_ym As String
    Dim row_data As Variant
    
    last_row = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    For row = 2 To last_row
        dispensing_code = ws.Cells(row, 2).Value
        dispensing_ym = ConvertToWesternDate(dispensing_code)
        
        If csv_yymm <> "" And Right(dispensing_code, 4) <> csv_yymm Then
            row_data = Array(ws.Cells(row, 4).Value, dispensing_ym, ws.Cells(row, 5).Value, ws.Cells(row, 10).Value)
            
            If InStr(LCase(csv_file_name), "fixf") > 0 Then
                late_dict(ws.Cells(row, 1).Value) = row_data
            ElseIf InStr(LCase(csv_file_name), "fmei") > 0 Then
                rebill_dict(ws.Cells(row, 1).Value) = row_data
            ElseIf InStr(LCase(csv_file_name), "zogn") > 0 Then
                unpaid_dict(ws.Cells(row, 1).Value) = row_data
            ElseIf InStr(LCase(csv_file_name), "henr") > 0 Then
                assessment_dict(ws.Cells(row, 1).Value) = row_data
            End If
        End If
    Next row
End Sub

Private Sub InsertAdditionalRows(ws As Worksheet, start_row_dict As Object, _
    rebill_count As Long, late_count As Long, assessment_count As Long)
    
    Dim a As Long, b As Long, c As Long
    
    If rebill_count > BASE_DETAIL_ROWS Then a = rebill_count - BASE_DETAIL_ROWS
    If late_count > BASE_DETAIL_ROWS Then b = late_count - BASE_DETAIL_ROWS
    If assessment_count > BASE_DETAIL_ROWS Then c = assessment_count - BASE_DETAIL_ROWS
    
    If a > 0 Then ws.Rows(start_row_dict("返戻再請求") + 1 & ":" & start_row_dict("返戻再請求") + a).Insert Shift:=xlDown
    If b > 0 Then ws.Rows(start_row_dict("月遅れ請求") + 1 & ":" & start_row_dict("月遅れ請求") + b).Insert Shift:=xlDown
    If c > 0 Then ws.Rows(start_row_dict("返戻・査定") + 1 & ":" & start_row_dict("返戻・査定") + c).Insert Shift:=xlDown
End Sub

Private Sub WriteDataToDetails(ws As Worksheet, start_row_dict As Object, _
    rebill_dict As Object, late_dict As Object, unpaid_dict As Object, assessment_dict As Object, payer_type As String)
    
    If rebill_dict.Count > 0 Then Call TransferData(rebill_dict, ws, start_row_dict("返戻再請求"), payer_type)
    If late_dict.Count > 0 Then Call TransferData(late_dict, ws, start_row_dict("月遅れ請求"), payer_type)
    If unpaid_dict.Count > 0 Then Call TransferData(unpaid_dict, ws, start_row_dict("未請求扱い"), payer_type)
    If assessment_dict.Count > 0 Then Call TransferData(assessment_dict, ws, start_row_dict("返戻・査定"), payer_type)
End Sub

Function TransferData(dataDict As Object, ws As Worksheet, start_row As Long, payer_type As String) As Boolean
    If dataDict.Count = 0 Then
        TransferData = False
        Exit Function
    End If

    Dim key As Variant, row_data As Variant
    Dim r As Long: r = start_row
    Dim payer_col As Long

    ' 社保はH列(8), 国保はI列(9)に種別を記載
    If payer_type = "社保" Then
        payer_col = 8
    ElseIf payer_type = "国保" Then
        payer_col = 9
    Else
        TransferData = False  ' その他（労災等）は対象外
        Exit Function
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
    
    TransferData = True
End Function

Function GetStartRow(ws As Worksheet, category_name As String) As Long
    Dim found_cell As Range
    Set found_cell = ws.Cells.Find(what:=category_name, LookAt:=xlWhole)
    If Not found_cell Is Nothing Then
        GetStartRow = found_cell.Row
    Else
        GetStartRow = 0
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

' ファイルコレクションをソートする関数
Function SortFileCollection(files As Collection, file_system As Object, file_type As String) As Collection
    Dim sorted_files As New Collection
    Dim file_array() As Object
    Dim i As Long, count As Long
    
    ' Collectionの要素数を取得
    count = files.Count
    If count = 0 Then
        Set SortFileCollection = sorted_files
        Exit Function
    End If
    
    ' 配列を初期化
    ReDim file_array(1 To count)
    
    ' CollectionをArrayにコピー
    For i = 1 To count
        Set file_array(i) = files(i)
    Next i
    
    ' バブルソートで年月順にソート
    Dim j As Long
    For i = 1 To count - 1
        For j = 1 To count - i
            Dim year1 As String, month1 As String
            Dim year2 As String, month2 As String
            
            If GetYearMonthFromFile(file_array(j).Path, file_type, year1, month1) And _
               GetYearMonthFromFile(file_array(j + 1).Path, file_type, year2, month2) Then
                
                ' 年月を結合して比較（例：202402）
                If (year1 & Format(CInt(month1), "00")) > (year2 & Format(CInt(month2), "00")) Then
                    ' 順序が逆の場合、要素を交換
                    Dim temp_obj As Object
                    Set temp_obj = file_array(j)
                    Set file_array(j) = file_array(j + 1)
                    Set file_array(j + 1) = temp_obj
                End If
            End If
        Next j
    Next i
    
    ' ソートされた配列を新しいCollectionに追加
    For i = 1 To count
        sorted_files.Add file_array(i)
    Next i
    
    Set SortFileCollection = sorted_files
End Function

Function GetYearMonthFromFile(file_path As String, file_type As String, ByRef year_str As String, ByRef month_str As String) As Boolean
    Dim file_name As String, base_name As String
    year_str = "": month_str = ""
    
    file_name = Right(file_path, Len(file_path) - InStrRev(file_path, "\"))
    base_name = file_name
    If InStr(file_name, ".") > 0 Then base_name = Left(file_name, InStrRev(file_name, ".") - 1)
    
    Select Case LCase(file_type)
        Case "fixf"
            ' fixfファイルの場合、18文字目以降からYYYYMMDDhhmmss形式で取得
            If Len(file_name) >= 25 Then
                year_str = Mid(file_name, 18, 4)
                month_str = Mid(file_name, 22, 2)
                GetYearMonthFromFile = True
            End If
            
        Case "fmei", "henr", "zogn"
            ' fmei/henr/zognファイルの場合、末尾5文字からGYYMM形式で取得
            Dim code_part As String
            code_part = Right(base_name, 5)
            If Len(code_part) = 5 And IsNumeric(code_part) Then
                Dim era_code As String, yy As String, mm As String
                era_code = Left(code_part, 1)
                yy = Mid(code_part, 2, 2)
                mm = Right(code_part, 2)
                
                ' 元号コードから西暦年を計算
                Select Case era_code
                    Case "5": year_str = CStr(2018 + CInt(yy))  ' 令和
                    Case "4": year_str = CStr(1988 + CInt(yy))  ' 平成
                    Case "3": year_str = CStr(1925 + CInt(yy))  ' 昭和
                    Case "2": year_str = CStr(1911 + CInt(yy))  ' 大正
                    Case "1": year_str = CStr(1867 + CInt(yy))  ' 明治
                End Select
                month_str = mm
                GetYearMonthFromFile = True
            End If
    End Select
End Function

' 長時間処理の進捗表示
Private Sub UpdateProgress(current As Long, total As Long, message As String)
    Application.StatusBar = message & " - " & current & "/" & total
End Sub

' 元号コードから元号名を取得する関数を追加
Private Function GetEraName(era_code As String) As String
    Select Case era_code
        Case "5": GetEraName = "令和"
        Case "4": GetEraName = "平成"
        Case "3": GetEraName = "昭和"
        Case "2": GetEraName = "大正"
        Case "1": GetEraName = "明治"
        Case Else: GetEraName = "不明"
    End Select
End Function

' 西暦から元号コードと年を取得する関数を追加
Private Function GetEraInfo(western_year As Integer, ByRef era_code As String, ByRef era_year As Integer) As Boolean
    If western_year >= 2019 Then
        era_code = "5": era_year = western_year - 2018   ' 令和
    ElseIf western_year >= 1989 Then
        era_code = "4": era_year = western_year - 1988   ' 平成
    ElseIf western_year >= 1926 Then
        era_code = "3": era_year = western_year - 1925   ' 昭和
    ElseIf western_year >= 1912 Then
        era_code = "2": era_year = western_year - 1911   ' 大正
    ElseIf western_year >= 1868 Then
        era_code = "1": era_year = western_year - 1867   ' 明治
    Else
        era_code = "0": era_year = 0
        GetEraInfo = False
        Exit Function
    End If
    GetEraInfo = True
End Function

Private Sub CreateBackup(file_path As String)
    ' ファイルのバックアップを作成
End Sub

' 西暦年から和暦年を計算する関数
Private Function CalculateEraYear(ByVal western_year As Integer) As Integer
    If western_year >= 2019 Then
        CalculateEraYear = western_year - 2018   ' 令和
    ElseIf western_year >= 1989 Then
        CalculateEraYear = western_year - 1988   ' 平成
    ElseIf western_year >= 1926 Then
        CalculateEraYear = western_year - 1925   ' 昭和
    ElseIf western_year >= 1912 Then
        CalculateEraYear = western_year - 1911   ' 大正
    ElseIf western_year >= 1868 Then
        CalculateEraYear = western_year - 1867   ' 明治
    Else
        CalculateEraYear = 0
    End If
End Function

' オブジェクト解放用の関数を追加
Private Sub CleanupObjects(ParamArray objects() As Variant)
    Dim obj As Variant
    For Each obj In objects
        If Not obj Is Nothing Then
            If TypeName(obj) = "Workbook" Then
                obj.Close SaveChanges:=False
            End If
            Set obj = Nothing
        End If
    Next obj
End Sub
