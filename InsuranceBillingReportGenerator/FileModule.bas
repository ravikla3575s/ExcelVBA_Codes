Option Explicit

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
    
    ' 変数の宣言
    Dim file As Object
    Dim billing_year As String, billing_month As String
    Dim era_letter As String, era_year_val As Integer
    Dim report_file_name As String, report_file_path As String
    
    Debug.Print "Starting CreateReportFiles"
    Debug.Print "Template path: " & template_path
    Debug.Print "Save path: " & save_path
    
    ' テンプレートファイルの存在確認を追加
    If Not file_system.FileExists(template_path) Then
        MsgBox "テンプレートファイルが見つかりません。" & vbCrLf & _
               "パス: " & template_path & vbCrLf & _
               "発生箇所: CreateReportFiles", _
               vbCritical, "エラー"
        Exit Function
    End If
    
    For Each file In files
        On Error GoTo ErrorHandler
        
        Debug.Print "----------------------------------------"
        Debug.Print "Processing file: " & file.Name
        
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
            ' 元号コードを設定
            era_code = Mid(file.Name, 18, 1)
            era_year_val = CInt(Mid(file.Name, 19, 2))
            billing_month = Mid(file.Name, 21, 2)
            
            Select Case era_code
                Case "5"  ' 令和
                    era_letter = "R"
                    billing_year = CStr(2018 + era_year_val)
                Case "4"  ' 平成
                    era_letter = "H"
                    billing_year = CStr(1988 + era_year_val)
                Case "3"  ' 昭和
                    era_letter = "S"
                    billing_year = CStr(1925 + era_year_val)
                Case "2"  ' 大正
                    era_letter = "T"
                    billing_year = CStr(1911 + era_year_val)
                Case "1"  ' 明治
                    era_letter = "M"
                    billing_year = CStr(1867 + era_year_val)
            End Select
        End If
        
        Debug.Print "File processing:"
        Debug.Print "File name: " & file.Name
        Debug.Print "Billing Year/Month: " & billing_year & "/" & billing_month
        
        If billing_year <> "" And billing_month <> "" Then
            ' 報告書ファイル名を生成（請求年月を使用）
            report_file_name = GenerateReportFileName(CInt(billing_year), CInt(billing_month))
            Debug.Print "Generated report file name: " & report_file_name
            
            If report_file_name = "" Then
                MsgBox "ファイル名の生成に失敗しました。", vbExclamation, "エラー"
                GoTo NextFile
            End If
            
            report_file_path = save_path & "\" & report_file_name
            
            ' ファイルが存在しない場合のみ新規作成
            If Not file_system.FileExists(report_file_path) Then
                Dim report_wb As Workbook
                Set report_wb = Workbooks.Add(template_path)
                
                If Not report_wb Is Nothing Then
                    ' テンプレート情報を設定（請求年月を渡す）
                    If SetTemplateInfo(report_wb, CInt(billing_year), CInt(billing_month)) Then
                        Application.DisplayAlerts = False
                        report_wb.SaveAs Filename:=report_file_path, _
                                       FileFormat:=xlOpenXMLWorkbookMacroEnabled, _
                                       Local:=True
                        Application.DisplayAlerts = True
                    End If
                    report_wb.Close SaveChanges:=True
                End If
            End If
        End If
NextFile:
    Next file
    Exit Function

ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in CreateReportFiles"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "Current file: " & IIf(Not file Is Nothing, file.Name, "Unknown")
    Debug.Print "Billing Year/Month: " & billing_year & "/" & billing_month
    Debug.Print "Report file name: " & report_file_name
    Debug.Print "=================================="
    
    Dim error_response As VbMsgBoxResult
    error_response = MsgBox("ファイル作成中にエラーが発生しました。変更を保存しますか？" & vbCrLf & _
                           "エラー番号: " & Err.Number & vbCrLf & _
                           "エラー内容: " & Err.Description & vbCrLf & _
                           "ファイル: " & IIf(Not file Is Nothing, file.Name, "不明"), _
                           vbYesNo + vbExclamation)
    
    If error_response = vbYes Then
        If Not report_wb Is Nothing Then
            Debug.Print "Cleaning up: Closing workbook"
            report_wb.Close SaveChanges:=True
            Set report_wb = Nothing
        End If
    End If
    Resume NextFile
End Function

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
            Dim year1 As Integer, month1 As Integer
            Dim year2 As Integer, month2 As Integer
            
            If GetYearMonthFromFile(file_array(j).Path, file_type, year1, month1) And _
               GetYearMonthFromFile(file_array(j + 1).Path, file_type, year2, month2) Then
                
                ' 年月を結合して比較（例：202402）
                If (CStr(year1) & Format(month1, "00")) > (CStr(year2) & Format(month2, "00")) Then
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

Function GetYearMonthFromFile(file_path As String, file_type As String, ByRef dispensing_year As Integer, ByRef dispensing_month As Integer) As Boolean
    Dim file_name As String, base_name As String
    dispensing_year = 0: dispensing_month = 0
    
    file_name = Right(file_path, Len(file_path) - InStrRev(file_path, "\"))
    base_name = file_name
    If InStr(file_name, ".") > 0 Then
        base_name = Left(file_name, InStrRev(file_name, ".") - 1)
    End If
    
    Debug.Print "Processing file: " & file_name
    
    Select Case file_type
        Case "請求確定状況"  ' fixfファイル
            If Len(file_name) >= 25 Then
                Dim billing_year As Integer, billing_month As Integer
                billing_year = CInt(Mid(file_name, 18, 4))
                billing_month = CInt(Mid(file_name, 22, 2))
                
                Debug.Print "Billing year/month from file: " & billing_year & "/" & billing_month
                
                ' 調剤月を請求月の1ヶ月前に設定
                If billing_month = 1 Then
                    dispensing_year = billing_year - 1
                    dispensing_month = 12
                Else
                    dispensing_year = billing_year
                    dispensing_month = billing_month - 1
                End If
                
                Debug.Print "Set dispensing year/month to: " & dispensing_year & "/" & dispensing_month
                GetYearMonthFromFile = True
            End If
            
        Case "振込額明細書", "返戻内訳書", "増減点連絡書"  ' fmei, henr, zognファイル
            If Len(base_name) >= 5 Then
                Dim code_part As String
                code_part = Right(base_name, 5)
                If Len(code_part) = 5 And IsNumeric(code_part) Then
                    Dim era_code As String, era_year As Integer
                    era_code = Left(code_part, 1)
                    era_year = CInt(Mid(code_part, 2, 2))
                    billing_month = CInt(Right(code_part, 2))
                    
                    ' 元号コードから西暦年を計算
                    Select Case era_code
                        Case "5": billing_year = 2018 + era_year  ' 令和
                        Case "4": billing_year = 1988 + era_year  ' 平成
                        Case "3": billing_year = 1925 + era_year  ' 昭和
                        Case "2": billing_year = 1911 + era_year  ' 大正
                        Case "1": billing_year = 1867 + era_year  ' 明治
                    End Select
                    
                    Debug.Print "Billing year/month from file: " & billing_year & "/" & billing_month
                    
                    ' 調剤月を請求月の1ヶ月前に設定
                    If billing_month = 1 Then
                        dispensing_year = billing_year - 1
                        dispensing_month = 12
                    Else
                        dispensing_year = billing_year
                        dispensing_month = billing_month - 1
                    End If
                    
                    Debug.Print "Set dispensing year/month to: " & dispensing_year & "/" & dispensing_month
                    GetYearMonthFromFile = True
                End If
            End If
    End Select

    Debug.Print "Final dispensing year/month: " & dispensing_year & "/" & dispensing_month
End Function

Private Sub CreateBackup(file_path As String)
    ' ファイルのバックアップを作成
    ' TODO: バックアップ機能の実装
End Sub

Function GenerateReportFileName(ByVal billing_year As Integer, ByVal billing_month As Integer) As String
    On Error GoTo ErrorHandler
    
    GenerateReportFileName = ""
    
    ' 入力値の検証
    If billing_year < 1900 Or billing_year > 9999 Then
        MsgBox "請求年が無効です。" & vbCrLf & _
               "年: " & billing_year & vbCrLf & _
               "発生箇所: GenerateReportFileName", _
               vbExclamation, "エラー"
        GenerateReportFileName = ""
        Exit Function
    End If
    
    If billing_month < 1 Or billing_month > 12 Then
        MsgBox "請求月が無効です。" & vbCrLf & _
               "月: " & billing_month & vbCrLf & _
               "発生箇所: GenerateReportFileName", _
               vbExclamation, "エラー"
        GenerateReportFileName = ""
        Exit Function
    End If
    
    Debug.Print "GenerateReportFileName input:"
    Debug.Print "Billing year: " & billing_year
    Debug.Print "Billing month: " & billing_month
    
    ' 元号情報を取得
    Dim era_info As Object
    Set era_info = ConvertEraYear(billing_year, True)
    
    If era_info Is Nothing Then
        MsgBox "元号の変換に失敗しました。" & vbCrLf & _
               "年: " & billing_year, _
               vbExclamation, "GenerateReportFileName - エラー"
        GenerateReportFileName = ""
        Exit Function
    End If
    
    ' ファイル名を生成
    GenerateReportFileName = "保険請求管理報告書_" & _
                            era_info("era") & _
                            Format(era_info("year"), "00") & "年" & _
                            Format(billing_month, "00") & "月.xlsm"
                            
    Debug.Print "Generated filename: " & GenerateReportFileName
    Exit Function
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error occurred in GenerateReportFileName"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    Dim error_response As VbMsgBoxResult
    error_response = MsgBox("ファイル名の生成中にエラーが発生しました。変更を保存しますか？" & vbCrLf & _
                           "エラー番号: " & Err.Number & vbCrLf & _
                           "エラー内容: " & Err.Description, _
                           vbYesNo + vbExclamation)
    
    If error_response = vbYes Then
        GenerateReportFileName = ""
    End If
End Function

Function ProcessCsvFilesByType(file_system As Object, csv_files As Collection, file_type_name As String)
    On Error GoTo ErrorHandler
    
    Dim file_obj As Object
    Dim report_file_name As String, report_file_path As String
    Dim base_name As String, sheet_name As String
    Dim report_wb As Workbook
    Dim sheet_exists As Boolean
    Dim dispensing_year As Integer, dispensing_month As Integer
    Dim insert_index As Long
    
    For Each file_obj In csv_files
        Dim save_successful As Boolean
        save_successful = False  ' 保存フラグを初期化
        
        Debug.Print "----------------------------------------"
        Debug.Print "Processing file: " & file_obj.Name
        Debug.Print "File type: " & file_type_name
        Debug.Print "File path: " & file_obj.Path
        
        ' CSVファイル名から調剤年月を取得
        If Not GetYearMonthFromFile(file_obj.Path, file_type_name, dispensing_year, dispensing_month) Then
            Debug.Print "ERROR: Failed to get year/month from file"
            MsgBox "ファイル " & file_obj.Name & " から調剤年月を取得できませんでした。", vbExclamation, "エラー"
            GoTo NextFile
        End If

        Debug.Print "Dispensing year/month: " & dispensing_year & "/" & dispensing_month
        
        ' 報告書ファイル名を生成
        report_file_name = GenerateReportFileName(dispensing_year, dispensing_month)
        Debug.Print "Generated report file name: " & report_file_name
        
        If report_file_name = "" Then
            Debug.Print "ERROR: Failed to generate report file name"
            GoTo NextFile
        End If

        report_file_path = save_path & "\" & report_file_name
        Debug.Print "Full report file path: " & report_file_path
        
        ' ファイルの存在確認
        If Not file_system.FileExists(report_file_path) Then
            Debug.Print "ERROR: Report file does not exist: " & report_file_path
            GoTo NextFile
        End If
        
        ' ワークブックを開く
        On Error Resume Next
        Set report_wb = Workbooks.Open(report_file_path, ReadOnly:=False, UpdateLinks:=False)
        If Err.Number <> 0 Then
            Debug.Print "ERROR: Failed to open workbook"
            Debug.Print "Error number: " & Err.Number
            Debug.Print "Error description: " & Err.Description
            On Error GoTo ErrorHandler
            GoTo NextFile
        End If
        On Error GoTo ErrorHandler
        
        If report_wb Is Nothing Then
            Debug.Print "ERROR: Failed to open workbook (report_wb is Nothing)"
            GoTo NextFile
        End If
        
        Debug.Print "Successfully opened workbook"
        
        ' CSVデータをインポートして新規シートに転記
        base_name = file_system.GetBaseName(file_obj.Name)
        sheet_name = base_name
        Debug.Print "Base sheet name: " & sheet_name
        
        ' シート名の重複チェックと一意の名前生成
        Dim sheet_index As Integer
        sheet_index = 1
        
        On Error Resume Next
        Do
            sheet_exists = False
            Dim test_ws As Worksheet
            Set test_ws = report_wb.Sheets(sheet_name)
            If Not test_ws Is Nothing Then
                sheet_exists = True
                sheet_name = base_name & "_" & Format(sheet_index, "00")
                sheet_index = sheet_index + 1
                Debug.Print "Sheet exists, trying new name: " & sheet_name
            End If
        Loop While sheet_exists
        On Error GoTo ErrorHandler
        
        Debug.Print "Final sheet name: " & sheet_name
        
        ' 新規シートの追加
        insert_index = Application.WorksheetFunction.Min(3, report_wb.Sheets.Count)
        Debug.Print "Insert index: " & insert_index
        
        On Error Resume Next
        Dim ws_csv As Worksheet
        Set ws_csv = report_wb.Sheets.Add(After:=report_wb.Sheets(insert_index))
        If Err.Number <> 0 Then
            Debug.Print "ERROR: Failed to add new sheet"
            Debug.Print "Error number: " & Err.Number
            Debug.Print "Error description: " & Err.Description
            GoTo NextFile
        End If
        On Error GoTo ErrorHandler
        
        If ws_csv Is Nothing Then
            Debug.Print "ERROR: Failed to create new sheet (ws_csv is Nothing)"
            GoTo NextFile
        End If
        
        ws_csv.Name = sheet_name
        Debug.Print "Successfully created and named new sheet"
        
        ' エラーが発生する可能性のある処理の前に On Error Resume Next
        On Error Resume Next
        
        ' 処理が成功したかどうかを確認
        Dim process_error As Boolean
        process_error = False
        
        ' CSVデータのインポート
        If file_type_name = "請求確定状況" Then
            ImportCsvData file_obj.Path, ws_csv, file_type_name, True
        Else
            ImportCsvData file_obj.Path, ws_csv, file_type_name, False
        End If
        
        If Err.Number <> 0 Then
            Debug.Print "ERROR in ImportCsvData: " & Err.Description
            process_error = True
        End If
        
        ' エラーをリセット
        Err.Clear
        
        ' 詳細データを詳細シートに反映
        If Not process_error Then
            Call TransferBillingDetails(report_wb, file_obj.Name, CStr(dispensing_year), _
                                      Format(dispensing_month, "00"), _
                                      (file_type_name = "請求確定状況"))
            
            If Err.Number <> 0 Then
                Debug.Print "ERROR in TransferBillingDetails: " & Err.Description
                process_error = True
            End If
        End If
        
        ' エラー処理を元に戻す
        On Error GoTo ErrorHandler
        
        ' 処理が成功した場合のみ保存
        If Not process_error Then
            Debug.Print "Processing completed successfully, saving workbook"
            report_wb.Save
            save_successful = True
        Else
            Debug.Print "Processing encountered errors, changes will not be saved"
        End If

NextFile:
        If Not report_wb Is Nothing Then
            Debug.Print "Cleaning up: Closing workbook"
            ' エラーが発生したが、重要な変更がある場合は保存するかどうかをユーザーに確認
            Dim error_response As VbMsgBoxResult
            error_response = MsgBox("エラーが発生しました。変更を保存しますか？" & vbCrLf & _
                                  "エラー内容: " & Err.Description, _
                                  vbYesNo + vbExclamation)
            report_wb.Close SaveChanges:=(error_response = vbYes)
            Set report_wb = Nothing
        End If
        Set ws_csv = Nothing
        Debug.Print "----------------------------------------"
    Next file_obj
    Exit Function

ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error occurred in ProcessCsvFilesByType"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "Current file: " & IIf(Not file_obj Is Nothing, file_obj.Name, "Unknown")
    Debug.Print "Current report file: " & report_file_name
    Debug.Print "Current report path: " & report_file_path
    Debug.Print "Current sheet name: " & sheet_name
    Debug.Print "Insert index: " & insert_index
    Debug.Print "Dispensing year: " & dispensing_year
    Debug.Print "Dispensing month: " & dispensing_month
    Debug.Print "File type: " & file_type_name
    Debug.Print "=================================="
    
    Dim error_response As VbMsgBoxResult
    error_response = MsgBox("処理中にエラーが発生しました。変更を保存しますか？" & vbCrLf & _
                           "エラー番号: " & Err.Number & vbCrLf & _
                           "エラー内容: " & Err.Description & vbCrLf & _
                           "ファイル: " & IIf(Not file_obj Is Nothing, file_obj.Name, "不明"), _
                           vbYesNo + vbExclamation)
    
    If error_response = vbYes Then
        If Not report_wb Is Nothing Then
            Debug.Print "Cleaning up: Closing workbook"
            report_wb.Close SaveChanges:=True
            Set report_wb = Nothing
        End If
    End If
    Resume NextFile
End Function

Private Function SetTemplateInfo(ByVal wb As Workbook, ByVal billing_year As Integer, ByVal billing_month As Integer) As Boolean
    On Error GoTo ErrorHandler
    
    ' メインシート（1枚目）を取得
    Dim ws_main As Worksheet
    Set ws_main = wb.Sheets(1)
    
    ' 令和年を計算
    Dim era_year As Integer
    era_year = billing_year - 2018
    
    ' テンプレートの年月を設定
    With ws_main
        ' 年月の設定（例：A1セルに "令和5年4月" のような形式で設定）
        .Range("A1").Value = "令和" & era_year & "年" & billing_month & "月"
        
        ' その他必要な初期設定があれば追加
        ' ...
    End With
    
    SetTemplateInfo = True
    Exit Function

ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in SetTemplateInfo"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "Billing year: " & billing_year
    Debug.Print "Billing month: " & billing_month
    Debug.Print "=================================="
    
    Dim error_response As VbMsgBoxResult
    error_response = MsgBox("テンプレート情報の設定中にエラーが発生しました。変更を保存しますか？" & vbCrLf & _
                           "エラー番号: " & Err.Number & vbCrLf & _
                           "エラー内容: " & Err.Description, _
                           vbYesNo + vbExclamation)
    
    If error_response = vbYes Then
        SetTemplateInfo = False
    End If
End Function 