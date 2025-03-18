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
    
    MsgBox "ファイル作成中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "ファイル: " & IIf(Not file Is Nothing, file.Name, "不明"), _
           vbCritical, "エラー"
    
    If Not report_wb Is Nothing Then
        report_wb.Close SaveChanges:=False
        Set report_wb = Nothing
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