Option Explicit

Sub ImportCsvData(csv_file_path As String, ws As Worksheet, file_type As String, Optional check_status As Boolean = False)
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
    row_index = 2  ' データは2行目から開始
    
    ' CSVの1行目と2行目（ヘッダー）を読み飛ばす
    If Not text_stream.AtEndOfStream Then
        text_stream.SkipLine  ' 1行目をスキップ
        If Not text_stream.AtEndOfStream Then
            text_stream.SkipLine  ' 2行目をスキップ
        End If
    End If
    
    ' 残りのデータを転記
    Do While Not text_stream.AtEndOfStream
        line_text = text_stream.ReadLine
        data_array = Split(line_text, ",")
        
        ' 請求確定状況のチェック（check_statusがTrueの場合）
        Dim should_transfer As Boolean
        should_transfer = True
        
        If check_status Then
            ' 請求確定状況は30列目（インデックス29）にある
            If UBound(data_array) >= 29 Then
                ' 請求確定状況が1以外の場合に転記
                should_transfer = (Trim(data_array(29)) <> "1")
                
                ' デバッグ出力を追加
                Debug.Print "Row " & row_index & " status: " & Trim(data_array(29)) & _
                          ", Transfer: " & should_transfer
            End If
        End If
        
        If should_transfer Then
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

Sub TransferBillingDetails(report_wb As Workbook, csv_file_name As String, dispensing_year As String, _
                         dispensing_month As String, Optional check_status As Boolean = False)
    On Error GoTo ErrorHandler
    
    Dim ws_main As Worksheet, ws_details As Worksheet
    Dim csv_yymm As String
    Dim payer_type As String
    Dim start_row_dict As Object
    Dim rebill_dict As Object, late_dict As Object, unpaid_dict As Object, assessment_dict As Object
    
    ' 丸付数字の月を取得
    Dim details_sheet_name As String
    details_sheet_name = ConvertToCircledNumber(CInt(dispensing_month))
    
    Debug.Print "Looking for details sheet: " & details_sheet_name
    
    ' 詳細シートの存在確認
    On Error Resume Next
    Set ws_details = report_wb.Sheets(details_sheet_name)
    On Error GoTo ErrorHandler
    
    If ws_details Is Nothing Then
        MsgBox "詳細シート '" & details_sheet_name & "' が見つかりません。", vbExclamation, "エラー"
        Exit Sub
    End If
    
    ' メインシートは存在確認せずに取得
    Set ws_main = report_wb.Sheets(1)
    
    ' 調剤年月と請求先区分の取得
    csv_yymm = GetDispenseYearMonth(ws_main)
    payer_type = GetPayerType(csv_file_name)
    
    If payer_type = "労災" Then
        Debug.Print "労災データのため、処理をスキップします。"
        Exit Sub
    End If
    
    ' 詳細シート上の各カテゴリ開始行を取得
    Set start_row_dict = GetCategoryStartRows(ws_details, payer_type)
    
    If start_row_dict.Count = 0 Then
        Debug.Print "WARNING: カテゴリの開始行が見つかりません: " & payer_type
        Exit Sub
    End If
    
    ' データの分類と辞書の作成
    Set rebill_dict = CreateObject("Scripting.Dictionary")
    Set late_dict = CreateObject("Scripting.Dictionary")
    Set unpaid_dict = CreateObject("Scripting.Dictionary")
    Set assessment_dict = CreateObject("Scripting.Dictionary")
    
    ' メインシートのデータを分類
    If check_status Then
        Call ClassifyMainSheetDataWithStatus(ws_main, csv_yymm, csv_file_name, _
                                           rebill_dict, late_dict, unpaid_dict, assessment_dict)
    Else
        Call ClassifyMainSheetData(ws_main, csv_yymm, csv_file_name, _
                                 rebill_dict, late_dict, unpaid_dict, assessment_dict)
    End If
    
    ' 行の追加処理
    Call InsertAdditionalRows(ws_details, start_row_dict, rebill_dict.Count, late_dict.Count, assessment_dict.Count)
    
    ' データの転記
    Call WriteDataToDetails(ws_details, start_row_dict, rebill_dict, late_dict, unpaid_dict, assessment_dict, payer_type)
    
    Exit Sub

ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error in TransferBillingDetails"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "Details sheet name: " & details_sheet_name
    Debug.Print "File name: " & csv_file_name
    Debug.Print "Payer type: " & payer_type
    Debug.Print "=================================="
    
    MsgBox "データ転記中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "詳細シート: " & details_sheet_name, _
           vbCritical, "エラー"
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

Private Sub ClassifyMainSheetDataWithStatus(ws As Worksheet, csv_yymm As String, csv_file_name As String, _
    ByRef rebill_dict As Object, ByRef late_dict As Object, ByRef unpaid_dict As Object, ByRef assessment_dict As Object)
    
    Dim last_row As Long, row As Long
    Dim dispensing_code As String, dispensing_ym As String
    Dim row_data As Variant
    
    last_row = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    For row = 2 To last_row
        ' 請求確定状況をチェック（AD列 = 30列目）
        If ws.Cells(row, 30).Value = "2" Then
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
        End If
    Next row
End Sub 