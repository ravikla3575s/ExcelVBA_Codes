Option Explicit

' 請求先の定数定義
Private Const BILLING_SHAHO As String = "社保"
Private Const BILLING_KOKUHO As String = "国保"

' レセプト状況の定数定義
Private Const STATUS_UNCLAIMED As Long = 1    ' 未請求
Private Const STATUS_RECLAIM As Long = 2      ' 再請求
Private Const STATUS_RETURN As Long = 3       ' 返戻
Private Const STATUS_ADJUSTMENT As Long = 4    ' 加減査定

' 各状況の開始行
Private Type StartRows
    Unclaimed As Long    ' 未請求開始行
    Reclaim As Long      ' 再請求開始行
    Return As Long       ' 返戻開始行
    Adjustment As Long   ' 加減査定開始行
End Type

' 請求先ごとのワークシート名
Private Const WS_SHAHO As String = "社保未請求一覧"
Private Const WS_KOKUHO As String = "国保未請求一覧"

' メイン処理関数
Private Function ProcessBillingData(ByVal dispensing_year As Integer, ByVal dispensing_month As Integer, _
                                  ByVal status As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' 社保・国保それぞれの配列を初期化
    Dim shahoData() As Variant
    Dim kuhoData() As Variant
    ReDim shahoData(1 To 8, 1 To 1)
    ReDim kuhoData(1 To 8, 1 To 1)
    
    ' カウンター初期化
    Dim shahoCount As Long: shahoCount = 1
    Dim kuhoCount As Long: kuhoCount = 1
    
    ' 開始行の取得
    Dim shahoStartRows As StartRows
    Dim kuhoStartRows As StartRows
    Call InitializeStartRows(shahoStartRows, kuhoStartRows)
    
    ' フォーム処理
    Dim billing_form As New UnclaimedBillingForm
    Dim continue_input As Boolean
    continue_input = True
    
    Do While continue_input
        billing_form.SetDispensingDate dispensing_year, dispensing_month
        billing_form.Show
        
        If Not billing_form.DialogResult Then
            If shahoCount = 1 And kuhoCount = 1 Then
                ' データ未入力でキャンセル
                ProcessBillingData = True
                Exit Function
            Else
                ' 既存データがある場合は確認
                If MsgBox("入力済みのデータを破棄してよろしいですか？", vbYesNo + vbQuestion) = vbYes Then
                    Exit Do
                End If
            End If
        Else
            ' 請求先に応じて適切な配列に格納
            If billing_form.BillingDestination = BILLING_SHAHO Then
                ' 社保配列の拡張チェック
                If shahoCount > UBound(shahoData, 2) Then
                    ReDim Preserve shahoData(1 To 8, 1 To shahoCount)
                End If
                Call StoreDataInArray(shahoData, shahoCount, billing_form, dispensing_year, dispensing_month)
                shahoCount = shahoCount + 1
            Else
                ' 国保配列の拡張チェック
                If kuhoCount > UBound(kuhoData, 2) Then
                    ReDim Preserve kuhoData(1 To 8, 1 To kuhoCount)
                End If
                Call StoreDataInArray(kuhoData, kuhoCount, billing_form, dispensing_year, dispensing_month)
                kuhoCount = kuhoCount + 1
            End If
            
            continue_input = billing_form.ContinueInput
        End If
    Loop
    
    ' データの転記処理
    If shahoCount > 1 Then
        Call WriteDataToWorksheet(shahoData, shahoCount - 1, WS_SHAHO, GetStartRow(shahoStartRows, status))
    End If
    
    If kuhoCount > 1 Then
        Call WriteDataToWorksheet(kuhoData, kuhoCount - 1, WS_KOKUHO, GetStartRow(kuhoStartRows, status))
    End If
    
    ProcessBillingData = True
    Exit Function
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    ProcessBillingData = False
End Function

' 開始行の初期化
Private Sub InitializeStartRows(ByRef shahoRows As StartRows, ByRef kuhoRows As StartRows)
    ' 社保の開始行
    With shahoRows
        .Unclaimed = 2      ' 未請求開始行
        .Reclaim = 8        ' 再請求開始行
        .Return = 14        ' 返戻開始行
        .Adjustment = 20    ' 加減査定開始行
    End With
    
    ' 国保の開始行
    With kuhoRows
        .Unclaimed = 2
        .Reclaim = 8
        .Return = 14
        .Adjustment = 20
    End With
End Sub

' 状態に応じた開始行の取得
Private Function GetStartRow(ByRef rows As StartRows, ByVal status As Long) As Long
    Select Case status
        Case STATUS_UNCLAIMED
            GetStartRow = rows.Unclaimed
        Case STATUS_RECLAIM
            GetStartRow = rows.Reclaim
        Case STATUS_RETURN
            GetStartRow = rows.Return
        Case STATUS_ADJUSTMENT
            GetStartRow = rows.Adjustment
    End Select
End Function

' 配列へのデータ格納
Private Sub StoreDataInArray(ByRef dataArray() As Variant, ByVal currentIndex As Long, _
                           ByVal form As UnclaimedBillingForm, ByVal year As Integer, ByVal month As Integer)
    With form
        dataArray(1, currentIndex) = .PatientName
        dataArray(2, currentIndex) = "R" & year & "." & Format(month, "00")
        dataArray(3, currentIndex) = .MedicalInstitution
        dataArray(4, currentIndex) = .UnclaimedReason
        dataArray(5, currentIndex) = .BillingDestination
        dataArray(6, currentIndex) = .InsuranceRatio
        dataArray(7, currentIndex) = .BillingPoints
        dataArray(8, currentIndex) = .Remarks
    End With
End Sub

' ワークシートへのデータ転記
Private Sub WriteDataToWorksheet(ByRef dataArray() As Variant, ByVal dataCount As Long, _
                               ByVal wsName As String, ByVal startRow As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(wsName)
    
    ' 現在の行数を確認
    Dim currentRows As Long
    currentRows = ws.Range("A" & startRow).End(xlDown).Row - startRow + 1
    
    ' 5行以上のデータがある場合、行を追加
    If currentRows >= 5 Then
        ws.Rows(startRow + 5).Resize(dataCount).Insert Shift:=xlDown
    End If
    
    ' データの転記
    With ws
        .Range(.Cells(startRow, 1), .Cells(startRow + dataCount - 1, 8)).Value = _
            WorksheetFunction.Transpose(WorksheetFunction.Transpose(dataArray))
        
        ' 書式設定
        .Range(.Cells(startRow, 1), .Cells(startRow + dataCount - 1, 8)).Borders.LineStyle = xlContinuous
    End With
End Sub

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
    details_sheet_name = UtilityModule.ConvertToCircledNumber(CInt(dispensing_month))
    
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
    Set start_row_dict = UtilityModule.GetCategoryStartRows(ws_details, payer_type)
    
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
    
    ' FIXFファイルの場合、未請求レセプトの確認（詳細シートを渡す）
    If InStr(LCase(csv_file_name), "fixf") > 0 Then
        Call CheckAndRegisterUnclaimedBilling(CInt(dispensing_year), CInt(dispensing_month), ws_details)
    End If
    
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
        dispensing_ym = UtilityModule.ConvertToWesternDate(dispensing_code)
        
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
            dispensing_ym = UtilityModule.ConvertToWesternDate(dispensing_code)
            
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

Private Type UnclaimedRecord
    PatientName As String
    DispensingDate As String
    MedicalInstitution As String
    UnclaimedReason As String
    BillingDestination As String
    InsuranceRatio As Integer
    BillingPoints As Long
    Remarks As String
End Type

Private Function CheckAndRegisterUnclaimedBilling(ByVal dispensing_year As Integer, ByVal dispensing_month As Integer, _
                                            Optional ByVal ws_details As Worksheet = Nothing) As Boolean
    On Error GoTo ErrorHandler
    
    Dim response As VbMsgBoxResult
    response = MsgBox("未請求レセプトの入力を開始しますか？", vbYesNo + vbQuestion)
    
    If response = vbYes Then
        ' 未請求レセプトデータを格納する二次元配列
        Dim unclaimedData() As Variant
        ReDim unclaimedData(1 To 8, 1 To 1)
        Dim currentColumn As Long
        currentColumn = 1
        
        Dim unclaimed_form As New UnclaimedBillingForm
        Dim continue_input As Boolean
        continue_input = True
        
        Do While continue_input
            ' 調剤年月を設定
            unclaimed_form.SetDispensingDate dispensing_year, dispensing_month
            
            ' 編集モードの場合、データをロード
            If unclaimed_form.CurrentIndex < currentColumn Then
                unclaimed_form.LoadData unclaimedData, unclaimed_form.CurrentIndex
            End If
            
            unclaimed_form.Show
            
            If Not unclaimed_form.DialogResult Then
                ' キャンセルボタンが押された場合
                If currentColumn = 1 Then
                    ' データ未入力でキャンセル
                    CheckAndRegisterUnclaimedBilling = True
                    Exit Function
                Else
                    ' 既存データがある場合は確認
                    If MsgBox("入力済みのデータを破棄してよろしいですか？", vbYesNo + vbQuestion) = vbYes Then
                        Exit Do
                    End If
                End If
            Else
                ' 配列のサイズを拡張（必要な場合）
                If currentColumn > UBound(unclaimedData, 2) Then
                    ReDim Preserve unclaimedData(1 To 8, 1 To currentColumn)
                End If
                
                ' データを配列に格納
                With unclaimed_form
                    unclaimedData(1, currentColumn) = .PatientName
                    unclaimedData(2, currentColumn) = "R" & dispensing_year & "." & Format(dispensing_month, "00")
                    unclaimedData(3, currentColumn) = .MedicalInstitution
                    unclaimedData(4, currentColumn) = .UnclaimedReason
                    unclaimedData(5, currentColumn) = .BillingDestination
                    unclaimedData(6, currentColumn) = .InsuranceRatio
                    unclaimedData(7, currentColumn) = .BillingPoints
                    unclaimedData(8, currentColumn) = .Remarks
                End With
                
                If .ContinueInput Then
                    ' 次へボタンが押された場合
                    currentColumn = currentColumn + 1
                    continue_input = True
                Else
                    ' 完了ボタンが押された場合
                    continue_input = False
                End If
            End If
        Loop
        
        ' データが1件以上入力されている場合のみ、Excelに転記
        If currentColumn > 0 Then
            If ws_details Is Nothing Then
                Set ws_details = ThisWorkbook.Worksheets("未請求一覧")
            End If
            
            ' 最終行の取得
            Dim lastRow As Long
            lastRow = ws_details.Cells(ws_details.Rows.Count, "A").End(xlUp).Row
            
            ' データの転記
            With ws_details
                .Range(.Cells(lastRow + 1, 1), .Cells(lastRow + currentColumn, 8)).Value = _
                    WorksheetFunction.Transpose(WorksheetFunction.Transpose(unclaimedData))
                
                ' 書式設定
                .Range(.Cells(lastRow + 1, 1), .Cells(lastRow + currentColumn, 8)).Borders.LineStyle = xlContinuous
            End With
        End If
    End If
    
    CheckAndRegisterUnclaimedBilling = True
    Exit Function

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    CheckAndRegisterUnclaimedBilling = False
End Function

Private Sub InsertAdditionalRows(ws As Worksheet, start_row_dict As Object, rebill_count As Long, late_count As Long, assessment_count As Long)
    Dim ws_details As Worksheet
    Set ws_details = ws
    
    Dim row_index As Long
    Dim start_row As Long
    Dim end_row As Long
    Dim i As Long
    
    ' 各カテゴリの開始行を取得
    For Each key In start_row_dict.Keys
        start_row = start_row_dict(key)
        end_row = start_row + 1
        
        ' 行の追加
        ws_details.Rows(end_row).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        ws_details.Cells(end_row, 1).Value = key
        
        ' データの転記
        If rebill_count > 0 Then
            ws_details.Cells(end_row, 2).Value = "再請求"
            rebill_count = rebill_count - 1
        ElseIf late_count > 0 Then
            ws_details.Cells(end_row, 2).Value = "遅請求"
            late_count = late_count - 1
        ElseIf assessment_count > 0 Then
            ws_details.Cells(end_row, 2).Value = "算定"
            assessment_count = assessment_count - 1
        End If
    Next key
End Sub

Private Sub WriteDataToDetails(ws As Worksheet, start_row_dict As Object, rebill_dict As Object, late_dict As Object, unpaid_dict As Object, assessment_dict As Object, payer_type As String)
    Dim ws_details As Worksheet
    Set ws_details = ws
    
    Dim row_index As Long
    Dim start_row As Long
    Dim end_row As Long
    Dim i As Long
    
    ' 各カテゴリの開始行を取得
    For Each key In start_row_dict.Keys
        start_row = start_row_dict(key)
        end_row = start_row + 1
        
        ' データの転記
        If rebill_dict.Exists(key) Then
            ws_details.Cells(end_row, 2).Value = rebill_dict(key)(0)
            ws_details.Cells(end_row, 3).Value = rebill_dict(key)(1)
            ws_details.Cells(end_row, 4).Value = rebill_dict(key)(2)
            ws_details.Cells(end_row, 5).Value = rebill_dict(key)(3)
        ElseIf late_dict.Exists(key) Then
            ws_details.Cells(end_row, 2).Value = late_dict(key)(0)
            ws_details.Cells(end_row, 3).Value = late_dict(key)(1)
            ws_details.Cells(end_row, 4).Value = late_dict(key)(2)
            ws_details.Cells(end_row, 5).Value = late_dict(key)(3)
        ElseIf unpaid_dict.Exists(key) Then
            ws_details.Cells(end_row, 2).Value = unpaid_dict(key)(0)
            ws_details.Cells(end_row, 3).Value = unpaid_dict(key)(1)
            ws_details.Cells(end_row, 4).Value = unpaid_dict(key)(2)
            ws_details.Cells(end_row, 5).Value = unpaid_dict(key)(3)
        ElseIf assessment_dict.Exists(key) Then
            ws_details.Cells(end_row, 2).Value = assessment_dict(key)(0)
            ws_details.Cells(end_row, 3).Value = assessment_dict(key)(1)
            ws_details.Cells(end_row, 4).Value = assessment_dict(key)(2)
            ws_details.Cells(end_row, 5).Value = assessment_dict(key)(3)
        End If
    Next key
End Sub