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
    folder_path = SAVE_PATH
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