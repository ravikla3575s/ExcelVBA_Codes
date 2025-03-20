Attribute VB_Name = "MainModule"
Option Explicit

' 定数定義
Public Const MAX_LINES_PER_SHEET As Long = 40
Public Const REQUIRED_SHEETS_COUNT As Integer = 6
Public Const BASE_DETAIL_ROWS As Integer = 4

' テンプレート・保存先パス
Public template_path As String
Public save_path As String

Sub CreateReportsFromCSV()
    On Error GoTo ErrorHandler
    
    ' パスの設定
    template_path = ThisWorkbook.Sheets(1).Range("B2").value & "\保険請求管理報告書テンプレート20250320.xltm"
    save_path = ThisWorkbook.Sheets(1).Range("B3").value
    
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
    For Each file_obj In file_system.GetFolder(csv_folder).files
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
    If fixf_files.count = 0 And fmei_files.count = 0 And henr_files.count = 0 And zogn_files.count = 0 Then
        MsgBox "選択したフォルダには処理対象のCSVファイルがありません。処理を中止します。", vbExclamation, "エラー"
        Exit Sub
    End If

    ' 7. fixfファイルとfmeiファイルの有無による処理分岐
    If fixf_files.count > 0 Then
        CreateReportFiles file_system, fixf_files, save_path, template_path
    End If
    If fmei_files.count > 0 Then
        CreateReportFiles file_system, fmei_files, save_path, template_path
    End If

    ' 8. 各種明細CSV（fmei, henr, zogn）の処理
    FileModule.ProcessCsvFilesByType file_system, fixf_files, "請求確定状況"
    FileModule.ProcessCsvFilesByType file_system, fmei_files, "振込額明細書"
    FileModule.ProcessCsvFilesByType file_system, henr_files, "返戻内訳書"
    FileModule.ProcessCsvFilesByType file_system, zogn_files, "増減点連絡書"
    
    ' 9. 完了メッセージ
    MsgBox "CSVファイルの処理が完了しました！", vbInformation, "完了"

    ' オブジェクトの解放処理を追加
    Set file_system = Nothing
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    Debug.Print "========== ERROR DETAILS =========="
    Debug.Print "Error occurred in CreateReportsFromCSV"
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error description: " & Err.Description
    Debug.Print "=================================="
    
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, _
           vbCritical, "エラー"
           
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

' 調剤年月を取得する関数（年と月を別々に取得）
Private Sub GetDispensingYearMonth(file_path As String, ByRef year_out As String, ByRef month_out As String)
    Dim file_name As String
    Dim billing_year As Integer, billing_month As Integer
    
    file_name = Right(file_path, Len(file_path) - InStrRev(file_path, "\"))
    year_out = ""
    month_out = ""
    
    ' FIXFファイルから請求年月を抽出
    If Len(file_name) >= 23 Then
        billing_year = CInt(Mid(file_name, 18, 4))
        billing_month = CInt(Mid(file_name, 22, 2))
        
        ' 請求月から調剤年月を計算（1月前）
        If billing_month = 1 Then
            year_out = CStr(billing_year - 1)
            month_out = "12"
        Else
            year_out = CStr(billing_year)
            month_out = Format(billing_month - 1, "00")
        End If
    End If
End Sub

