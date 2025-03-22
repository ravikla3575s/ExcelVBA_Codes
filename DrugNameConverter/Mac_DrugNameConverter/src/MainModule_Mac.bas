Option Explicit

' 医薬品マスターシートの設定
Const MASTER_SHEET_NAME As String = "薬品マスター"
Const DRUG_CODE_COLUMN As Integer = 1  ' A列
Const DRUG_NAME_COLUMN As Integer = 2  ' B列
Const FIRST_DATA_ROW As Integer = 2    ' 2行目からデータ開始

' ============================================================
' 医薬品情報検索・操作関数
' ============================================================

' 医薬品コードを14桁に整形する関数
Public Function FormatDrugCode(ByVal drugCode As String) As String
    ' 空文字や数値以外の文字を除去
    Dim cleanCode As String
    Dim i As Long
    
    cleanCode = ""
    For i = 1 To Len(drugCode)
        If IsNumeric(Mid(drugCode, i, 1)) Then
            cleanCode = cleanCode & Mid(drugCode, i, 1)
        End If
    Next i
    
    ' 14桁に調整
    If Len(cleanCode) > 14 Then
        ' 14桁を超える場合は左から14桁を使用
        FormatDrugCode = Left(cleanCode, 14)
    ElseIf Len(cleanCode) < 14 Then
        ' 14桁未満の場合は右寄せでゼロ埋め
        FormatDrugCode = String(14 - Len(cleanCode), "0") & cleanCode
    Else
        ' ちょうど14桁の場合はそのまま
        FormatDrugCode = cleanCode
    End If
End Function

' 医薬品コードから医薬品名を検索する関数
Public Function FindDrugNameByCode(ByVal drugCode As String) As String
    On Error GoTo ErrorHandler
    
    ' 引数チェック
    If Len(drugCode) = 0 Then
        FindDrugNameByCode = ""
        Exit Function
    End If
    
    ' 薬品マスターシートの存在確認
    Dim masterSheet As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' シートが存在するか確認
    Dim sheetExists As Boolean
    sheetExists = False
    
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = MASTER_SHEET_NAME Then
            sheetExists = True
            Set masterSheet = ws
            Exit For
        End If
    Next ws
    
    If Not sheetExists Then
        FindDrugNameByCode = "[マスターシートなし]"
        Exit Function
    End If
    
    ' 医薬品コードを整形
    drugCode = FormatDrugCode(drugCode)
    
    ' 検索範囲の設定
    Dim lastRow As Long
    lastRow = masterSheet.Cells(masterSheet.Rows.Count, DRUG_CODE_COLUMN).End(xlUp).Row
    
    Dim searchRange As Range
    Set searchRange = masterSheet.Range(masterSheet.Cells(FIRST_DATA_ROW, DRUG_CODE_COLUMN), masterSheet.Cells(lastRow, DRUG_CODE_COLUMN))
    
    ' 医薬品コードで検索
    Dim foundCell As Range
    Set foundCell = searchRange.Find(What:=drugCode, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
    ' 検索結果の処理
    If Not foundCell Is Nothing Then
        ' 同じ行のB列から医薬品名を取得
        FindDrugNameByCode = masterSheet.Cells(foundCell.Row, DRUG_NAME_COLUMN).Value
    Else
        FindDrugNameByCode = "[コード未登録]"
    End If
    
    Exit Function
    
ErrorHandler:
    FindDrugNameByCode = "[エラー]"
End Function

' 医薬品コードを元に医薬品名を設定する関数
Public Sub FillDrugNamesByCode()
    On Error GoTo ErrorHandler
    
    ' ワークシート参照の取得
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' 最終行の取得
    Dim lastRow As Long
    lastRow = settingsSheet.Cells(settingsSheet.Rows.Count, "A").End(xlUp).Row
    
    ' 各行の医薬品コードから医薬品名を取得してC列に表示
    Dim i As Long
    For i = 7 To lastRow
        Application.StatusBar = "医薬品名取得中: " & (i - 6) & "/" & (lastRow - 6) & "..."
        DoEvents
        
        Dim drugCode As String
        drugCode = settingsSheet.Cells(i, "A").Value
        
        If Len(drugCode) > 0 Then
            ' 医薬品コードを14桁に整形
            drugCode = FormatDrugCode(drugCode)
            settingsSheet.Cells(i, "A").Value = drugCode
            
            ' 医薬品名を取得してC列に設定
            Dim drugName As String
            drugName = FindDrugNameByCode(drugCode)
            
            settingsSheet.Cells(i, "C").Value = drugName
        End If
    Next i
    
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "医薬品名の設定中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 医薬品名から最も一致する医薬品を検索する関数（包装形態も考慮）
Public Function FindBestMatchingDrug(ByVal searchTerm As String, ByRef targetDrugs() As String, Optional ByVal packageType As String = "") As String
    ' 引数チェック
    If Len(searchTerm) = 0 Then
        FindBestMatchingDrug = ""
        Exit Function
    End If
    
    ' 検索対象の配列が空の場合
    If UBound(targetDrugs) < LBound(targetDrugs) Then
        FindBestMatchingDrug = ""
        Exit Function
    End If
    
    ' キーワードを抽出
    Dim keywords As Variant
    keywords = DrugNameParser_Mac.ExtractKeywords(searchTerm)
    
    ' 最高スコアの初期化
    Dim bestScore As Double
    bestScore = 0
    
    Dim bestMatch As String
    bestMatch = ""
    
    ' 各ターゲット薬品との一致度を計算
    Dim i As Long, j As Long
    For i = LBound(targetDrugs) To UBound(targetDrugs)
        Dim target As String
        target = targetDrugs(i)
        
        If Len(target) > 0 Then
            ' 包装形態のチェック
            Dim targetParts As DrugNameParts
            targetParts = DrugNameParser_Mac.ParseDrugString(target)
            
            Dim packageMatch As Boolean
            packageMatch = True
            
            ' 包装形態が指定されている場合、一致するか確認
            If Len(packageType) > 0 And Len(targetParts.Package) > 0 Then
                ' パッケージタイプが一致するかチェック
                If InStr(1, packageType, "バラ", vbTextCompare) > 0 Then
                    ' バラ包装系なら"バラ"または"調剤用"が含まれているか確認
                    packageMatch = (InStr(1, targetParts.Package, "バラ", vbTextCompare) > 0 Or _
                                  InStr(1, targetParts.Package, "調剤用", vbTextCompare) > 0)
                ElseIf InStr(1, packageType, "PTP", vbTextCompare) > 0 Then
                    ' PTP包装系なら"PTP"が含まれているか確認
                    packageMatch = (InStr(1, targetParts.Package, "PTP", vbTextCompare) > 0)
                ElseIf InStr(1, packageType, "分包", vbTextCompare) > 0 Then
                    ' 分包系なら"分包"が含まれているか確認
                    packageMatch = (InStr(1, targetParts.Package, "分包", vbTextCompare) > 0)
                ElseIf InStr(1, packageType, "SP", vbTextCompare) > 0 Then
                    ' SP包装系なら"SP"が含まれているか確認
                    packageMatch = (InStr(1, targetParts.Package, "SP", vbTextCompare) > 0)
                End If
            End If
            
            ' 包装形態が一致する場合のみスコア計算
            If packageMatch Then
                ' スコアの計算
                Dim score As Double
                score = 0
                
                ' 完全一致の場合は最高スコア
                If searchTerm = target Then
                    score = 1
                Else
                    ' キーワードごとの一致度を計算
                    For j = LBound(keywords) To UBound(keywords)
                        Dim keyword As String
                        keyword = keywords(j)
                        
                        If InStr(1, target, keyword, vbTextCompare) > 0 Then
                            ' キーワードの重みに応じてスコアを加算
                            score = score + (1 / (UBound(keywords) - LBound(keywords) + 1))
                        End If
                    Next j
                    
                    ' さらに詳細な一致度を計算
                    score = score + DrugNameParser_Mac.CompareDrugStringsWithRate(searchTerm, target)
                End If
                
                ' 最高スコアの更新
                If score > bestScore Then
                    bestScore = score
                    bestMatch = target
                End If
            End If
        End If
    Next i
    
    ' 十分な一致度がある場合のみ結果を返す
    If bestScore >= 0.5 Then
        FindBestMatchingDrug = bestMatch
    Else
        FindBestMatchingDrug = ""
    End If
End Function

' 一致マーカーを処理中のセルに追加する関数
Public Sub AddMatchMarker(ByVal targetCell As Range, ByVal matchType As String)
    With targetCell
        Select Case matchType
            Case "完全一致"
                .Interior.Color = RGB(198, 239, 206) ' 薄い緑
            Case "部分一致"
                .Interior.Color = RGB(255, 235, 156) ' 薄い黄色
            Case "不一致"
                .Interior.Color = RGB(255, 199, 206) ' 薄い赤
        End Select
    End With
End Sub

' ファイル選択ダイアログを表示し、選択されたファイルパスを返す関数
Public Function GetFilePathFromDialog(Optional ByVal fileFilter As String = "Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", Optional ByVal dialogTitle As String = "ファイルを選択") As String
    Dim filePath As String
    filePath = Application.GetOpenFilename(fileFilter, , dialogTitle, , False)
    
    ' キャンセルされた場合は空文字を返す
    If filePath = "False" Then
        GetFilePathFromDialog = ""
    Else
        GetFilePathFromDialog = filePath
    End If
End Function

' データのバックアップを作成する関数
Public Sub BackupWorksheet(ByVal sourceSheet As Worksheet, ByVal backupName As String)
    ' バックアップシートがすでに存在する場合は削除
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(backupName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' シートをコピーして名前を変更
    sourceSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ActiveSheet.Name = backupName
    
    ' バックアップ日時を記録
    ActiveSheet.Cells(1, 1).Value = "バックアップ: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

' ============================================================
' 設定と初期化関数
' ============================================================

' 指定されたシートの内容をクリアする関数
Public Sub ClearSheet(ByVal ws As Worksheet, Optional ByVal startRow As Long = 1, Optional ByVal preserveFormatting As Boolean = True)
    If preserveFormatting Then
        ' フォーマットを保持しながらクリア
        ws.Cells.ClearContents
    Else
        ' すべてクリア
        ws.Cells.Clear
    End If
End Sub

' 使用方法の説明を追加する関数
Public Sub AddInstructions()
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' A35から下の内容をクリア（既存の指示があれば削除）
    settingsSheet.Range("A35:E50").ClearContents
    
    ' 使用方法の説明を追加
    Dim instructions As Variant
    instructions = Array("■ 使用方法", _
                         "1. 医薬品コードシートに医薬品コードを入力します（A7セルから下）", _
                         "2. 設定シートで包装形態を選択します（B4セル）", _
                         "3. メニューから「ツール」→「マクロ」→「ProcessDrugCodesAndCompare」を実行します", _
                         "", _
                         "■ 動作内容", _
                         "* 処理中はステータスバーに進捗状況が表示されます", _
                         "* 最初の医薬品名から自動的に包装形態が判定され、最適な結果が得られるように処理されます", _
                         "* 一致した医薬品名はB列に表示されます", _
                         "", _
                         "■ パッケージタイプについて", _
                         "* 「バラ包装」：バラや調剤用の薬品を優先的に検索します", _
                         "* 「分包品」：PTP、分包、SP包装の薬品を優先的に検索します", _
                         "", _
                         "■ 処理のパフォーマンス", _
                         "* 最初の医薬品から包装規格を自動判定し、処理を最適化します", _
                         "* 処理された結果は、完了時にメッセージボックスで表示されます")
    
    ' 説明文を表示
    Dim i As Long
    For i = LBound(instructions) To UBound(instructions)
        settingsSheet.Cells(35 + i, 1).Value = instructions(i)
    Next i
    
    ' 書式設定
    With settingsSheet.Range("A35")
        .Font.Bold = True
        .Font.Size = 11
    End With
End Sub 