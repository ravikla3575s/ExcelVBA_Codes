Option Explicit

' ラッパーモジュール - 基本機能

' ステータスバーを表示するクラス
Private Type StatusBarType
    Message As String
    ProgressMsg As MSForms.Label
    ProgressBar As MSForms.Label
    ProgressForm As MSForms.UserForm
    MaxValue As Long
    CurrentValue As Long
End Type

Private StatusBar As StatusBarType

' 進捗バーを初期化する関数
Private Sub InitProgressBar(ByVal title As String, ByVal maxValue As Long)
    ' ユーザーフォームを作成
    Set StatusBar.ProgressForm = New MSForms.UserForm
    With StatusBar.ProgressForm
        .Caption = title
        .Width = 300
        .Height = 100
        .StartPosition = 0 ' 手動
        .Left = Application.Left + (Application.Width - .Width) / 2
        .Top = Application.Top + (Application.Height - .Height) / 2
    End With
    
    ' メッセージラベルを追加
    Set StatusBar.ProgressMsg = StatusBar.ProgressForm.Controls.Add("Forms.Label.1")
    With StatusBar.ProgressMsg
        .Left = 10
        .Top = 10
        .Width = 280
        .Height = 20
        .Caption = "処理を開始しています..."
    End With
    
    ' プログレスバーを追加
    Set StatusBar.ProgressBar = StatusBar.ProgressForm.Controls.Add("Forms.Label.1")
    With StatusBar.ProgressBar
        .Left = 10
        .Top = 40
        .Width = 0
        .Height = 20
        .BackColor = RGB(0, 120, 215)
        .BorderStyle = 1 ' 実線
    End With
    
    ' 背景のプログレスバーフレームを追加
    Dim progressFrame As MSForms.Label
    Set progressFrame = StatusBar.ProgressForm.Controls.Add("Forms.Label.1")
    With progressFrame
        .Left = 10
        .Top = 40
        .Width = 280
        .Height = 20
        .BackColor = RGB(240, 240, 240)
        .BorderStyle = 1 ' 実線
        .ZOrder (1) ' 背景に配置
    End With
    
    ' 値を初期化
    StatusBar.MaxValue = maxValue
    StatusBar.CurrentValue = 0
    
    ' フォームを表示（モードレス）
    StatusBar.ProgressForm.Show vbModeless
    DoEvents
End Sub

' 進捗バーを更新する関数
Private Sub UpdateProgressBar(ByVal message As String, ByVal currentValue As Long)
    StatusBar.CurrentValue = currentValue
    StatusBar.ProgressMsg.Caption = message
    
    ' プログレスバーの幅を更新
    Dim percentage As Double
    percentage = StatusBar.CurrentValue / StatusBar.MaxValue
    StatusBar.ProgressBar.Width = 280 * percentage
    
    StatusBar.ProgressForm.Repaint
    DoEvents
End Sub

' 進捗バーを閉じる関数
Private Sub CloseProgressBar()
    If Not StatusBar.ProgressForm Is Nothing Then
        Unload StatusBar.ProgressForm
        Set StatusBar.ProgressForm = Nothing
        Set StatusBar.ProgressMsg = Nothing
        Set StatusBar.ProgressBar = Nothing
    End If
End Sub

' メイン処理を呼び出すラッパー関数（選択された包装形態に応じて複数の包装単位で処理）
Public Sub RunDrugNameComparison()
    On Error GoTo ErrorHandler
    
    ' ワークシート参照の取得
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' 包装形態の取得
    Dim packageSelection As String
    packageSelection = settingsSheet.Range("B4").Value
    
    ' 包装単位の配列を選択に応じて設定
    Dim packageTypes As Variant
    Dim totalProcessed As Long, totalSkipped As Long
    totalProcessed = 0
    totalSkipped = 0
    
    Application.ScreenUpdating = False
    
    If packageSelection = "バラ包装" Then
        packageTypes = Array("バラ", "調剤用")
    ElseIf packageSelection = "分包品" Then
        packageTypes = Array("PTP", "分包", "SP", "包装小", "PTP(患者用)")
    Else
        MsgBox "有効な包装形態を選択してください。", vbExclamation
        GoTo CleanExit
    End If
    
    ' 進捗バーを初期化
    InitProgressBar "医薬品名比較処理", UBound(packageTypes) + 2 ' 医薬品名取得 + 包装単位の数
    UpdateProgressBar "医薬品コードから医薬品名を取得しています...", 1
    
    ' 医薬品コードから取得した医薬品名を配列に保存
    Dim drugNames() As String
    Dim drugCodes() As String
    Call GetDrugNamesFromCodes(drugNames, drugCodes)
    
    ' 各包装単位で処理を実行
    Dim i As Long, processed As Long, skipped As Long
    Dim resultMsg As String
    resultMsg = "処理結果:" & vbCrLf & vbCrLf
    
    For i = LBound(packageTypes) To UBound(packageTypes)
        ' 現在処理中の包装単位を表示
        settingsSheet.Range("E4").Value = packageTypes(i)
        
        ' 進捗バーを更新
        UpdateProgressBar packageTypes(i) & "形態での医薬品名を比較しています...", i + 2
        
        ' 包装単位ごとの処理
        Call ProcessPackageType(packageTypes(i), drugNames, drugCodes, processed, skipped)
        
        totalProcessed = totalProcessed + processed
        totalSkipped = totalSkipped + skipped
        
        resultMsg = resultMsg & packageTypes(i) & ": " & processed & "件一致、" & skipped & "件不一致" & vbCrLf
    Next i
    
    settingsSheet.Range("E4").Value = "完了"
    
    resultMsg = resultMsg & vbCrLf & "合計: " & totalProcessed & "件一致、" & totalSkipped & "件不一致"
    
CleanExit:
    ' 進捗バーを閉じる
    CloseProgressBar
    Application.ScreenUpdating = True
    MsgBox resultMsg, vbInformation
    Exit Sub
    
ErrorHandler:
    ' 進捗バーを閉じる
    CloseProgressBar
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' 医薬品コードから医薬品名を取得して配列に保存する関数
Private Sub GetDrugNamesFromCodes(ByRef drugNames() As String, ByRef drugCodes() As String)
    On Error GoTo ErrorHandler
    
    ' ワークシート参照の取得
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' 最終行の取得
    Dim lastRow As Long
    lastRow = settingsSheet.Cells(settingsSheet.Rows.Count, "A").End(xlUp).Row
    
    ' 配列のサイズを設定
    Dim count As Long
    count = lastRow - 6 ' 7行目から始まるため
    If count <= 0 Then Exit Sub
    
    ReDim drugNames(1 To count)
    ReDim drugCodes(1 To count)
    
    ' プログレスバーのサブステップ用に進捗バーを作成（モードレス）
    Dim progressStep As Long
    progressStep = 0
    
    ' 各行のコードを処理
    Dim i As Long, idx As Long
    idx = 1
    
    For i = 7 To lastRow
        ' 進捗状況の更新（10%ごとに表示）
        If (i - 7) Mod (Application.WorksheetFunction.Max(1, (lastRow - 7) / 10)) = 0 Then
            progressStep = progressStep + 1
            StatusBar.ProgressMsg.Caption = "医薬品コード " & (i - 6) & "/" & count & " を処理中..."
            DoEvents
        End If
        
        Dim drugCode As String
        drugCode = settingsSheet.Cells(i, "A").Value
        
        If Len(drugCode) > 0 Then
            ' 医薬品コードを14桁に整形
            drugCode = MainModule.FormatDrugCode(drugCode)
            settingsSheet.Cells(i, "A").Value = drugCode
            
            ' コードを元に医薬品名を検索
            Dim drugName As String
            drugName = MainModule.FindDrugNameByCode(drugCode)
            
            ' 配列に保存
            drugNames(idx) = drugName
            drugCodes(idx) = drugCode
            idx = idx + 1
        End If
    Next i
    
    ' 配列のサイズを実際に使用した分に調整
    If idx > 1 Then
        ReDim Preserve drugNames(1 To idx - 1)
        ReDim Preserve drugCodes(1 To idx - 1)
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "医薬品名の取得中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 包装形態ごとの処理を行う関数
Private Sub ProcessPackageType(ByVal packageType As String, ByRef drugNames() As String, ByRef drugCodes() As String, ByRef processedCount As Long, ByRef skippedCount As Long)
    ' 初期設定
    processedCount = 0
    skippedCount = 0
    
    ' ワークシート参照の取得
    Dim settingsSheet As Worksheet, targetSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    Set targetSheet = ThisWorkbook.Worksheets(2)
    
    ' 最終行の取得
    Dim lastRowTarget As Long
    lastRowTarget = targetSheet.Cells(targetSheet.Rows.Count, "B").End(xlUp).Row
    
    ' 比較対象薬品名を配列に格納
    Dim targetDrugs() As String
    ReDim targetDrugs(1 To lastRowTarget - 1)
    
    Dim i As Long
    For i = 2 To lastRowTarget
        targetDrugs(i - 1) = targetSheet.Cells(i, "B").Value
    Next i
    
    ' 医薬品名の比較と転記
    Dim bestMatch As String
    Dim drugCount As Long
    drugCount = UBound(drugNames)
    Dim progressStep As Long
    progressStep = 0
    
    For i = LBound(drugNames) To UBound(drugNames)
        ' 進捗状況の更新（10%ごとに表示）
        If (i - LBound(drugNames)) Mod (Application.WorksheetFunction.Max(1, drugCount / 10)) = 0 Then
            progressStep = progressStep + 1
            StatusBar.ProgressMsg.Caption = packageType & ": 医薬品名 " & i & "/" & drugCount & " を比較中..."
            DoEvents
        End If
        
        ' 医薬品名が配列にあれば処理
        If Len(drugNames(i)) > 0 Then
            ' 包装規格も自動的に考慮
            bestMatch = MainModule.FindBestMatchingDrug(drugNames(i), targetDrugs, packageType)
            
            ' 一致する結果があれば転記、なければスキップ
            If Len(bestMatch) > 0 Then
                ' 対応する行を特定してB列に直接転記
                Dim rowIndex As Long
                rowIndex = i + 6 ' 配列のインデックスは1から始まり、データは7行目から始まるため
                
                ' B列に値がなければ、または既存の値よりも良い一致があれば上書き
                If Len(settingsSheet.Cells(rowIndex, "B").Value) = 0 Then
                    settingsSheet.Cells(rowIndex, "B").Value = bestMatch
                    processedCount = processedCount + 1
                End If
            Else
                ' 一致しない場合は何もしない
                skippedCount = skippedCount + 1
            End If
        End If
    Next i
End Sub

' 医薬品コードの検索と医薬品名比較を一括で実行する関数
Public Sub ProcessDrugCodesAndCompare()
    On Error GoTo ErrorHandler
    
    ' 進捗バーを初期化
    InitProgressBar "医薬品名コード処理", 1
    UpdateProgressBar "医薬品名の比較処理を開始しています...", 1
    
    ' ワークシート参照の取得
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' 医薬品コードから医薬品名を取得して直接比較を実行
    Call RunDrugNameComparison
    
    Exit Sub
    
ErrorHandler:
    ' 進捗バーを閉じる
    CloseProgressBar
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

' 医薬品コードの書式設定と医薬品名の検索を実行する関数（B列への転記用）
Public Sub FormatCodesAndFillDrugNames()
    On Error GoTo ErrorHandler
    
    ' ワークシート参照の取得
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' 最終行の取得
    Dim lastRow As Long
    lastRow = settingsSheet.Cells(settingsSheet.Rows.Count, "A").End(xlUp).Row
    
    ' A列の医薬品コードを14桁に整形
    Dim i As Long
    For i = 7 To lastRow ' 7行目から開始
        Dim drugCode As String
        drugCode = settingsSheet.Cells(i, "A").Value
        
        If Len(drugCode) > 0 Then
            ' 医薬品コードを14桁に整形
            settingsSheet.Cells(i, "A").Value = MainModule.FormatDrugCode(drugCode)
        End If
    Next i
    
    ' 医薬品コードから医薬品名を検索して設定
    MainModule.FillDrugNamesByCode
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

' B4セルに包装形態の選択肢をドロップダウンリストとして設定する関数
Public Sub SetupPackageTypeDropdown()
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' B4セルにドロップダウンリストを設定
    With settingsSheet.Range("B4").Validation
        .Delete ' 既存の入力規則を削除
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="バラ包装,分包品"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "包装形態の選択"
        .ErrorTitle = "無効な選択"
        .InputMessage = "「バラ包装」または「分包品」を選択してください"
        .ErrorMessage = "リストから有効な包装形態を選択してください"
    End With
    
    ' B4セルの書式設定
    With settingsSheet.Range("B4")
        .Value = "バラ包装" ' デフォルト値を設定
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242) ' 薄い青色の背景
    End With
    
    ' A4セルにラベルを設定
    With settingsSheet.Range("A4")
        .Value = "包装形態:"
        .Font.Bold = True
    End With
    
    ' 現在の包装単位表示用のセルを追加
    With settingsSheet.Range("D4")
        .Value = "現在の処理:"
        .Font.Bold = True
    End With
    
    ' 現在処理中の包装単位を表示するセル
    With settingsSheet.Range("E4")
        .Value = ""
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
    End With
    
    MsgBox "包装形態のドロップダウンリストを設定しました。", vbInformation
End Sub

' シート1にインストラクションを追加する関数
Public Sub AddInstructions()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    
    ' 既存の指示内容を削除（A35以降）
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow > 35 Then
        ws.Range("A35:A" & lastRow).ClearContents
    End If
    
    ' 新しい指示内容を追加
    Dim note As String
    note = "※使用方法:" & vbCrLf & vbCrLf & _
           "1. 医薬品コードシート（シート3）に医薬品名とコードを入力します。" & vbCrLf & _
           "   F列: 医薬品名" & vbCrLf & _
           "   G列: 医薬品コード（0から始まる14桁）" & vbCrLf & _
           "   H列: 包装コード（1から始まる14桁）" & vbCrLf & _
           "   I列: 販売コード（2から始まる14桁）" & vbCrLf & vbCrLf & _
           "2. 設定シート（シート1）のB4セルで包装形態を選択します。" & vbCrLf & _
           "   バラ包装: 「バラ」と「調剤用」の包装単位で検索" & vbCrLf & _
           "   分包品: 「PTP」「分包」「SP」「包装小」「PTP(患者用)」の包装単位で検索" & vbCrLf & vbCrLf & _
           "3. 設定シート（シート1）のA7以降に医薬品コードを入力します。" & vbCrLf & _
           "   コードは自動的に14桁に整形されます。" & vbCrLf & vbCrLf & _
           "4. メニューから「ツール」→「マクロ」を選択し、" & vbCrLf & _
           "   「ProcessDrugCodesAndCompare」を選んで「実行」ボタンをクリックします。" & vbCrLf & vbCrLf
    
    note = note & "※処理の流れ:" & vbCrLf & _
           "1. 実行するとプログレスバーが表示され、処理の進捗状況が確認できます。" & vbCrLf & _
           "2. A列の医薬品コードから対応する医薬品名を検索します。" & vbCrLf & _
           "3. 各包装形態に対して順番に処理を行い、一致する結果をB列に表示します。" & vbCrLf & _
           "4. すべての包装形態での処理が終わると結果がメッセージボックスに表示されます。" & vbCrLf & vbCrLf & _
           "※括弧内の包装規格（例：0.5g/包、5mL）も自動的に認識され、" & vbCrLf & _
           "同じ薬品名でも包装規格が一致する医薬品を優先的に選択します。" & vbCrLf & _
           "例）レボセチリジン塩酸塩DS0.5%「タカタ」分包 0.5gと" & vbCrLf & _
           "レボセチリジン塩酸塩DS0.5%「タカタ」分包 0.25gは区別されます。"
    
    ws.Range("A" & (ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 2)).Value = note
    
    MsgBox "使用方法の指示を追加しました。メニューから「ツール」→「マクロ」を選択し、" & vbCrLf & _
           "「ProcessDrugCodesAndCompare」を選んで処理を実行してください。", vbInformation
End Sub

' ワークブックの初期化関数
Public Sub InitWorkbook()
    On Error GoTo ErrorHandler
    
    ' 必要なシートを確保
    Dim settingsSheet As Worksheet, targetSheet As Worksheet, codeSheet As Worksheet
    Dim i As Long
    
    ' シート1を設定シートとして使用
    If ThisWorkbook.Worksheets.Count < 1 Then
        Set settingsSheet = ThisWorkbook.Worksheets.Add
    Else
        Set settingsSheet = ThisWorkbook.Worksheets(1)
    End If
    settingsSheet.Name = "設定シート"
    
    ' シート2を比較対象シートとして使用
    If ThisWorkbook.Worksheets.Count < 2 Then
        Set targetSheet = ThisWorkbook.Worksheets.Add(After:=settingsSheet)
    Else
        Set targetSheet = ThisWorkbook.Worksheets(2)
    End If
    targetSheet.Name = "比較対象医薬品"
    
    ' シート3を医薬品コードシートとして使用
    If ThisWorkbook.Worksheets.Count < 3 Then
        Set codeSheet = ThisWorkbook.Worksheets.Add(After:=targetSheet)
    Else
        Set codeSheet = ThisWorkbook.Worksheets(3)
    End If
    codeSheet.Name = "医薬品コード"
    
    ' シート1（設定シート）の設定
    With settingsSheet
        ' タイトル設定
        .Range("A1:E1").Merge
        .Range("A1").Value = "医薬品名比較ツール"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.Color = RGB(180, 198, 231)
        
        ' 使い方
        .Range("A2").Value = "【使い方】"
        .Range("A2").Font.Bold = True
        .Range("A3").Value = "1. B4の包装形態を選択"
        .Range("A4").Value = "包装形態:"
        .Range("A4").Font.Bold = True
        
        ' 現在の包装単位表示用のセルを追加
        .Range("D4").Value = "現在の処理:"
        .Range("D4").Font.Bold = True
        .Range("E4").Value = ""
        .Range("E4").Font.Bold = True
        .Range("E4").Interior.Color = RGB(217, 225, 242)
        
        ' ドロップダウンリスト設定
        With .Range("B4").Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, _
                 Formula1:="バラ包装,分包品"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        
        ' B4セル設定
        .Range("B4").Value = "バラ包装"
        .Range("B4").Font.Bold = True
        .Range("B4").Interior.Color = RGB(217, 225, 242)
        
        ' 手順
        .Range("A5").Value = "2. A7以降に医薬品コードを入力（自動的に14桁に整形されます）"
        .Range("A5").Font.Bold = True
        
        ' ヘッダー
        .Range("A6").Value = "医薬品コード"
        .Range("B6").Value = "一致医薬品名"
        .Range("C6").Value = "備考"
        .Range("A6:C6").Font.Bold = True
        .Range("A6:C6").Interior.Color = RGB(221, 235, 247)
        
        ' 列幅
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 40
        .Columns("C").ColumnWidth = 30
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 15
        
        ' 行番号と空の医薬品コードセル
        Dim i As Long
        For i = 7 To 30
            .Cells(i, "A").Value = "" ' 医薬品コード用に空欄を設定
        Next i
    
        ' 実行方法の案内
        .Range("A32").Value = "※実行方法: メニューから「ツール」→「マクロ」を選択し、「ProcessDrugCodesAndCompare」を実行"
        .Range("A33").Value = "A列の医薬品コードから直接医薬品名を検索し、バラ包装「バラ」「調剤用」または分包品「PTP」「分包」などで比較します。"
        .Range("A34").Value = "※括弧内の包装規格（例：0.5g/包、5mL）も自動的に認識され、同じ薬品名でも包装規格が一致する医薬品を優先選択します。"
        .Range("A32:A34").Font.Italic = True
    End With
    
    ' シート2の設定
    With targetSheet
        ' タイトル
        .Range("A1:B1").Merge
        .Range("A1").Value = "比較対象医薬品リスト"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.Color = RGB(180, 198, 231)
        
        ' ヘッダー
        .Range("A2").Value = "No."
        .Range("B2").Value = "医薬品名"
        .Range("A2:B2").Font.Bold = True
        .Range("A2:B2").Interior.Color = RGB(221, 235, 247)
        
        ' 列幅
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 50
        
        ' 行番号
        For i = 3 To 30
            .Cells(i, "A").Value = i - 2
        Next i
    End With
    
    ' シート3の設定（医薬品コードシート）
    With codeSheet
        ' タイトル
        .Range("A1:I1").Merge
        .Range("A1").Value = "医薬品コードシート"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.Color = RGB(180, 198, 231)
        
        ' ヘッダー
        .Range("A2").Value = "No."
        .Range("F2").Value = "医薬品名"
        .Range("G2").Value = "医薬品コード"
        .Range("H2").Value = "包装コード"
        .Range("I2").Value = "販売コード"
        .Range("A2:I2").Font.Bold = True
        .Range("A2:I2").Interior.Color = RGB(221, 235, 247)
        
        ' 列幅
        .Columns("A").ColumnWidth = 5
        .Columns("F").ColumnWidth = 50
        .Columns("G").ColumnWidth = 20
        .Columns("H").ColumnWidth = 20
        .Columns("I").ColumnWidth = 20
        
        ' 行番号
        For i = 3 To 30
            .Cells(i, "A").Value = i - 2
        Next i
    End With
    
    MsgBox "ワークブックを初期化しました。" & vbNewLine & _
           "1. 医薬品コードシート（シート3）にデータを入力してください" & vbNewLine & _
           "2. 設定シート（シート1）のB4セルで包装形態を選択（バラ包装または分包品）" & vbNewLine & _
           "3. 設定シート（シート1）のA7以降に医薬品コードを入力" & vbNewLine & _
           "4. メニューの「ツール」→「マクロ」から「ProcessDrugCodesAndCompare」を実行", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub
