Option Explicit

' ラッパーモジュール - Mac互換のコード

' メイン処理を呼び出すラッパー関数（7行目以降の医薬品名比較）
Public Sub RunDrugNameComparison()
    ' MainModuleの関数を呼び出し
    MainModule.CompareAndTransferDrugNamesFromRow7
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
             Formula1:="(未定義),その他(なし),包装小,調剤用,PTP,分包,バラ,SP,PTP(患者用)"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "包装形態の選択"
        .ErrorTitle = "無効な選択"
        .InputMessage = "リストから包装形態を選択してください"
        .ErrorMessage = "リストから有効な包装形態を選択してください"
    End With
    
    ' B4セルの書式設定
    With settingsSheet.Range("B4")
        .Value = "PTP" ' デフォルト値を設定
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242) ' 薄い青色の背景
    End With
    
    ' A4セルにラベルを設定
    With settingsSheet.Range("A4")
        .Value = "包装形態:"
        .Font.Bold = True
    End With
    
    MsgBox "包装形態のドロップダウンリストを設定しました。", vbInformation
End Sub

' シート1にインストラクションを追加する関数
Public Sub AddInstructions()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    
    ' 既存の指示を削除
    ws.Range("A2:C3").ClearContents
    
    ' 指示を追加
    ws.Range("A2").Value = "【使い方】"
    ws.Range("A3").Value = "1. B4の包装形態を選択して下さい"
    ws.Range("A4").Value = "包装形態:"
    ws.Range("B4").Font.Bold = True
    
    ' セルの書式設定
    ws.Range("A2").Font.Bold = True
    ws.Range("A2").Font.Size = 12
    
    ' 実行方法の指示
    ws.Range("A5").Value = "2. B7以降に検索する医薬品名を入力"
    ws.Range("A6").Value = "No."
    ws.Range("B6").Value = "検索医薬品名"
    ws.Range("C6").Value = "一致医薬品名"
    
    With ws.Range("A6:C6")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247) ' 薄い青色の背景
    End With
    
    ' 実行方法の案内
    Dim note As String
    note = "※実行方法: メニューから「ツール」→「マクロ」を選択し、" & vbCrLf & _
           "「RunDrugNameComparison」を選んで「実行」ボタンをクリックします。"
    
    ws.Range("A" & (ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 2)).Value = note
    
    MsgBox "使用方法の指示を追加しました。Mac版Excelでは、メニューから「ツール」→「マクロ」を選択し、" & vbCrLf & _
           "「RunDrugNameComparison」を選んで処理を実行してください。", vbInformation
End Sub

' 初期化用関数を追加（Mac互換版）
Public Sub InitializeMacWorkbook()
    On Error GoTo ErrorHandler
    
    ' ワークシートの参照を取得
    Dim settingsSheet As Worksheet
    Dim targetSheet As Worksheet
    
    Set settingsSheet = ThisWorkbook.Worksheets(1) ' 設定シート
    Set targetSheet = ThisWorkbook.Worksheets(2)   ' 比較対象のシート
    
    ' シート1の設定
    With settingsSheet
        ' タイトル設定
        .Range("A1:C1").Merge
        .Range("A1").Value = "医薬品名比較ツール（Mac版）"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Interior.Color = RGB(180, 198, 231) ' 青色の背景
        
        ' 使い方
        .Range("A2").Value = "【使い方】"
        .Range("A2").Font.Bold = True
        .Range("A3").Value = "1. B4の包装形態を選択"
        .Range("A4").Value = "包装形態:"
        .Range("A4").Font.Bold = True
        
        ' ドロップダウンリスト設定
        With .Range("B4").Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, _
                 Formula1:="(未定義),その他(なし),包装小,調剤用,PTP,分包,バラ,SP,PTP(患者用)"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        
        ' B4セル設定
        .Range("B4").Value = "PTP"
        .Range("B4").Font.Bold = True
        .Range("B4").Interior.Color = RGB(217, 225, 242)
        
        ' 手順
        .Range("A5").Value = "2. B7以降に検索する医薬品名を入力"
        .Range("A5").Font.Bold = True
        
        ' ヘッダー
        .Range("A6").Value = "No."
        .Range("B6").Value = "検索医薬品名"
        .Range("C6").Value = "一致医薬品名"
        .Range("A6:C6").Font.Bold = True
        .Range("A6:C6").Interior.Color = RGB(221, 235, 247)
        
        ' 列幅
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 40
        
        ' 行番号
        Dim i As Long
        For i = 7 To 30
            .Cells(i, "A").Value = i - 6
        Next i
        
        ' 実行方法の案内
        .Range("A32").Value = "※実行方法: メニューから「ツール」→「マクロ」を選択し、「RunDrugNameComparison」を実行"
        .Range("A32").Font.Italic = True
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
    
    MsgBox "ワークブックを初期化しました。" & vbNewLine & _
           "1. 設定シートのB4セルで包装形態を選択" & vbNewLine & _
           "2. シート2に比較対象の医薬品名を入力" & vbNewLine & _
           "3. 設定シートのB7以降に検索する医薬品名を入力" & vbNewLine & _
           "4. メニューの「ツール」→「マクロ」から「RunDrugNameComparison」を実行", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

' Mac版で7行目以降を処理するための専用関数
Public Sub RunDrugNameComparisonMac()
    ' Applicationオブジェクトを使わない簡易版
    Dim settingsSheet As Worksheet, targetSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    Set targetSheet = ThisWorkbook.Worksheets(2)
    
    ' 包装形態と医薬品名を取得
    Dim packageType As String
    packageType = settingsSheet.Range("B4").Value
    
    ' 比較対象の医薬品名を取得
    Dim lastRowTarget As Long
    lastRowTarget = targetSheet.Range("B" & Rows.Count).End(xlUp).Row
    
    Dim targetDrugs() As String
    ReDim targetDrugs(1 To lastRowTarget - 1)
        
    Dim i As Long
    For i = 2 To lastRowTarget
        targetDrugs(i - 1) = targetSheet.Range("B" & i).Value
    Next i
    
    ' 7行目以降の処理
    Dim processedCount As Long
    processedCount = 0
    
    Dim lastRowSettings As Long
    lastRowSettings = settingsSheet.Range("B" & Rows.Count).End(xlUp).Row
    
    For i = 7 To lastRowSettings
        Dim searchDrug As String
        searchDrug = settingsSheet.Range("B" & i).Value
        
        If Len(searchDrug) > 0 Then
            ' 最適な一致の検索
            Dim bestMatch As String
            bestMatch = MainModule.FindBestMatchingDrugForMac(searchDrug, targetDrugs, packageType)
            
            ' 結果を転記
            settingsSheet.Range("C" & i).Value = bestMatch
            
            If Len(bestMatch) > 0 Then
                processedCount = processedCount + 1
            End If
        End If
    Next i
    
    MsgBox "処理が完了しました。" & vbNewLine & _
           processedCount & "件の医薬品名が一致しました。", vbInformation
End Sub

