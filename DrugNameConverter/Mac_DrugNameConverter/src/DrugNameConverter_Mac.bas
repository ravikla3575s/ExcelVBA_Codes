Option Explicit

' ラッパーモジュール - 基本機能 (Mac版)
' Mac版ではMSFormsが利用できないため、ステータスバーを使用した進捗表示に変更しています

' メイン処理を呼び出すラッパー関数（選択された包装形態に応じて複数の包装単位で処理）
Public Sub RunDrugNameComparison()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' ワークシート参照の取得
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' 全体で使用する空の配列を一度だけ宣言
    Dim emptyStringArray() As String
    Dim emptyLongArray() As Long
    ReDim emptyStringArray(0 To 0)
    ReDim emptyLongArray(0 To 0)
    
    ' 包装形態の取得
    Dim packageSelection As String
    packageSelection = settingsSheet.Range("B4").Value
    
    ' 包装単位の配列を選択に応じて設定
    Dim packageTypes As Variant
    Dim totalProcessed As Long, totalSkipped As Long
    totalProcessed = 0
    totalSkipped = 0
    
    If packageSelection = "バラ包装" Then
        packageTypes = Array("バラ", "調剤用")
    ElseIf packageSelection = "分包品" Then
        packageTypes = Array("PTP", "分包", "SP", "包装小", "PTP(患者用)")
    Else
        MsgBox "有効な包装形態を選択してください。", vbExclamation
        GoTo CleanExit
    End If
    
    ' 進捗状況をステータスバーに表示
    Application.StatusBar = "医薬品コードから医薬品名を取得しています..."
    
    ' 医薬品コードから取得した医薬品名を配列に保存
    Dim drugNames() As String
    Dim drugCodes() As String
    Dim rowIndices() As Long
    Call GetDrugNamesFromCodes(drugNames, drugCodes, rowIndices)
    
    ' 処理結果メッセージ
    Dim resultMsg As String
    resultMsg = "処理結果:" & vbCrLf & vbCrLf
    
    ' ここから包装形態の自動判定と処理の実装
    If UBound(drugNames) >= LBound(drugNames) Then
        ' 最初の医薬品から包装形態を判定
        Application.StatusBar = "最初の医薬品から包装形態を自動判定しています..."
        
        Dim firstDrugName As String
        firstDrugName = drugNames(LBound(drugNames))
        
        ' 包装形態を判定する
        Dim detectedPackageType As String
        detectedPackageType = DetectPackageType(firstDrugName)
        
        If Len(detectedPackageType) > 0 Then
            ' 包装形態が判定できた場合
            settingsSheet.Range("E4").Value = detectedPackageType & "(自動判定)"
            
            ' 自動判定された包装形態に属する包装単位を処理
            Dim matchingPackageTypes As Variant
            If InStr(1, "バラ 調剤用", detectedPackageType, vbTextCompare) > 0 Then
                matchingPackageTypes = Array("バラ", "調剤用")
            Else
                matchingPackageTypes = Array("PTP", "分包", "SP", "包装小", "PTP(患者用)")
            End If
            
            ' 最初に判定された包装形態で処理
            Dim primaryPackageType As String
            primaryPackageType = detectedPackageType
            
            ' ステータスバーを更新
            Application.StatusBar = primaryPackageType & "形態で医薬品名を比較しています..."
            
            ' 処理用の変数
            Dim processed As Long, skipped As Long
            Dim skippedDrugNames() As String
            Dim skippedDrugCodes() As String
            Dim skippedRowIndices() As Long
            
            ' 主包装形態での処理
            Call ProcessPackageType(primaryPackageType, drugNames, drugCodes, rowIndices, processed, skipped, skippedDrugNames, skippedDrugCodes, skippedRowIndices)
            
            totalProcessed = totalProcessed + processed
            totalSkipped = totalSkipped + skipped
            
            resultMsg = resultMsg & primaryPackageType & "(自動判定): " & processed & "件一致、" & skipped & "件不一致" & vbCrLf
            
            ' 不一致だった項目を他の包装形態で処理
            If Not Not skippedDrugNames Then ' 配列が初期化されているか確認
                If UBound(skippedDrugNames) >= LBound(skippedDrugNames) Then
                    Dim j As Long, otherProcessed As Long, otherSkipped As Long
                    Dim secondaryPackageTypes As Variant
                    
                    ' バラ包装と分包品で異なる配列を設定
                    If InStr(1, "バラ 調剤用", primaryPackageType, vbTextCompare) > 0 Then
                        secondaryPackageTypes = Array("PTP", "分包", "SP", "包装小")
                    Else
                        secondaryPackageTypes = Array("バラ", "調剤用")
                    End If
                    
                    ' 他の包装形態で処理
                    For j = LBound(secondaryPackageTypes) To UBound(secondaryPackageTypes)
                        Dim secondaryPackageType As String
                        secondaryPackageType = secondaryPackageTypes(j)
                        
                        ' ステータスバーを更新
                        Application.StatusBar = secondaryPackageType & "形態でスキップされた医薬品を再処理中..."
                        
                        ' スキップされた項目を処理
                        Call ProcessPackageType(secondaryPackageType, skippedDrugNames, skippedDrugCodes, skippedRowIndices, otherProcessed, otherSkipped, emptyStringArray, emptyStringArray, emptyLongArray)
                        
                        totalProcessed = totalProcessed + otherProcessed
                        
                        resultMsg = resultMsg & secondaryPackageType & "(再処理): " & otherProcessed & "件一致" & vbCrLf
                    Next j
                End If
            End If
        Else
            ' 包装形態が判定できない場合は、通常の処理を実行
            Application.StatusBar = "包装形態を自動判定できませんでした。通常の処理を実行します。"
            
            ' 各包装単位で処理を実行
            Dim i As Long
            For i = LBound(packageTypes) To UBound(packageTypes)
                ' 現在処理中の包装単位を表示
                settingsSheet.Range("E4").Value = packageTypes(i)
                
                ' ステータスバーを更新
                Application.StatusBar = packageTypes(i) & "形態での医薬品名を比較しています..."
                
                ' 包装単位ごとの処理
                Call ProcessPackageType(packageTypes(i), drugNames, drugCodes, rowIndices, processed, skipped, emptyStringArray, emptyStringArray, emptyLongArray)
                
                totalProcessed = totalProcessed + processed
                totalSkipped = totalSkipped + skipped
                
                resultMsg = resultMsg & packageTypes(i) & ": " & processed & "件一致、" & skipped & "件不一致" & vbCrLf
            Next i
        End If
    End If
    
    settingsSheet.Range("E4").Value = "完了"
    
    resultMsg = resultMsg & vbCrLf & "合計: " & totalProcessed & "件一致、" & totalSkipped & "件不一致"
    
CleanExit:
    ' ステータスバーをクリア
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox resultMsg, vbInformation
    Exit Sub
    
ErrorHandler:
    ' ステータスバーをクリア
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' 医薬品コードから医薬品名を取得して配列に保存する関数
Private Sub GetDrugNamesFromCodes(ByRef drugNames() As String, ByRef drugCodes() As String, ByRef rowIndices() As Long)
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
    ReDim rowIndices(1 To count)
    
    ' 各行のコードを処理
    Dim i As Long, idx As Long
    idx = 1
    
    For i = 7 To lastRow
        ' 進捗状況を更新（10%ごとに表示）
        If (i - 7) Mod (Application.WorksheetFunction.Max(1, (lastRow - 7) / 10)) = 0 Then
            Application.StatusBar = "医薬品コード " & (i - 6) & "/" & count & " を処理中..."
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
            rowIndices(idx) = i
            idx = idx + 1
        End If
    Next i
    
    ' 配列のサイズを実際に使用した分に調整
    If idx > 1 Then
        ReDim Preserve drugNames(1 To idx - 1)
        ReDim Preserve drugCodes(1 To idx - 1)
        ReDim Preserve rowIndices(1 To idx - 1)
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "医薬品名の取得中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 包装形態ごとの処理を行う関数
Private Sub ProcessPackageType(ByVal packageType As String, ByRef drugNames() As String, ByRef drugCodes() As String, ByRef rowIndices() As Long, ByRef processedCount As Long, ByRef skippedCount As Long, ByRef skippedDrugNames() As String, ByRef skippedDrugCodes() As String, ByRef skippedRowIndices() As Long)
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
    
    ' スキップされた項目を保存するための一時配列
    Dim tempSkippedDrugNames() As String
    Dim tempSkippedDrugCodes() As String
    Dim tempSkippedRowIndices() As Long
    
    ' 処理するアイテム数に基づいて一時配列のサイズを設定
    Dim tempSkippedCount As Long
    tempSkippedCount = 0
    
    If Not Not drugNames Then ' 配列が初期化されているか確認
        If UBound(drugNames) >= LBound(drugNames) Then
            ReDim tempSkippedDrugNames(LBound(drugNames) To UBound(drugNames))
            ReDim tempSkippedDrugCodes(LBound(drugNames) To UBound(drugNames))
            ReDim tempSkippedRowIndices(LBound(drugNames) To UBound(drugNames))
        End If
    End If
    
    ' 医薬品名の比較と転記
    Dim bestMatch As String
    Dim drugCount As Long
    
    If Not Not drugNames Then
        drugCount = UBound(drugNames)
        
        For i = LBound(drugNames) To UBound(drugNames)
            ' 進捗状況を更新（10%ごとに表示）
            If (i - LBound(drugNames)) Mod (Application.WorksheetFunction.Max(1, drugCount / 10)) = 0 Then
                Application.StatusBar = packageType & ": 医薬品名 " & i & "/" & drugCount & " を比較中..."
                DoEvents
            End If
            
            ' 医薬品名が配列にあれば処理
            If Len(drugNames(i)) > 0 Then
                ' 包装規格も自動的に考慮
                bestMatch = MainModule.FindBestMatchingDrug(drugNames(i), targetDrugs, packageType)
                
                ' 一致する結果があれば転記、なければスキップ
                If Len(bestMatch) > 0 Then
                    ' 対応する行を取得
                    Dim rowIndex As Long
                    rowIndex = rowIndices(i)
                    
                    ' B列に値がなければ、または既存の値よりも良い一致があれば上書き
                    If Len(settingsSheet.Cells(rowIndex, "B").Value) = 0 Then
                        settingsSheet.Cells(rowIndex, "B").Value = bestMatch
                        processedCount = processedCount + 1
                    End If
                Else
                    ' スキップされた項目を一時配列に保存
                    If Not IsMissing(skippedDrugNames) Then ' 配列が渡されているか確認
                        tempSkippedDrugNames(tempSkippedCount) = drugNames(i)
                        tempSkippedDrugCodes(tempSkippedCount) = drugCodes(i)
                        tempSkippedRowIndices(tempSkippedCount) = rowIndices(i)
                        tempSkippedCount = tempSkippedCount + 1
                    End If
                    
                    ' 一致しない場合は何もせずスキップカウントを増やす
                    skippedCount = skippedCount + 1
                End If
            End If
        Next i
    End If
    
    ' スキップされた項目を配列にコピー
    If Not IsMissing(skippedDrugNames) And tempSkippedCount > 0 Then
        ReDim skippedDrugNames(0 To tempSkippedCount - 1)
        ReDim skippedDrugCodes(0 To tempSkippedCount - 1)
        ReDim skippedRowIndices(0 To tempSkippedCount - 1)
        
        For i = 0 To tempSkippedCount - 1
            skippedDrugNames(i) = tempSkippedDrugNames(i)
            skippedDrugCodes(i) = tempSkippedDrugCodes(i)
            skippedRowIndices(i) = tempSkippedRowIndices(i)
        Next i
    ElseIf Not IsMissing(skippedDrugNames) Then
        ' スキップされた項目がない場合は空の配列を作成
        ReDim skippedDrugNames(0 To 0)
        ReDim skippedDrugCodes(0 To 0)
        ReDim skippedRowIndices(0 To 0)
    End If
End Sub

' 医薬品名から包装形態を判定する関数
Private Function DetectPackageType(ByVal drugName As String) As String
    Dim packages As Variant
    Dim i As Long
    
    ' バラ包装と分包品の包装単位
    Dim bulkPackages As Variant
    Dim unitPackages As Variant
    
    bulkPackages = Array("バラ", "調剤用")
    unitPackages = Array("PTP", "分包", "SP", "包装小", "PTP(患者用)")
    
    ' 医薬品名を解析
    Dim parts As DrugNameParts
    parts = ParseDrugString(drugName)
    
    ' 包装形態が明確に含まれている場合
    If Len(parts.Package) > 0 Then
        ' バラ包装系かチェック
        For i = LBound(bulkPackages) To UBound(bulkPackages)
            If InStr(1, parts.Package, bulkPackages(i), vbTextCompare) > 0 Then
                DetectPackageType = bulkPackages(i)
                Exit Function
            End If
        Next i
        
        ' 分包品系かチェック
        For i = LBound(unitPackages) To UBound(unitPackages)
            If InStr(1, parts.Package, unitPackages(i), vbTextCompare) > 0 Then
                DetectPackageType = unitPackages(i)
                Exit Function
            End If
        Next i
    End If
    
    ' 医薬品名全体で探す
    For i = LBound(bulkPackages) To UBound(bulkPackages)
        If InStr(1, drugName, bulkPackages(i), vbTextCompare) > 0 Then
            DetectPackageType = bulkPackages(i)
            Exit Function
        End If
    Next i
    
    For i = LBound(unitPackages) To UBound(unitPackages)
        If InStr(1, drugName, unitPackages(i), vbTextCompare) > 0 Then
            DetectPackageType = unitPackages(i)
            Exit Function
        End If
    Next i
    
    ' 判定できない場合
    DetectPackageType = ""
End Function

' 医薬品コードの検索と医薬品名比較を一括で実行する関数
Public Sub ProcessDrugCodesAndCompare()
    On Error GoTo ErrorHandler
    
    ' ステータスバーに開始メッセージを表示
    Application.StatusBar = "医薬品名の比較処理を開始しています..."
    
    ' ワークシート参照の取得
    Dim settingsSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1)
    
    ' 医薬品コードから医薬品名を取得して直接比較を実行
    Call RunDrugNameComparison
    
    Exit Sub
    
ErrorHandler:
    ' ステータスバーをクリア
    Application.StatusBar = False
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