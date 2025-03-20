Option Explicit

' メインの処理関数：薬品名の一致率に基づいて転記
Public Sub MainProcess()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets(1)
    Set ws2 = ThisWorkbook.Worksheets(2)
    Set ws3 = ThisWorkbook.Worksheets(3)
    
    Dim lastRow1 As Long, lastRow2 As Long
    lastRow1 = ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row
    
    Const MATCH_THRESHOLD As Double = 80 ' 一致率のしきい値（80%）
    
    Dim i As Long, j As Long
    For i = 2 To lastRow1
        Dim sourceStr As String
        sourceStr = ws1.Cells(i, "B").Value
        
        If Len(sourceStr) > 0 Then
            Dim maxMatchRate As Double
            Dim bestMatchIndex As Long
            maxMatchRate = 0
            bestMatchIndex = 0
            
            For j = 2 To lastRow2
                Dim targetStr As String
                targetStr = ws2.Cells(j, "B").Value
                
                Dim currentMatchRate As Double
                currentMatchRate = CompareDrugStringsWithRate(sourceStr, targetStr)
                
                If currentMatchRate > maxMatchRate Then
                    maxMatchRate = currentMatchRate
                    bestMatchIndex = j
                End If
            Next j
            
            ' 結果の出力
            If maxMatchRate >= MATCH_THRESHOLD Then
                ws1.Cells(i, "C").Value = ws2.Cells(bestMatchIndex, "B").Value
                ws1.Cells(i, "D").Value = maxMatchRate & "%"
                
                ' 一致した各要素の詳細を出力（デバッグ用）
                Dim sourceParts As DrugNameParts
                Dim targetParts As DrugNameParts
                sourceParts = ParseDrugString(sourceStr)
                targetParts = ParseDrugString(ws2.Cells(bestMatchIndex, "B").Value)
                
                ws1.Cells(i, "E").Value = "基本名:" & sourceParts.BaseName & _
                                         " 剤型:" & sourceParts.formType & _
                                         " 規格:" & sourceParts.strength & _
                                         " メーカー:" & sourceParts.maker
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "処理が完了しました。"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description
End Sub

' 薬品名の検索と転記関数
Public Sub SearchAndTransferDrugData()
    On Error GoTo ErrorHandler
    
    '画面更新を一時停止してパフォーマンス向上
    Application.ScreenUpdating = False
    
    'ワークシートの設定
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets(1)
    Set ws2 = ThisWorkbook.Worksheets(2)
    Set ws3 = ThisWorkbook.Worksheets(3)
    
    '最終行の取得
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row
    lastRow3 = ws3.Cells(ws3.Rows.Count, "F").End(xlUp).Row
    
    Dim i As Long
    Dim inputValue As Variant
    
    '各行のA列の値を処理
    For i = 2 To lastRow1  'ヘッダーをスキップ
        inputValue = ws1.Cells(i, "A").Value
        
        '入力値を処理する関数を呼び出し
        ProcessInputValue inputValue, ws1, ws2, ws3, i, lastRow2, lastRow3
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "処理が完了しました。"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description
End Sub

' 入力値を処理する関数
Private Sub ProcessInputValue(ByVal inputValue As Variant, _
                            ByRef ws1 As Worksheet, _
                            ByRef ws2 As Worksheet, _
                            ByRef ws3 As Worksheet, _
                            ByVal currentRow As Long, _
                            ByVal lastRow2 As Long, _
                            ByVal lastRow3 As Long)
                            
    Dim drugNameFromSheet3 As String
    Dim drugNameFromSheet2 As String
    Dim packageType As String
    Dim j As Long, k As Long
    
    'Sheet3から薬剤名を検索
    For k = 2 To lastRow3
        drugNameFromSheet3 = ws3.Cells(k, "F").Value
        If InStr(1, inputValue, drugNameFromSheet3) > 0 Then
            'Sheet2から対応する薬剤名を検索
            For j = 2 To lastRow2
                drugNameFromSheet2 = ws2.Cells(j, "B").Value
                If drugNameFromSheet2 = drugNameFromSheet3 Then
                    '包装タイプを取得
                    packageType = GetPackageType(inputValue)
                    
                    'データを転記
                    ws1.Cells(currentRow, "B").Value = ws2.Cells(j, "A").Value
                    ws1.Cells(currentRow, "C").Value = packageType
                    Exit For
                End If
            Next j
            Exit For
        End If
    Next k
End Sub

' 一致率計算による薬品名処理関数
Public Sub ProcessDrugNamesWithMatchRate()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets(1)
    Set ws2 = ThisWorkbook.Worksheets(2)
    
    Dim i As Long, j As Long
    Dim lastRow1 As Long, lastRow2 As Long
    Const MATCH_THRESHOLD As Double = 80 ' 一致率のしきい値（80%）
    
    lastRow1 = ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row
    
    For i = 2 To lastRow1
        Dim sourceStr As String
        Dim maxMatchRate As Double
        Dim bestMatchIndex As Long
        
        sourceStr = ws1.Cells(i, "B").Value
        maxMatchRate = 0
        bestMatchIndex = 0
        
        For j = 2 To lastRow2
            Dim targetStr As String
            Dim currentMatchRate As Double
            
            targetStr = ws2.Cells(j, "B").Value
            currentMatchRate = CompareDrugStringsWithRate(sourceStr, targetStr)
            
            If currentMatchRate > maxMatchRate Then
                maxMatchRate = currentMatchRate
                bestMatchIndex = j
            End If
        Next j
        
        ' しきい値以上の一致率があった場合のみ転記
        If maxMatchRate >= MATCH_THRESHOLD Then
            ws1.Cells(i, "C").Value = ws2.Cells(bestMatchIndex, "B").Value
            ws1.Cells(i, "D").Value = maxMatchRate & "%"
        End If
    Next i
    
    MsgBox "処理が完了しました。"
End Sub

' 設定シートの包装形態を考慮した医薬品名比較と転記を行う
Public Sub CompareAndTransferDrugNamesByPackage()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' ワークシートの設定
    Dim wsSettings As Worksheet, wsTarget As Worksheet
    Set wsSettings = ThisWorkbook.Worksheets(1) ' 設定シート
    Set wsTarget = ThisWorkbook.Worksheets(2)   ' 比較対象のシート
    
    ' B4セルから包装形態を取得
    Dim packageType As String
    packageType = wsSettings.Range("B4").Value
    
    ' 最終行を取得
    Dim lastRowSettings As Long, lastRowTarget As Long
    lastRowSettings = wsSettings.Cells(wsSettings.Rows.Count, "B").End(xlUp).Row
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "B").End(xlUp).Row
    
    ' 検索対象と比較対象の医薬品名を配列に格納
    Dim searchDrugs() As String
    Dim targetDrugs() As String
    Dim i As Long, j As Long
    
    ' 検索医薬品用の配列を初期化
    ReDim searchDrugs(1 To lastRowSettings - 1) ' ヘッダー行を除く
    For i = 2 To lastRowSettings
        searchDrugs(i - 1) = wsSettings.Cells(i, "B").Value
    Next i
    
    ' 比較対象用の配列を初期化
    ReDim targetDrugs(1 To lastRowTarget - 1) ' ヘッダー行を除く
    For i = 2 To lastRowTarget
        targetDrugs(i - 1) = wsTarget.Cells(i, "B").Value
    Next i
    
    ' 各検索医薬品に対して比較処理
    For i = 2 To lastRowSettings
        Dim searchDrug As String
        searchDrug = wsSettings.Cells(i, "B").Value
        
        If Len(searchDrug) > 0 Then
            Dim bestMatch As String
            bestMatch = FindBestMatchWithPackage(searchDrug, targetDrugs, packageType)
            
            If Len(bestMatch) > 0 Then
                ' 一致した医薬品名をC列に転記
                wsSettings.Cells(i, "C").Value = bestMatch
            Else
                ' 一致しなかった場合は空欄にする
                wsSettings.Cells(i, "C").Value = ""
            End If
        End If
    Next i
    
CleanExit:
    Application.ScreenUpdating = True
    MsgBox "処理が完了しました。", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

' 医薬品名の成分、規格、単位の一致度を計算
Public Function CalculateMatchScore(ByRef searchParts As DrugNameParts, ByRef targetParts As DrugNameParts) As Double
    Dim score As Double
    Dim totalWeight As Double
    
    score = 0
    totalWeight = 0
    
    ' 成分名の比較（重み: 50%）
    If StrComp(searchParts.BaseName, targetParts.BaseName, vbTextCompare) = 0 Then
        score = score + 50
    End If
    totalWeight = totalWeight + 50
    
    ' 剤型の比較（重み: 20%）
    If StrComp(searchParts.formType, targetParts.formType, vbTextCompare) = 0 Then
        score = score + 20
    End If
    totalWeight = totalWeight + 20
    
    ' 規格の比較（重み: 30%）
    If CompareStrength(searchParts.strength, targetParts.strength) Then
        score = score + 30
    End If
    totalWeight = totalWeight + 30
    
    ' スコアの正規化（百分率）
    If totalWeight > 0 Then
        CalculateMatchScore = (score / totalWeight) * 100
    Else
        CalculateMatchScore = 0
    End If
End Function

' 包装形態を考慮した最適な医薬品名の一致を検索する
Private Function FindBestMatchWithPackage(ByVal searchDrug As String, ByRef targetDrugs() As String, ByVal requiredPackage As String) As String
    Dim i As Long
    Dim bestMatchScore As Double
    Dim bestMatchIndex As Long
    Dim searchParts As DrugNameParts
    
    ' 検索対象の医薬品名を分解
    searchParts = ParseDrugString(searchDrug)
    bestMatchScore = 0
    bestMatchIndex = -1
    
    ' 各比較対象に対してスコアを計算
    For i = LBound(targetDrugs) To UBound(targetDrugs)
        Dim targetParts As DrugNameParts
        Dim currentScore As Double
        Dim hasRequiredPackage As Boolean
        
        ' 比較対象の医薬品名を分解
        targetParts = ParseDrugString(targetDrugs(i))
        
        ' 包装形態の確認
        hasRequiredPackage = (InStr(1, targetParts.Package, requiredPackage, vbTextCompare) > 0)
        
        If hasRequiredPackage Then
            ' 成分名、規格、単位の一致度を計算
            currentScore = CalculateMatchScore(searchParts, targetParts)
            
            If currentScore > bestMatchScore Then
                bestMatchScore = currentScore
                bestMatchIndex = i
            End If
        End If
    Next i
    
    ' 一定以上のスコアがある場合のみ結果を返す
    If bestMatchScore >= 70 And bestMatchIndex >= 0 Then ' 70%以上の一致率
        FindBestMatchWithPackage = targetDrugs(bestMatchIndex)
    Else
        FindBestMatchWithPackage = ""
    End If
End Function

' 7行目以降の医薬品名比較と転記を行う関数
Public Sub CompareAndTransferDrugNamesFromRow7()
    On Error GoTo ErrorHandler
    
    ' 初期設定
    Application.ScreenUpdating = False
    
    ' ワークシート参照の取得
    Dim settingsSheet As Worksheet, targetSheet As Worksheet
    Set settingsSheet = ThisWorkbook.Worksheets(1) ' 設定シート
    Set targetSheet = ThisWorkbook.Worksheets(2)   ' 比較対象のシート
    
    ' 包装形態の取得と確認
    Dim packageType As String
    packageType = settingsSheet.Range("B4").Value
    
    ' 有効な包装形態かチェック
    Dim validPackageTypes As Variant
    validPackageTypes = Array("(未定義)", "その他(なし)", "包装小", "調剤用", "PTP", "分包", "バラ", "SP", "PTP(患者用)")
    
    Dim isValidPackage As Boolean
    Dim i As Long
    isValidPackage = False
    
    For i = LBound(validPackageTypes) To UBound(validPackageTypes)
        If packageType = validPackageTypes(i) Then
            isValidPackage = True
            Exit For
        End If
    Next i
    
    If Not isValidPackage Then
        MsgBox "B4セルに有効な包装形態を設定してください。" & vbCrLf & _
               "有効な値: (未定義), その他(なし), 包装小, 調剤用, PTP, 分包, バラ, SP, PTP(患者用)", vbExclamation
        GoTo CleanExit
    End If
    
    ' 最終行の取得
    Dim lastRowSettings As Long, lastRowTarget As Long
    lastRowSettings = settingsSheet.Cells(settingsSheet.Rows.Count, "B").End(xlUp).Row
    lastRowTarget = targetSheet.Cells(targetSheet.Rows.Count, "B").End(xlUp).Row
    
    ' 比較対象薬品名を配列に格納
    Dim targetDrugs() As String
    ReDim targetDrugs(1 To lastRowTarget - 1)
    
    For i = 2 To lastRowTarget
        targetDrugs(i - 1) = targetSheet.Cells(i, "B").Value
    Next i
    
    ' 医薬品名の比較と転記（7行目から開始）
    Dim searchDrug As String, bestMatch As String
    Dim processedCount As Long
    processedCount = 0
    
    For i = 7 To lastRowSettings ' ここで7行目以降を処理
        searchDrug = settingsSheet.Cells(i, "B").Value
        
        If Len(searchDrug) > 0 Then
            ' 最適な一致を検索
            bestMatch = FindBestMatchingDrug(searchDrug, targetDrugs, packageType)
            
            ' C列に転記
            settingsSheet.Cells(i, "C").Value = bestMatch
            
            ' 処理数をカウント
            If Len(bestMatch) > 0 Then
                processedCount = processedCount + 1
            End If
        End If
    Next i
    
CleanExit:
    Application.ScreenUpdating = True
    MsgBox "処理が完了しました。" & vbCrLf & _
           processedCount & "件の医薬品名が一致しました。", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' 最も一致する医薬品名を検索する関数
Private Function FindBestMatchingDrug(ByVal searchDrug As String, ByRef targetDrugs() As String, ByVal packageType As String) As String
    Dim i As Long
    Dim bestMatchIndex As Long, bestMatchScore As Long, currentScore As Long
    
    bestMatchIndex = -1
    bestMatchScore = 0
    
    ' 検索対象をキーワードに分解
    Dim keywords As Variant
    keywords = ExtractKeywords(searchDrug)
    
    ' 包装形態の特別処理
    Dim skipPackageCheck As Boolean
    skipPackageCheck = (packageType = "(未定義)" Or packageType = "その他(なし)")
    
    ' 各比較対象に対して処理
    For i = LBound(targetDrugs) To UBound(targetDrugs)
        If Len(targetDrugs(i)) > 0 Then
            ' 包装形態チェック
            Dim matchesPackage As Boolean
            
            If skipPackageCheck Then
                ' 未定義またはその他の場合は包装形態チェックをスキップ
                matchesPackage = True
            Else
                ' 包装形態が一致するか確認
                matchesPackage = CheckPackageTypeMatch(targetDrugs(i), packageType)
            End If
            
            If matchesPackage Then
                ' キーワード一致率を計算
                currentScore = CalculateMatchingScore(keywords, targetDrugs(i))
                
                ' より高いスコアを記録
                If currentScore > bestMatchScore Then
                    bestMatchScore = currentScore
                    bestMatchIndex = i
                End If
            End If
        End If
    Next i
    
    ' 結果を返す（閾値以上のスコアの場合のみ）
    If bestMatchScore >= 50 And bestMatchIndex >= 0 Then
        FindBestMatchingDrug = targetDrugs(bestMatchIndex)
    Else
        FindBestMatchingDrug = ""
    End If
End Function

' 包装形態が一致するかチェックする関数
Private Function CheckPackageTypeMatch(ByVal drugName As String, ByVal packageType As String) As Boolean
    ' 包装形態のバリエーションを定義
    Dim packageVariations As Object
    Set packageVariations = CreateObject("Scripting.Dictionary")
    
    ' 各包装形態の異表記を登録
    packageVariations.Add "PTP", Array("PTP", "ＰＴＰ", "P.T.P.", "P.T.P")
    packageVariations.Add "バラ", Array("バラ", "ﾊﾞﾗ", "BARA", "バラ錠")
    packageVariations.Add "SP", Array("SP", "ＳＰ", "S.P")
    packageVariations.Add "分包", Array("分包", "ぶんぽう", "分包品")
    packageVariations.Add "包装小", Array("包装小", "小包装")
    packageVariations.Add "調剤用", Array("調剤用", "調剤", "調剤用包装")
    packageVariations.Add "PTP(患者用)", Array("PTP(患者用)", "患者用PTP", "患者用")
    
    ' 包装形態が定義されているか確認
    If Not packageVariations.Exists(packageType) Then
        ' 定義されていない場合は文字列完全一致で確認
        CheckPackageTypeMatch = (InStr(1, drugName, packageType, vbTextCompare) > 0)
        Exit Function
    End If
    
    ' 各バリエーションで確認
    Dim variations As Variant
    variations = packageVariations(packageType)
    
    Dim j As Long
    For j = LBound(variations) To UBound(variations)
        If InStr(1, drugName, variations(j), vbTextCompare) > 0 Then
            CheckPackageTypeMatch = True
            Exit Function
        End If
    Next j
    
    CheckPackageTypeMatch = False
End Function

' 医薬品名からキーワードを抽出する関数
Private Function ExtractKeywords(ByVal drugName As String) As Variant
    ' 全角スペースを半角に変換
    drugName = Replace(drugName, "　", " ")
    
    ' スペースで分割して配列に格納
    Dim words As Variant, result() As String
    Dim i As Long, validCount As Long
    
    words = Split(drugName, " ")
    ReDim result(UBound(words))
    validCount = 0
    
    ' 空でない要素のみ取得
    For i = 0 To UBound(words)
        If Trim(words(i)) <> "" Then
            result(validCount) = LCase(Trim(words(i)))
            validCount = validCount + 1
        End If
    Next i
    
    ' 結果が空の場合の処理
    If validCount = 0 Then
        ReDim result(0)
        result(0) = LCase(Trim(drugName))
        validCount = 1
    End If
    
    ReDim Preserve result(validCount - 1)
    ExtractKeywords = result
End Function

' キーワードの一致率を計算する関数
Private Function CalculateMatchingScore(ByRef keywords As Variant, ByVal targetDrug As String) As Long
    Dim i As Long, matchCount As Long
    Dim lowerTargetDrug As String
    
    lowerTargetDrug = LCase(targetDrug)
    matchCount = 0
    
    ' 各キーワードが含まれているかチェック
    For i = 0 To UBound(keywords)
        If InStr(1, lowerTargetDrug, keywords(i), vbTextCompare) > 0 Then
            matchCount = matchCount + 1
        End If
    Next i
    
    ' 一致率を計算（百分率）
    If UBound(keywords) >= 0 Then
        CalculateMatchingScore = (matchCount * 100) / (UBound(keywords) + 1)
    Else
        CalculateMatchingScore = 0
    End If
End Function

' Mac対応のための専用関数を追加
Public Function FindBestMatchingDrugForMac(ByVal searchDrug As String, ByRef targetDrugs() As String, ByVal packageType As String) As String
    Dim i As Long
    Dim bestMatchIndex As Long, bestMatchScore As Long, currentScore As Long
    
    bestMatchIndex = -1
    bestMatchScore = 0
    
    ' 検索対象をキーワードに分解
    Dim keywords As Variant
    keywords = StringUtils.ExtractKeywords(searchDrug)
    
    ' 包装形態の特別処理
    Dim skipPackageCheck As Boolean
    skipPackageCheck = (packageType = "(未定義)" Or packageType = "その他(なし)")
    
    ' 各比較対象に対して処理
    For i = LBound(targetDrugs) To UBound(targetDrugs)
        If Len(targetDrugs(i)) > 0 Then
            ' 包装形態チェック
            Dim matchesPackage As Boolean
            
            If skipPackageCheck Then
                ' 未定義またはその他の場合は包装形態チェックをスキップ
                matchesPackage = True
            Else
                ' 包装形態が一致するか確認
                matchesPackage = StringUtils.CheckPackageTypeMatch(targetDrugs(i), packageType)
            End If
            
            If matchesPackage Then
                ' キーワード一致率を計算
                currentScore = StringUtils.CalculateMatchingScore(keywords, targetDrugs(i))
                
                ' より高いスコアを記録
                If currentScore > bestMatchScore Then
                    bestMatchScore = currentScore
                    bestMatchIndex = i
                End If
            End If
        End If
    Next i
    
    ' 結果を返す（閾値以上のスコアの場合のみ）
    If bestMatchScore >= 50 And bestMatchIndex >= 0 Then
        FindBestMatchingDrugForMac = targetDrugs(bestMatchIndex)
    Else
        FindBestMatchingDrugForMac = ""
    End If
End Function

