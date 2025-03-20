Option Explicit

' 「」で囲まれたテキストを抽出する関数（正規表現を使わないバージョン）
Public Function ExtractBetweenQuotes(ByVal text As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(1, text, "「")
    If startPos > 0 Then
        endPos = InStr(startPos + 1, text, "」")
        If endPos > startPos Then
            ExtractBetweenQuotes = Mid(text, startPos + 1, endPos - startPos - 1)
        Else
            ExtractBetweenQuotes = ""
        End If
    Else
        ExtractBetweenQuotes = ""
    End If
End Function

' 規格（強度）を抽出する関数（正規表現を使わない版）
Public Function ExtractStrength(ByVal text As String) As String
    Dim i As Long
    Dim numStart As Long
    Dim result As String
    Dim inNumber As Boolean
    Dim units As Variant
    
    units = Array("mg", "g", "ml", "μg")
    inNumber = False
    numStart = 0
    
    For i = 1 To Len(text)
        Dim c As String
        c = Mid(text, i, 1)
        
        If IsNumeric(c) Or c = "." Then
            If Not inNumber Then
                inNumber = True
                numStart = i
            End If
        ElseIf c = " " And inNumber Then
            ' スペースは許容
        Else
            If inNumber Then
                ' 数字の後に単位があるか確認
                Dim j As Long
                Dim found As Boolean
                found = False
                
                For j = 0 To UBound(units)
                    If LCase(Mid(text, i, Len(units(j)))) = LCase(units(j)) Then
                        result = Mid(text, numStart, i - numStart + Len(units(j)))
                        found = True
                        Exit For
                    End If
                Next j
                
                If found Then
                    ExtractStrength = result
                    Exit Function
                End If
                
                inNumber = False
            End If
        End If
    Next i
    
    ExtractStrength = ""
End Function

' 数値と単位を分離する関数（正規表現を使わないバージョン）
Public Sub ExtractNumberAndUnit(ByVal str As String, ByRef num As Double, ByRef unit As String)
    Dim i As Long
    Dim numStr As String
    Dim unitStr As String
    Dim numStart As Long
    Dim inNumber As Boolean
    
    inNumber = False
    numStart = 0
    numStr = ""
    unitStr = ""
    
    For i = 1 To Len(str)
        Dim c As String
        c = Mid(str, i, 1)
        
        If IsNumeric(c) Or c = "." Then
            If Not inNumber Then
                inNumber = True
                numStart = i
            End If
        ElseIf c = " " And inNumber Then
            ' スペースは数字と見なす
        Else
            If inNumber Then
                numStr = Mid(str, numStart, i - numStart)
                unitStr = Mid(str, i)
                Exit For
            End If
        End If
    Next i
    
    ' 単位から不要な文字を削除
    unitStr = Trim(unitStr)
    
    ' 単位の標準化
    If LCase(Left(unitStr, 2)) = "mg" Then
        unitStr = "mg"
    ElseIf LCase(Left(unitStr, 1)) = "g" Then
        unitStr = "g"
    ElseIf LCase(Left(unitStr, 2)) = "ml" Then
        unitStr = "ml"
    ElseIf LCase(Left(unitStr, 2)) = "μg" Then
        unitStr = "μg"
    End If
    
    ' 結果を設定
    If Len(numStr) > 0 Then
        On Error Resume Next
        num = CDbl(numStr)
        If Err.Number <> 0 Then
            num = 0
        End If
        On Error GoTo 0
        unit = LCase(unitStr)
    Else
        num = 0
        unit = ""
    End If
End Sub

' 規格（強度）を比較する関数
Public Function CompareStrength(ByVal str1 As String, ByVal str2 As String) As Boolean
    ' 数値と単位を分離して比較
    Dim num1 As Double, num2 As Double
    Dim unit1 As String, unit2 As String
    
    ' 数値と単位を抽出
    ExtractNumberAndUnit str1, num1, unit1
    ExtractNumberAndUnit str2, num2, unit2
    
    ' 数値と単位が両方一致する場合のみTrue
    CompareStrength = (num1 = num2) And (StrComp(unit1, unit2, vbTextCompare) = 0)
End Function

' 設定シートのB列7行目以降の医薬品名をB4セルの包装形態に基づいて比較し、
' 一致するものをC列に転記する
Public Sub CompareAndTransferDrugNames()
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
    
    For i = 7 To lastRowSettings
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
    ' 包装形態のバリエーションを定義（CreateObjectを使わない実装）
    Dim PTPVariations As Variant
    Dim BulkVariations As Variant
    Dim SPVariations As Variant
    Dim DividedVariations As Variant
    Dim SmallPackageVariations As Variant
    Dim DispensingVariations As Variant
    Dim PatientPTPVariations As Variant
    
    ' 各包装形態の異表記を配列で定義
    PTPVariations = Array("PTP", "ＰＴＰ", "P.T.P.", "P.T.P")
    BulkVariations = Array("バラ", "ﾊﾞﾗ", "BARA", "バラ錠")
    SPVariations = Array("SP", "ＳＰ", "S.P")
    DividedVariations = Array("分包", "ぶんぽう", "分包品")
    SmallPackageVariations = Array("包装小", "小包装")
    DispensingVariations = Array("調剤用", "調剤", "調剤用包装")
    PatientPTPVariations = Array("PTP(患者用)", "患者用PTP", "患者用")
    
    ' 包装形態に応じた変数を選択
    Dim variations As Variant
    
    Select Case packageType
        Case "PTP"
            variations = PTPVariations
        Case "バラ"
            variations = BulkVariations
        Case "SP"
            variations = SPVariations
        Case "分包"
            variations = DividedVariations
        Case "包装小"
            variations = SmallPackageVariations
        Case "調剤用"
            variations = DispensingVariations
        Case "PTP(患者用)"
            variations = PatientPTPVariations
        Case Else
            ' 定義されていない場合は文字列完全一致で確認
            CheckPackageTypeMatch = (InStr(1, drugName, packageType, vbTextCompare) > 0)
            Exit Function
    End Select
    
    ' 各バリエーションで確認
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
    
    ' B3セルにタイトルを設定
    With settingsSheet.Range("A1:C1")
        .Merge
        .Value = "医薬品名比較ツール"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(180, 198, 231) ' 青色の背景
    End With
    
    ' 列ヘッダーを設定
    settingsSheet.Range("A6").Value = "No."
    settingsSheet.Range("B6").Value = "検索医薬品名"
    settingsSheet.Range("C6").Value = "一致医薬品名"
    
    With settingsSheet.Range("A6:C6")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247) ' 薄い青色の背景
    End With
    
    ' 列幅を調整
    settingsSheet.Columns("A").ColumnWidth = 5
    settingsSheet.Columns("B").ColumnWidth = 30
    settingsSheet.Columns("C").ColumnWidth = 40
    
    ' 行番号を設定（7行目から30行目まで）
    Dim i As Long
    For i = 7 To 30
        settingsSheet.Cells(i, "A").Value = i - 6
    Next i
    
    MsgBox "包装形態のドロップダウンリストを設定しました。", vbInformation
End Sub
