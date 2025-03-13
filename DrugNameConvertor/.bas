' モジュールの先頭に追加
Option Explicit

' ユーザー定義型の宣言（モジュールレベルで定義）
Private Type DrugNameParts
    BaseName As String    ' 基本名（例：タモキシフェン）
    FormType As String    ' 剤型（例：錠、カプセル）
    Strength As String    ' 規格（例：20mg）
    Maker As String       ' メーカー名（例：MYL、サワイ）
    Package As String     ' 包装形態（例：PTP）
End Type

Sub SearchAndTransferDrugData()
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
    
    Dim i As Long, j As Long, k As Long
    Dim inputValue As Variant
    Dim drugNameFromSheet3 As String
    Dim drugNameFromSheet2 As String
    Dim packageType As String
    
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

'入力値を処理する関数
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

'スラッシュで囲まれた文字列を取得する関数
Private Function GetPackageType(ByVal text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = "/([^/]+)/"
    regex.Global = True
    
    Dim matches As Object
    Set matches = regex.Execute(text)
    
    If matches.Count > 0 Then
        GetPackageType = matches(0).SubMatches(0)
    Else
        GetPackageType = ""
    End If
End Function

Private Function CompareDrugStringsWithRate(ByVal sourceStr As String, ByVal targetStr As String) As Double
    Dim sourceParts As DrugNameParts
    Dim targetParts As DrugNameParts
    Dim matchCount As Integer
    Dim totalItems As Integer
    
    sourceParts = ParseDrugString(sourceStr)
    targetParts = ParseDrugString(targetStr)
    
    totalItems = 5 ' 基本名、剤型、規格、メーカー、包装の5項目
    matchCount = 0
    
    ' 基本名の比較（完全一致）
    If StrComp(sourceParts.BaseName, targetParts.BaseName, vbTextCompare) = 0 Then
        matchCount = matchCount + 1
    End If
    
    ' 剤型の比較（完全一致）
    If StrComp(sourceParts.FormType, targetParts.FormType, vbTextCompare) = 0 Then
        matchCount = matchCount + 1
    End If
    
    ' 規格の比較（数値と単位を正規化して比較）
    If CompareStrength(sourceParts.Strength, targetParts.Strength) Then
        matchCount = matchCount + 1
    End If
    
    ' メーカー名の比較（完全一致）
    If StrComp(sourceParts.Maker, targetParts.Maker, vbTextCompare) = 0 Then
        matchCount = matchCount + 1
    End If
    
    ' 包装形態の比較（ある程度の揺れを許容）
    If ComparePackageType(sourceParts.Package, targetParts.Package) Then
        matchCount = matchCount + 1
    End If
    
    ' 一致率を計算（百分率）
    CompareDrugStringsWithRate = (matchCount / totalItems) * 100
End Function

Private Function ParseDrugString(ByVal drugStr As String) As DrugNameParts
    Dim result As DrugNameParts
    Dim tempStr As String
    
    ' 全角文字を半角に変換
    tempStr = StrConv(drugStr, vbNarrow)
    
    ' メーカー名を抽出 (「」内)
    Dim makerMatch As String
    makerMatch = ExtractBetweenQuotes(tempStr)
    result.Maker = makerMatch
    
    ' 規格を抽出 (数字+単位)
    Dim strengthMatch As String
    strengthMatch = ExtractStrength(tempStr)
    result.Strength = strengthMatch
    
    ' 剤型を抽出
    Dim formMatch As String
    formMatch = ExtractFormType(tempStr)
    result.FormType = formMatch
    
    ' 包装形態を抽出
    result.Package = ExtractPackageType(tempStr)
    
    ' 基本名を抽出（メーカー名と規格の前まで）
    result.BaseName = ExtractBaseName(tempStr, result.Maker, result.Strength, result.FormType)
    
    ParseDrugString = result
End Function

Private Function ExtractBetweenQuotes(ByVal text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = "「([^」]+)」"
    Dim matches As Object
    Set matches = regex.Execute(text)
    
    If matches.Count > 0 Then
        ExtractBetweenQuotes = matches(0).SubMatches(0)
    Else
        ExtractBetweenQuotes = ""
    End If
End Function

Private Function CompareStrength(ByVal str1 As String, ByVal str2 As String) As Boolean
    ' 数値と単位を分離して比較
    Dim num1 As Double, num2 As Double
    Dim unit1 As String, unit2 As String
    
    ' 数値と単位を抽出
    ExtractNumberAndUnit str1, num1, unit1
    ExtractNumberAndUnit str2, num2, unit2
    
    ' 数値と単位が両方一致する場合のみTrue
    CompareStrength = (num1 = num2) And (StrComp(unit1, unit2, vbTextCompare) = 0)
End Function

Private Sub ExtractNumberAndUnit(ByVal str As String, ByRef num As Double, ByRef unit As String)
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' 数字と単位を分離
    regex.Pattern = "(\d+\.?\d*)\s*(mg|g|ml|μg)"
    regex.IgnoreCase = True
    
    Dim matches As Object
    Set matches = regex.Execute(str)
    
    If matches.Count > 0 Then
        num = CDbl(matches(0).SubMatches(0))
        unit = LCase(matches(0).SubMatches(1))
    Else
        num = 0
        unit = ""
    End If
End Sub

Private Function ComparePackageType(ByVal type1 As String, ByVal type2 As String) As Boolean
    ' 包装形態の表記揺れに対応
    Dim normalizedType1 As String, normalizedType2 As String
    
    normalizedType1 = NormalizePackageType(type1)
    normalizedType2 = NormalizePackageType(type2)
    
    ComparePackageType = (StrComp(normalizedType1, normalizedType2, vbTextCompare) = 0)
End Function

Private Function NormalizePackageType(ByVal packageType As String) As String
    ' 包装形態の表記を標準化
    Select Case UCase(Trim(packageType))
        Case "PTP", "ＰＴＰ", "P.T.P.", "P.T.P"
            NormalizePackageType = "PTP"
        Case "ﾊﾞﾗ", "バラ", "BARA"
            NormalizePackageType = "バラ"
        Case Else
            NormalizePackageType = UCase(Trim(packageType))
    End Select
End Function

Sub ProcessDrugNamesWithMatchRate()
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

' 不足している関数の追加
Private Function ExtractStrength(ByVal text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = "(\d+\.?\d*\s*(?:mg|g|ml|μg))"
    regex.IgnoreCase = True
    
    Dim matches As Object
    Set matches = regex.Execute(text)
    
    If matches.Count > 0 Then
        ExtractStrength = matches(0).Value
    Else
        ExtractStrength = ""
    End If
End Function

Private Function ExtractFormType(ByVal text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' 剤型のパターン
    regex.Pattern = "(錠|カプセル|細粒|顆粒|散|シロップ|ドライシロップ|注射液|注射用|軟膏|クリーム|ゲル|テープ|パップ|点眼液)"
    
    Dim matches As Object
    Set matches = regex.Execute(text)
    
    If matches.Count > 0 Then
        ExtractFormType = matches(0).Value
    Else
        ExtractFormType = ""
    End If
End Function

Private Function ExtractBaseName(ByVal text As String, _
                               ByVal maker As String, _
                               ByVal strength As String, _
                               ByVal formType As String) As String
    Dim result As String
    result = text
    
    ' メーカー名を除去
    If maker <> "" Then
        result = Replace(result, "「" & maker & "」", "")
    End If
    
    ' 規格を除去
    If strength <> "" Then
        result = Replace(result, strength, "")
    End If
    
    ' 剤型を除去
    If formType <> "" Then
        result = Replace(result, formType, "")
    End If
    
    ' 数量表現を除去（例：10錠）
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\d+\s*(?:錠|カプセル|包|個|枚|本|管|筒|組|袋)"
    regex.Global = True
    result = regex.Replace(result, "")
    
    ' 特殊文字と余分な空白を除去
    result = Replace(result, "　", " ")  ' 全角スペースを半角に
    result = Trim(result)
    
    ExtractBaseName = result
End Function

' メイン処理の修正
Sub MainProcess()
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
                                         " 剤型:" & sourceParts.FormType & _
                                         " 規格:" & sourceParts.Strength & _
                                         " メーカー:" & sourceParts.Maker
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

Private Function ExtractPackageType(ByVal text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' 包装形態のパターンを定義
    regex.Pattern = "(PTP|ＰＴＰ|P\.T\.P\.|P\.T\.P|バラ|ﾊﾞﾗ|BARA|分包|SP|ＳＰ|瓶|ボトル|管|アンプル|シリンジ)"
    regex.IgnoreCase = True
    
    Dim matches As Object
    Set matches = regex.Execute(text)
    
    If matches.Count > 0 Then
        ' 見つかった包装形態を標準化
        ExtractPackageType = NormalizePackageType(matches(0).Value)
    Else
        ' スラッシュで囲まれた部分を検索（既存のGetPackageType関数の処理）
        regex.Pattern = "/([^/]+)/"
        Set matches = regex.Execute(text)
        
        If matches.Count > 0 Then
            ExtractPackageType = NormalizePackageType(matches(0).SubMatches(0))
        Else
            ExtractPackageType = ""
        End If
    End If
End Function