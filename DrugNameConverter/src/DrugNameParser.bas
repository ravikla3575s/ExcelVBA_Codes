Option Explicit

' 薬品名の構造体
Public Type DrugNameParts
    BaseName As String
    formType As String
    strength As String
    maker As String
    Package As String
    PackageSize As String ' 追加：包装規格（0.5g、5mLなど）
End Type

' 薬品名を解析して構造体に分解する関数
Public Function ParseDrugString(ByVal drugStr As String) As DrugNameParts
    Dim result As DrugNameParts
    Dim tempStr As String
    
    ' 全角文字を半角に変換
    tempStr = StrConv(drugStr, vbNarrow)
    
    ' メーカー名を抽出 (「」内)
    Dim makerMatch As String
    makerMatch = ExtractBetweenQuotes(tempStr)
    result.maker = makerMatch
    
    ' 規格を抽出 (数字+単位)
    Dim strengthMatch As String
    strengthMatch = ExtractStrengthSimple(tempStr)
    result.strength = strengthMatch
    
    ' 剤型を抽出
    Dim formMatch As String
    formMatch = ExtractFormTypeSimple(tempStr)
    result.formType = formMatch
    
    ' 包装形態を抽出
    result.Package = ExtractPackageTypeSimple(tempStr)
    
    ' 包装規格を抽出
    result.PackageSize = ExtractPackageSizeSimple(tempStr)
    
    ' 基本名を抽出（メーカー名と規格の前まで）
    result.BaseName = ExtractBaseNameSimple(tempStr, result.maker, result.strength, result.formType)
    
    ParseDrugString = result
End Function

' 薬品名の基本部分を抽出する関数（正規表現を使わないバージョン）
Public Function ExtractBaseNameSimple(ByVal text As String, _
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
    
    ' 数量表現を除去（例：10錠）- 正規表現を使わないバージョン
    Dim i As Long
    Dim parts() As String
    parts = Split(result, " ")
    
    For i = 0 To UBound(parts)
        If IsNumericWithSuffix(parts(i)) Then
            parts(i) = ""
        End If
    Next i
    
    result = Join(parts, " ")
    
    ' 特殊文字と余分な空白を除去
    result = Replace(result, "　", " ")  ' 全角スペースを半角に
    result = Trim(result)
    
    ' 連続するスペースを単一のスペースに置換
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    ExtractBaseNameSimple = result
End Function

' 数字+単位かどうかをチェックする（例：10錠）
Private Function IsNumericWithSuffix(ByVal text As String) As Boolean
    Dim i As Long
    Dim hasDigit As Boolean
    Dim hasSuffix As Boolean
    
    hasDigit = False
    
    For i = 1 To Len(text)
        If IsNumeric(Mid(text, i, 1)) Then
            hasDigit = True
        End If
    Next i
    
    ' 単位のリスト
    Dim units As Variant
    units = Array("錠", "カプセル", "包", "個", "枚", "本", "管", "筒", "組", "袋")
    
    hasSuffix = False
    For i = 0 To UBound(units)
        If InStr(text, units(i)) > 0 Then
            hasSuffix = True
            Exit For
        End If
    Next i
    
    IsNumericWithSuffix = hasDigit And hasSuffix
End Function

' 規格（強度）を抽出する関数（正規表現を使わないバージョン）
Public Function ExtractStrengthSimple(ByVal text As String) As String
    Dim i As Long, j As Long
    Dim numStart As Long
    Dim result As String
    Dim inNumber As Boolean
    Dim units As Variant
    
    units = Array("mg", "g", "ml", "μg")
    inNumber = False
    numStart = 0
    result = ""
    
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
                For j = 0 To UBound(units)
                    If LCase(Mid(text, i, Len(units(j)))) = LCase(units(j)) Then
                        result = Mid(text, numStart, i - numStart + Len(units(j)))
                        Exit For
                    End If
                Next j
                
                If result <> "" Then
                    Exit For
                End If
                
                inNumber = False
            End If
        End If
    Next i
    
    ExtractStrengthSimple = result
End Function

' 剤型を抽出する関数（正規表現を使わないバージョン）
Public Function ExtractFormTypeSimple(ByVal text As String) As String
    Dim forms As Variant
    Dim i As Long
    
    forms = Array("錠", "カプセル", "細粒", "顆粒", "散", "シロップ", "ドライシロップ", _
                  "注射液", "注射用", "軟膏", "クリーム", "ゲル", "テープ", "パップ", "点眼液")
    
    For i = 0 To UBound(forms)
        If InStr(text, forms(i)) > 0 Then
            ExtractFormTypeSimple = forms(i)
            Exit Function
        End If
    Next i
    
    ExtractFormTypeSimple = ""
End Function

' 包装形態を抽出する関数（正規表現を使わないバージョン）
Public Function ExtractPackageTypeSimple(ByVal text As String) As String
    Dim packages As Variant
    Dim i As Long
    
    packages = Array("(未定義)", "その他(なし)", "包装小", "調剤用", "PTP", "分包", "バラ", "SP", "PTP(患者用)")
    
    For i = 0 To UBound(packages)
        If InStr(1, text, packages(i), vbTextCompare) > 0 Then
            ' 見つかった包装形態をそのまま返す（NormalizePackageTypeは使わない）
            ExtractPackageTypeSimple = packages(i)
            Exit Function
        End If
    Next i
    
    ' スラッシュで囲まれた部分を検索
    Dim startPos As Long, endPos As Long
    startPos = InStr(1, text, "/")
    If startPos > 0 Then
        endPos = InStr(startPos + 1, text, "/")
        If endPos > startPos Then
            ' スラッシュ間の文字列をそのまま返す（NormalizePackageTypeは使わない）
            ExtractPackageTypeSimple = Mid(text, startPos + 1, endPos - startPos - 1)
            Exit Function
        End If
    End If
    
    ExtractPackageTypeSimple = ""
End Function

' パッケージタイプ取得（スラッシュ間の文字列）
Public Function GetPackageType(ByVal text As String) As String
    Dim startPos As Long, endPos As Long
    
    startPos = InStr(1, text, "/")
    If startPos > 0 Then
        endPos = InStr(startPos + 1, text, "/")
        If endPos > startPos Then
            GetPackageType = Mid(text, startPos + 1, endPos - startPos - 1)
        Else
            GetPackageType = ""
        End If
    Else
        GetPackageType = ""
    End If
End Function

' 包装規格を抽出する関数（括弧内の数値+単位）
Public Function ExtractPackageSizeSimple(ByVal text As String) As String
    ' 括弧内のテキストを抽出
    Dim bracketContent As String
    bracketContent = ExtractBracketContent(text)
    
    ' 抽出されたテキストがなければ
    If bracketContent = "" Then
        ' スペースで区切られた部分から探す
        Dim parts As Variant
        parts = Split(text, " ")
        Dim i As Long
        For i = 0 To UBound(parts)
            If IsPackageSize(parts(i)) Then
                ExtractPackageSizeSimple = parts(i)
                Exit Function
            End If
        Next i
        
        ExtractPackageSizeSimple = ""
        Exit Function
    End If
    
    ' 括弧内から数字+単位を抽出
    Dim j As Long, k As Long
    Dim numStart As Long
    Dim inNumber As Boolean
    Dim result As String
    Dim units As Variant
    
    units = Array("mg", "g", "ml", "mL", "μg", "包", "枚", "管", "錠")
    inNumber = False
    numStart = 0
    result = ""
    
    For j = 1 To Len(bracketContent)
        Dim c As String
        c = Mid(bracketContent, j, 1)
        
        If IsNumeric(c) Or c = "." Then
            If Not inNumber Then
                inNumber = True
                numStart = j
            End If
        ElseIf c = " " Or c = "/" And inNumber Then
            ' スペースとスラッシュは許容
        Else
            If inNumber Then
                ' 数字の後に単位があるか確認
                For k = 0 To UBound(units)
                    If j + Len(units(k)) - 1 <= Len(bracketContent) Then
                        If LCase(Mid(bracketContent, j, Len(units(k)))) = LCase(units(k)) Then
                            result = Mid(bracketContent, numStart, j - numStart + Len(units(k)))
                            ExtractPackageSizeSimple = result
                            Exit Function
                        End If
                    End If
                Next k
                
                inNumber = False
            End If
        End If
    Next j
    
    ExtractPackageSizeSimple = result
End Function

' 括弧内のテキストを抽出する関数
Private Function ExtractBracketContent(ByVal text As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(1, text, "(")
    If startPos = 0 Then startPos = InStr(1, text, "（")
    
    If startPos > 0 Then
        endPos = InStr(startPos + 1, text, ")")
        If endPos = 0 Then endPos = InStr(startPos + 1, text, "）")
        
        If endPos > startPos Then
            ExtractBracketContent = Mid(text, startPos + 1, endPos - startPos - 1)
        Else
            ExtractBracketContent = ""
        End If
    Else
        ExtractBracketContent = ""
    End If
End Function

' 数字+単位が包装規格かどうかを判断する関数
Private Function IsPackageSize(ByVal text As String) As Boolean
    Dim units As Variant
    units = Array("g", "mg", "ml", "mL", "μg", "包", "枚", "管", "錠")
    
    Dim hasDigit As Boolean
    Dim hasUnit As Boolean
    Dim i As Long
    
    ' 数字を含むか確認
    hasDigit = False
    For i = 1 To Len(text)
        If IsNumeric(Mid(text, i, 1)) Then
            hasDigit = True
            Exit For
        End If
    Next i
    
    If Not hasDigit Then
        IsPackageSize = False
        Exit Function
    End If
    
    ' 単位を含むか確認
    hasUnit = False
    For i = 0 To UBound(units)
        If InStr(1, text, units(i), vbTextCompare) > 0 Then
            hasUnit = True
            Exit For
        End If
    Next i
    
    IsPackageSize = hasDigit And hasUnit
End Function

' 薬品名の比較関数
Public Function CompareDrugStringsWithRate(ByVal sourceStr As String, ByVal targetStr As String) As Double
    Dim sourceParts As DrugNameParts
    Dim targetParts As DrugNameParts
    Dim matchCount As Integer
    Dim totalItems As Integer
    
    sourceParts = ParseDrugString(sourceStr)
    targetParts = ParseDrugString(targetStr)
    
    totalItems = 6 ' 基本名、剤型、規格、メーカー、包装、包装規格の6項目
    matchCount = 0
    
    ' 基本名の比較（完全一致）
    If StrComp(sourceParts.BaseName, targetParts.BaseName, vbTextCompare) = 0 Then
        matchCount = matchCount + 1
    End If
    
    ' 剤型の比較（完全一致）
    If StrComp(sourceParts.formType, targetParts.formType, vbTextCompare) = 0 Then
        matchCount = matchCount + 1
    End If
    
    ' 規格の比較（数値と単位を正規化して比較）
    If CompareStrength(sourceParts.strength, targetParts.strength) Then
        matchCount = matchCount + 1
    End If
    
    ' メーカー名の比較（完全一致）
    If StrComp(sourceParts.maker, targetParts.maker, vbTextCompare) = 0 Then
        matchCount = matchCount + 1
    End If
    
    ' 包装形態の比較（ある程度の揺れを許容）
    If StrComp(sourceParts.Package, targetParts.Package, vbTextCompare) = 0 Then
        matchCount = matchCount + 1
    End If
    
    ' 包装規格の比較
    If ComparePackageSize(sourceParts.PackageSize, targetParts.PackageSize) Then
        matchCount = matchCount + 1
    End If
    
    ' 一致率を計算（百分率）
    CompareDrugStringsWithRate = (matchCount / totalItems) * 100
End Function

' 包装規格を比較する関数
Public Function ComparePackageSize(ByVal size1 As String, ByVal size2 As String) As Boolean
    ' 空文字列の場合は不一致とする
    If size1 = "" Or size2 = "" Then
        ComparePackageSize = False
        Exit Function
    End If
    
    ' 数値と単位を分離して比較
    Dim num1 As Double, num2 As Double
    Dim unit1 As String, unit2 As String
    
    ' 数値と単位を抽出
    ExtractNumberAndUnit size1, num1, unit1
    ExtractNumberAndUnit size2, num2, unit2
    
    ' 数値と単位が両方一致する場合のみTrue
    ComparePackageSize = (num1 = num2) And (StrComp(unit1, unit2, vbTextCompare) = 0)
End Function



