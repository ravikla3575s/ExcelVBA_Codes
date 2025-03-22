Option Explicit

' 薬品名の構成要素を格納するための型定義
Public Type DrugNameParts
    BaseName As String     ' 基本名称
    FormType As String     ' 剤形
    Strength As String     ' 規格・含量
    Maker As String        ' メーカー
    Package As String      ' 包装形態
    PackageSize As String  ' 包装単位
End Type

' 文字列を解析して構成要素に分解する関数
Public Function ParseDrugString(ByVal drugString As String) As DrugNameParts
    Dim result As DrugNameParts
    
    ' 空文字列のチェック
    If Len(drugString) = 0 Then
        ParseDrugString = result
        Exit Function
    End If
    
    ' 括弧内容の抽出
    Dim bracketsContent As String
    bracketsContent = ExtractBracketsContent(drugString)
    
    ' 基本名称（括弧を除いた部分）
    result.BaseName = RemoveBracketsContent(drugString)
    
    ' 剤形の抽出
    result.FormType = ExtractFormType(result.BaseName)
    
    ' 規格・含量の抽出
    result.Strength = ExtractStrength(result.BaseName)
    
    ' メーカーの抽出（括弧内から）
    result.Maker = ExtractMaker(bracketsContent)
    
    ' 包装形態の抽出（括弧内から）
    result.Package = ExtractPackageTypeSimple(bracketsContent)
    
    ' 包装単位の抽出（括弧内から）
    result.PackageSize = ExtractPackageSize(bracketsContent)
    
    ParseDrugString = result
End Function

' 括弧内の内容を抽出する関数
Private Function ExtractBracketsContent(ByVal drugString As String) As String
    Dim result As String
    Dim startPos As Long, endPos As Long
    
    ' 最初の開き括弧を検索
    startPos = InStr(1, drugString, "(")
    If startPos = 0 Then
        ' 丸括弧がない場合は角括弧をチェック
        startPos = InStr(1, drugString, "[")
        If startPos = 0 Then
            ' どちらの括弧もない場合
            ExtractBracketsContent = ""
            Exit Function
        Else
            ' 対応する閉じ角括弧を検索
            endPos = InStr(startPos + 1, drugString, "]")
        End If
    Else
        ' 対応する閉じ丸括弧を検索
        endPos = InStr(startPos + 1, drugString, ")")
    End If
    
    ' 閉じ括弧が見つからない場合
    If endPos = 0 Then
        ExtractBracketsContent = ""
        Exit Function
    End If
    
    ' 括弧内の内容を抽出（括弧自体は含まない）
    result = Mid(drugString, startPos + 1, endPos - startPos - 1)
    
    ExtractBracketsContent = result
End Function

' 括弧内の内容を除去する関数（Mac環境向け実装）
Private Function RemoveBracketsContent(ByVal drugString As String) As String
    Dim result As String
    Dim pos1 As Long, pos2 As Long
    result = drugString
    
    ' すべての括弧とその内容を繰り返し削除
    Do
        ' 丸括弧を検索
        pos1 = InStr(1, result, "(")
        If pos1 > 0 Then
            pos2 = InStr(pos1, result, ")")
            If pos2 > 0 Then
                result = Left(result, pos1 - 1) & Mid(result, pos2 + 1)
                ' 処理を継続
                GoTo ContinueLoop
            End If
        End If
        
        ' 角括弧を検索
        pos1 = InStr(1, result, "[")
        If pos1 > 0 Then
            pos2 = InStr(pos1, result, "]")
            If pos2 > 0 Then
                result = Left(result, pos1 - 1) & Mid(result, pos2 + 1)
                ' 処理を継続
                GoTo ContinueLoop
            End If
        End If
        
        ' 括弧が見つからなければループを抜ける
        Exit Do
        
ContinueLoop:
    Loop
    
    RemoveBracketsContent = Trim(result)
End Function

' 剤形を抽出する関数
Private Function ExtractFormType(ByVal drugName As String) As String
    Dim formTypes As Variant
    formTypes = Array("錠", "カプセル", "細粒", "顆粒", "散", "シロップ", "液", "注射液", "注", "軟膏", "クリーム", "ゲル", "ローション", "点眼液", "目薬", "点鼻液", "吸入液", "貼付剤", "パッチ", "坐剤")
    
    Dim i As Long
    For i = LBound(formTypes) To UBound(formTypes)
        If InStr(drugName, formTypes(i)) > 0 Then
            ExtractFormType = formTypes(i)
            Exit Function
        End If
    Next i
    
    ExtractFormType = ""
End Function

' 規格・含量を抽出する関数（Mac環境向け実装）
Private Function ExtractStrength(ByVal drugName As String) As String
    Dim result As String
    result = ""
    
    ' 数字+単位のパターンを検索（正規表現の代わりに文字列操作で実装）
    Dim i As Long, j As Long
    Dim units As Variant
    units = Array("mg", "g", "mL", "μg", "単位", "IU", "%")
    
    ' 数字の開始位置を探す
    For i = 1 To Len(drugName)
        If IsNumeric(Mid(drugName, i, 1)) Then
            ' 数字が見つかったら、その後の単位を確認
            Dim numEnd As Long
            numEnd = i
            
            ' 数字部分の終了位置を特定
            Do While numEnd <= Len(drugName) And (IsNumeric(Mid(drugName, numEnd, 1)) Or Mid(drugName, numEnd, 1) = ".")
                numEnd = numEnd + 1
            Loop
            
            ' 単位を探す
            For j = LBound(units) To UBound(units)
                If InStr(numEnd, drugName, units(j)) = numEnd Then
                    ' 数値+単位を取得
                    result = Mid(drugName, i, numEnd - i + Len(units(j)))
                    ExtractStrength = Trim(result)
                    Exit Function
                End If
            Next j
        End If
    Next i
    
    ExtractStrength = ""
End Function

' メーカー名を抽出する関数
Private Function ExtractMaker(ByVal bracketsContent As String) As String
    Dim makers As Variant
    makers = Array("武田", "第一三共", "アステラス", "エーザイ", "田辺三菱", "大塚", "アストラゼネカ", "ノバルティス", "ファイザー", "MSD", "バイエル", "大正", "中外", "参天", "久光", "杏林", "沢井", "東和", "日医工", "あすか", "ニプロ", "サンド", "陽進堂", "科研", "キョーリン", "ツムラ", "日本ケミファ", "トーアエイヨー", "共和", "明治", "救急", "持田", "ゼリア", "小野", "協和", "Meiji Seika", "テバ", "富士", "マイラン", "ヤンセン", "ギリアド", "シオノギ", "塩野義", "アッヴィ", "ブリストル", "テルモ", "帝人", "キッセイ", "ロシュ", "グラクソ", "サノフィ", "大日本住友", "興和", "鳥居")
    
    Dim i As Long
    For i = LBound(makers) To UBound(makers)
        If InStr(bracketsContent, makers(i)) > 0 Then
            ExtractMaker = makers(i)
            Exit Function
        End If
    Next i
    
    ExtractMaker = ""
End Function

' シンプルな包装形態を抽出する関数
Private Function ExtractPackageTypeSimple(ByVal bracketsContent As String) As String
    Dim packageTypes As Variant
    packageTypes = Array("PTP", "分包", "バラ", "SP", "調剤用", "包装小", "PTP(患者用)")
    
    Dim i As Long
    For i = LBound(packageTypes) To UBound(packageTypes)
        If InStr(bracketsContent, packageTypes(i)) > 0 Then
            ExtractPackageTypeSimple = packageTypes(i)
            Exit Function
        End If
    Next i
    
    ExtractPackageTypeSimple = ""
End Function

' 包装単位を抽出する関数（Mac環境向け実装）
Private Function ExtractPackageSize(ByVal bracketsContent As String) As String
    Dim result As String
    result = ""
    
    ' 数字+単位のパターンを検索（正規表現の代わりに文字列操作で実装）
    Dim i As Long, j As Long
    Dim units As Variant
    units = Array("錠", "カプセル", "個", "枚", "包", "本", "管", "瓶", "袋")
    
    ' 数字の開始位置を探す
    For i = 1 To Len(bracketsContent)
        If IsNumeric(Mid(bracketsContent, i, 1)) Then
            ' 数字が見つかったら、その後の単位を確認
            Dim numEnd As Long
            numEnd = i
            
            ' 数字部分の終了位置を特定
            Do While numEnd <= Len(bracketsContent) And IsNumeric(Mid(bracketsContent, numEnd, 1))
                numEnd = numEnd + 1
            Loop
            
            ' 数字の後の空白をスキップ
            Do While numEnd <= Len(bracketsContent) And Mid(bracketsContent, numEnd, 1) = " "
                numEnd = numEnd + 1
            Loop
            
            ' 単位を探す
            For j = LBound(units) To UBound(units)
                If InStr(numEnd, bracketsContent, units(j)) = numEnd Then
                    ' 数値+単位を取得
                    result = Mid(bracketsContent, i, numEnd - i + Len(units(j)))
                    ExtractPackageSize = Trim(result)
                    Exit Function
                End If
            Next j
        End If
    Next i
    
    ExtractPackageSize = ""
End Function

' 薬品名から特定のキーワードを抽出する関数（Mac環境向け実装）
Public Function ExtractKeywords(ByVal drugName As String) As Variant
    ' キーワードを格納する配列
    Dim keywords() As String
    ReDim keywords(0 To 9)  ' 最大10個のキーワードを想定
    Dim keywordCount As Long
    keywordCount = 0
    
    ' 括弧内の内容を抽出
    Dim bracketsContent As String
    bracketsContent = ExtractBracketsContent(drugName)
    
    ' メーカー名を抽出
    Dim maker As String
    maker = ExtractMaker(bracketsContent)
    If Len(maker) > 0 Then
        keywords(keywordCount) = maker
        keywordCount = keywordCount + 1
    End If
    
    ' 包装形態を抽出
    Dim packageType As String
    packageType = ExtractPackageTypeSimple(bracketsContent)
    If Len(packageType) > 0 Then
        keywords(keywordCount) = packageType
        keywordCount = keywordCount + 1
    End If
    
    ' 薬品名から括弧内容を除去
    Dim nameWithoutBrackets As String
    nameWithoutBrackets = RemoveBracketsContent(drugName)
    
    ' 剤形を抽出
    Dim formType As String
    formType = ExtractFormType(nameWithoutBrackets)
    If Len(formType) > 0 Then
        keywords(keywordCount) = formType
        keywordCount = keywordCount + 1
    End If
    
    ' 規格・含量を抽出
    Dim strength As String
    strength = ExtractStrength(nameWithoutBrackets)
    If Len(strength) > 0 Then
        keywords(keywordCount) = strength
        keywordCount = keywordCount + 1
    End If
    
    ' 結果を適切なサイズの配列に調整
    If keywordCount > 0 Then
        ReDim Preserve keywords(0 To keywordCount - 1)
    Else
        ReDim keywords(0 To 0)
    End If
    
    ExtractKeywords = keywords
End Function

' 薬品文字列を比較して類似度を計算する関数
Public Function CompareDrugStringsWithRate(ByVal string1 As String, ByVal string2 As String) As Double
    Dim parts1 As DrugNameParts
    Dim parts2 As DrugNameParts
    
    parts1 = ParseDrugString(string1)
    parts2 = ParseDrugString(string2)
    
    ' 各コンポーネントの重み付け
    Const BASENAME_WEIGHT As Double = 0.5
    Const FORMTYPE_WEIGHT As Double = 0.15
    Const STRENGTH_WEIGHT As Double = 0.15
    Const MAKER_WEIGHT As Double = 0.1
    Const PACKAGE_WEIGHT As Double = 0.1
    
    Dim totalScore As Double
    totalScore = 0
    
    ' 基本名称の比較
    If Len(parts1.BaseName) > 0 And Len(parts2.BaseName) > 0 Then
        totalScore = totalScore + (GetSimilarity(parts1.BaseName, parts2.BaseName) * BASENAME_WEIGHT)
    End If
    
    ' 剤形の比較
    If Len(parts1.FormType) > 0 And Len(parts2.FormType) > 0 Then
        If parts1.FormType = parts2.FormType Then
            totalScore = totalScore + FORMTYPE_WEIGHT
        End If
    End If
    
    ' 規格・含量の比較
    If Len(parts1.Strength) > 0 And Len(parts2.Strength) > 0 Then
        If parts1.Strength = parts2.Strength Then
            totalScore = totalScore + STRENGTH_WEIGHT
        End If
    End If
    
    ' メーカーの比較
    If Len(parts1.Maker) > 0 And Len(parts2.Maker) > 0 Then
        If parts1.Maker = parts2.Maker Then
            totalScore = totalScore + MAKER_WEIGHT
        End If
    End If
    
    ' 包装形態の比較
    If Len(parts1.Package) > 0 And Len(parts2.Package) > 0 Then
        If parts1.Package = parts2.Package Then
            totalScore = totalScore + PACKAGE_WEIGHT
        End If
    End If
    
    CompareDrugStringsWithRate = totalScore
End Function

' 2つの文字列の類似度を計算する関数
Public Function GetSimilarity(ByVal string1 As String, ByVal string2 As String) As Double
    ' 両方の文字列が空の場合は完全一致とみなす
    If Len(string1) = 0 And Len(string2) = 0 Then
        GetSimilarity = 1
        Exit Function
    End If
    
    ' どちらかの文字列が空の場合は類似度0
    If Len(string1) = 0 Or Len(string2) = 0 Then
        GetSimilarity = 0
        Exit Function
    End If
    
    ' 完全一致の場合
    If string1 = string2 Then
        GetSimilarity = 1
        Exit Function
    End If
    
    ' レーベンシュタイン距離を計算
    Dim distance As Long
    distance = LevenshteinDistance(string1, string2)
    
    ' 長い方の文字列の長さを基準に正規化
    Dim maxLength As Long
    maxLength = Application.WorksheetFunction.Max(Len(string1), Len(string2))
    
    ' 類似度を計算 (0に近いほど違い、1に近いほど似ている)
    GetSimilarity = 1 - (distance / maxLength)
End Function

' レーベンシュタイン距離を計算する関数（2つの文字列間の編集距離）
Private Function LevenshteinDistance(ByVal s1 As String, ByVal s2 As String) As Long
    Dim i As Long, j As Long
    Dim cost As Long
    
    ' 文字列の長さを取得
    Dim len1 As Long, len2 As Long
    len1 = Len(s1)
    len2 = Len(s2)
    
    ' 距離行列を初期化
    Dim d() As Long
    ReDim d(0 To len1, 0 To len2)
    
    ' ベースケースを初期化
    For i = 0 To len1
        d(i, 0) = i
    Next i
    
    For j = 0 To len2
        d(0, j) = j
    Next j
    
    ' 距離行列を埋める
    For i = 1 To len1
        For j = 1 To len2
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            
            d(i, j) = Application.WorksheetFunction.Min( _
                d(i - 1, j) + 1, _      ' 削除
                d(i, j - 1) + 1, _      ' 挿入
                d(i - 1, j - 1) + cost) ' 置換
        Next j
    Next i
    
    ' 結果を返す
    LevenshteinDistance = d(len1, len2)
End Function