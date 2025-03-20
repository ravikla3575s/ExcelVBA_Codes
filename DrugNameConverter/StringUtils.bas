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

' 包装形態の比較関数
Public Function ComparePackageType(ByVal type1 As String, ByVal type2 As String) As Boolean
    ' 包装形態の表記揺れに対応
    Dim normalizedType1 As String, normalizedType2 As String
    
    normalizedType1 = NormalizePackageType(type1)
    normalizedType2 = NormalizePackageType(type2)
    
    ComparePackageType = (StrComp(normalizedType1, normalizedType2, vbTextCompare) = 0)
End Function

' 包装形態の表記を標準化する関数
Public Function NormalizePackageType(ByVal packageType As String) As String
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