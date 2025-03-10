' 日付変換・計算用の専用モジュール

' 請求年月から調剤年月を計算（文字列→数値→計算→文字列）
Private Function ConvertBillingToDispensing(billing_year As String, billing_month As String, _
    ByRef dispensing_year As String, ByRef dispensing_month As String) As Boolean
    
    Dim western_year As Integer, western_month As Integer
    
    On Error GoTo ErrorHandler
    
    ' 文字列から数値への変換
    western_year = CInt(billing_year)
    western_month = CInt(billing_month)
    
    ' 調剤月の計算（請求月の前月）
    western_month = western_month - 1
    If western_month < 1 Then
        western_month = 12
        western_year = western_year - 1
    End If
    
    ' 結果を文字列として設定
    dispensing_year = CStr(western_year)
    dispensing_month = Format(western_month, "00")
    
    ConvertBillingToDispensing = True
    Exit Function
    
ErrorHandler:
    ConvertBillingToDispensing = False
End Function

' 調剤年月から請求年月を計算（文字列→数値→計算→文字列）
Private Function ConvertDispensingToBilling(dispensing_year As String, dispensing_month As String, _
    ByRef billing_year As String, ByRef billing_month As String) As Boolean
    
    Dim western_year As Integer, western_month As Integer
    
    On Error GoTo ErrorHandler
    
    ' 文字列から数値への変換
    western_year = CInt(dispensing_year)
    western_month = CInt(dispensing_month)
    
    ' 請求月の計算（調剤月の翌月）
    western_month = western_month + 1
    If western_month > 12 Then
        western_month = 1
        western_year = western_year + 1
    End If
    
    ' 結果を文字列として設定
    billing_year = CStr(western_year)
    billing_month = Format(western_month, "00")
    
    ConvertDispensingToBilling = True
    Exit Function
    
ErrorHandler:
    ConvertDispensingToBilling = False
End Function

' 西暦年月を和暦年月に変換
Private Function ConvertWesternToEra(western_year As String, western_month As String, _
    ByRef era_name As String, ByRef era_year As String, ByRef era_month As String) As Boolean
    
    Dim year_num As Integer
    
    On Error GoTo ErrorHandler
    
    year_num = CInt(western_year)
    
    ' 元号の判定
    If year_num >= 2019 Then
        era_name = "令和"
        era_year = Format(year_num - 2018, "00")
    ElseIf year_num >= 1989 Then
        era_name = "平成"
        era_year = Format(year_num - 1988, "00")
    ElseIf year_num >= 1926 Then
        era_name = "昭和"
        era_year = Format(year_num - 1925, "00")
    ElseIf year_num >= 1912 Then
        era_name = "大正"
        era_year = Format(year_num - 1911, "00")
    ElseIf year_num >= 1868 Then
        era_name = "明治"
        era_year = Format(year_num - 1867, "00")
    Else
        ConvertWesternToEra = False
        Exit Function
    End If
    
    era_month = western_month
    
    ConvertWesternToEra = True
    Exit Function
    
ErrorHandler:
    ConvertWesternToEra = False
End Function

' 和暦年月を西暦年月に変換
Private Function ConvertEraToWestern(era_code As String, era_year As String, era_month As String, _
    ByRef western_year As String, ByRef western_month As String) As Boolean
    
    Dim year_num As Integer
    
    On Error GoTo ErrorHandler
    
    ' 元号コードから西暦年を計算
    Select Case era_code
        Case "5"  ' 令和
            year_num = 2018 + CInt(era_year)
        Case "4"  ' 平成
            year_num = 1988 + CInt(era_year)
        Case "3"  ' 昭和
            year_num = 1925 + CInt(era_year)
        Case "2"  ' 大正
            year_num = 1911 + CInt(era_year)
        Case "1"  ' 明治
            year_num = 1867 + CInt(era_year)
        Case Else
            ConvertEraToWestern = False
            Exit Function
    End Select
    
    western_year = CStr(year_num)
    western_month = era_month
    
    ConvertEraToWestern = True
    Exit Function
    
ErrorHandler:
    ConvertEraToWestern = False
End Function

' 使用例：
Private Sub ExampleUsage()
    Dim billing_year As String, billing_month As String
    Dim dispensing_year As String, dispensing_month As String
    Dim era_name As String, era_year As String, era_month As String
    
    ' 請求年月から調剤年月を計算
    billing_year = "2024"
    billing_month = "03"
    If ConvertBillingToDispensing(billing_year, billing_month, dispensing_year, dispensing_month) Then
        Debug.Print "調剤年月: " & dispensing_year & "年" & dispensing_month & "月"
    End If
    
    ' 調剤年月を和暦に変換
    If ConvertWesternToEra(dispensing_year, dispensing_month, era_name, era_year, era_month) Then
        Debug.Print "和暦: " & era_name & era_year & "年" & era_month & "月"
    End If
End Sub

' ファイルコレクションをソートする関数を修正
Function SortFileCollection(files As Collection, file_system As Object, file_type As String) As Collection
    Dim sorted_array() As String
    Dim sorted_files As New Collection
    Dim file_obj As Object
    Dim i As Long, count As Long
    
    ' コレクションの要素数を取得
    count = files.Count
    If count = 0 Then
        Set SortFileCollection = files
        Exit Function
    End If
    
    ' 配列の初期化
    ReDim sorted_array(1 To count, 1 To 2)  ' 1列目:ファイルパス、2列目:ソート用キー
    
    ' コレクションの内容を配列にコピー
    i = 1
    For Each file_obj In files
        Dim year_str As String, month_str As String
        sorted_array(i, 1) = file_obj.Path
        
        ' ソートキーを生成（YYYYMM形式）
        If GetYearMonthFromFile(file_obj.Path, file_type, year_str, month_str) Then
            sorted_array(i, 2) = year_str & Format(CInt(month_str), "00")
        Else
            sorted_array(i, 2) = "999999"  ' エラー時は最後尾に
        End If
        i = i + 1
    Next file_obj
    
    ' クイックソートで配列を並び替え
    Call QuickSort(sorted_array, 1, count)
    
    ' ソート済みの配列から新しいコレクションを作成
    Set sorted_files = New Collection
    For i = 1 To count
        sorted_files.Add file_system.GetFile(sorted_array(i, 1))
    Next i
    
    Set SortFileCollection = sorted_files
End Function

' クイックソートの実装
Private Sub QuickSort(ByRef arr() As String, ByVal left As Long, ByVal right As Long)
    Dim i As Long, j As Long
    Dim pivot As String
    Dim temp_path As String, temp_key As String
    
    If left >= right Then Exit Sub
    
    i = left
    j = right
    pivot = arr((left + right) \ 2, 2)
    
    Do While i <= j
        Do While arr(i, 2) < pivot And i < right
            i = i + 1
        Loop
        
        Do While arr(j, 2) > pivot And j > left
            j = j - 1
        Loop
        
        If i <= j Then
            ' パスの交換
            temp_path = arr(i, 1)
            arr(i, 1) = arr(j, 1)
            arr(j, 1) = temp_path
            
            ' キーの交換
            temp_key = arr(i, 2)
            arr(i, 2) = arr(j, 2)
            arr(j, 2) = temp_key
            
            i = i + 1
            j = j - 1
        End If
    Loop
    
    If left < j Then Call QuickSort(arr, left, j)
    If i < right Then Call QuickSort(arr, i, right)
End Sub

' 年月取得関数（既存の関数を使用）
Private Function GetYearMonthFromFile(file_path As String, file_type As String, _
    ByRef year_str As String, ByRef month_str As String) As Boolean
    
    ' ... 既存のコード ...
End Function 