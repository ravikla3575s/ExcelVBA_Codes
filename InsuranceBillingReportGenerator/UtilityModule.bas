Attribute VB_Name = "UtilityModule"
Option Explicit

' 長時間処理の進捗表示
Private Sub UpdateProgress(current As Long, total As Long, message As String)
    Application.StatusBar = message & " - " & current & "/" & total
End Sub

' オブジェクト解放用の関数
Private Sub CleanupObjects(ParamArray objects() As Variant)
    Dim obj As Variant
    For Each obj In objects
        If Not obj Is Nothing Then
            If TypeName(obj) = "Workbook" Then
                obj.Close SaveChanges:=False
            End If
            Set obj = Nothing
        End If
    Next obj
End Sub

' シートの存在チェック用の関数
Private Function SheetExists(wb As Workbook, sheet_name As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = wb.Sheets(sheet_name)
    On Error GoTo 0
    
    SheetExists = Not ws Is Nothing
End Function

' 数字を丸付き数字に変換する関数
Function ConvertToCircledNumber(ByVal month As Integer) As String
    Dim circled_numbers As Variant
    circled_numbers = Array("", "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫")  ' インデックス0に空文字を追加
    If month >= 1 And month <= 12 Then
        ConvertToCircledNumber = circled_numbers(month)  ' そのままのmonthを使用
    Else
        ConvertToCircledNumber = CStr(month)
    End If
End Function

' カテゴリの開始行を取得する関数
Public Function GetStartRow(ws As Worksheet, category_name As String) As Long
    Dim found_cell As Range
    Set found_cell = ws.Cells.Find(what:=category_name, LookAt:=xlWhole)
    If Not found_cell Is Nothing Then
        GetStartRow = found_cell.row
    Else
        GetStartRow = 0
    End If
End Function

' マーキングされた開始行を検索する関数
Public Function FindMarkedRow(ws As Worksheet, marker As String) As Long
    Dim found_cell As Range
    Dim search_marker As String
    
    ' マーキングのフォーマットを確認（既に<<>>が含まれているかどうか）
    If Left(marker, 2) <> "<<" Then
        search_marker = "<<" & marker & ">>"
    Else
        search_marker = marker
    End If
    
    ' シート全体を検索
    Set found_cell = ws.Cells.Find(what:=search_marker, LookAt:=xlPart, MatchCase:=False)
    
    If Not found_cell Is Nothing Then
        FindMarkedRow = found_cell.Row
        Debug.Print "マーカー '" & search_marker & "' を行 " & FindMarkedRow & " で発見しました"
    Else
        FindMarkedRow = 0
        Debug.Print "マーカー '" & search_marker & "' は見つかりませんでした"
    End If
End Function

' カテゴリの開始行を取得する関数 - マーキングベースの新バージョン
Public Function GetCategoryStartRowsFromMarkers(ws As Worksheet, payer_type As String) As Object
    Dim start_row_dict As Object
    Set start_row_dict = CreateObject("Scripting.Dictionary")
    
    Debug.Print "マーキングから開始行を検索しています: " & payer_type
    
    ' 社保関連のマーカーを検索
    Dim shaho_saiseikyu As Long
    Dim shaho_tsukiokure As Long
    Dim shaho_tsukinokuri As Long
    Dim shaho_henrei As Long
    Dim shaho_miseikyuu As Long
    
    ' 国保関連のマーカーを検索
    Dim kokuho_saiseikyu As Long
    Dim kokuho_tsukiokure As Long
    Dim kokuho_tsukinokuri As Long
    Dim kokuho_henrei As Long
    Dim kokuho_miseikyuu As Long
    
    ' 介護関連のマーカーを検索
    Dim kaigo_henrei As Long
    
    ' その他のマーカーを検索
    Dim sonota As Long
    
    ' 各マーカーの行を検索
    shaho_saiseikyu = FindMarkedRow(ws, "社保再請求")
    shaho_tsukiokure = FindMarkedRow(ws, "社保月遅れ")
    shaho_tsukinokuri = FindMarkedRow(ws, "社保月送り")
    shaho_henrei = FindMarkedRow(ws, "社保返戻")
    shaho_miseikyuu = FindMarkedRow(ws, "社保未請求扱い")
    
    kokuho_saiseikyu = FindMarkedRow(ws, "国保再請求")
    kokuho_tsukiokure = FindMarkedRow(ws, "国保月遅れ")
    kokuho_tsukinokuri = FindMarkedRow(ws, "国保月送り")
    kokuho_henrei = FindMarkedRow(ws, "国保返戻")
    kokuho_miseikyuu = FindMarkedRow(ws, "国保未請求扱い")
    
    kaigo_henrei = FindMarkedRow(ws, "介護返戻")
    
    sonota = FindMarkedRow(ws, "その他")
    
    ' 請求先タイプに基づいてディクショナリを構築
    If payer_type = "社保" Then
        If shaho_saiseikyu > 0 Then
            start_row_dict.Add "再請求", shaho_saiseikyu
        End If
        
        If shaho_tsukiokure > 0 Then
            start_row_dict.Add "月遅れ請求", shaho_tsukiokure
        End If
        
        If shaho_tsukinokuri > 0 Then
            start_row_dict.Add "月送り", shaho_tsukinokuri
        End If
        
        If shaho_henrei > 0 Then
            start_row_dict.Add "返戻・査定", shaho_henrei
        End If
        
        If shaho_miseikyuu > 0 Then
            start_row_dict.Add "未請求扱い", shaho_miseikyuu
        End If
    ElseIf payer_type = "国保" Then
        If kokuho_saiseikyu > 0 Then
            start_row_dict.Add "再請求", kokuho_saiseikyu
        End If
        
        If kokuho_tsukiokure > 0 Then
            start_row_dict.Add "月遅れ請求", kokuho_tsukiokure
        End If
        
        If kokuho_tsukinokuri > 0 Then
            start_row_dict.Add "月送り", kokuho_tsukinokuri
        End If
        
        If kokuho_henrei > 0 Then
            start_row_dict.Add "返戻・査定", kokuho_henrei
        End If
        
        If kokuho_miseikyuu > 0 Then
            start_row_dict.Add "未請求扱い", kokuho_miseikyuu
        End If
    ElseIf payer_type = "介護" Then
        If kaigo_henrei > 0 Then
            start_row_dict.Add "返戻", kaigo_henrei
        End If
    End If
    
    ' その他は共通
    If sonota > 0 Then
        start_row_dict.Add "その他", sonota
    End If
    
    ' ディクショナリが空の場合（マーカーが見つからない場合）
    If start_row_dict.Count = 0 Then
        Debug.Print "WARNING: マーキングが見つかりませんでした。従来の方法でカテゴリ開始行を取得します。"
        Set start_row_dict = GetCategoryStartRows(ws, payer_type)
    End If
    
    Set GetCategoryStartRowsFromMarkers = start_row_dict
End Function

' 従来のカテゴリの開始行を取得する関数（バックアップとして残す）
Public Function GetCategoryStartRows(ws As Worksheet, payer_type As String) As Object
    Dim start_row_dict As Object
    Set start_row_dict = CreateObject("Scripting.Dictionary")
    
    Debug.Print "従来の方法で開始行を検索しています: " & payer_type
    
    If payer_type = "社保" Then
        Dim social_start_row As Long
        social_start_row = GetStartRow(ws, "社保返戻再請求")
        
        If social_start_row > 0 Then
            start_row_dict.Add "再請求", social_start_row
            start_row_dict.Add "月遅れ請求", GetStartRow(ws, "社保月遅れ請求")
            start_row_dict.Add "返戻・査定", GetStartRow(ws, "社保返戻・査定")
            start_row_dict.Add "未請求扱い", GetStartRow(ws, "社保未請求扱い")
        Else
            ' 見出しが見つからない場合のデフォルト値を設定
            Debug.Print "社保の見出しが見つかりません。デフォルト値を使用します。"
            start_row_dict.Add "再請求", 3  ' デフォルト開始行
            start_row_dict.Add "月遅れ請求", 8
            start_row_dict.Add "返戻・査定", 13
            start_row_dict.Add "未請求扱い", 18
        End If
    ElseIf payer_type = "国保" Then
        Dim kokuho_start_row As Long
        kokuho_start_row = GetStartRow(ws, "国保返戻再請求")
        
        If kokuho_start_row > 0 Then
            start_row_dict.Add "再請求", kokuho_start_row
            start_row_dict.Add "月遅れ請求", GetStartRow(ws, "国保月遅れ請求")
            start_row_dict.Add "返戻・査定", GetStartRow(ws, "国保返戻・査定")
            start_row_dict.Add "未請求扱い", GetStartRow(ws, "国保未請求扱い")
        Else
            ' 見出しが見つからない場合のデフォルト値を設定
            Debug.Print "国保の見出しが見つかりません。デフォルト値を使用します。"
            start_row_dict.Add "再請求", 23  ' デフォルト開始行
            start_row_dict.Add "月遅れ請求", 28
            start_row_dict.Add "返戻・査定", 33
            start_row_dict.Add "未請求扱い", 38
        End If
    End If
    
    ' ディクショナリが空の場合（想定外の請求先タイプなど）
    If start_row_dict.Count = 0 Then
        Debug.Print "WARNING: カテゴリの開始行が設定できませんでした。請求先: " & payer_type
        start_row_dict.Add "再請求", 3
        start_row_dict.Add "月遅れ請求", 8
        start_row_dict.Add "返戻・査定", 13
        start_row_dict.Add "未請求扱い", 18
    End If
    
    Set GetCategoryStartRows = start_row_dict
End Function

' 追加行の挿入処理
Public Sub InsertAdditionalRows(ws As Worksheet, start_row_dict As Object, _
    rebill_count As Long, late_count As Long, assessment_count As Long)
    
    Dim a As Long, b As Long, c As Long
    
    If rebill_count > BASE_DETAIL_ROWS Then a = rebill_count - BASE_DETAIL_ROWS
    If late_count > BASE_DETAIL_ROWS Then b = late_count - BASE_DETAIL_ROWS
    If assessment_count > BASE_DETAIL_ROWS Then c = assessment_count - BASE_DETAIL_ROWS
    
    If a > 0 Then ws.rows(start_row_dict("再請求") + 1 & ":" & start_row_dict("再請求") + a).Insert Shift:=xlDown
    If b > 0 Then ws.rows(start_row_dict("月遅れ請求") + 1 & ":" & start_row_dict("月遅れ請求") + b).Insert Shift:=xlDown
    If c > 0 Then ws.rows(start_row_dict("返戻・査定") + 1 & ":" & start_row_dict("返戻・査定") + c).Insert Shift:=xlDown
End Sub

' データを詳細シートに転記する関数
Function TransferData(dataDict As Object, ws As Worksheet, start_row As Long, payer_type As String) As Boolean
    If dataDict.count = 0 Then
        TransferData = False
        Exit Function
    End If

    Dim key As Variant, row_data As Variant
    Dim r As Long: r = start_row
    Dim payer_col As Long

    ' 社保はH列(8), 国保はI列(9)に種別を記載
    If payer_type = "社保" Then
        payer_col = 8
    ElseIf payer_type = "国保" Then
        payer_col = 9
    Else
        TransferData = False  ' その他（労災等）は対象外
        Exit Function
    End If

    ' 各レコードを書き込み
    For Each key In dataDict.Keys
        row_data = dataDict(key)
        ws.Cells(r, 4).value = row_data(0)          ' 患者氏名
        ws.Cells(r, 5).value = row_data(1)          ' 調剤年月 (YY.MM形式)
        ws.Cells(r, 6).value = row_data(2)          ' 医療機関名
        ws.Cells(r, payer_col).value = payer_type   ' 請求先種別 (社保/国保)
        ws.Cells(r, payer_col).Font.Bold = True     ' 強調表示
        ws.Cells(r, 10).value = row_data(3)         ' 請求点数
        r = r + 1
    Next key
    
    TransferData = True
End Function

' 全角数字を半角数字に変換する関数
Public Function ConvertToHankaku(ByVal strText As String) As String
    Dim i As Long
    Dim result As String
    Dim c As String
    
    result = ""
    For i = 1 To Len(strText)
        c = Mid(strText, i, 1)
        Select Case c
            Case "０": result = result & "0"
            Case "１": result = result & "1"
            Case "２": result = result & "2"
            Case "３": result = result & "3"
            Case "４": result = result & "4"
            Case "５": result = result & "5"
            Case "６": result = result & "6"
            Case "７": result = result & "7"
            Case "８": result = result & "8"
            Case "９": result = result & "9"
            Case Else: result = result & c
        End Select
    Next i
    
    ConvertToHankaku = result
End Function

