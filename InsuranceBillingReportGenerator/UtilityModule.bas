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

' カテゴリの開始行を取得する関数
Public Function GetCategoryStartRows(ws As Worksheet, payer_type As String) As Object
    Dim start_row_dict As Object
    Set start_row_dict = CreateObject("Scripting.Dictionary")
    
    Debug.Print "Getting category start rows for: " & payer_type
    
    If payer_type = "社保" Then
        Dim social_start_row As Long
        social_start_row = GetStartRow(ws, "社保返戻再請求")
        
        If social_start_row > 0 Then
            start_row_dict.Add "返戻再請求", social_start_row
            start_row_dict.Add "月遅れ請求", GetStartRow(ws, "社保月遅れ請求")
            start_row_dict.Add "返戻・査定", GetStartRow(ws, "社保返戻・査定")
            start_row_dict.Add "未請求扱い", GetStartRow(ws, "社保未請求扱い")
        Else
            ' 見出しが見つからない場合のデフォルト値を設定
            Debug.Print "社保の見出しが見つかりません。デフォルト値を使用します。"
            start_row_dict.Add "返戻再請求", 3  ' デフォルト開始行
            start_row_dict.Add "月遅れ請求", 8
            start_row_dict.Add "返戻・査定", 13
            start_row_dict.Add "未請求扱い", 18
        End If
    ElseIf payer_type = "国保" Then
        Dim kokuho_start_row As Long
        kokuho_start_row = GetStartRow(ws, "国保返戻再請求")
        
        If kokuho_start_row > 0 Then
            start_row_dict.Add "返戻再請求", kokuho_start_row
            start_row_dict.Add "月遅れ請求", GetStartRow(ws, "国保月遅れ請求")
            start_row_dict.Add "返戻・査定", GetStartRow(ws, "国保返戻・査定")
            start_row_dict.Add "未請求扱い", GetStartRow(ws, "国保未請求扱い")
        Else
            ' 見出しが見つからない場合のデフォルト値を設定
            Debug.Print "国保の見出しが見つかりません。デフォルト値を使用します。"
            start_row_dict.Add "返戻再請求", 23  ' デフォルト開始行
            start_row_dict.Add "月遅れ請求", 28
            start_row_dict.Add "返戻・査定", 33
            start_row_dict.Add "未請求扱い", 38
        End If
    End If
    
    ' ディクショナリが空の場合（想定外の請求先タイプなど）
    If start_row_dict.Count = 0 Then
        Debug.Print "WARNING: カテゴリの開始行が設定できませんでした。請求先: " & payer_type
        start_row_dict.Add "返戻再請求", 3
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
    
    If a > 0 Then ws.rows(start_row_dict("返戻再請求") + 1 & ":" & start_row_dict("返戻再請求") + a).Insert Shift:=xlDown
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

' マーキングからカテゴリの開始行を取得する関数
Public Function FindMarkedRow(ws As Worksheet, marker As String) As Long
    Dim found_cell As Range
    Dim search_text As String
    search_text = "<<" & marker & ">>"
    
    ' D列（4列目）に限定して検索
    Set found_cell = ws.Columns(4).Find(What:=search_text, LookIn:=xlValues, LookAt:=xlPart)
    
    If Not found_cell Is Nothing Then
        Debug.Print "Found marker '" & marker & "' at row " & found_cell.Row
        FindMarkedRow = found_cell.Row
    Else
        Debug.Print "WARNING: Marker '" & marker & "' not found"
        FindMarkedRow = 0
    End If
End Function

' マーキングを基にカテゴリ開始行の辞書を作成する関数
Public Function GetMarkedCategoryRows(ws As Worksheet) As Object
    Dim category_dict As Object
    Set category_dict = CreateObject("Scripting.Dictionary")
    
    ' すべてのカテゴリとそれに対応するマーカーを定義
    Dim categories As Object
    Set categories = CreateObject("Scripting.Dictionary")
    
    categories.Add "社保再請求", "社保再請求"
    categories.Add "国保再請求", "国保再請求"
    categories.Add "社保月遅れ", "社保月遅れ"
    categories.Add "国保月遅れ", "国保月遅れ"
    categories.Add "社保月送り", "社保月送り"
    categories.Add "国保月送り", "国保月送り"
    categories.Add "社保返戻", "社保返戻"
    categories.Add "社保未請求", "社保未請求扱い"
    categories.Add "国保返戻", "国保返戻"
    categories.Add "国保未請求", "国保未請求扱い"
    categories.Add "介護返戻", "介護返戻"
    categories.Add "その他", "その他"
    
    ' 各カテゴリのマーカーを検索
    Dim cat_key As Variant
    Dim row_num As Long
    
    For Each cat_key In categories.Keys
        row_num = FindMarkedRow(ws, categories(cat_key))
        
        ' マーカーが見つかった場合のみ辞書に追加
        If row_num > 0 Then
            category_dict.Add cat_key, row_num
            Debug.Print "Added category '" & cat_key & "' with row " & row_num
        End If
    Next cat_key
    
    ' マーカーが一つも見つからなかった場合、従来の方法で検索
    If category_dict.Count = 0 Then
        Debug.Print "WARNING: No markers found, using default method"
        ' 社保カテゴリ
        category_dict.Add "社保再請求", 3
        category_dict.Add "社保月遅れ", 8
        category_dict.Add "社保返戻", 13
        category_dict.Add "社保未請求", 18
        ' 国保カテゴリ
        category_dict.Add "国保再請求", 23
        category_dict.Add "国保月遅れ", 28
        category_dict.Add "国保返戻", 33
        category_dict.Add "国保未請求", 38
    End If
    
    Set GetMarkedCategoryRows = category_dict
End Function

' 請求先タイプ別にカテゴリ開始行を取得する関数（マーキング対応版）
Public Function GetCategoryStartRowsByMarker(ws As Worksheet, payer_type As String) As Object
    Dim all_category_rows As Object
    Dim filtered_dict As Object
    
    Set all_category_rows = GetMarkedCategoryRows(ws)
    Set filtered_dict = CreateObject("Scripting.Dictionary")
    
    Dim cat_key As Variant
    
    If payer_type = "社保" Then
        ' 社保関連のカテゴリのみを抽出
        For Each cat_key In all_category_rows.Keys
            If InStr(cat_key, "社保") > 0 Then
                ' キー名を標準化
                If cat_key = "社保再請求" Then
                    filtered_dict.Add "返戻再請求", all_category_rows(cat_key)
                ElseIf cat_key = "社保月遅れ" Then
                    filtered_dict.Add "月遅れ請求", all_category_rows(cat_key)
                ElseIf cat_key = "社保返戻" Then
                    filtered_dict.Add "返戻・査定", all_category_rows(cat_key)
                ElseIf cat_key = "社保未請求" Then
                    filtered_dict.Add "未請求扱い", all_category_rows(cat_key)
                End If
            End If
        Next cat_key
    ElseIf payer_type = "国保" Then
        ' 国保関連のカテゴリのみを抽出
        For Each cat_key In all_category_rows.Keys
            If InStr(cat_key, "国保") > 0 Then
                ' キー名を標準化
                If cat_key = "国保再請求" Then
                    filtered_dict.Add "返戻再請求", all_category_rows(cat_key)
                ElseIf cat_key = "国保月遅れ" Then
                    filtered_dict.Add "月遅れ請求", all_category_rows(cat_key)
                ElseIf cat_key = "国保返戻" Then
                    filtered_dict.Add "返戻・査定", all_category_rows(cat_key)
                ElseIf cat_key = "国保未請求" Then
                    filtered_dict.Add "未請求扱い", all_category_rows(cat_key)
                End If
            End If
        Next cat_key
    End If
    
    ' 必要なカテゴリが見つからなかった場合は従来の方法で取得
    If filtered_dict.Count = 0 Then
        Debug.Print "WARNING: No " & payer_type & " categories found, using default method"
        Set filtered_dict = GetCategoryStartRows(ws, payer_type)
    End If
    
    Set GetCategoryStartRowsByMarker = filtered_dict
End Function

