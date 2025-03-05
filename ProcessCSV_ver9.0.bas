Option Explicit

' グローバル変数定義
Dim gOlderList As New Collection         ' 過去月データのリスト
Dim gRebillData As New Collection        ' 返戻再請求データ
Dim gLateData As New Collection          ' 月遅れ請求データ
Dim gUnclaimedList As New Collection     ' 未請求扱いデータ（前月分）
Dim gSelectedUnclaimed As New Collection ' ユーザー選択された未請求データ
Dim wsDetails As Worksheet               ' 保険請求管理報告書Bシート参照

Public Sub ProcessCSV()
    Dim fso As Object
    Dim csvFolder As String, savePath As String, templatePath As String
    Dim fixfFiles As Collection
    Dim fixfFile As Object
    Dim targetYear As String, targetMonth As String
    Dim targetFile As String
    Dim newBook As Workbook
    Dim payerType As String
    
    On Error GoTo ProcErr  ' エラーハンドリング開始
    
    ' ファイルシステムオブジェクトの作成
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 対象CSVフォルダ・出力保存先・テンプレートパスの設定（必要に応じて適切なパスを指定）
    csvFolder = "＜CSVフォルダのパス＞"
    savePath = "＜出力保存先フォルダのパス＞"
    templatePath = "＜帳票テンプレートExcelのパス＞"
    
    ' フォルダ内のfixfファイル（請求確定ファイル）一覧を取得
    Set fixfFiles = FindAllFixfFiles(fso, csvFolder)
    If fixfFiles Is Nothing Or fixfFiles.Count = 0 Then
        MsgBox "指定フォルダ内に処理対象のCSVファイル（fixfファイル）がありません。", vbExclamation
        Exit Sub
    End If
    
    ' 各fixfファイルに対する処理ループ
    For Each fixfFile In fixfFiles
        ' 請求年月をファイル名から取得（GYYMM形式をRYYMM形式に変換）
        Call GetYearMonthFromFixf(fixfFile.Path, targetYear, targetMonth)
        
        ' 出力先ファイルパスの組み立てと存在チェック
        targetFile = FindOrCreateReport(savePath, targetYear, targetMonth, templatePath)
        If fso.FileExists(targetFile) Then
            ' 既に同じ請求年月のファイルが存在する場合はスキップ
            MsgBox "請求年月 " & targetYear & targetMonth & " は既に処理済みのためスキップします。", vbInformation
            ' 処理済みCSVを「Processed」フォルダへ移動
            MoveProcessedFiles fso, csvFolder, targetYear, targetMonth
            ' 次のfixfファイルへ
            GoTo NextFixf
        End If
        
        ' 帳票ブック（新規または既存）を開く
        Set newBook = Workbooks.Open(targetFile)
        
        ' 帳票ブックのヘッダー情報等を設定
        Call SetTemplateInfo(newBook, targetYear, targetMonth)
        
        ' フォルダ内CSVを読み込み（fixfファイル以外）
        Call ProcessAllCSVFiles(fso, newBook, csvFolder)
        
        ' 請求確定ファイルの詳細データを帳票ブックへ転記（保険請求管理報告書A/Bシート以外のデータをもとにBシートを埋める）
        Call TransferBillingDetails(newBook, fixfFile.Path)
        
        ' 帳票ブックを保存して閉じる
        newBook.Save
        newBook.Close SaveChanges:=False
        
        ' 処理完了メッセージ表示（社保/国保で切替え）
        If InStr(1, fixfFile.Name, "社保", vbTextCompare) > 0 Or InStr(1, fixfFile.Name, "sh", vbTextCompare) > 0 Then
            payerType = "社保"
        Else
            payerType = "国保"
        End If
        MsgBox payerType & " のデータ転記完了", vbInformation
        
        ' 処理済みCSVをProcessedフォルダへ移動
        MoveProcessedFiles fso, csvFolder, targetYear, targetMonth
        
NextFixf:
        ' 次のfixfFileへループ継続
    Next fixfFile
    
    ' 全ての処理完了メッセージ
    MsgBox "全てのCSVファイルの処理が完了しました！", vbInformation
    Exit Sub

ProcErr:
    ' エラー発生時の処理
    MsgBox "ファイル " & (IIf(fixfFile Is Nothing, "(不明)", fixfFile.Name)) & " の処理中にエラーが発生しました: " & Err.Description, vbExclamation
    ' ブックが開いていれば閉じる（保存せず）
    If Not newBook Is Nothing Then
        On Error Resume Next
        newBook.Close SaveChanges:=False
        On Error GoTo 0
    End If
    ' 新規作成途中の出力ファイルがあれば削除して次へ
    If fso.FileExists(targetFile) Then
        On Error Resume Next
        fso.DeleteFile targetFile
        On Error GoTo 0
    End If
    Resume NextFixf  ' 次のfixfファイルの処理に移行
End Sub

Private Function FindAllFixfFiles(fso As Object, csvFolder As String) As Collection
    Dim result As New Collection
    Dim folder As Object, file As Object
    
    If Not fso.FolderExists(csvFolder) Then
        MsgBox "CSVフォルダが見つかりません: " & csvFolder, vbExclamation
        Set FindAllFixfFiles = Nothing
        Exit Function
    End If
    Set folder = fso.GetFolder(csvFolder)
    For Each file In folder.Files
        ' ファイル名に "fixf" を含むCSVを収集
        If InStr(1, LCase(file.Name), "fixf") > 0 And LCase(fso.GetExtensionName(file.Path)) = "csv" Then
            result.Add file
        End If
    Next
    If result.Count = 0 Then
        Set FindAllFixfFiles = Nothing
    Else
        Set FindAllFixfFiles = result
    End If
End Function

Private Sub GetYearMonthFromFixf(fixfFilePath As String, ByRef targetYear As String, ByRef targetMonth As String)
    ' ファイル名からGYYMM形式の5桁コードを取得し、RYYMM形式に変換
    Dim fileName As String, code As String
    fileName = Dir(fixfFilePath)  ' ファイル名のみ取得
    If InStr(fileName, ".") > 0 Then
        fileName = Left$(fileName, InStrRev(fileName, ".") - 1)
    End If
    If Len(fileName) >= 5 Then
        code = Right$(fileName, 5)
    Else
        code = fileName
    End If
    Dim eraCode As String, yy As String, mm As String
    If Len(code) = 5 Then
        eraCode = Left$(code, 1)
        yy = Mid$(code, 2, 2)
        mm = Right$(code, 2)
    Else
        eraCode = ""
        yy = ""
        mm = ""
    End If
    Select Case eraCode
        Case "5": targetYear = "R" & Format$(Val(yy), "00")
        Case "4": targetYear = "H" & Format$(Val(yy), "00")
        Case "3": targetYear = "S" & Format$(Val(yy), "00")
        Case "2": targetYear = "T" & Format$(Val(yy), "00")
        Case "1": targetYear = "M" & Format$(Val(yy), "00")
        Case Else: targetYear = eraCode & yy  ' 想定外: そのまま結合
    End Select
    targetMonth = mm
End Sub

Private Function FindOrCreateReport(savePath As String, targetYear As String, targetMonth As String, templatePath As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim circledMonth As String
    circledMonth = ConvertToCircledNumber(targetMonth)
    ' 出力ファイル名の組み立て（例: "保険請求管理報告書_R05④.xlsm"）
    Dim fileName As String
    fileName = "保険請求管理報告書_" & targetYear & circledMonth & ".xlsm"
    Dim targetFile As String
    targetFile = fso.BuildPath(savePath, fileName)
    ' ファイルが存在しなければテンプレートからコピーして作成
    If Not fso.FileExists(targetFile) Then
        fso.CopyFile templatePath, targetFile
    End If
    FindOrCreateReport = targetFile
End Function

Private Function ConvertToCircledNumber(month As String) As String
    Dim n As Integer: n = Val(month)
    Dim circledChar As String
    If n >= 1 And n <= 20 Then
        circledChar = ChrW(&H245F + n)  ' ①(0x2460)～⑳(0x2473)に対応
    Else
        circledChar = month
    End If
    ConvertToCircledNumber = circledChar
End Function

Private Sub SetTemplateInfo(newBook As Workbook, targetYear As String, targetMonth As String)
    ' 帳票ブックのシート名やヘッダー情報を設定
    Dim wsA As Worksheet, wsB As Worksheet
    Set wsA = newBook.Sheets("保険請求管理報告書A")
    Set wsB = newBook.Sheets("保険請求管理報告書B")
    ' タイトル行などに請求年月を設定（例：R05年4月）
    wsA.Range("A1").Value = targetYear & "年" & CInt(targetMonth) & "月度 保険請求管理報告書A"
    wsB.Range("A1").Value = targetYear & "年" & CInt(targetMonth) & "月度 保険請求管理報告書B"
End Sub

Private Sub ProcessAllCSVFiles(fso As Object, newBook As Workbook, csvFolder As String)
    Dim folder As Object, file As Object
    Dim wsCSV As Worksheet
    
    Set folder = fso.GetFolder(csvFolder)
    For Each file In folder.Files
        Dim fName As String
        fName = LCase(file.Name)
        ' fixfファイルおよびCSV以外はスキップ
        If InStr(fName, "fixf") > 0 Or LCase(fso.GetExtensionName(file.Path)) <> "csv" Then
            Continue For
        End If
        ' CSV種類に応じて新規シートを作成しインポート
        If InStr(fName, "fmei") > 0 Then
            Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(2))
            wsCSV.Name = "振込額明細書"
            On Error Resume Next
            ImportCSVData file.Path, wsCSV
            If Err.Number <> 0 Then
                MsgBox "ファイル " & file.Name & " の読み込みでエラー: " & Err.Description, vbExclamation
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf InStr(fName, "zogn") > 0 Then
            Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(2))
            wsCSV.Name = "増減点連絡書"
            On Error Resume Next
            ImportCSVData file.Path, wsCSV
            If Err.Number <> 0 Then
                MsgBox "ファイル " & file.Name & " の読み込みでエラー: " & Err.Description, vbExclamation
                Err.Clear
            End If
            On Error GoTo 0
        ElseIf InStr(fName, "henr") > 0 Then
            Set wsCSV = newBook.Sheets.Add(After:=newBook.Sheets(2))
            wsCSV.Name = "返戻内訳書"
            On Error Resume Next
            ImportCSVData file.Path, wsCSV
            If Err.Number <> 0 Then
                MsgBox "ファイル " & file.Name & " の読み込みでエラー: " & Err.Description, vbExclamation
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next file
End Sub

Private Sub ImportCSVData(filePath As String, ws As Worksheet)
    ' CSVファイル内容を指定シートに転記
    Dim fso As Object, ts As Object
    Dim line As String, values As Variant
    Dim r As Long, c As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' テキストストリームを開く（システム既定のエンコーディングで読み込み）
    Set ts = fso.OpenTextFile(filePath, 1)
    r = 1
    On Error GoTo ImportErr  ' エラー発生時は終了処理へ
    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        If Len(line) > 0 Then
            values = Split(line, ",")
            For c = LBound(values) To UBound(values)
                ws.Cells(r, c + 1).Value = values(c)
            Next c
        End If
        r = r + 1
    Loop
    ts.Close
    Exit Sub

ImportErr:
    If Not ts Is Nothing Then ts.Close
    Err.Raise Err.Number, Err.Source, Err.Description  ' エラーを呼び出し元に再送出
End Sub

Private Function IsPastMonth(dispensingMonth As String, targetYYMM As String) As Boolean
    ' 調剤年月dispensingMonth（GYYMM形式）が今回請求分targetYYMMより過去か判定
    If Len(dispensingMonth) < 5 Or Len(targetYYMM) < 5 Then
        IsPastMonth = False
    Else
        ' 元号コードを比較（小さい＝過去の元号）
        Dim eraDisp As String, eraTarget As String
        eraDisp = Left$(dispensingMonth, 1)
        eraTarget = Left$(targetYYMM, 1)
        If IsNumeric(eraDisp) And IsNumeric(eraTarget) Then
            If CInt(eraDisp) < CInt(eraTarget) Then
                IsPastMonth = True
                Exit Function
            ElseIf CInt(eraDisp) > CInt(eraTarget) Then
                IsPastMonth = False
                Exit Function
            End If
        ElseIf eraDisp <> eraTarget Then
            ' コードが数字以外（想定外）や異なる場合は、とりあえずFalse
            IsPastMonth = (eraDisp < eraTarget)
            Exit Function
        End If
        ' 同一元号の場合、年＋月の下4桁の数値で比較
        If Val(Right$(dispensingMonth, 4)) < Val(Right$(targetYYMM, 4)) Then
            IsPastMonth = True Else IsPastMonth = False
        End If
    End If
End Function

Private Sub TransferBillingDetails(newBook As Workbook, fixfFilePath As String)
    Dim wsMain As Worksheet
    Set wsMain = newBook.Sheets("保険請求管理報告書A")
    Set wsDetails = newBook.Sheets("保険請求管理報告書B")
    
    ' fixf（請求確定）CSVをAシートに読み込み
    ImportCSVData fixfFilePath, wsMain
    
    ' 対象請求年月のGYYMMコードを作成（比較用）
    Dim targetCode As String
    Select Case Left$(wsMain.Cells(1, 2).Value, 1)
        Case "5": targetCode = "5" & Mid$(wsMain.Cells(1, 2).Value, 2, 4)   ' Reiwa (令和)
        Case "4": targetCode = "4" & Mid$(wsMain.Cells(1, 2).Value, 2, 4)   ' Heisei (平成) 等
        Case Else: targetCode = wsMain.Cells(1, 2).Value
    End Select
    
    ' 過去月データを収集
    gOlderList.Clear
    Dim lastRow As Long, i As Long
    lastRow = wsMain.Cells(wsMain.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastRow
        Dim dispCode As String
        dispCode = wsMain.Cells(i, 2).Value  ' 仮定：2列目に調剤年月コード(GYYMM)がある
        If Len(dispCode) > 0 And IsPastMonth(dispCode, targetCode) Then
            ' 必要項目を収集（仮定：氏名や点数等の列番号を適宜設定）
            Dim rec(3) As Variant
            rec(0) = dispCode                                 ' 調剤年月コード
            rec(1) = wsMain.Cells(i, 13).Value                ' 氏名（例：13列目）
            rec(2) = wsMain.Cells(i, 3).Value                 ' 医療機関名（例：3列目）
            rec(3) = wsMain.Cells(i, 20).Value                ' 請求点数（例：20列目）
            gOlderList.Add rec
        End If
    Next i
    
    ' ユーザーに過去月データの再請求/月遅れ区分を選択させる処理（ここでは簡略化）
    gRebillData.Clear
    gLateData.Clear
    If gOlderList.Count > 0 Then
        ' （簡易処理: 全件を「月遅れ請求」として扱う）
        Dim idx As Integer: idx = 1
        Dim itm As Variant
        For Each itm In gOlderList
            gLateData.Add itm   ' 全件を月遅れ請求リストへ
            idx = idx + 1
        Next itm
    End If
    
    ' 返戻内訳書シートのデータを収集（査定・返戻データ）
    Dim assessmentDict As Object: Set assessmentDict = CreateObject("Scripting.Dictionary")
    Dim wsHenr As Worksheet
    On Error Resume Next
    Set wsHenr = newBook.Sheets("返戻内訳書")
    On Error GoTo 0
    If Not wsHenr Is Nothing Then
        Dim lastRowH As Long: lastRowH = wsHenr.Cells(wsHenr.Rows.Count, 1).End(xlUp).Row
        Dim j As Long, key As String
        j = 1
        Dim recAss(3) As Variant
        Dim r As Long
        For r = 2 To lastRowH
            recAss(0) = wsHenr.Cells(r, 1).Value  ' 例: 返戻理由等
            recAss(1) = wsHenr.Cells(r, 2).Value
            recAss(2) = wsHenr.Cells(r, 3).Value
            recAss(3) = wsHenr.Cells(r, 4).Value
            key = CStr(j)
            assessmentDict(key) = recAss
            j = j + 1
        Next r
    End If
    
    ' 前月未請求扱いデータの追加（ユーザーフォームによる選択）
    Dim payerType As String
    If InStr(1, fixfFilePath, "社保", vbTextCompare) > 0 Or InStr(1, fixfFilePath, "sh", vbTextCompare) > 0 Then
        payerType = "社保" Else payerType = "国保"
    Dim prevYearMonth As String
    prevYearMonth = GetPrevYearMonth(targetCode)  ' 前月の年月コードを取得
    Dim selectedUnclaimedDict As Object
    Set selectedUnclaimedDict = AddUnclaimedRecords(payerType, prevYearMonth)
    If Not selectedUnclaimedDict Is Nothing Then
        Dim k As Variant
        For Each k In selectedUnclaimedDict.keys
            assessmentDict(k) = selectedUnclaimedDict(k)  ' （例：未請求分は査定扱いに追加）
        Next k
    End If
    
    ' 月遅れ請求・返戻再請求の辞書を作成（ユーザー選択結果に基づく）
    Dim rebillDict As Object: Set rebillDict = CreateObject("Scripting.Dictionary")
    Dim lateDict As Object: Set lateDict = CreateObject("Scripting.Dictionary")
    Dim it As Variant, keyLate As String, keyRe As String
    Dim n As Long
    n = 1
    For Each it In gRebillData  ' 返戻再請求データを辞書にセット
        keyRe = "R" & CStr(n)
        rebillDict(keyRe) = it
        n = n + 1
    Next it
    n = 1
    For Each it In gLateData    ' 月遅れ請求データを辞書にセット
        keyLate = "L" & CStr(n)
        lateDict(keyLate) = it
        n = n + 1
    Next it
    
    ' 必要に応じて各カテゴリ枠の追加
    If rebillDict.Count > 5 Then
        wsDetails.Rows(startRow("返戻再請求") + 5 & ":" & startRow("返戻再請求") + rebillDict.Count - 1).Insert Shift:=xlDown
    End If
    If lateDict.Count > 5 Then
        wsDetails.Rows(startRow("月遅れ請求") + 5 & ":" & startRow("月遅れ請求") + lateDict.Count - 1).Insert Shift:=xlDown
    End If
    If assessmentDict.Count > 5 Then
        wsDetails.Rows(startRow("返戻・査定") + 5 & ":" & startRow("返戻・査定") + assessmentDict.Count - 1).Insert Shift:=xlDown
    End If
    
    ' Bシートに各カテゴリのデータを転記
    Call TransferData(rebillDict, wsDetails, startRow("返戻再請求"), payerType)
    Call TransferData(lateDict, wsDetails, startRow("月遅れ請求"), payerType)
    Call TransferData(assessmentDict, wsDetails, startRow("返戻・査定"), payerType)
End Sub

Private Function GetPrevYearMonth(targetCode As String) As String
    ' 対象年月コード（GYYMM形式）の前月コードを取得
    If Len(targetCode) <> 5 Or Not IsNumeric(Left$(targetCode, 1)) Then
        GetPrevYearMonth = ""
    Else
        Dim eraNum As Integer, yy As Integer, mm As Integer
        eraNum = CInt(Left$(targetCode, 1))
        yy = CInt(Mid$(targetCode, 2, 2))
        mm = CInt(Right$(targetCode, 2))
        If mm = 1 Then
            mm = 12
            If yy = 1 Then
                ' 元号切替（例：R01→H31）
                If eraNum = 5 Then  ' Reiwa1の前月はHeisei31
                    eraNum = 4
                    yy = 31
                Else
                    yy = yy - 1
                End If
            Else
                yy = yy - 1
            End If
        Else
            mm = mm - 1
        End If
        GetPrevYearMonth = CStr(eraNum) & Format$(yy, "00") & Format$(mm, "00")
    End If
End Function

Private Function AddUnclaimedRecords(payerType As String, prevYearMonth As String) As Object
    ' 前月未請求扱いのデータを取得（ダミー実装: 常に空辞書を返す）
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Set AddUnclaimedRecords = dict
End Function

Private Function startRow(sectionName As String) As Long
    ' Bシート内の各セクション開始行を取得（セクション見出しの次行を返す）
    Dim rng As Range
    Set rng = wsDetails.Cells.Find(sectionName, LookAt:=xlWhole)
    If Not rng Is Nothing Then
        startRow = rng.Row + 1
    Else
        startRow = 1  ' 見出しが見つからない場合は1行目を返す
    End If
End Function

Private Sub TransferData(dataDict As Object, ws As Worksheet, startRowIndex As Long, payerType As String)
    ' データ辞書の内容を指定シートに転記
    Dim i As Long, key As Variant, rec As Variant
    i = 0
    For Each key In dataDict.Keys
        rec = dataDict(key)
        If IsArray(rec) Then
            ' 配列データを各列に配置
            ws.Cells(startRowIndex + i, 1).Value = rec(0)
            ws.Cells(startRowIndex + i, 2).Value = rec(1)
            ws.Cells(startRowIndex + i, 3).Value = rec(2)
            ws.Cells(startRowIndex + i, 4).Value = rec(3)
        Else
            ws.Cells(startRowIndex + i, 1).Value = rec
        End If
        i = i + 1
    Next key
End Sub

Private Sub MoveProcessedFiles(fso As Object, csvFolder As String, targetYear As String, targetMonth As String)
    On Error Resume Next
    Dim processedFolderPath As String
    processedFolderPath = fso.BuildPath(csvFolder, "Processed")
    If Not fso.FolderExists(processedFolderPath) Then
        fso.CreateFolder processedFolderPath
    End If
    ' 対象年月コード（RYYMM形式）からGYYMM形式の部分文字列を生成
    Dim eraCodeChar As String, eraLetter As String, yearNum As String, code As String
    eraLetter = Left$(targetYear, 1)
    yearNum = Mid$(targetYear, 2)  ' "05" 等
    Select Case eraLetter
        Case "R": eraCodeChar = "5"
        Case "H": eraCodeChar = "4"
        Case "S": eraCodeChar = "3"
        Case "T": eraCodeChar = "2"
        Case "M": eraCodeChar = "1"
        Case Else: eraCodeChar = ""
    End Select
    code = eraCodeChar & yearNum & targetMonth  ' 例: "50504"
    Dim file As Object
    For Each file In fso.GetFolder(csvFolder).Files
        If InStr(1, file.Name, code, vbTextCompare) > 0 Then
            fso.MoveFile file.Path, fso.BuildPath(processedFolderPath, file.Name)
        End If
    Next file
End Sub