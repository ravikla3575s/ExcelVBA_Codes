Sub UpdateShelfNumbersWithShelfInfo()
    Dim wsTana As Worksheet
    Dim wbMed As Workbook
    Dim wsMed As Worksheet
    Dim lastRowTana As Long
    Dim lastRowMed As Long
    Dim i As Long, j As Long
    Dim medName As String
    Dim medCodeName As String
    Dim shelf1 As String, shelf2 As String, shelf3 As String
    Dim outputFilePath As String
    
    ' 対象のシートを設定
    Set wsTana = ThisWorkbook.Worksheets("tmp_tana")
    
    ' 棚名情報を取得（A1:B3に棚名1〜3の情報があると仮定）
    shelf1 = ThisWorkbook.Sheets(1).Cells(1, 2).Value ' 棚名1
    shelf2 = ThisWorkbook.Sheets(1).Cells(2, 2).Value ' 棚名2
    shelf3 = ThisWorkbook.Sheets(1).Cells(3, 2).Value ' 棚名3
    
    ' tmp_tanaシートの最終行を取得
    lastRowTana = wsTana.Cells(wsTana.Rows.Count, 1).End(xlUp).Row
    
    ' 医薬品コードファイルを開く
    Set wbMed = Workbooks.Open("/Users/yoshipc/Desktop/医薬品コード.xlsx")
    Set wsMed = wbMed.Sheets("シート1 - 医薬品コード")
    lastRowMed = wsMed.Cells(wsMed.Rows.Count, 1).End(xlUp).Row
    
    ' A4以降のセルに記載された医薬品名リストでtmp_tanaの医薬品を部分一致検索
    Dim readRow As Long
    readRow = 4 ' A4から読み取り開始と仮定
    
    Do While ThisWorkbook.Sheets(1).Cells(readRow, 3).Value <> ""
        medName = ThisWorkbook.Sheets(1).Cells(readRow, 3).Value
        
        ' tmp_tanaの各薬品名と部分一致検索
        For i = 2 To lastRowTana
            If InStr(1, wsTana.Cells(i, 2).Value, medName, vbTextCompare) > 0 Then
                ' 部分一致した行に棚番を設定（空欄の場合は変更しない）
                If shelf1 <> "" Then wsTana.Cells(i, 7).Value = "[" & shelf1 & "]"
                If shelf2 <> "" Then wsTana.Cells(i, 8).Value = "[" & shelf2 & "]"
                If shelf3 <> "" Then wsTana.Cells(i, 9).Value = "[" & shelf3 & "]"
                Exit For
            End If
        Next i
        
        readRow = readRow + 1
    Loop
    
    ' CSVファイルとして出力する
    outputFilePath = Application.ThisWorkbook.Path & Application.PathSeparator & "updated_tmp_tana.csv"
    Call ExportToCSV(wsTana, outputFilePath)
    
    ' 医薬品コードファイルを閉じる
    wbMed.Close SaveChanges:=False
    
    MsgBox "棚番の更新が完了し、CSVファイルとして保存しました。"
End Sub

' シートをCSVファイルとして出力するサブプロシージャ
Sub ExportToCSV(ws As Worksheet, filePath As String)
    Dim csvData As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    
    ' シートの最終行と最終列を取得
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' データをCSV形式に変換
    For i = 1 To lastRow
        For j = 1 To lastCol
            csvData = csvData & ws.Cells(i, j).Value
            If j < lastCol Then csvData = csvData & ","
        Next j
        csvData = csvData & vbNewLine
    Next i
    
    ' ファイルに書き込み
    Open filePath For Output As #1
    Print #1, csvData
    Close #1
End Sub
