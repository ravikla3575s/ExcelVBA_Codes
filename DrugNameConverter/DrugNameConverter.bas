Option Explicit

' ラッパーモジュール - Mac互換のコード

' メイン処理を呼び出すラッパー関数（7行目以降の医薬品名比較）
Public Sub RunDrugNameComparison()
    ' MainModuleの関数を呼び出し
    MainModule.CompareAndTransferDrugNamesFromRow7
End Sub

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
    
    MsgBox "包装形態のドロップダウンリストを設定しました。", vbInformation
End Sub

' シート1にインストラクションを追加する関数
Public Sub AddInstructions()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    
    ' 既存の指示を削除
    ws.Range("A2:C3").ClearContents
    
    ' 指示を追加
    ws.Range("A2").Value = "【使い方】"
    ws.Range("A3").Value = "1. B4の包装形態を選択して下さい"
    ws.Range("A4").Value = "包装形態:"
    ws.Range("B4").Font.Bold = True
    
    ' セルの書式設定
    ws.Range("A2").Font.Bold = True
    ws.Range("A2").Font.Size = 12
    
    ' 実行方法の指示
    ws.Range("A5").Value = "2. B7以降に検索する医薬品名を入力"
    ws.Range("A6").Value = "No."
    ws.Range("B6").Value = "検索医薬品名"
    ws.Range("C6").Value = "一致医薬品名"
    
    With ws.Range("A6:C6")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247) ' 薄い青色の背景
    End With
    
    ' 実行方法の案内
    Dim note As String
    note = "※実行方法: メニューから「ツール」→「マクロ」を選択し、" & vbCrLf & _
           "「RunDrugNameComparison」を選んで「実行」ボタンをクリックします。"
    
    ws.Range("A" & (ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 2)).Value = note
    
    MsgBox "使用方法の指示を追加しました。Mac版Excelでは、メニューから「ツール」→「マクロ」を選択し、" & vbCrLf & _
           "「RunDrugNameComparison」を選んで処理を実行してください。", vbInformation
End Sub

