Option Explicit

' ラッパーモジュール - 各種機能を呼び出す入口となるモジュール

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

' ボタンの作成とマクロの割り当てを行うラッパー関数
Public Sub CreateComparisonButton()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    
    ' 既存のボタンを削除
    On Error Resume Next
    ws.Shapes.SelectAll
    Selection.Delete
    On Error GoTo 0
    
    ' フォームコントロールボタンを作成
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRectangle, 200, 30, 120, 30)
    
    With btn
        .Fill.ForeColor.RGB = RGB(221, 235, 247) ' 薄い青色
        .Line.ForeColor.RGB = RGB(91, 155, 213) ' 濃い青色
        .TextFrame.Characters.Text = "医薬品名比較"
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Name = "CompareButton"
        .OnAction = "DrugNameConverter.RunDrugNameComparison"
    End With
    
    MsgBox "比較ボタンを作成しました。このボタンをクリックすると7行目以降の医薬品名比較処理が実行されます。", vbInformation
End Sub

