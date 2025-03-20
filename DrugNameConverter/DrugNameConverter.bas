Option Explicit

' ラッパーモジュール - 各種機能を呼び出す入口となるモジュール

' 薬品名の一致率に基づくメイン処理を呼び出す関数
Public Sub RunDrugNameMainProcess()
    ' MainModuleの関数を呼び出し
    MainModule.MainProcess  ' フルパスで指定
End Sub

' 薬品名の検索と転記処理を呼び出す関数
Public Sub RunSearchAndTransferDrugData()
    ' MainModuleの関数を呼び出し
    MainModule.SearchAndTransferDrugData  ' フルパスで指定
End Sub

' 一致率計算による薬品名処理を呼び出す関数
Public Sub RunProcessDrugNamesWithMatchRate()
    ' MainModuleの関数を呼び出し
    MainModule.ProcessDrugNamesWithMatchRate  ' フルパスで指定
End Sub

' 包装形態を考慮した医薬品名比較と転記処理を呼び出す関数
Public Sub RunCompareAndTransferDrugNamesByPackage()
    ' MainModuleの関数を呼び出し
    MainModule.CompareAndTransferDrugNamesByPackage
End Sub

' 設定シートのB列7行目以降を比較して転記する処理を呼び出す関数
Public Sub RunCompareAndTransferDrugNamesFromRow7()
    ' MainModuleの関数を呼び出し
    CompareAndTransferDrugNamesFromRow7
End Sub

' メイン処理を呼び出すラッパー関数
Public Sub RunDrugNameComparison()
    ' MainModuleの関数を呼び出し
    CompareAndTransferDrugNames
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