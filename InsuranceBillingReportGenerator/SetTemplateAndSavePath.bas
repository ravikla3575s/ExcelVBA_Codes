Attribute VB_Name = "SetTemplateAndSavePath"
Sub SetTemplateAndSavePath()
    Dim ws As Worksheet
    Dim templatePath As String
    Dim saveFolder As String
    Dim storeName As String

    ' シートを指定（必要に応じて変更）
    Set ws = ThisWorkbook.Sheets(1)

    ' B1: 店舗名をユーザーに入力させる
    storeName = InputBox("店舗名を入力してください", "店舗名の設定")
    If storeName = "" Then
        MsgBox "店舗名が入力されていません。処理を中止します。", vbExclamation
        Exit Sub
    End If
    ws.Range("B1").value = storeName

    ' B2: テンプレート保存フォルダを選択
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "テンプレート保存フォルダを選択してください"
        If .Show = -1 Then
            templatePath = .SelectedItems(1)
        Else
            MsgBox "テンプレート保存フォルダが選択されていません。処理を中止します。", vbExclamation
            Exit Sub
        End If
    End With
    ws.Range("B2").value = templatePath

    ' B3: 新規ファイルの保存フォルダを選択
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "新規ファイルの保存フォルダを選択してください"
        If .Show = -1 Then
            saveFolder = .SelectedItems(1)
        Else
            MsgBox "新規ファイルの保存フォルダが選択されていません。処理を中止します。", vbExclamation
            Exit Sub
        End If
    End With
    ws.Range("B3").value = saveFolder

    ' 設定完了メッセージ
    MsgBox "設定が完了しました。" & vbCrLf & _
           "店舗名: " & storeName & vbCrLf & _
           "テンプレート保存先: " & templatePath & vbCrLf & _
           "新規ファイル保存先: " & saveFolder, vbInformation
End Sub

