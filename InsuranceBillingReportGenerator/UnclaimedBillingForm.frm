VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UnclaimedBillingForm
   Caption         =   "未請求レセプト登録"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6615
   OleObjectBlob   =   "UnclaimedBillingForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UnclaimedBillingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 調剤年月を保持する変数
Private m_era_year As Integer
Private m_month As Integer

' 全角数字を半角数字に変換する関数
Private Function ConvertToHankaku(ByVal strText As String) As String
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

' 請求点数の入力時の処理
Private Sub txtBillingPoints_Change()
    Dim cursorPos As Long
    cursorPos = txtBillingPoints.SelStart
    txtBillingPoints.Text = UtilityModule.ConvertToHankaku(txtBillingPoints.Text)
    txtBillingPoints.SelStart = cursorPos
End Sub

' 保険割合の入力時の処理
Private Sub txtInsuranceRatio_Change()
    Dim cursorPos As Long
    cursorPos = txtInsuranceRatio.SelStart
    txtInsuranceRatio.Text = UtilityModule.ConvertToHankaku(txtInsuranceRatio.Text)
    txtInsuranceRatio.SelStart = cursorPos
End Sub

Public Sub SetDispensingDate(ByVal era_year As Integer, ByVal month As Integer)
    m_era_year = era_year
    m_month = month
End Sub

Private Sub UserForm_Initialize()
    ' フォームの初期化処理
    Me.Caption = "未請求レセプト登録"
    
    ' コントロールの配置とプロパティの設定
    With Me
        .Width = 400
        .Height = 350  ' フォームの高さを調整
        
        ' ラベルの追加
        With .Controls.Add("Forms.Label.1", "lblPatientName")
            .Caption = "患者氏名："
            .Left = 20
            .Top = 20
            .Width = 80
            .Height = 20
        End With
        
        ' テキストボックスの追加（患者氏名）
        With .Controls.Add("Forms.TextBox.1", "txtPatientName")
            .Left = 100
            .Top = 20
            .Width = 200
            .Height = 20
        End With
        
        ' ラベルの追加（調剤年月）
        With .Controls.Add("Forms.Label.1", "lblDispensingDate")
            .Caption = "調剤年月："
            .Left = 20
            .Top = 50
            .Width = 80
            .Height = 20
        End With
        
        ' テキストボックスの追加（調剤年月）
        With .Controls.Add("Forms.TextBox.1", "txtDispensingDate")
            .Left = 100
            .Top = 50
            .Width = 100
            .Height = 20
            ' 初期値を設定
            If m_era_year > 0 And m_month > 0 Then
                .Value = "令和" & m_era_year & "年" & m_month & "月"
            End If
        End With
        
        ' ラベルの追加（医療機関）
        With .Controls.Add("Forms.Label.1", "lblMedicalInstitution")
            .Caption = "医療機関："
            .Left = 20
            .Top = 80
            .Width = 80
            .Height = 20
        End With
        
        ' テキストボックスの追加（医療機関）
        With .Controls.Add("Forms.TextBox.1", "txtMedicalInstitution")
            .Left = 100
            .Top = 80
            .Width = 200
            .Height = 20
        End With
        
        ' ラベルの追加（請求点数）
        With .Controls.Add("Forms.Label.1", "lblBillingPoints")
            .Caption = "請求点数："
            .Left = 20
            .Top = 110
            .Width = 80
            .Height = 20
        End With
        
        ' テキストボックスの追加（請求点数）
        With .Controls.Add("Forms.TextBox.1", "txtBillingPoints")
            .Left = 100
            .Top = 110
            .Width = 100
            .Height = 20
        End With
        
        ' ラベルの追加（請求保険割合）
        With .Controls.Add("Forms.Label.1", "lblInsuranceRatio")
            .Caption = "保険割合："
            .Left = 20
            .Top = 140
            .Width = 80
            .Height = 20
        End With
        
        ' テキストボックスの追加（請求保険割合）
        With .Controls.Add("Forms.TextBox.1", "txtInsuranceRatio")
            .Left = 100
            .Top = 140
            .Width = 40
            .Height = 20
        End With
        
        ' 割合の単位（割）ラベルを追加
        With .Controls.Add("Forms.Label.1", "lblPercentage")
            .Caption = "割"
            .Left = 145
            .Top = 140
            .Width = 20
            .Height = 20
        End With
        
        ' 登録ボタンの追加
        With .Controls.Add("Forms.CommandButton.1", "btnRegister")
            .Caption = "登録"
            .Left = 100
            .Top = 180  ' ボタンの位置を下に調整
            .Width = 80
            .Height = 30
        End With
        
        ' キャンセルボタンの追加
        With .Controls.Add("Forms.CommandButton.1", "btnCancel")
            .Caption = "キャンセル"
            .Left = 200
            .Top = 180  ' ボタンの位置を下に調整
            .Width = 80
            .Height = 30
        End With
    End With
End Sub

Private Sub btnRegister_Click()
    ' 入力値の検証
    If Trim(Me.txtPatientName.Value) = "" Then
        MsgBox "患者氏名を入力してください。", vbExclamation
        Exit Sub
    End If
    
    If Trim(Me.txtDispensingDate.Value) = "" Then
        MsgBox "調剤年月を入力してください。", vbExclamation
        Exit Sub
    End If
    
    If Trim(Me.txtMedicalInstitution.Value) = "" Then
        MsgBox "医療機関を入力してください。", vbExclamation
        Exit Sub
    End If
    
    ' 請求点数の検証（自動的に半角変換済み）
    If Not IsNumeric(Me.txtBillingPoints.Value) Then
        MsgBox "請求点数を数値で入力してください。", vbExclamation
        Exit Sub
    End If
    
    ' 保険割合の検証（自動的に半角変換済み）
    If Not IsNumeric(Me.txtInsuranceRatio.Value) Then
        MsgBox "保険割合を数値で入力してください。", vbExclamation
        Exit Sub
    End If
    
    ' 保険割合の範囲チェック
    Dim ratio As Integer
    If Not IsNumeric(Me.txtInsuranceRatio.Value) Or _
       InStr(Me.txtInsuranceRatio.Value, ".") > 0 Then
        MsgBox "保険割合は1～10の整数で入力してください。", vbExclamation
        Exit Sub
    End If
    
    ratio = CInt(Me.txtInsuranceRatio.Value)
    If ratio < 1 Or ratio > 10 Then
        MsgBox "保険割合は1割から10割の間で入力してください。", vbExclamation
        Exit Sub
    End If
    
    ' フォームを閉じる（OKで閉じる）
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    ' フォームをキャンセルで閉じる
    Me.Hide
End Sub 