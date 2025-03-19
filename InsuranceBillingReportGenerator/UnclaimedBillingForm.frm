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

' フォームの結果を保持するプライベート変数
Private m_DialogResult As Boolean
Private m_PatientName As String
Private m_DispensingDate As String
Private m_MedicalInstitution As String
Private m_UnclaimedReason As String
Private m_BillingDestination As String
Private m_InsuranceRatio As Integer
Private m_BillingPoints As Long
Private m_Remarks As String
Private m_ContinueInput As Boolean  ' 続けて入力するかどうか

' 配列操作用のプロパティを追加
Private m_CurrentIndex As Long
Private m_IsEditMode As Boolean

' DialogResult プロパティ
Public Property Get DialogResult() As Boolean
    DialogResult = m_DialogResult
End Property

' PatientName プロパティ
Public Property Get PatientName() As String
    PatientName = m_PatientName
End Property

' DispensingDate プロパティ
Public Property Get DispensingDate() As String
    DispensingDate = m_DispensingDate
End Property

' MedicalInstitution プロパティ
Public Property Get MedicalInstitution() As String
    MedicalInstitution = m_MedicalInstitution
End Property

' UnclaimedReason プロパティ
Public Property Get UnclaimedReason() As String
    UnclaimedReason = m_UnclaimedReason
End Property

' BillingDestination プロパティ
Public Property Get BillingDestination() As String
    BillingDestination = m_BillingDestination
End Property

' InsuranceRatio プロパティ
Public Property Get InsuranceRatio() As Integer
    InsuranceRatio = m_InsuranceRatio
End Property

' BillingPoints プロパティ
Public Property Get BillingPoints() As Long
    BillingPoints = m_BillingPoints
End Property

' Remarks プロパティ
Public Property Get Remarks() As String
    Remarks = m_Remarks
End Property

' ContinueInput プロパティを追加
Public Property Get ContinueInput() As Boolean
    ContinueInput = m_ContinueInput
End Property

' 調剤年月を設定するメソッド
Public Sub SetDispensingDate(ByVal era_year As Integer, ByVal month As Integer)
    m_era_year = era_year
    m_month = month
End Sub

' 配列データを保持するプロパティ
Public Property Let CurrentIndex(ByVal value As Long)
    m_CurrentIndex = value
End Property

Public Property Get CurrentIndex() As Long
    CurrentIndex = m_CurrentIndex
End Property

' フォームにデータをロードするメソッド
Public Sub LoadData(ByRef unclaimedData() As Variant, ByVal index As Long)
    m_IsEditMode = True
    m_CurrentIndex = index
    
    With Me
        .txtPatientName.Value = unclaimedData(1, index)
        .txtMedicalInstitution.Value = IIf(IsNull(unclaimedData(3, index)), "", unclaimedData(3, index))
        .txtUnclaimedReason.Value = IIf(IsNull(unclaimedData(4, index)), "", unclaimedData(4, index))
        
        ' 請求先の設定
        If unclaimedData(5, index) = "社保" Then
            .frameBillingDestination.Controls("optShaho").Value = True
        Else
            .frameBillingDestination.Controls("optKokuho").Value = True
        End If
        
        .txtInsuranceRatio.Value = IIf(IsNull(unclaimedData(6, index)), "", unclaimedData(6, index))
        .txtBillingPoints.Value = IIf(IsNull(unclaimedData(7, index)), "", unclaimedData(7, index))
        .txtRemarks.Value = IIf(IsNull(unclaimedData(8, index)), "", unclaimedData(8, index))
    End With
End Sub

Private Sub UserForm_Initialize()
    ' フォームの初期化処理
    With Me
        .Width = 400
        .Height = 350
        
        ' D列: 患者氏名
        With .Controls.Add("Forms.Label.1", "lblPatientName")
            .Caption = "患者氏名："
            .Left = 20
            .Top = 20
            .Width = 80
            .Height = 20
        End With
        
        With .Controls.Add("Forms.TextBox.1", "txtPatientName")
            .Left = 100
            .Top = 20
            .Width = 200
            .Height = 20
        End With
        
        ' E列: 調剤年月
        With .Controls.Add("Forms.Label.1", "lblDispensingDate")
            .Caption = "調剤年月："
            .Left = 20
            .Top = 50
            .Width = 80
            .Height = 20
        End With
        
        With .Controls.Add("Forms.TextBox.1", "txtDispensingDate")
            .Left = 100
            .Top = 50
            .Width = 100
            .Height = 20
            If m_era_year > 0 And m_month > 0 Then
                .Value = "令和" & m_era_year & "年" & m_month & "月"
            End If
            .Enabled = False  ' 編集不可
        End With
        
        ' F列: 医療機関
        With .Controls.Add("Forms.Label.1", "lblMedicalInstitution")
            .Caption = "医療機関："
            .Left = 20
            .Top = 80
            .Width = 80
            .Height = 20
        End With
        
        With .Controls.Add("Forms.TextBox.1", "txtMedicalInstitution")
            .Left = 100
            .Top = 80
            .Width = 200
            .Height = 20
        End With
        
        ' G列: 未請求理由
        With .Controls.Add("Forms.Label.1", "lblUnclaimedReason")
            .Caption = "未請求理由："
            .Left = 20
            .Top = 110
            .Width = 80
            .Height = 20
        End With
        
        With .Controls.Add("Forms.TextBox.1", "txtUnclaimedReason")
            .Left = 100
            .Top = 110
            .Width = 200
            .Height = 20
        End With
        
        ' H列: 請求先（ラジオボタンに変更）
        With .Controls.Add("Forms.Label.1", "lblBillingDestination")
            .Caption = "請求先："
            .Left = 20
            .Top = 140
            .Width = 80
            .Height = 20
        End With
        
        ' フレームを追加してラジオボタンをグループ化
        With .Controls.Add("Forms.Frame.1", "frameBillingDestination")
            .Caption = ""
            .Left = 100
            .Top = 135
            .Width = 200
            .Height = 30
            
            ' 社保ラジオボタン
            With .Controls.Add("Forms.OptionButton.1", "optShaho")
                .Caption = "社保"
                .Left = 10
                .Top = 5
                .Width = 60
                .Height = 20
                .Value = True  ' デフォルトで社保を選択
            End With
            
            ' 国保ラジオボタン
            With .Controls.Add("Forms.OptionButton.1", "optKokuho")
                .Caption = "国保"
                .Left = 80
                .Top = 5
                .Width = 60
                .Height = 20
            End With
        End With
        
        ' I列: 保険割合
        With .Controls.Add("Forms.Label.1", "lblInsuranceRatio")
            .Caption = "保険割合："
            .Left = 20
            .Top = 170
            .Width = 80
            .Height = 20
        End With
        
        With .Controls.Add("Forms.TextBox.1", "txtInsuranceRatio")
            .Left = 100
            .Top = 170
            .Width = 40
            .Height = 20
        End With
        
        With .Controls.Add("Forms.Label.1", "lblPercentage")
            .Caption = "割"
            .Left = 145
            .Top = 170
            .Width = 20
            .Height = 20
        End With
        
        ' J列: 請求点数
        With .Controls.Add("Forms.Label.1", "lblBillingPoints")
            .Caption = "請求点数："
            .Left = 20
            .Top = 200
            .Width = 80
            .Height = 20
        End With
        
        With .Controls.Add("Forms.TextBox.1", "txtBillingPoints")
            .Left = 100
            .Top = 200
            .Width = 100
            .Height = 20
        End With
        
        ' M列: 備考
        With .Controls.Add("Forms.Label.1", "lblRemarks")
            .Caption = "備考："
            .Left = 20
            .Top = 230
            .Width = 80
            .Height = 20
        End With
        
        With .Controls.Add("Forms.TextBox.1", "txtRemarks")
            .Left = 100
            .Top = 230
            .Width = 200
            .Height = 40
            .MultiLine = True
            .ScrollBars = fmScrollBarsVertical
        End With
        
        ' ボタンの配置を変更
        With .Controls.Add("Forms.CommandButton.1", "btnPrevious")
            .Caption = "前へ戻る"
            .Left = 20
            .Top = 280
            .Width = 80
            .Height = 30
            .Enabled = False  ' 初期状態では無効
        End With
        
        With .Controls.Add("Forms.CommandButton.1", "btnNext")
            .Caption = "次へ進む"
            .Left = 110
            .Top = 280
            .Width = 80
            .Height = 30
        End With
        
        With .Controls.Add("Forms.CommandButton.1", "btnComplete")
            .Caption = "完了"
            .Left = 200
            .Top = 280
            .Width = 80
            .Height = 30
        End With
        
        With .Controls.Add("Forms.CommandButton.1", "btnCancel")
            .Caption = "キャンセル"
            .Left = 290
            .Top = 280
            .Width = 80
            .Height = 30
        End With
    End With
End Sub

Private Sub ClearForm()
    ' フォームの入力値をクリア
    txtPatientName.Value = ""
    txtMedicalInstitution.Value = ""
    txtUnclaimedReason.Value = ""
    frameBillingDestination.Controls("optShaho").Value = True
    txtInsuranceRatio.Value = ""
    txtBillingPoints.Value = ""
    txtRemarks.Value = ""
    
    ' フォーカスを患者氏名に設定
    txtPatientName.SetFocus
End Sub

Private Function ValidateAndSaveData() As Boolean
    ' 患者名のみ必須チェック
    If Trim(Me.txtPatientName.Value) = "" Then
        MsgBox "患者名は必須項目です。", vbExclamation
        Me.txtPatientName.SetFocus
        Exit Function
    End If
    
    ValidateAndSaveData = True
End Function

Private Sub btnNext_Click()
    If ValidateAndSaveData() Then
        m_DialogResult = True
        m_ContinueInput = True
        Me.Hide
    End If
End Sub

Private Sub btnPrevious_Click()
    m_DialogResult = True
    m_ContinueInput = True
    m_CurrentIndex = m_CurrentIndex - 1  ' 前のインデックスに戻る
    Me.Hide
End Sub

Private Sub btnComplete_Click()
    If ValidateAndSaveData() Then
        m_DialogResult = True
        m_ContinueInput = False  ' 完了を示すフラグ
        Me.Hide
    End If
End Sub

Private Sub btnCancel_Click()
    m_DialogResult = False
    m_ContinueInput = False
    Me.Hide
End Sub

Private Sub txtBillingPoints_Change()
    ' 数値のみ許可（全角数字を半角に変換）
    Dim cursorPos As Long
    cursorPos = txtBillingPoints.SelStart
    txtBillingPoints.Text = UtilityModule.ConvertToHankaku(txtBillingPoints.Text)
    txtBillingPoints.SelStart = cursorPos
End Sub

Private Sub txtInsuranceRatio_Change()
    ' 数値のみ許可（全角数字を半角に変換）
    Dim cursorPos As Long
    cursorPos = txtInsuranceRatio.SelStart
    txtInsuranceRatio.Text = UtilityModule.ConvertToHankaku(txtInsuranceRatio.Text)
    txtInsuranceRatio.SelStart = cursorPos
End Sub

Private Sub btnRegister_Click()
    ' 患者氏名は必須
    If Trim(Me.txtPatientName.Value) = "" Then
        MsgBox "患者氏名を入力してください。", vbExclamation
        Exit Sub
    End If
    
    ' 医療機関は必須
    If Trim(Me.txtMedicalInstitution.Value) = "" Then
        MsgBox "医療機関を入力してください。", vbExclamation
        Exit Sub
    End If
    
    ' 未請求理由は必須
    If Trim(Me.txtUnclaimedReason.Value) = "" Then
        MsgBox "未請求理由を入力してください。", vbExclamation
        Exit Sub
    End If
    
    ' 請求先は必須（ラジオボタンなのでnullにはならない）
    
    ' 保険割合の検証（入力がある場合のみ）
    If Trim(Me.txtInsuranceRatio.Value) <> "" Then
        If Not IsNumeric(Me.txtInsuranceRatio.Value) Then
            MsgBox "保険割合は数値で入力してください。", vbExclamation
            Exit Sub
        End If
        
        Dim ratio As Integer
        ratio = CInt(Me.txtInsuranceRatio.Value)
        If ratio < 1 Or ratio > 10 Then
            MsgBox "保険割合は1割から10割の間で入力してください。", vbExclamation
            Exit Sub
        End If
        m_InsuranceRatio = ratio
    Else
        m_InsuranceRatio = Null
    End If
    
    ' 請求点数の検証（入力がある場合のみ）
    If Trim(Me.txtBillingPoints.Value) <> "" Then
        If Not IsNumeric(Me.txtBillingPoints.Value) Then
            MsgBox "請求点数は数値で入力してください。", vbExclamation
            Exit Sub
        End If
        m_BillingPoints = CLng(Me.txtBillingPoints.Value)
    Else
        m_BillingPoints = Null
    End If
    
    ' 値を保存
    m_PatientName = Trim(Me.txtPatientName.Value)
    m_DispensingDate = Trim(Me.txtDispensingDate.Value)
    m_MedicalInstitution = Trim(Me.txtMedicalInstitution.Value)
    m_UnclaimedReason = Trim(Me.txtUnclaimedReason.Value)
    m_BillingDestination = IIf(Me.frameBillingDestination.Controls("optShaho").Value, "社保", "国保")
    m_Remarks = IIf(Trim(Me.txtRemarks.Value) = "", Null, Trim(Me.txtRemarks.Value))
    
    m_DialogResult = True
    Me.Hide
End Sub 