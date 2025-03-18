' ファイル名から年月を取得する関数
Function ParseReportFileName(ByVal file_name As String, ByRef western_year As Integer, ByRef month As Integer) As Boolean
    Dim matches As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' 正規表現パターン: 保険請求管理報告書_(令和|平成|昭和|大正|明治)(\d{2})年(\d{2})月調剤分
    regex.Pattern = "保険請求管理報告書_(令和|平成|昭和|大正|明治)(\d{2})年(\d{2})月調剤分"
    regex.Global = False
    
    If regex.Test(file_name) Then
        Set matches = regex.Execute(file_name)
        Dim era_name As String, era_year As String, month_str As String
        
        era_name = matches(0).SubMatches(0)
        era_year = matches(0).SubMatches(1)
        month_str = matches(0).SubMatches(2)
        
        ' 元号から西暦を計算
        Select Case era_name
            Case "令和": western_year = 2018 + CInt(era_year)
            Case "平成": western_year = 1988 + CInt(era_year)
            Case "昭和": western_year = 1925 + CInt(era_year)
            Case "大正": western_year = 1911 + CInt(era_year)
            Case "明治": western_year = 1867 + CInt(era_year)
            Case Else: ParseReportFileName = False: Exit Function
        End Select
        
        month = CInt(month_str)
        ParseReportFileName = True
    Else
        ParseReportFileName = False
    End If
End Function