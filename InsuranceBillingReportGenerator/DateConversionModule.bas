Attribute VB_Name = "DateConversionModule"
Option Explicit

Function ConvertEraYear(ByVal western_year As Integer, Optional ByVal return_dict As Boolean = False) As Variant
    Dim era As String
    Dim era_year As Integer
    
    If western_year >= 2019 Then
        era = "�ߘa"
        era_year = western_year - 2018
    ElseIf western_year >= 1989 Then
        era = "����"
        era_year = western_year - 1988
    ElseIf western_year >= 1926 Then
        era = "���a"
        era_year = western_year - 1925
    ElseIf western_year >= 1912 Then
        era = "�吳"
        era_year = western_year - 1911
    ElseIf western_year >= 1868 Then
        era = "����"
        era_year = western_year - 1867
    Else
        era = ""
        era_year = 0
    End If
    
    If return_dict Then
        ' Dictionary �I�u�W�F�N�g��Ԃ�
        Dim result As Object
        Set result = CreateObject("Scripting.Dictionary")
        result.Add "era", era
        result.Add "year", era_year
        Set ConvertEraYear = result
    Else
        ' �����������Ԃ�
        ConvertEraYear = era
    End If
End Function

Private Function GetEraInfo(western_year As Integer, ByRef era_code As String, ByRef era_year As Integer) As Boolean
    If western_year >= 2019 Then
        era_code = "5": era_year = western_year - 2018   ' �ߘa
    ElseIf western_year >= 1989 Then
        era_code = "4": era_year = western_year - 1988   ' ����
    ElseIf western_year >= 1926 Then
        era_code = "3": era_year = western_year - 1925   ' ���a
    ElseIf western_year >= 1912 Then
        era_code = "2": era_year = western_year - 1911   ' �吳
    ElseIf western_year >= 1868 Then
        era_code = "1": era_year = western_year - 1867   ' ����
    Else
        era_code = "0": era_year = 0
        GetEraInfo = False
        Exit Function
    End If
    GetEraInfo = True
End Function

Private Function CalculateEraYear(ByVal western_year As Integer) As Integer
    If western_year >= 2019 Then
        CalculateEraYear = western_year - 2018   ' �ߘa
    ElseIf western_year >= 1989 Then
        CalculateEraYear = western_year - 1988   ' ����
    ElseIf western_year >= 1926 Then
        CalculateEraYear = western_year - 1925   ' ���a
    ElseIf western_year >= 1912 Then
        CalculateEraYear = western_year - 1911   ' �吳
    ElseIf western_year >= 1868 Then
        CalculateEraYear = western_year - 1867   ' ����
    Else
        CalculateEraYear = 0
    End If
End Function

Function ConvertToWesternDate(dispensing_code As String) As String
    Dim era_code As String, year_num As Integer, western_year As Integer, month_part As String
    If Len(dispensing_code) < 5 Then
        ConvertToWesternDate = ""
        Exit Function
    End If
    era_code = Left(dispensing_code, 1)
    year_num = CInt(Mid(dispensing_code, 2, 2))
    month_part = Right(dispensing_code, 2)
    Select Case era_code
        Case "5": western_year = 2018 + year_num   ' �ߘa
        Case "4": western_year = 1988 + year_num   ' ����
        Case "3": western_year = 1925 + year_num   ' ���a
        Case "2": western_year = 1911 + year_num   ' �吳
        Case "1": western_year = 1867 + year_num   ' ����
        Case Else: western_year = 2000 + year_num
    End Select
    ConvertToWesternDate = Right(CStr(western_year), 2) & "." & month_part
End Function
