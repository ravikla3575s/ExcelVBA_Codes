Function GetUniqueSheetName(workbook_obj As Workbook, base_name As String) As String
    Dim new_name As String, counter As Integer
    Dim ws As Worksheet, exists As Boolean
    new_name = base_name
    counter = 1
    Do
        exists = False
        For Each ws In workbook_obj.Sheets
            If LCase(ws.Name) = LCase(new_name) Then
                exists = True
                Exit For
            End If
        Next ws
        If exists Then
            new_name = base_name & "_" & counter
            counter = counter + 1
        End If
    Loop While exists
    GetUniqueSheetName = new_name
End Function