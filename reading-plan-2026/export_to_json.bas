Sub ExportPlanToJSON()
    Dim ws As Worksheet
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets("Plan")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    Dim json As String
    json = "["
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).Value) <> "" Then
            If json <> "[" Then json = json & ","
            json = json & vbCrLf & "  {"
            json = json & """day"":" & CLng(ws.Cells(i, 1).Value) & ","
            json = json & """date"":" & """" & Replace(ws.Cells(i, 2).Text, """", ""'"" ) & """" & ","
            json = json & """topic"":" & """" & Replace(CStr(ws.Cells(i, 3).Value), """", """" ) & """" & ","
            json = json & """ot"":" & """" & Replace(CStr(ws.Cells(i, 4).Value), """", """" ) & """" & ","
            json = json & """nt"":" & """" & Replace(CStr(ws.Cells(i, 5).Value), """", """" ) & """" & ","
            json = json & """psalm"":" & """" & Replace(CStr(ws.Cells(i, 6).Value), """", """" ) & """" & ","
            json = json & """proverbs"":" & """" & Replace(CStr(ws.Cells(i, 7).Value), """", """" ) & """" & ","
            json = json & """ot_url"":" & """" & Replace(CStr(ws.Cells(i, 8).Value), """", """" ) & """" & ","
            json = json & """nt_url"":" & """" & Replace(CStr(ws.Cells(i, 9).Value), """", """" ) & """" & ","
            json = json & """psalm_url"":" & """" & Replace(CStr(ws.Cells(i, 10).Value), """", """" ) & """" & ","
            json = json & """proverbs_url"":" & """" & Replace(CStr(ws.Cells(i, 11).Value), """", """" ) & """" & ","
            json = json & """message"":" & """" & Replace(CStr(ws.Cells(i, 12).Value), """", """" ) & """" & "}"
        End If
    Next i
    json = json & vbCrLf & "]"
    Dim fname As String
    Dim yearVal As String
    On Error Resume Next
    yearVal = CStr(ThisWorkbook.Worksheets("Settings").Range("B1").Value)
    If yearVal = "" Then yearVal = Format(Date, "yyyy")
    On Error GoTo ErrHandler
    fname = ThisWorkbook.Path & Application.PathSeparator & "قراءة الكتاب المقدس كاملا في عام واحد بالموضوعات لعام " & yearVal & ".json"
    Dim fnum As Integer
    fnum = FreeFile
    Open fname For Output As #fnum
    Print #fnum, json
    Close #fnum
    MsgBox "تم التصدير إلى: " & fname, vbInformation, "Export Complete"
    Exit Sub
ErrHandler:
    MsgBox "حدث خطأ أثناء التصدير: " & Err.Description, vbExclamation, "Export Error"
End Sub
