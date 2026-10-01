Option Explicit

Sub Main()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    BuildMonthCalendar wb
End Sub

Sub BuildMonthCalendar(targetWb As Workbook)
    Dim ws As Worksheet
    Dim y As Long
    Dim m As Long
    Dim firstDay As Date
    Dim daysInMonth As Long
    Dim startCol As Long
    Dim headerArr As Variant
    Dim i As Long
    Dim pos As Long
    Dim r As Long
    Dim c As Long
    Dim lastRow As Long

    If targetWb Is Nothing Then Exit Sub

    Set ws = targetWb.Worksheets(1)

    y = Year(Date)
    m = Month(Date)
    firstDay = DateSerial(y, m, 1)
    daysInMonth = Day(DateSerial(y, m + 1, 0))
    startCol = Weekday(firstDay, vbSunday)
    headerArr = Array("日", "一", "二", "三", "四", "五", "六")

    ws.Cells.Clear

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 7))
        .Merge
        .Value = y & "年" & m & "月"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 16
    End With

    ws.Rows(1).RowHeight = 28

    For i = 0 To 6
        With ws.Cells(2, i + 1)
            .Value = headerArr(i)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
        End With
    Next i

    ws.Rows(2).RowHeight = 20

    lastRow = 3 + (startCol - 1 + daysInMonth - 1) \ 7

    For i = 1 To daysInMonth
        pos = startCol - 1 + (i - 1)
        r = 3 + pos \ 7
        c = (pos Mod 7) + 1

        With ws.Cells(r, c)
            .Value = i
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        If DateSerial(y, m, i) = Date Then
            With ws.Cells(r, c)
                .Interior.Color = RGB(255, 235, 156)
                .Font.Bold = True
            End With
        End If

        If Weekday(DateSerial(y, m, i), vbSunday) = 1 Then
            ws.Cells(r, c).Font.Color = RGB(192, 0, 0)
        ElseIf Weekday(DateSerial(y, m, i), vbSunday) = 7 Then
            ws.Cells(r, c).Font.Color = RGB(192, 0, 0)
        End If
    Next i

    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 7))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    For i = 1 To 7
        ws.Columns(i).ColumnWidth = 6
    Next i

    For r = 3 To lastRow
        ws.Rows(r).RowHeight = 22
    Next r

    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 7)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 7)).VerticalAlignment = xlCenter

    ws.Cells(1, 1).Select
End Sub

' ===== [LeeExcel 自动生成的受控调用入口包装器 - 保持模型正文源码零篡改] =====
Sub LeeHostRunner_5131_412_3bcc(targetWb As Workbook)
    targetWb.Activate
    Call Main
End Sub