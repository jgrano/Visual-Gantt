Attribute VB_Name = "GanttChart"
'===============================================================================
' Visual Gantt - Excel VBA Timeline Generator
' Creates professional Gantt chart with draggable/resizable shapes
'===============================================================================

Option Explicit

' Constants
Private Const DATA_SHEET_NAME As String = "Project Data"
Private Const TIMELINE_SHEET_NAME As String = "Timeline View"
Private Const HEADER_ROWS As Integer = 3
Private Const LABEL_COLS As Integer = 3
Private Const WEEK_COL_WIDTH As Double = 80
Private Const TASK_ROW_HEIGHT As Double = 40
Private Const BAR_HEIGHT As Double = 25
Private Const BAR_TOP_MARGIN As Double = 7.5

' Column indices in Project Data (1-based)
Private Const COL_TASK_ID As Integer = 1
Private Const COL_TASK_NAME As Integer = 2
Private Const COL_START_DATE As Integer = 3
Private Const COL_END_DATE As Integer = 4
Private Const COL_OWNER As Integer = 6
Private Const COL_PERCENT As Integer = 7
Private Const COL_PROJECT As Integer = 14
Private Const COL_ORIG_END As Integer = 16

'===============================================================================
' Main Entry Point - Call this to generate the timeline
'===============================================================================
Public Sub GenerateTimeline()
    Dim wsData As Worksheet
    Dim wsTimeline As Worksheet
    Dim startDate As Date
    Dim endDate As Date
    Dim totalWeeks As Integer
    Dim lastRow As Long
    Dim barsCreated As Integer

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Get or create sheets
    Set wsData = GetDataSheet()
    If wsData Is Nothing Then
        MsgBox "Could not find or create '" & DATA_SHEET_NAME & "' sheet.", vbExclamation
        GoTo Cleanup
    End If

    Set wsTimeline = GetOrCreateTimelineSheet()

    ' Get date range (default: -1 month to +2 months from today)
    startDate = DateSerial(Year(Date), Month(Date) - 1, 1)
    endDate = DateSerial(Year(Date), Month(Date) + 3, 0)

    ' Calculate weeks
    totalWeeks = Int((endDate - startDate) / 7) + 1

    ' Clear existing content and shapes
    wsTimeline.Cells.Clear
    ClearAllShapes wsTimeline

    ' Setup layout
    SetupTimelineLayout wsTimeline, totalWeeks

    ' Build headers
    BuildHeaders wsTimeline, startDate, totalWeeks

    ' Get task count
    lastRow = wsData.Cells(wsData.Rows.Count, COL_TASK_NAME).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No tasks found in '" & DATA_SHEET_NAME & "' sheet.", vbInformation
        GoTo Cleanup
    End If

    ' Build task rows and bars
    barsCreated = BuildTaskBars(wsData, wsTimeline, startDate, endDate, totalWeeks, lastRow)

    ' Add today marker
    AddTodayMarker wsTimeline, startDate, totalWeeks, lastRow - 1

    ' Freeze panes
    wsTimeline.Activate
    wsTimeline.Range("D4").Select
    ActiveWindow.FreezePanes = True

    ' Final message
    MsgBox "Timeline Generated!" & vbCrLf & vbCrLf & _
           "Task bars created: " & barsCreated & " of " & (lastRow - 1) & vbCrLf & vbCrLf & _
           "- Click bars to select" & vbCrLf & _
           "- Drag to move" & vbCrLf & _
           "- Drag corners to resize", vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

'===============================================================================
' Get or create the data sheet
'===============================================================================
Private Function GetDataSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Create with headers
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = DATA_SHEET_NAME

        ' Add headers
        ws.Range("A1").Value = "Task ID"
        ws.Range("B1").Value = "Task Name"
        ws.Range("C1").Value = "Start Date"
        ws.Range("D1").Value = "End Date"
        ws.Range("E1").Value = "Duration (Days)"
        ws.Range("F1").Value = "Owner"
        ws.Range("G1").Value = "% Complete"
        ws.Range("H1").Value = "Priority"
        ws.Range("I1").Value = "Jira Ticket"
        ws.Range("J1").Value = "Parent Task"
        ws.Range("K1").Value = "Dependencies"
        ws.Range("L1").Value = "Task Type"
        ws.Range("M1").Value = "Category"
        ws.Range("N1").Value = "Project"
        ws.Range("O1").Value = "Original Start"
        ws.Range("P1").Value = "Original End"
        ws.Range("Q1").Value = "Slip"

        ws.Range("A1:Q1").Font.Bold = True
    End If

    Set GetDataSheet = ws
End Function

'===============================================================================
' Get or create timeline sheet
'===============================================================================
Private Function GetOrCreateTimelineSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TIMELINE_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(1))
        ws.Name = TIMELINE_SHEET_NAME
    End If

    Set GetOrCreateTimelineSheet = ws
End Function

'===============================================================================
' Clear all shapes from sheet
'===============================================================================
Private Sub ClearAllShapes(ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
End Sub

'===============================================================================
' Setup timeline layout (columns and rows)
'===============================================================================
Private Sub SetupTimelineLayout(ws As Worksheet, totalWeeks As Integer)
    Dim i As Integer

    ' Label columns
    ws.Columns(1).ColumnWidth = 12  ' Project
    ws.Columns(2).ColumnWidth = 30  ' Task Name
    ws.Columns(3).ColumnWidth = 18  ' Owner

    ' Week columns
    For i = 1 To totalWeeks
        ws.Columns(LABEL_COLS + i).ColumnWidth = WEEK_COL_WIDTH / 7  ' Convert pixels to approx column width
    Next i

    ' Header row heights
    ws.Rows(1).RowHeight = 30  ' Title
    ws.Rows(2).RowHeight = 20  ' Month
    ws.Rows(3).RowHeight = 20  ' Week
End Sub

'===============================================================================
' Build headers (title, months, weeks)
'===============================================================================
Private Sub BuildHeaders(ws As Worksheet, startDate As Date, totalWeeks As Integer)
    Dim i As Integer
    Dim weekStart As Date
    Dim currentMonth As String
    Dim monthStartCol As Integer
    Dim prevMonth As String

    ' Title
    ws.Cells(1, 1).Value = "Project Timeline"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True

    ws.Cells(1, 2).Value = Format(startDate, "mmm d, yyyy") & " - " & Format(DateAdd("ww", totalWeeks, startDate), "mmm d, yyyy")
    ws.Cells(1, 2).Font.Color = RGB(102, 102, 102)

    ' Column headers
    ws.Cells(2, 1).Value = "Project"
    ws.Cells(2, 2).Value = "Task"
    ws.Cells(2, 3).Value = "Owner"
    ws.Range("A2:C3").Font.Bold = True
    ws.Range("A2:C3").Interior.Color = RGB(248, 249, 250)

    ' Week headers and month tracking
    prevMonth = ""
    monthStartCol = LABEL_COLS + 1

    For i = 0 To totalWeeks - 1
        weekStart = DateAdd("ww", i, startDate)
        currentMonth = Format(weekStart, "mmm yyyy")

        ' Week number
        ws.Cells(3, LABEL_COLS + 1 + i).Value = "W" & Format(weekStart, "ww")
        ws.Cells(3, LABEL_COLS + 1 + i).HorizontalAlignment = xlCenter
        ws.Cells(3, LABEL_COLS + 1 + i).Font.Size = 9
        ws.Cells(3, LABEL_COLS + 1 + i).Interior.Color = RGB(232, 240, 254)

        ' Track month changes
        If currentMonth <> prevMonth Then
            If prevMonth <> "" Then
                ' Merge previous month cells
                If LABEL_COLS + i > monthStartCol Then
                    ws.Range(ws.Cells(2, monthStartCol), ws.Cells(2, LABEL_COLS + i)).Merge
                End If
                ws.Cells(2, monthStartCol).Value = prevMonth
                ws.Cells(2, monthStartCol).HorizontalAlignment = xlCenter
                ws.Cells(2, monthStartCol).Interior.Color = RGB(66, 133, 244)
                ws.Cells(2, monthStartCol).Font.Color = RGB(255, 255, 255)
                ws.Cells(2, monthStartCol).Font.Bold = True
            End If
            monthStartCol = LABEL_COLS + 1 + i
            prevMonth = currentMonth
        End If
    Next i

    ' Last month
    If LABEL_COLS + totalWeeks >= monthStartCol Then
        ws.Range(ws.Cells(2, monthStartCol), ws.Cells(2, LABEL_COLS + totalWeeks)).Merge
    End If
    ws.Cells(2, monthStartCol).Value = prevMonth
    ws.Cells(2, monthStartCol).HorizontalAlignment = xlCenter
    ws.Cells(2, monthStartCol).Interior.Color = RGB(66, 133, 244)
    ws.Cells(2, monthStartCol).Font.Color = RGB(255, 255, 255)
    ws.Cells(2, monthStartCol).Font.Bold = True

    ' Add borders
    ws.Range(ws.Cells(2, 1), ws.Cells(3, LABEL_COLS + totalWeeks)).Borders.LineStyle = xlContinuous
    ws.Range(ws.Cells(2, 1), ws.Cells(3, LABEL_COLS + totalWeeks)).Borders.Color = RGB(200, 200, 200)
End Sub

'===============================================================================
' Build task rows and create bar shapes
'===============================================================================
Private Function BuildTaskBars(wsData As Worksheet, wsTimeline As Worksheet, _
                                startDate As Date, endDate As Date, _
                                totalWeeks As Integer, lastRow As Long) As Integer
    Dim i As Long
    Dim row As Integer
    Dim taskName As String
    Dim taskStart As Variant
    Dim taskEnd As Variant
    Dim origEnd As Variant
    Dim owner As String
    Dim project As String
    Dim percentComplete As Double
    Dim barsCreated As Integer
    Dim barColor As Long
    Dim startWeek As Double
    Dim endWeek As Double
    Dim clippedStart As Double
    Dim clippedEnd As Double
    Dim barLeft As Double
    Dim barWidth As Double
    Dim barTop As Double
    Dim shp As Shape
    Dim hasSlip As Boolean

    barsCreated = 0

    For i = 2 To lastRow
        row = HEADER_ROWS + (i - 1)

        ' Set row height
        wsTimeline.Rows(row).RowHeight = TASK_ROW_HEIGHT

        ' Get task data
        taskName = wsData.Cells(i, COL_TASK_NAME).Value
        project = wsData.Cells(i, COL_PROJECT).Value
        owner = wsData.Cells(i, COL_OWNER).Value
        taskStart = wsData.Cells(i, COL_START_DATE).Value
        taskEnd = wsData.Cells(i, COL_END_DATE).Value
        origEnd = wsData.Cells(i, COL_ORIG_END).Value

        On Error Resume Next
        percentComplete = wsData.Cells(i, COL_PERCENT).Value
        If Err.Number <> 0 Then percentComplete = 0
        On Error GoTo 0

        ' Write labels
        wsTimeline.Cells(row, 1).Value = project
        wsTimeline.Cells(row, 1).Font.Size = 9

        wsTimeline.Cells(row, 2).Value = taskName
        wsTimeline.Cells(row, 2).Font.Bold = True
        wsTimeline.Cells(row, 2).Font.Size = 10

        wsTimeline.Cells(row, 3).Value = owner
        wsTimeline.Cells(row, 3).Font.Color = RGB(102, 102, 102)
        wsTimeline.Cells(row, 3).Font.Size = 9

        ' Alternate row background
        If (i Mod 2) = 0 Then
            wsTimeline.Range(wsTimeline.Cells(row, 1), wsTimeline.Cells(row, LABEL_COLS + totalWeeks)).Interior.Color = RGB(250, 250, 250)
        End If

        ' Skip if no dates
        If IsEmpty(taskStart) Or IsEmpty(taskEnd) Then GoTo NextTask
        If Not IsDate(taskStart) Or Not IsDate(taskEnd) Then GoTo NextTask

        ' Calculate week positions
        startWeek = (CDate(taskStart) - startDate) / 7
        endWeek = (CDate(taskEnd) - startDate) / 7

        ' Check if completely outside range
        If endWeek <= 0 Or startWeek >= totalWeeks Then GoTo NextTask

        ' Clip to visible range
        clippedStart = Application.Max(0, startWeek)
        clippedEnd = Application.Min(totalWeeks, endWeek)

        If clippedEnd <= clippedStart Then GoTo NextTask

        ' Calculate bar position in pixels
        barLeft = GetColumnLeft(wsTimeline, LABEL_COLS + 1) + (clippedStart * GetColumnWidth(wsTimeline, LABEL_COLS + 1))
        barWidth = (clippedEnd - clippedStart) * GetColumnWidth(wsTimeline, LABEL_COLS + 1)
        barTop = GetRowTop(wsTimeline, row) + BAR_TOP_MARGIN

        ' Minimum bar width
        If barWidth < 20 Then barWidth = 20

        ' Choose color based on project or default
        barColor = GetProjectColor(project)

        ' Check for slip
        hasSlip = False
        If Not IsEmpty(origEnd) And IsDate(origEnd) Then
            If CDate(taskEnd) > CDate(origEnd) Then
                hasSlip = True
            End If
        End If

        ' Create main bar shape
        Set shp = wsTimeline.Shapes.AddShape(msoShapeRoundedRectangle, barLeft, barTop, barWidth, BAR_HEIGHT)
        With shp
            .Fill.ForeColor.RGB = barColor
            .Line.ForeColor.RGB = DarkenColor(barColor, 0.2)
            .Line.Weight = 1

            ' Add task name as tooltip
            .AlternativeText = taskName & vbCrLf & _
                              Format(taskStart, "mmm d") & " - " & Format(taskEnd, "mmm d") & vbCrLf & _
                              percentComplete & "% complete"

            ' Show percentage if between 0 and 100
            If percentComplete > 0 And percentComplete < 100 Then
                .TextFrame2.TextRange.Text = percentComplete & "%"
                .TextFrame2.TextRange.Font.Size = 8
                .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .TextFrame2.TextRange.Font.Bold = msoTrue
                .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                .TextFrame2.VerticalAnchor = msoAnchorMiddle
            End If
        End With

        ' Add slip indicator if applicable
        If hasSlip Then
            Dim slipStart As Double
            Dim slipWidth As Double
            Dim slipLeft As Double
            Dim slipShp As Shape

            slipStart = (CDate(origEnd) - startDate) / 7
            If slipStart < totalWeeks And slipStart >= 0 Then
                slipLeft = GetColumnLeft(wsTimeline, LABEL_COLS + 1) + (Application.Max(0, slipStart) * GetColumnWidth(wsTimeline, LABEL_COLS + 1))
                slipWidth = barLeft + barWidth - slipLeft

                If slipWidth > 5 Then
                    Set slipShp = wsTimeline.Shapes.AddShape(msoShapeRoundedRectangle, slipLeft, barTop, slipWidth, BAR_HEIGHT)
                    With slipShp
                        .Fill.ForeColor.RGB = LightenColor(barColor, 0.5)
                        .Line.ForeColor.RGB = barColor
                        .Line.Weight = 2
                        .Line.DashStyle = msoLineDash
                    End With
                End If
            End If
        End If

        barsCreated = barsCreated + 1

NextTask:
    Next i

    BuildTaskBars = barsCreated
End Function

'===============================================================================
' Add today marker
'===============================================================================
Private Sub AddTodayMarker(ws As Worksheet, startDate As Date, totalWeeks As Integer, taskCount As Long)
    Dim todayWeek As Double
    Dim markerLeft As Double
    Dim markerTop As Double
    Dim markerHeight As Double
    Dim shp As Shape

    todayWeek = (Date - startDate) / 7

    If todayWeek >= 0 And todayWeek < totalWeeks Then
        markerLeft = GetColumnLeft(ws, LABEL_COLS + 1) + (todayWeek * GetColumnWidth(ws, LABEL_COLS + 1))
        markerTop = GetRowTop(ws, HEADER_ROWS + 1)
        markerHeight = taskCount * TASK_ROW_HEIGHT

        Set shp = ws.Shapes.AddShape(msoShapeRectangle, markerLeft, markerTop, 3, markerHeight)
        With shp
            .Fill.ForeColor.RGB = RGB(255, 87, 34)
            .Line.Visible = msoFalse
            .AlternativeText = "Today: " & Format(Date, "mmm d, yyyy")
        End With

        ' Add label
        ws.Cells(3, LABEL_COLS + 1 + Int(todayWeek)).Value = Chr(9660)  ' Down arrow
        ws.Cells(3, LABEL_COLS + 1 + Int(todayWeek)).Font.Color = RGB(255, 87, 34)
    End If
End Sub

'===============================================================================
' Helper: Get column left position in points
'===============================================================================
Private Function GetColumnLeft(ws As Worksheet, col As Integer) As Double
    Dim i As Integer
    Dim left As Double
    left = 0
    For i = 1 To col - 1
        left = left + ws.Columns(i).Width
    Next i
    GetColumnLeft = left
End Function

'===============================================================================
' Helper: Get column width in points
'===============================================================================
Private Function GetColumnWidth(ws As Worksheet, col As Integer) As Double
    GetColumnWidth = ws.Columns(col).Width
End Function

'===============================================================================
' Helper: Get row top position in points
'===============================================================================
Private Function GetRowTop(ws As Worksheet, row As Integer) As Double
    Dim i As Integer
    Dim top As Double
    top = 0
    For i = 1 To row - 1
        top = top + ws.Rows(i).Height
    Next i
    GetRowTop = top
End Function

'===============================================================================
' Helper: Get color for project
'===============================================================================
Private Function GetProjectColor(project As String) As Long
    Select Case LCase(Trim(project))
        Case "c1", "c1, c2", "all products"
            GetProjectColor = RGB(66, 133, 244)   ' Blue
        Case "c2"
            GetProjectColor = RGB(52, 168, 83)    ' Green
        Case "v2g ac"
            GetProjectColor = RGB(251, 188, 5)    ' Yellow/Orange
        Case "vue"
            GetProjectColor = RGB(234, 67, 53)    ' Red
        Case "cm transfer"
            GetProjectColor = RGB(156, 39, 176)   ' Purple
        Case Else
            GetProjectColor = RGB(74, 101, 114)   ' Default gray-blue
    End Select
End Function

'===============================================================================
' Helper: Darken a color
'===============================================================================
Private Function DarkenColor(baseColor As Long, amount As Double) As Long
    Dim r As Integer, g As Integer, b As Integer
    r = baseColor Mod 256
    g = (baseColor \ 256) Mod 256
    b = (baseColor \ 65536) Mod 256

    r = Int(r * (1 - amount))
    g = Int(g * (1 - amount))
    b = Int(b * (1 - amount))

    DarkenColor = RGB(r, g, b)
End Function

'===============================================================================
' Helper: Lighten a color
'===============================================================================
Private Function LightenColor(baseColor As Long, amount As Double) As Long
    Dim r As Integer, g As Integer, b As Integer
    r = baseColor Mod 256
    g = (baseColor \ 256) Mod 256
    b = (baseColor \ 65536) Mod 256

    r = Int(r + (255 - r) * amount)
    g = Int(g + (255 - g) * amount)
    b = Int(b + (255 - b) * amount)

    LightenColor = RGB(r, g, b)
End Function
