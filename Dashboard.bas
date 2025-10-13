Sub CreatePivotChartWithSlicerPanel()
    Dim tStart As Double
    tStart = Timer

    Dim wsData As Worksheet, wsPivot As Worksheet, wsDash As Worksheet
    Dim ptCache As PivotCache, pt As PivotTable
    Dim tblRange As Range
    Dim lastCol As Long, lastRow As Long, i As Long
    Dim chartObj As ChartObject
    Dim headerName As String
    Dim dateField As String, resourceField As String
    Dim colIndex As Long
    Dim slicerTop As Long, slicerLeft As Long
    Dim timelineCache As SlicerCache
    Dim timelineSlicer As Slicer

    ' Optimize performance
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Set up sheets
    Set wsData = ThisWorkbook.Sheets("Export")
    Set wsDash = ThisWorkbook.Sheets("Dashboard")

    ' Create or clear Pivot sheet
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Sheets("Pivot")
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Sheets.Add
        wsPivot.Name = "Pivot"
    Else
        wsPivot.Cells.Clear
    End If
    wsPivot.Visible = xlSheetHidden
    wsData.Visible = xlSheetVisible
    On Error GoTo 0

    ' Delete existing charts
    For Each chartObj In wsDash.ChartObjects
        chartObj.Delete
    Next chartObj

    ' Define table range precisely
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    Set tblRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))

    ' Create PivotCache and PivotTable
    Set ptCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=tblRange)
    Set pt = ptCache.CreatePivotTable(TableDestination:=wsPivot.Range("A3"), TableName:="AutoPivot")

    ' Find "Date" and "Resource" headers
    For colIndex = 1 To lastCol
        Select Case tblRange.Cells(1, colIndex).Value
            Case "Date": dateField = "Date"
            Case "Resource": resourceField = "Resource"
        End Select
    Next colIndex

    If dateField = "" Or resourceField = "" Then
        MsgBox "Required fields 'Date' or 'Resource' not found.", vbCritical
        GoTo Cleanup
    End If

    ' Add fields to PivotTable
    With pt.PivotFields(dateField)
        .Orientation = xlRowField
        .Position = 1
    End With

    With pt.PivotFields(resourceField)
        .Orientation = xlDataField
        .Function = xlSum
        .Calculation = xlRunningTotal
        .BaseField = dateField
    End With

    ' Create PivotChart
    Set chartObj = wsDash.ChartObjects.Add(Left:=50, Width:=700, Top:=50, Height:=450)
    chartObj.Name = "PivotChart"
    chartObj.Chart.SetSourceData pt.TableRange2
    chartObj.Chart.ChartType = xlLine
    chartObj.Chart.ChartStyle = 230
    chartObj.Chart.HasTitle = True
    chartObj.Chart.ChartTitle.Text = "Production Over Time"

    ' Slicer panel positioning
    slicerTop = 50
    slicerLeft = chartObj.Left + chartObj.Width + 50

    ' Add slicers for fields with multiple unique values
    For i = 1 To lastCol
        headerName = tblRange.Cells(1, i).Value
        If headerName <> "Date" And headerName <> "Resource" Then
            Dim uniqueCount As Long
            uniqueCount = WorksheetFunction.CountA(wsData.Range(wsData.Cells(2, i), wsData.Cells(lastRow, i)))
            If uniqueCount > 1 Then
                On Error Resume Next
                With pt.PivotFields(headerName)
                    .Orientation = xlPageField
                    .Position = pt.PageFields.Count + 1
                End With
                ActiveWorkbook.SlicerCaches.Add(pt, headerName).Slicers.Add _
                    SlicerDestination:=wsDash, Name:=headerName, Caption:=headerName, _
                    Top:=slicerTop, Left:=slicerLeft, Width:=200, Height:=60
                On Error GoTo 0
                slicerTop = slicerTop + 70
            End If
        End If
    Next i

    ' Add timeline slicer for Date
    On Error Resume Next
    Set timelineCache = ActiveWorkbook.SlicerCaches.Add2(pt, dateField, xlTimeline)
    Set timelineSlicer = timelineCache.Slicers.Add(wsDash, , dateField, "Date Timeline", slicerLeft, slicerTop, 200, 100)
    On Error GoTo 0

Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    Debug.Print "CreatePivotChartWithSlicerPanel took " & Format(Timer - tStart, "0.000") & " seconds"
End Sub

Sub PivotChartSytleChecker()
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    
    On Error Resume Next
    Set chartObj = wsDash.ChartObjects("PivotChart")
    On Error GoTo 0

    If chartObj Is Nothing Then
        MsgBox "Chart named 'PivotChart' not found.", vbExclamation
        Exit Sub
    End If
    
    Dim currentStyle As Long
    currentStyle = chartObj.Chart.ChartStyle
    Debug.Print "Current ChartStyle is: " & currentStyle
End Sub

Sub AddLabelsAndTrendlineToPivotChart(ShowDataLabels As Boolean, ShowTrendLine As Boolean)
    Dim ws As Worksheet
    Dim cht As ChartObject
    Dim srs As Series
    Dim lbl As DataLabel
    Dim trnd As trendline

    ' Set your worksheet and chart name
    Set ws = ThisWorkbook.Sheets("Dashboard") ' Change to your sheet name
    Set cht = ws.ChartObjects("PivotChart") ' Change to your chart name

    ' Loop through each series in the chart
    For Each srs In cht.Chart.SeriesCollection
        ' Remove existing data labels
        If srs.HasDataLabels Then
            srs.HasDataLabels = False
        End If

        ' Remove existing trendlines
        Do While srs.Trendlines.Count > 0
            srs.Trendlines(1).Delete
        Loop

        ' Add new data labels if requested
        If ShowDataLabels Then
            srs.HasDataLabels = True
            With srs.DataLabels
                .ShowValue = True
                .NumberFormat = "0"
                .Position = xlLabelPositionAbove
                .AutoScaleFont = False
                For Each lbl In srs.DataLabels
                    lbl.Orientation = 90
                    lbl.Font.Size = 10
                    lbl.Font.Name = "Calibri"
                Next lbl
            End With
        End If

        ' Add new trendline if requested
        If ShowTrendLine Then
            srs.Trendlines.Add Type:=xlLinear, Forward:=0, Backward:=0, DisplayEquation:=True, DisplayRSquared:=True
        End If
    Next srs
    cht.Chart.HasLegend = False

    'MsgBox "PivotChart updated with new data labels and trendlines."
End Sub

