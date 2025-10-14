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

Function ExportPivotTableDataToNewWorkbook()
    Dim ptSheet As Worksheet
    Dim pt As PivotTable
    Dim dataRange As Range
    Dim exportSheet As Worksheet
    Dim newWb As Workbook

    ' Prompt for the new sheet name
    sheetName = InputBox("Enter a name for the export sheet:", "Export Sheet Name")
    If sheetName = "" Then
        MsgBox "Export cancelled. No sheet name provided.", vbExclamation
        Exit Function
    End If

    ' Set your PivotTable sheet and PivotTable name
    Set ptSheet = ThisWorkbook.Sheets("Pivot") ' Change to your sheet name
    Set pt = ptSheet.PivotTables("AutoPivot")     ' Change to your PivotTable name

    ' Get the visible data range of the PivotTable
    Set dataRange = pt.TableRange1

    ' Create a new sheet in the current workbook
    Set exportSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    On Error Resume Next
    exportSheet.Name = sheetName
    If Err.Number <> 0 Then
        MsgBox "Invalid or duplicate sheet name. Please try again.", vbCritical
        ExportPivotTableDataToNewWorkbook = ""
        Exit Function
    End If
    On Error GoTo 0

    ' Copy the PivotTable data to the new sheet
    dataRange.Copy Destination:=exportSheet.Range("A1")

    MsgBox "PivotTable data exported successfully to new tab" & sheetName
    ExportPivotTableDataToNewWorkbook = sheetName
End Function

Sub CreateChartOnSheet(sheetName As String)
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim lastRow As Long
    Dim chartRangeX As Range, chartRangeY As Range

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Find the last row with data in column A (assumes both columns have same number of rows)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Define the ranges for X and Y values
    Set chartRangeX = ws.Range("A2:A" & lastRow)
    Set chartRangeY = ws.Range("B2:B" & lastRow)

    ' Add a chart object to the sheet
    Set chartObj = ws.ChartObjects.Add(Left:=300, Width:=500, Top:=50, Height:=300)

    With chartObj.Chart
        .ChartType = xlLine
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "=""Production Over Time"""
            .XValues = chartRangeX
            .Values = chartRangeY
        End With
        .HasTitle = True
        .ChartTitle.Text = "Production Over Time"
    End With
End Sub

Sub RunExportAndChart()
    Dim exportSheetName As String

    exportSheetName = ExportPivotTableDataToNewWorkbook
    If exportSheetName <> "" Then
        Call CreateChartOnSheet(exportSheetName)
    End If
End Sub


Sub MoveSheetsExceptSpecified()
    Dim ws As Worksheet
    Dim excludeSheets As Variant
    Dim newWbName As String
    Dim tempWb As Workbook

    ' Prompt for the new workbook name
    newWbName = InputBox("Enter a name for the new workbook (without extension):", "New Workbook Name")
    If newWbName = "" Then
        MsgBox "Operation cancelled. No workbook name provided.", vbExclamation
        Exit Sub
    End If

    ' List of sheet names to exclude from moving
    excludeSheets = Array("Gantt", "Distributions", "Export", "Dashboard", "Pivot", "Calendars", "Holidays", "ProductionChart", "Plot", "Distribution Tables")

    ' Add a new workbook to hold the moved sheets
    Set tempWb = Workbooks.Add(xlWBATWorksheet) ' Starts with one sheet

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Loop through sheets and copy all except the excluded ones, then delete them from the original workbook.
    For Each ws In ThisWorkbook.Worksheets
        If Not IsInArray(ws.Name, excludeSheets) Then
            ws.Copy After:=tempWb.Sheets(tempWb.Sheets.Count)
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    
    ' Delete the default sheet
    Application.DisplayAlerts = False
    tempWb.Sheets(1).Delete
    Application.DisplayAlerts = True

    ' Save the new workbook
    tempWb.SaveAs fileName:=ThisWorkbook.Path & "\" & newWbName & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    ' tempWb.Close SaveChanges:=False

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub

Function IsInArray(valToCheck As String, arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If element = valToCheck Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function


Function ExportSheetsExist() As Collection
    Dim ws As Worksheet
    Dim excludeSheets As Variant
    Dim exportableSheets As New Collection
    
    excludeSheets = Array("Gantt", "Distributions", "Export", "Dashboard", "Pivot", "Calendars", "Holidays", "ProductionChart", "Plot", "Distribution Tables")
    
    For Each ws In ThisWorkbook.Worksheets
        If Not IsInArray(ws.Name, excludeSheets) Then
            exportableSheets.Add ws.Name
        End If
    Next ws
    
    Set ExportSheetsExist = exportableSheets
End Function