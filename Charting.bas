
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Function ActivatePlotSheet()
    Worksheets(5).Activate
End Function

Sub SumValuesByDate(ByVal appendNew As Boolean)
    Dim dateRange As Range
    Dim resultArray() As Variant
    Dim i As Long, colIndex As Long
    Dim ws As Worksheet
    Dim targetWS As Worksheet
    Dim lastRow As Long
    Dim cumulativeSum As Double
    cumulativeSum = 0
    Dim startingCol As Long
    startingCol = ColumnIndex("GanttStart") ' Column J

    Set ws = ThisWorkbook.Sheets("Gantt")
    Set dateRange = GetLastColumnRange()

    If SheetExists("Plot") Then
        Set targetWS = ThisWorkbook.Sheets("Plot")
    Else
        MsgBox "Sheet 'Plot' does not exist!"
        Exit Sub
    End If

    ReDim resultArray(0 To dateRange.Columns.Count - 1, 1 To 2)

    ' Variables for stats
    Dim maxVal As Double, minVal As Double, totalSum As Double
    Dim MaxDate As Variant, MinDate As Variant
    maxVal = 0
    minVal = 1E+99
    totalSum = 0

    For i = 0 To dateRange.Columns.Count - 1
        colIndex = startingCol + i
        lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row

        Dim currentSum As Double
        currentSum = 0

        Dim r As Long
        For r = 2 To lastRow
            If IsNumeric(ws.Cells(r, colIndex).Value) Then
                currentSum = currentSum + ws.Cells(r, colIndex).Value
            End If
        Next r

        If currentSum > maxVal Then
            maxVal = currentSum
            MaxDate = ws.Cells(1, colIndex).Value
        End If

        If currentSum < minVal And currentSum <> 0 Then
            minVal = currentSum
            MinDate = ws.Cells(1, colIndex).Value
        End If
        
        cumulativeSum = cumulativeSum + currentSum
        totalSum = totalSum + currentSum

        resultArray(i, 1) = ws.Cells(1, colIndex).Value ' Date
        resultArray(i, 2) = cumulativeSum               ' Cumulative Sum

    Next i

    ' Output result to Plot sheet starting at A2
    Dim lastUsedCol As Long
    lastUsedCol = GetLastUsedColumnPair(targetWS)

    Dim xCol As Long, yCol As Long
    If appendNew Then
        xCol = lastUsedCol + 2
    Else
        xCol = lastUsedCol
    End If
    yCol = xCol + 1

    ' Write X-axis (dates) and Y-axis (cumulative sums)
    targetWS.Cells(2, xCol).Resize(UBound(resultArray), 1).Value = Application.Index(resultArray, 0, 1)
    targetWS.Cells(2, yCol).Resize(UBound(resultArray), 1).Value = Application.Index(resultArray, 0, 2)

    ' Optional: Add headers
    targetWS.Cells(1, xCol).Value = "Date"
    targetWS.Cells(1, yCol).Value = ws.Cells(2, ColumnIndex("Category")).Value

    ' Write directly to TextBoxes on Sheet1
    With ThisWorkbook.Sheets("Gantt")
        .OLEObjects("AveBox").Object.Text = Round(totalSum / dateRange.Columns.Count)
        .OLEObjects("MaxBox").Object.Text = Round(maxVal)
        .OLEObjects("MaxDate").Object.Text = MaxDate
        .OLEObjects("MinBox").Object.Text = Round(minVal)
        .OLEObjects("MinDate").Object.Text = MinDate
    End With
End Sub

Function GetLastUsedColumnPair(ws As Worksheet) As Long
    Dim LastCol As Long
    LastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' If row 1 is completely empty, set lastCol to 0
    If Application.WorksheetFunction.CountA(ws.Rows(1)) = 0 Then
        LastCol = 0
    End If

    ' Ensure lastCol is never negative
    If LastCol < 1 Then
        GetLastUsedColumnPair = 1 ' Or 0, depending on how you want to handle "no data"
        Exit Function
    End If

    ' Return the first column of the last pair
    If LastCol Mod 2 = 0 Then
        GetLastUsedColumnPair = LastCol - 1
    Else
        GetLastUsedColumnPair = LastCol
    End If
End Function

Sub CreateChartOnNewSheet()
    Dim chartSheet As Chart

    ' Delete existing chart sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("ProductionChart").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Add a new chart sheet
    Set chartSheet = ThisWorkbook.Charts.Add
    chartSheet.Name = "ProductionChart"
    chartSheet.ChartType = xlLine
    
    ' Move the chart sheet to the end
    chartSheet.Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

    With chartSheet
            ' Delete any auto-generated series
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Now add your custom series
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "=""Production Over Time"""
            .XValues = "='Plot'!ChartDates"
            .Values = "='Plot'!ChartSums"
        End With

    End With
    'Call FormatChart(chartSheet, PlotForm.ShowDataLables.Value, PlotForm.ShowTrendLine.Value)
    Call UpdateChartFormatting(chartSheet)
End Sub

Sub UpdateChartFormatting(chartObj As Chart)
    Dim showLabels As Boolean
    Dim showTrend As Boolean
    

    With Worksheets("Gantt")
        showLabels = .OLEObjects("ShowDataLabels").Object.Value
        showTrend = .OLEObjects("ShowTrendLine").Object.Value
        
    End With

    Call FormatChart(chartObj, showLabels, showTrend)
End Sub

Sub FormatChart(chartObj As Chart, ShowDataLabels As Boolean, ShowTrendLine As Boolean)

    Dim yVals As Variant
    yVals = chartObj.SeriesCollection(1).Values
    maxYValue = WorksheetFunction.Max(yVals)

    With chartObj
        ' === Chart Title ===
        .HasTitle = True
        .ChartTitle.Text = "Production Over Time"

        ' === Axes Titles ===
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Date"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Cumulative Sum"

        ' === Axes Scaling ===
        With .Axes(xlValue)
            .MinimumScale = 0 ' Prevent negative values
            .MaximumScale = maxYValue * 1.1 ' 10% Buffer
            
        End With
        
        ' === Gridlines ===
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).HasMinorGridlines = True
        .Axes(xlCategory).HasMajorGridlines = True
        .Axes(xlCategory).HasMinorGridlines = True

        ' === Legend ===
        .HasLegend = False 'True to show legend
        '.Legend.IncludeInLayout = False ' False to not include in layout
        '.Legend.Position = xlLegendPositionBottom

        ' === Series Styling ===
        With .SeriesCollection(1)
            .Format.Line.ForeColor.RGB = RGB(0, 112, 192) ' Blue line
            .Format.Line.Weight = 2.25
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 6
            
            
            If ShowDataLabels Then
                .ApplyDataLabels
                .DataLabels.ShowValue = True
                .DataLabels.NumberFormat = "0"
                .DataLabels.Position = xlLabelPositionAbove
                Dim lbl As DataLabel
                For Each lbl In .DataLabels
                    lbl.Orientation = 45
                Next lbl
            End If
                      
            If ShowTrendLine Then
                .Trendlines.Add Type:=xlLinear, Forward:=0, Backward:=0, DisplayEquation:=True, DisplayRSquared:=False
            End If
            
        End With

    End With
End Sub

Sub GetTrendlineStats()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Gantt")
    
    Dim xRange As Range, yRange As Range
    Set xRange = ThisWorkbook.Sheets("Plot").Range("ChartDates")
    Set yRange = ThisWorkbook.Sheets("Plot").Range("ChartSums")
    
    Dim stats As Variant
    stats = WorksheetFunction.LinEst(yRange, xRange, True, True)
    
    Dim slope As Double, intercept As Double, rsq As Double
    slope = stats(1, 1)
    intercept = stats(1, 2)
    rsq = stats(3, 1) ' Store in TextBoxes
    
    ws.OLEObjects("SlopeBox").Object.Text = Round(slope)
    ws.OLEObjects("RSqBox").Object.Text = Round(rsq, 2)
End Sub

Sub RunChartting()
    ' This subroutine will run all the necessary subroutines in order
    'Call SumValuesByDate
    Call CreateChartOnNewSheet
    'Call GetTrendlineStats
End Sub


'Author: Steve Hall P.E.
'Email: shall@austin-ind.com
'Date: 2025-09-11
'Version: 1.5

