Public ColumnIndex As Scripting.Dictionary

Sub InitializeColumnIndex()
    Dim ws As Worksheet
    Dim ColumnNames() As Variant
    Dim i As Long, lastCol As Long
    
    Set ws = ThisWorkbook.Sheets("Gantt")
    Set ColumnIndex = New Scripting.Dictionary
    
    ' Find the last column before "Error"
    i = 1
    Do While ws.Cells(1, i).Value <> "Error" And ws.Cells(1, i).Value <> ""
        i = i + 1
    Loop
    lastCol = i - 1
    
    ' Resize and populate ColumnNames array
    ReDim ColumnNames(0 To lastCol - 1)
    For i = 0 To lastCol - 1
        ColumnNames(i) = ws.Cells(1, i + 1).Value
    Next i
    
    ' Build the index
    BuildColumnIndex ColumnNames, ColumnIndex
    
    ColumnIndex.Add "GanttStart", lastCol + 2 ' Add GanttStart column index
    ' Optional: Print the dictionary to Immediate Window
    Dim key As Variant
    'For Each key In ColumnIndex.Keys
        'Debug.Print key & "-> " & ColumnIndex(key)
    'Next key
End Sub

Sub BuildColumnIndex(ColumnNames As Variant, ColumnIndex As Scripting.Dictionary)
    Dim i As Long
    ColumnIndex.RemoveAll ' Clear existing entries
    
    For i = LBound(ColumnNames) To UBound(ColumnNames)
        ColumnIndex.Add ColumnNames(i), i + 1 ' VBA is 1-based for Excel columns
    Next i
End Sub

Function ColumnLetter(colNum As Long) As String
    ColumnLetter = Split(Cells(1, colNum).Address(True, False), "$")(0)
End Function

Function ActivateGanttSheet()
    Worksheets(1).Activate
End Function

'This function will return the range of the distribution table or calendar based on the name provided. It will be used to get the correct range for the distribution values.
Function GetTableRange(rangeName As String) As Range
    Dim cell As Range
    Dim tableRange As Range
    Dim foundInDist As Boolean
    Dim foundInCal As Boolean

    foundInDist = False
    foundInCal = False

    ' Check if rangeName is in DistType
    On Error Resume Next
    For Each cell In Range("DistType")
        If StrComp(cell.Value, rangeName, vbTextCompare) = 0 Then
            foundInDist = True
            Exit For
        End If
    Next cell

    ' Check if rangeName is in CalType
    If Not foundInDist Then
        For Each cell In Range("CalType")
            If StrComp(cell.Value, rangeName, vbTextCompare) = 0 Then
                foundInCal = True
                Exit For
            End If
        Next cell
    End If
    On Error GoTo 0

    ' Assign the appropriate table range
    If foundInDist Then
        Set tableRange = ThisWorkbook.Names("DistTab").RefersToRange
        'Debug.Print "Table Range Address: " & tableRange.Address
        'Debug.Print "Sheet Name: " & tableRange.Worksheet.Name

    ElseIf foundInCal Then
        Set tableRange = ThisWorkbook.Names("Calendar").RefersToRange
        'Debug.Print "Table Range Address: " & tableRange.Address
        'Debug.Print "Sheet Name: " & tableRange.Worksheet.Name

    Else
        Set tableRange = Nothing ' Or handle error
    End If
    
    'Debug.Print "Found in DistType: " & foundInDist
    'Debug.Print "Found in CalType: " & foundInCal


    Set GetTableRange = tableRange
    'Debug.Print "Table Range " & GetTableRange.Address
    'Debug.Print "Sheet Name: " & GetTableRange.Worksheet.Name
    
End Function

Function GetColumn(Table As String) As Variant
    Dim tableRange As Range
    Dim colIndex As Long
    Dim result() As Double
    Dim i As Long
    Dim dataRange As Range

    ' Get the full table range (including Table row)
    Set tableRange = GetTableRange(Table)
    

    ' Find the column index for the Table
    For colIndex = 1 To tableRange.Columns.Count
        If tableRange.Cells(1, colIndex).Value = Table Then
            Exit For
        End If
    Next colIndex

    ' If Table not found
    If colIndex > tableRange.Columns.Count Then
        GetColumn = CVErr(xlErrNA)
        Exit Function
    End If

    ' Define the data range (excluding the Table row)
    Set dataRange = tableRange.Offset(1, 0).Resize(tableRange.Rows.Count - 1)

    ' Resize result array
    ReDim result(0 To dataRange.Rows.Count - 1)

    ' Fill result array with values from the column
    For i = 0 To dataRange.Rows.Count - 1
        result(i) = dataRange.Cells(i + 1, colIndex).Value
    Next i

    ' Return the result array
    GetColumn = result
End Function

Function SumByDistributions(Distribution As String, duration As Double) As Variant
    Dim outputArray() As Double
    Dim inputArray() As Double
    Dim totalItems As Long
    Dim baseGroupSize As Long
    Dim remainder As Long
    Dim startIndex As Long
    Dim currentGroupSize As Long
    Dim i As Long, j As Long
    Dim groupSum As Double

    ' Get the input data
    inputArray = GetColumn(Distribution)
    totalItems = UBound(inputArray) + 1
    ' Calculate base group size and remainder
    baseGroupSize = totalItems \ duration
    remainder = totalItems Mod duration

    ' Prepare output array
    ReDim outputArray(0 To duration - 1)

    startIndex = 0
    For i = 0 To duration - 1
        groupSum = 0
        currentGroupSize = baseGroupSize
        If i < remainder Then currentGroupSize = currentGroupSize + 1

        For j = 0 To currentGroupSize - 1
            groupSum = groupSum + inputArray(startIndex + j)
        Next j

        outputArray(i) = groupSum
        startIndex = startIndex + currentGroupSize
    Next i

    SumByDistributions = outputArray
End Function

'This function takes the product of Summed Distribution and resource and creates an array.
Function CreateDistributionArray(duration As Double, resource As Double, Distribution As String) As Variant
    Dim inputArray As Variant
    Dim outputArray() As Double
    Dim i As Long

    inputArray = SumByDistributions(Distribution, duration)
    ' Start with an empty array
    ReDim outputArray(0 To 0)

    For i = 0 To UBound(inputArray)
        ' Resize if needed
        If i > UBound(outputArray) Then
            ReDim Preserve outputArray(0 To i)
        End If

        ' Multiply each summed distribution value by the resource value
        outputArray(i) = inputArray(i) * resource
    Next i
    'Debug.Print "Create Dist Array" & LBound(outputArray) & "; " & UBound(outputArray)

    CreateDistributionArray = outputArray
End Function

Function AssignDistributionToDates(Start As Date, duration As Double, DistributionArray As Variant, calendar As String) As Variant
    Dim outputArray() As Double
    Dim i As Long
    Dim workingDayCount As Integer
    Dim calendarRange As Double
    Dim holidays As Range
    Dim calendarWorkDays As Variant
    Dim currentDateIndex As Integer
    Dim isWorkingDay As Boolean

    ' ReDim outputArray(0 To duration * (7 / 5))
        
    ReDim outputArray(0 To 0)
    
    currentDate = Start
    holidayRange = Range("Holidays")
    calendarWorkDays = GetColumn(calendar)

    i = 0
    Do While workingDayCount < duration
        currentDateIndex = Weekday(currentDate, vbMonday) - 1 ' 1 = Monday, 7 = Sunday
        If i > UBound(outputArray) Then
            ReDim Preserve outputArray(0 To i)
        End If
    
        ' Default value
        outputArray(i) = 0
    
        ' Check if it's a working weekday
        If calendarWorkDays(currentDateIndex) = 1 Then
            ' Check if it's a holiday
            isHoliday = Not IsError(Application.Match(currentDate, holidayRange, 0))
    
            If Not isHoliday Then
                ' It's a working day and not a holiday
                outputArray(i) = DistributionArray(workingDayCount)
                workingDayCount = workingDayCount + 1
            End If
        End If
   
        ' Move to the next date regardless
        currentDate = currentDate + 1
        i = i + 1
    Loop
    'Debug.Print "Assign Dist Dates " & LBound(outputArray) & "; " & UBound(outputArray)

    'Return the output array with assigned distribution values for each date in the duration
    AssignDistributionToDates = outputArray
End Function

'This function will return the last column range starting from column J (10), where my dates start to the last used column in the first row. Should return the same row everytime.
Public Function GetLastColumnRange() As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Gantt")

    Dim startCol As Long
    startCol = ColumnIndex("GanttStart")

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If startCol = 0 Or lastCol = 0 Then
        MsgBox "Invalid column index."
        Exit Function
    End If

    Set GetLastColumnRange = ws.Range(ws.Cells(1, startCol), ws.Cells(1, lastCol))
End Function

'This function will find the cell that contains the target date in the specified search range. We will already be on the correct row so we want to return the column value of the cell that contains the date.
Function FindDateColumn(startDate As Date) As Variant
    Dim matchIndex As Variant
    Dim searchRange As Range
    Dim lastCol As Long
    
    ' Ensure startDate is a pure date
    startDate = DateValue(startDate)
    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column

    ' Define the range (dates in row 1, columns J to DE)
    Set searchRange = Range(Cells(1, ColumnIndex("GanttStart")), Cells(1, lastCol)) ' J1 to DE1

    ' Use Match with exact match
    matchIndex = Application.Match(CLng(startDate), searchRange, 0)

    If IsError(matchIndex) Then
        FindDateColumn = CVErr(xlErrNA)
    Else
        ' Add 9 to align with your indexing logic
        FindDateColumn = ColumnIndex("GanttStart") + matchIndex - 1
    End If
End Function

Sub WriteToCells()
    Dim currentCell As Range
    Dim distArray As Variant
    Dim resultArray As Variant
    Dim startingCol As Variant

    Dim duration As Double
    Dim startDate As Date
    Dim finishDate As Date
    Dim resource As Double
    Dim distributionType As String
    Dim calendarType As String
    Dim exportRow As Long
    exportRow = 2

    Dim tStart As Double, tLoopStart As Double
    tStart = Timer

    'Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim outputSheet As Worksheet
    Set outputSheet = ThisWorkbook.Sheets("Export")

    ' Write headers once
    Dim outputKeys As Collection
    Set outputKeys = New Collection
    Dim key As Variant
    Dim startIndex As Long, endIndex As Long
    startIndex = ColumnIndex("Activity ID")
    endIndex = ColumnIndex("Duration")

    For Each key In ColumnIndex.Keys
        If ColumnIndex(key) > startIndex And ColumnIndex(key) < endIndex Then
            outputKeys.Add key
        End If
    Next key

    Dim outputActivityDescCol As Long
    Dim outputDateCol As Long
    Dim outputDataCol As Long
    outputActivityDescCol = outputKeys.Count
    outputDateCol = outputActivityDescCol + 1
    outputDataCol = outputDateCol + 1

    ' Write headers once
    For keyIndex = 1 To outputKeys.Count
        outputSheet.Cells(1, keyIndex).Value = outputKeys(keyIndex)
    Next keyIndex
    outputSheet.Cells(1, outputDateCol).Value = "Date"
    outputSheet.Cells(1, outputDataCol).Value = "Resource"

    ' Loop through rows
    For Each currentCell In Range("A2:A" & Cells(Rows.Count, 1).End(xlUp).Row)
        If Not isEmpty(currentCell.Value) Then
            tLoopStart = Timer

            Dim currentRow As Long
            currentRow = currentCell.Row

            duration = Cells(currentRow, ColumnIndex("Duration")).Value
            startDate = Cells(currentRow, ColumnIndex("Start")).Value
            finishDate = Cells(currentRow, ColumnIndex("Finish")).Value
            resource = Cells(currentRow, ColumnIndex("Resource")).Value
            distributionType = Cells(currentRow, ColumnIndex("Distribution")).Value
            calendarType = Cells(currentRow, ColumnIndex("Calendar")).Value

            distArray = CreateDistributionArray(duration, resource, distributionType)
            resultArray = AssignDistributionToDates(startDate, duration, distArray, calendarType)
            startingCol = FindDateColumn(startDate)

            ' Write resultArray to main sheet in one go
            ' Range(Cells(currentRow, startingCol), Cells(currentRow, startingCol + UBound(resultArray))).Value = Application.WorksheetFunction.Transpose(resultArray)

            ' Write to Export sheet
            Dim i As Long
            For i = LBound(resultArray) To UBound(resultArray)
                Cells(currentRow, startingCol + i).Value = resultArray(i)
                For j = 1 To outputKeys.Count
                    outputSheet.Cells(exportRow, j).Value = Cells(currentRow, ColumnIndex(outputKeys(j))).Value
                Next j
                outputSheet.Cells(exportRow, outputDateCol).Value = Cells(1, startingCol + i).Value
                outputSheet.Cells(exportRow, outputDataCol).Value = resultArray(i)
                exportRow = exportRow + 1
            Next i

            Debug.Print "Row " & currentRow & " took " & Format(Timer - tLoopStart, "0.000") & " seconds"
        End If
    Next currentCell

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    Debug.Print "Total WriteToCells time: " & Format(Timer - tStart, "0.000") & " seconds"
End Sub

Function CalculateStartAndFinish(flag As Boolean) As Boolean
    Dim i As Long
    Dim lag As Double
    Dim relationship As String
    Dim relationshipType As String
    Dim activityID As String
    Dim lastRow As Long
    Dim calculateStart As Boolean

    calculateStart = (Worksheets("Gantt").RelationBox.Value = True)
    lastRow = Cells(Rows.Count, ColumnIndex("Activity ID")).End(xlUp).Row
    CalculateStartAndFinish = True

    ' Ensure first activity has a start date
    If isEmpty(Cells(2, ColumnIndex("Start")).Value) Then
        Cells(2, ColumnIndex("Start")).Interior.Color = RGB(255, 200, 200)
        MsgBox "Invalid data for Activity ID: " & Cells(2, ColumnIndex("Activity ID")).Value & ". Initial Start date required.", vbExclamation
        CalculateStartAndFinish = False
        Exit Function
    End If

    For i = 2 To lastRow
        activityID = Cells(i, ColumnIndex("Activity ID")).Value

        ' Validate Duration
        If isEmpty(Cells(i, ColumnIndex("Duration")).Value) Then
            Cells(i, ColumnIndex("Duration")).Interior.Color = RGB(255, 200, 200)
            MsgBox "Invalid data for Activity ID: " & activityID & ". Activity duration required.", vbExclamation
            CalculateStartAndFinish = False
            Exit Function
        End If

        ' Validate Resource
        If isEmpty(Cells(i, ColumnIndex("Resource")).Value) Then
            Cells(i, ColumnIndex("Resource")).Interior.Color = RGB(255, 200, 200)
            MsgBox "Invalid data for Activity ID: " & activityID & ". Resources required.", vbExclamation
            CalculateStartAndFinish = False
            Exit Function
        End If

        If calculateStart Then
            ' Relationship Type
            If Not isEmpty(Cells(i, ColumnIndex("Rel.")).Value) Then
                relationshipType = Cells(i, ColumnIndex("Rel.")).Value
            Else
                Cells(i, ColumnIndex("Rel.")).Value = "FS"
                relationshipType = "FS"
            End If

            ' Relationship Target
            If Not isEmpty(Cells(i, ColumnIndex("w/")).Value) Then
                relationship = Cells(i, ColumnIndex("w/")).Value
            Else
                Cells(i, ColumnIndex("w/")).Value = "Prev"
                relationship = "Prev"
            End If

            ' Lag
            If Not isEmpty(Cells(i, ColumnIndex("Lag")).Value) Then
                lag = Cells(i, ColumnIndex("Lag")).Value
            Else
                Cells(i, ColumnIndex("Lag")).Value = 0
                lag = 0
            End If

            ' Calculate Start
            If i <> 2 Then
                If relationshipType = "FS" Then
                    If relationship = "Prev" Then
                        Cells(i, ColumnIndex("Start")).Value = GetDateCustom(Cells(i - 1, ColumnIndex("Finish")).Value, Cells(i, ColumnIndex("Lag")).Value, Cells(i, ColumnIndex("Calendar")).Value)
                    Else
                        Cells(i, ColumnIndex("Start")).Value = GetDateCustom(IDFinishDate(relationship), Cells(i, ColumnIndex("Lag")).Value, Cells(i, ColumnIndex("Calendar")).Value)
                    End If
                ElseIf relationshipType = "SS" Then
                    If relationship = "Prev" Then
                        Cells(i, ColumnIndex("Start")).Value = GetDateCustom(Cells(i - 1, ColumnIndex("Start")).Value, Cells(i, ColumnIndex("Lag")).Value, Cells(i, ColumnIndex("Calendar")).Value)
                    Else
                        Cells(i, ColumnIndex("Start")).Value = GetDateCustom(IDStartDate(relationship), Cells(i, ColumnIndex("Lag")).Value, Cells(i, ColumnIndex("Calendar")).Value)
                    End If
                End If
            End If
        End If

        ' Calculate Finish
        If isEmpty(Cells(i, ColumnIndex("Start")).Value) Or isEmpty(Cells(i, ColumnIndex("Duration")).Value) Then
            Cells(i, ColumnIndex("Duration")).Interior.Color = RGB(255, 200, 200)
            Cells(i, ColumnIndex("Start")).Interior.Color = RGB(255, 200, 200)
            MsgBox "Invalid data for Activity ID: " & activityID & ". Start or Duration is empty.", vbExclamation
            CalculateStartAndFinish = False
            Exit Function
        Else
            Cells(i, ColumnIndex("Finish")).Value = GetDateCustom(Cells(i, ColumnIndex("Start")).Value, Cells(i, ColumnIndex("Duration")).Value, Cells(i, ColumnIndex("Calendar")).Value) - 1
        End If
    Next i
End Function

Function IDStartDate(activityID As String) As Date
    Dim rowIndex As Variant
    Dim activityCol As Long
    Dim lastRow As Long

    activityCol = ColumnIndex("Activity ID")
    lastRow = Cells(Rows.Count, activityCol).End(xlUp).Row

    rowIndex = Application.Match(activityID, Range(Cells(2, activityCol), Cells(lastRow, activityCol)), 0)

    If Not IsError(rowIndex) Then
        rowIndex = rowIndex + 1 ' Adjust for header row
        IDStartDate = CDate(Cells(rowIndex, ColumnIndex("Start")).Value)
    Else
        IDStartDate = 0 ' Or use CVDate("") to return an empty date
    End If
End Function

Function IDFinishDate(activityID As String) As Date
    Dim rowIndex As Variant
    Dim activityCol As Long
    Dim lastRow As Long

    activityCol = ColumnIndex("Activity ID")
    lastRow = Cells(Rows.Count, activityCol).End(xlUp).Row

    rowIndex = Application.Match(activityID, Range(Cells(2, activityCol), Cells(lastRow, activityCol)), 0)

    If Not IsError(rowIndex) Then
        rowIndex = rowIndex + 1 ' Adjust for header row
        IDFinishDate = CDate(Cells(rowIndex, ColumnIndex("Finish")).Value)
    Else
        IDFinishDate = 0
    End If
End Function

Function GetDateCustom(Start As Date, duration As Double, calendar As String) As Date
    Dim workingDayCount As Long
    Dim currentDate As Date
    Dim currentDateIndex As Integer
    Dim isHoliday As Boolean
    Dim calendarWorkDays As Variant

    currentDate = Start
    holidayRange = Range("Holidays")
    calendarWorkDays = GetColumn(calendar)

    workingDayCount = 0

    ' Count working days based on calendar and holidays
    Do While workingDayCount < duration
        currentDateIndex = Weekday(currentDate, vbMonday) - 1 ' 1=Monday, 7=Sunday

        If calendarWorkDays(currentDateIndex) = 1 Then
            isHoliday = Not IsError(Application.Match(currentDate, holidayRange, 0))
            
            If Not isHoliday Then
                workingDayCount = workingDayCount + 1
            End If
        End If
        currentDate = currentDate + 1
    Loop

    ' Ensure the final date is also a valid working day
    Do
        currentDateIndex = Weekday(currentDate, vbMonday) - 1

        If calendarWorkDays(currentDateIndex) = 1 Then
            isHoliday = Not IsError(Application.Match(currentDate + 1, holidayRange, 0))
            
            If Not isHoliday Then
                Exit Do
            End If
        End If
        currentDate = currentDate + 1
    Loop

    GetDateCustom = currentDate
End Function

Sub PopulateDateHeaders()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Gantt")
    Dim earliestStart As Date
    Dim latestFinish As Date
    Dim i As Long
    Dim lastRow As Long
    Dim col As Long
    Dim headerDate As Date
    Dim calendar As String

    col = ColumnIndex("GanttStart")
    earliestStart = ws.Cells(2, ColumnIndex("Start")).Value
    latestFinish = earliestStart
    lastRow = ws.Cells(ws.Rows.Count, ColumnIndex("Start")).End(xlUp).Row
    

    For i = 2 To lastRow
        If IsDate(ws.Cells(i, ColumnIndex("Start")).Value) And IsNumeric(ws.Cells(i, ColumnIndex("Duration")).Value) Then
            Dim startDate As Date
            Dim duration As Double
            Dim finishDate As Date
            Dim checkDate As Date
            Dim holidayCount As Integer

            startDate = ws.Cells(i, ColumnIndex("Start")).Value
            duration = ws.Cells(i, ColumnIndex("Duration")).Value
            calendar = ws.Cells(i, ColumnIndex("Calendar")).Value
            finishDate = GetDateCustom(startDate, duration, calendar)

            If startDate < earliestStart Then earliestStart = startDate
            If finishDate > latestFinish Then latestFinish = finishDate
        End If
    Next i

    headerDate = earliestStart
    Dim outputSheet As Worksheet
    Set outputSheet = ThisWorkbook.Sheets("Export")
    'outputSheet.Cells(1, 1).Value = "Date"

    Do While headerDate <= latestFinish
        ws.Cells(1, col).Value = headerDate
        ws.Cells(1, col).NumberFormat = "mm/dd/yyyy"
        'outputSheet.Cells(col - 12, 1).Value = headerDate
        headerDate = headerDate + 1
        col = col + 1
    Loop
End Sub

Sub DebugDateRange()
    Dim cell As Range
    Dim rng As Range
    Set rng = GetLastColumnRange()

    Debug.Print "Range is: " & rng.Address(False, False)

    For Each cell In rng
        Debug.Print "Cell Value: " & cell.Value & " | Type: " & TypeName(cell.Value)
    Next cell
End Sub

Sub FormatDateRange()
    Dim cell As Range
    Dim lastCol As Long
    Dim rng As Range
    
    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    Set rng = Range(Cells(1, ColumnIndex("GanttStart")), Cells(1, lastCol))
    For Each cell In rng
        If IsDate(cell.Value) Then
            cell.Value = DateValue(cell.Value)
        End If
    Next cell
End Sub

Sub ClearFromJ2()
    Dim ws As Worksheet
    Dim plotWS As Worksheet

    Set ws = ThisWorkbook.Sheets("Gantt") ' Main sheet
    Set plotWS = ThisWorkbook.Sheets("Plot") ' Sheet to clear additionally
    Set exportSheet = ThisWorkbook.Sheets("Export")

    Call ClearAllInteriorColor

    With ws
        .Range(.Cells(1, ColumnIndex("GanttStart")), .Cells(.Rows.Count, .Columns.Count)).ClearContents

        ' Clear ActiveX TextBoxes
        .OLEObjects("MaxBox").Object.Text = ""
        .OLEObjects("MinBox").Object.Text = ""
        .OLEObjects("AveBox").Object.Text = ""
        .OLEObjects("MaxDate").Object.Text = ""
        .OLEObjects("MinDate").Object.Text = ""
        .OLEObjects("SlopeBox").Object.Text = ""
        .OLEObjects("RSqBox").Object.Text = ""
    End With

    'plotWS.Cells.ClearContents
    exportSheet.Cells.ClearContents
End Sub

Sub ClearAllInteriorColor()
    Cells.Interior.ColorIndex = xlColorIndexNone
End Sub

Sub DiagnoseClearContentsIssue()
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim lastRow As Long
    Dim rng As Range

    Set ws = ThisWorkbook.Sheets("Gantt")

    With ws
        If isEmpty(.Range(.Cells(1, ColumnIndex("GanttStart")))) Then
            Debug.Print "J1 is empty — nothing to clear.", vbInformation
            Exit Sub
        End If
        
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        lastRow = .Cells(.Rows.Count, 2).End(xlUp).Row  ' Using column K (11) for lastRow

        Debug.Print "Last Column: " & lastCol
        Debug.Print "Last Row: " & lastRow

        Set rng = .Range(.Cells(1, ColumnIndex("GanttStart")), .Cells(lastRow, lastCol))

        Debug.Print "Range Address: " & rng.Address
        Debug.Print "Merged Cells: " & rng.MergeCells
        Debug.Print "Sheet Protected: " & .ProtectContents

        On Error Resume Next
        rng.ClearContents
        If Err.Number <> 0 Then
            MsgBox "Error 1004: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End With
End Sub

Sub DeleteEmptyRows(ws As Worksheet, Optional checkCols As Variant)
    Dim lastRow As Long, i As Long, isEmpty As Boolean
    Dim col As Variant

    Application.ScreenUpdating = False

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = lastRow To 1 Step -1  ' Loop from bottom up to avoid skipping rows after deletion
        isEmpty = True

        If IsMissing(checkCols) Then
            ' Check entire row
            If Application.WorksheetFunction.CountA(ws.Rows(i)) > 0 Then
                isEmpty = False
            End If
        Else
            ' Check only specified columns
            For Each col In checkCols
                If Trim(ws.Cells(i, col).Value) <> "" Then
                    isEmpty = False
                    Exit For
                End If
            Next col
        End If

        If isEmpty Then ws.Rows(i).Delete
    Next i

    Application.ScreenUpdating = True
End Sub

Sub ShowInfoImageIfMissing()
    Dim imgName As String
    imgName = "Info"

    Dim img As Shape
    On Error Resume Next
    Set img = ActiveSheet.Shapes(imgName)
    On Error GoTo 0

    ' If image exists, just make sure it's visible
    If Not img Is Nothing Then
        img.Visible = True
        Exit Sub
    End If

    ' If image is missing, copy it from the Holidays sheet
    Dim imgSource As Shape
    On Error Resume Next
    Set imgSource = Sheets("Calendars").Shapes("StoredInfo")
    On Error GoTo 0

    If imgSource Is Nothing Then
        'MsgBox "Image named 'Info' not found on Holidays sheet.", vbExclamation
        Exit Sub
    End If

    imgSource.Copy
    ActiveSheet.Paste
    Set img = ActiveSheet.Shapes(ActiveSheet.Shapes.Count)
    img.Name = imgName
    img.Top = Range("A1").Top
    img.Left = Range("A1").Left
    img.Placement = xlMoveAndSize
    img.ZOrder msoSendToBack
End Sub


Sub RunResources()
    Dim success As Boolean
    Dim targetWS As Worksheet
  
    Set targetWS = ThisWorkbook.Sheets("Gantt")
    
    If Not ActiveSheet Is targetWS Then
        targetWS.Activate
    End If

    Call InitializeColumnIndex
    Call ClearFromJ2
    
    success = CalculateStartAndFinish(True)
    If Not success Then
        Exit Sub
    End If
    Call PopulateDateHeaders
    Call WriteToCells
    Call SumValuesByDate(False)
    Call GetTrendlineStats
    Call DeleteEmptyRows(ThisWorkbook.Sheets("Export"))
    Call CreatePivotChartWithSlicerPanel
    
End Sub

Sub ShowMyForm()
    PlotForm.Show
End Sub

'Author: Steven Hall
'Email: shall@austin-ind.com
'Date: 2025-09-11
'Version: 1.5




