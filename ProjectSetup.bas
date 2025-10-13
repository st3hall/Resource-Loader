Sub ClearCriteriaSelection()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Criteria Selection")

    With ws
        .Range("A12", .Cells(.Rows.Count, .Columns.Count)).Clear
    End With
End Sub

Sub CreateCriteriaSelection()
    Dim wsWeights As Worksheet
    Dim wsCriteria As Worksheet
    Dim wsReference As Worksheet
    Dim wsSubs As Worksheet
    Dim buildingList As Collection
    Dim subcontractorList As Collection
    Dim criteriaList As Collection
    Dim buildingName As String
    Dim subcontractorName As String
    Dim criteriaName As String
    Dim i As Long, j As Long, k As Long
    
    Set wsWeights = ThisWorkbook.Sheets("Weights")
    Set wsCriteria = ThisWorkbook.Sheets("Criteria Selection")
    Set wsReference = ThisWorkbook.Sheets("Reference List")
    Set wsSubs = ThisWorkbook.Sheets("Subs Required")
    
    ' Step 1: Create Building List
    Set buildingList = New Collection
    col = 1
    Do While wsSubs.Cells(10, col).Value <> ""
        buildingList.Add wsSubs.Cells(10, col).Value
        'Debug.Print " - " & wsSubs.Cells(10, col).Value
        col = col + 1
    Loop

    ' Step 2: Loop through each building
    For Each building In buildingList
        ' Create Subcontractor List for current building
        Set subcontractorList = New Collection
        col = 1
        Do While wsWeights.Cells(11, col).Value <> ""
            If wsWeights.Cells(11, col).Value = building Then
                row = 13
                Do While wsWeights.Cells(row, col).Value <> ""
                    subcontractorList.Add wsWeights.Cells(row, col).Value
                    'Debug.Print "   - " & wsWeights.Cells(row, col).Value
                    row = row + 1
                Loop
                Exit Do
            End If
            col = col + 1
        Loop

        ' Step 3: Loop through each subcontractor
        For Each subcontractor In subcontractorList
            ' Create Criteria List for current subcontractor
            'Debug.Print "    Processing Subcontractor: " & subcontractor
            Set criteriaList = New Collection
            lastRow = wsReference.Cells(wsReference.Rows.Count, "A").End(xlUp).Row
            For i = 2 To lastRow
                If wsReference.Cells(i, "A").Value = subcontractor Then ' Assuming "FORM" is column A and "Criteria" is column B
                    criteriaList.Add wsReference.Cells(i, "B").Value
                    'Debug.Print "      Criteria: " & wsReference.Cells(i, "B").Value
                End If
            Next i

            ' Step 4: Write to Criteria Selection sheet
            ' Find the column for the current building
            col = 1
            Do While wsCriteria.Cells(11, col).Value <> ""
                If wsCriteria.Cells(11, col).Value = building Then Exit Do
                col = col + 1
            Loop

            ' Find the first empty row under the building column
            row = 12
            Do While wsCriteria.Cells(row, col).Value <> ""
                row = row + 1
            Loop
            ' Write criteria starting from next column
            Dim critCol As Long
            critCol = col + 1
            For Each criteria In criteriaList
                wsCriteria.Cells(row, col).Value = subcontractor
                wsCriteria.Cells(row, critCol).Value = criteria
                row = row + 1
            Next criteria
        Next subcontractor
    Next building
    
    ' Disable text wrapping for entire output table
    wsCriteria.Range("A12", wsCriteria.Cells(wsCriteria.Rows.Count, wsCriteria.Columns.Count)).WrapText = False

    MsgBox "Criteria Selection sheet updated successfully!"
End Sub