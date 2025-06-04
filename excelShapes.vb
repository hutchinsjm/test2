Sub DrawEventShapes()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' Change if needed
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    Dim startRow As Long: startRow = 7
    Dim teamStartRow As Long: teamStartRow = 9
    Dim teamCol As Long: teamCol = 8 ' Column H
    Dim dateRow As Long: dateRow = 8
    Dim dateStartCol As Long: dateStartCol = 10 ' Column J

    Dim eventName As String, assignedUnit As String
    Dim startDate As Date, endDate As Date
    Dim teamRow As Long, startCol As Long, endCol As Long
    Dim cell As Range, shp As Shape
    Dim i As Long, j As Long

    ' Clear existing shapes
    For Each shp In ws.Shapes
        If Left(shp.Name, 5) = "Event" Then shp.Delete
    Next shp

    ' Loop through events
    For i = startRow To lastRow
        eventName = ws.Cells(i, 2).Value
        assignedUnit = ws.Cells(i, 3).Value
        startDate = ws.Cells(i, 4).Value
        endDate = ws.Cells(i, 5).Value

        ' Find team row
        teamRow = 0
        For j = teamStartRow To teamStartRow + 10
            If ws.Cells(j, teamCol).Value = assignedUnit Then
                teamRow = j
                Exit For
            End If
        Next j
        If teamRow = 0 Then GoTo SkipEvent

        ' Find start and end columns based on date headers
        startCol = 0: endCol = 0
        For j = dateStartCol To ws.Cells(dateRow, ws.Columns.Count).End(xlToLeft).Column
            If ws.Cells(dateRow, j).Value = startDate Then startCol = j
            If ws.Cells(dateRow, j).Value = endDate Then endCol = j
        Next j
        If startCol = 0 Or endCol = 0 Then GoTo SkipEvent

        ' Draw rectangle
        Dim topLeft As Range, bottomRight As Range
        Set topLeft = ws.Cells(teamRow, startCol)
        Set bottomRight = ws.Cells(teamRow, endCol)
        
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
            topLeft.Left, topLeft.Top, _
            bottomRight.Left + bottomRight.width - topLeft.Left, topLeft.height)
        
        With shp
            .Name = "Event_" & i
            .TextFrame2.TextRange.Text = eventName
            .TextFrame2.HorizontalAnchor = msoAnchorCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .TextFrame2.TextRange.Font.Size = 8
            .Fill.ForeColor.RGB = RGB(0, 112, 192)
            .Line.Visible = msoFalse
        End With

SkipEvent:
    Next i

    MsgBox "Shapes drawn.", vbInformation
End Sub


