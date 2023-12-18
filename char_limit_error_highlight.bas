Sub CopyAndHighlightExceedingChars()
    Dim sourceWorkbook As Workbook
    Dim destinationWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim destinationWorksheet As Worksheet
    Dim lastRowSource As Long
    Dim i As Long
    Dim MAX_CHAR_LIMIT As Integer ' Set the maximum character limit
    MAX_CHAR_LIMIT = 10 ' Change to your desired character limit

    ' Set the source workbook and worksheet
    Set sourceWorkbook = Workbooks("SourceWorkbook.xlsx") ' Change to the name of your source workbook
    Set sourceWorksheet = sourceWorkbook.Sheets("Sheet1")

    ' Set the destination workbook and worksheet
    Set destinationWorkbook = Workbooks("DestinationWorkbook.xlsx") ' Change to the name of your destination workbook
    Set destinationWorksheet = destinationWorkbook.Sheets("Sheet1")

    ' Find the last row in the source worksheet
    lastRowSource = sourceWorksheet.Cells(sourceWorksheet.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row
    For i = 1 To lastRowSource
        ' Copy the value from the source to the destination worksheet
        destinationWorksheet.Cells(i, 1).Value = sourceWorksheet.Cells(i, 1).Value

        ' Check if the length exceeds the character limit
        If Len(destinationWorksheet.Cells(i, 1).Value) > MAX_CHAR_LIMIT Then
            ' Highlight the cell with a yellow background
            destinationWorksheet.Cells(i, 1).Interior.Color = RGB(255, 255, 0) ' Yellow
        End If
    Next i

    ' Save changes in the destination workbook
    destinationWorkbook.Save

    ' Close workbooks
    sourceWorkbook.Close
    destinationWorkbook.Close
End Sub
