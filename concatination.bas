Sub ConcatenateColumns()
    Dim sourceWorkbook As Workbook
    Dim destinationWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim destinationWorksheet1 As Worksheet
    Dim destinationWorksheet2 As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim concatenatedAB As String
    Dim concatAdditional As String

    ' Set the source workbook and worksheet
    Set sourceWorkbook = Workbooks("Source_Workbook.xlsx") ' Change to the name of your source workbook
    Set sourceWorksheet = sourceWorkbook.Sheets("ConcatInput")

    ' Set the destination workbook
    Set destinationWorkbook = Workbooks("Destination_Workbook.xlsx") ' Change to the name of your destination workbook

    ' Set the destination worksheets
    Set destinationWorksheet1 = destinationWorkbook.Sheets("ConcatOutput1")
    Set destinationWorksheet2 = destinationWorkbook.Sheets("ConcatOutput2")

    ' Find the last row in the source worksheet
    lastRow = sourceWorksheet.Cells(sourceWorksheet.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row
    For i = 1 To lastRow
        ' Concatenate columns B and C
        concatenatedAB = sourceWorksheet.Cells(i, 2).Value & sourceWorksheet.Cells(i, 3).Value

        'Additional info to Column B
        concatAdditional = sourceWorksheet.Cells(i, 2).Value & " is a Company"


        ' Write the concatenated values to the destination worksheets
        destinationWorksheet1.Cells(i, 2).Value = concatenatedAB
        destinationWorksheet2.Cells(i,2).Value = concatAdditional
    Next i

    ' Save changes in the destination workbook
    destinationWorkbook.Save

    ' Close workbooks
    sourceWorkbook.Close
    destinationWorkbook.Close
End Sub
