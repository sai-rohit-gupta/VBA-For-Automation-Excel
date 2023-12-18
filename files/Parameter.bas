Sub TransferData()
    Dim sourceWorkbook As Workbook
    Dim destinationWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim destinationWorksheet As Worksheet
    Dim devConstantsWs As Worksheet
    Dim currentWs As Worksheet
    Dim TargetSheetName As Range
    Dim MappingInputRows As Long
    Dim lastRowSource As Long
    Dim lastRowDestination As Long
    Dim i As Long
    Dim sourceWBName As String
    Dim DestinationWBName As String

    'reading the source and destination workbook names from the Dev-constants
    Set currentWs = ThisWorkbook.Sheets("Parameter")
    Set devConstantsWs = ThisWorkbook.Sheets("Dev-Constants")

    ' Set the source & Destination workbook and worksheet
    Set sourceWorkbook = Workbooks(devConstantsWs.Cells(2,2)) ' Change to the name of your source workbook
    Set destinationWorkbook = Workbooks(devConstantsWs.Cells(3,2)) ' Change to the name of your destination workbook

    ' Set the source worksheet
    Set sourceWorksheet = sourceWorkbook.Sheets("Parameter")

    'Finding the Input Worksheet end row
    MappingInputRows = currentWs.Cells(sourceWorksheet.Rows.Count, "A").End(xlUp).Row

    for i=2 to MappingInputRows
        Set TargetSheetName = ThisWorkbook.Sheets("Parameter").Cells(i,4)
        if Not IsEmpty(TargetSheetName.Value)
            Set destinationWorksheet = destinationWorkbook.Sheets()
            
            lastRowSource = 
            for j=1 to MappingInputRows-1+3
                destinationWorksheet.Cells(j, 1).Value = sourceWorksheet.Cells(j-3+2, 1).Value
        End If
    Next i

    for 




    Set destinationWorksheet = destinationWorkbook.Sheets("ColumnShift")
End Sub