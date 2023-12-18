Sub TransferData()
    Dim sourceWorkbook As Workbook
    Dim destinationWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim destinationWorksheet As Worksheet
    Dim lastRowSource As Long
    Dim lastRowDestination As Long
    Dim i As Long
    'list to maintain all the source columns to map in destination.TBD


    'dictionary mapping to say what Column in Source to be mapped to what column in destination.TBD

    ' Dim mappingDict As Object
    ' Set mappingDict = CreateObject("Scripting.Dictionary")


    
    ' Set the source workbook and worksheet
    Set sourceWorkbook = Workbooks("Source_Workbook.xlsx") ' Change to the name of your source workbook
    Set sourceWorksheet = sourceWorkbook.Sheets("ColumnShift")
    
    ' Set the destination workbook and worksheet
    Set destinationWorkbook = Workbooks("Destination_Workbook.xlsx") ' Change to the name of your destination workbook
    Set destinationWorksheet = destinationWorkbook.Sheets("ColumnShift")
    
    ' Find the last row in the source worksheet
    lastRowSource = sourceWorksheet.Cells(sourceWorksheet.Rows.Count, "C").End(xlUp).Row
    
    ' Find the last row in the destination worksheet
    lastRowDestination = destinationWorksheet.Cells(destinationWorksheet.Rows.Count, "C").End(xlUp).Row
    
    ' Copy data from source to destination
    For i = 1 To lastRowSource
        lastRowDestination = lastRowDestination + 1
        destinationWorksheet.Cells(lastRowDestination, 1).Value = sourceWorksheet.Cells(i, 1).Value
    Next i
    
    ' Save changes in the destination workbook
    destinationWorkbook.Save
    
    ' Close workbooks
    sourceWorkbook.Close
    destinationWorkbook.Close
End Sub

Sub 