Sub MapLongForms()
    Dim sourceWorkbook As Workbook
    Dim destinationWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim destinationWorksheet As Worksheet
    Dim longToShortDictionary As Object
    Dim lastRow As Long
    Dim i As Long
    Dim longForm As String
    Dim shortForm As String

    ' Set the source workbook and worksheet
    Set sourceWorkbook = Workbooks("Source_Workbook.xlsx") ' Change to the name of your source workbook
    Set sourceWorksheet = sourceWorkbook.Sheets("LongShort")

    ' Set the destination workbook and worksheet
    Set destinationWorkbook = Workbooks("Destination_Workbook.xlsx") ' Change to the name of your destination workbook
    Set destinationWorksheet = destinationWorkbook.Sheets("LongShort")

    ' Assuming you have a dictionary named longToShortDictionary
    Set longToShortDictionary = CreateObject("Scripting.Dictionary")
    ' Populate your dictionary with long to short form mappings
    longToShortDictionary.Add "Yes", "Y"
    longToShortDictionary.Add "No", "N"
    ' longToShortDictionary.Add "LongForm2", "Short2"
    ' Add more mappings as needed

    ' Find the last row in the source worksheet
    lastRow = sourceWorksheet.Cells(sourceWorksheet.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row
    For i = 1 To lastRow
        ' Read the long form from column A
        longForm = sourceWorksheet.Cells(i, 1).Value

        ' Check if the long form is in the dictionary
        If longToShortDictionary.Exists(longForm) Then
            ' If it exists, get the corresponding short form
            shortForm = longToShortDictionary(longForm)

            ' Write the short form to the destination worksheet in column A
            destinationWorksheet.Cells(i, 1).Value = shortForm
        Else
            ' If the long form is not in the dictionary, you may want to handle it accordingly
            ' For now, just copy the long form to the destination worksheet
            destinationWorksheet.Cells(i, 1).Value = longForm
        End If
    Next i

    ' Save changes in the destination workbook
    destinationWorkbook.Save

    ' Close workbooks
    sourceWorkbook.Close
    destinationWorkbook.Close
End Sub
