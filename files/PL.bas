Sub Pl()
    Dim sourceWorkbook As Workbook
    Dim destinationWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim destinationWorksheet As Worksheet
    Dim devConstantsWs As Worksheet
    Dim currentWs As Worksheet
    Dim TargetSheetName As Range
    Dim TargetColumnName As Range
    Dim transformationRuleCell As Range
    Dim longToShortDictionary As Object
    Dim MappingInputRows As Long
    Dim lastRowSource As Long
    Dim i As Long
    Dim j As Long
    Dim sourceResultColumn As Long
    Dim DestinationResultColumn As Long
    Dim CurrentWsName As String
    Dim longForm As String
    Dim shortForm As String

    ' Assuming you have a dictionary named longToShortDictionary
    Set longToShortDictionary = CreateObject("Scripting.Dictionary")
    longToShortDictionary.Add "Yes", "Y"
    longToShortDictionary.Add "No", "N"

    'Setting the WorkSheet name in scope for the Sub process
    Set CurrentWsName = "PL"

    'reading the source and destination workbook names from the Dev-constants
    Set currentWs = ThisWorkbook.Sheets(CurrentWsName)
    Set devConstantsWs = ThisWorkbook.Sheets("Dev-Constants")

    ' Set the source & Destination workbook
    Set sourceWorkbook = Workbooks(devConstantsWs.Cells(2, 2).Value) ' Change to the name of your source workbook
    Set destinationWorkbook = Workbooks(devConstantsWs.Cells(3, 2).Value) ' Change to the name of your destination workbook

    ' Set the source worksheet
    Set sourceWorksheet = sourceWorkbook.Sheets("Parameter list")

    'Finding the Input Worksheet end row
    MappingInputRows = currentWs.Cells(currentWs.Rows.Count, "A").End(xlUp).Row

    'Looping throught each element of column A of the driver work book Parameter sheet.
    For i = 2 To MappingInputRows
        'setting the source column, target sheet name, target column name cells.
        Set SourceColumnName = ThisWorkbook.Sheets(CurrentWsName).Cells(i, 2)
        Set TargetSheetName = ThisWorkbook.Sheets(CurrentWsName).Cells(i, 3)
        Set TargetColumnName = ThisWorkbook.Sheets(CurrentWsName).Cells(i, 4)
        Set transformationRuleCell = ThisWorkbook.Sheets(CurrentWsName).Cells(i,5)
        'check if the target sheet name or the column name is empty if both not empty only then execute.
        If Not IsEmpty(TargetSheetName.Value) And Not IsEmpty(TargetColumnName.Value) Then
            Set destinationWorksheet = destinationWorkbook.Sheets(TargetSheetName.Value)
            sourceResultColumn = FindColumnByKeyword(sourceWorkbook, "Parameter", SourceColumnName.Value, 2)
            DestinationResultColumn = FindColumnByKeyword(destinationWorkbook, TargetSheetName.Value, TargetColumnName.Value, 3)
            lastRowSource = sourceWorksheet.Cells(sourceWorksheet.Rows.Count, sourceResultColumn).End(xlUp).Row

            If Not IsEmpty(transformationRuleCell.Value) Then
                'actions to perform the transformation rules.
                longForm = sourceWorksheet.Cells(j, sourceResultColumn).Value

                'If the transform exists then put the transformed data else the original.
                If longToShortDictionary.Exists(longForm) Then
                    destinationWorksheet.Cells(j + 1, DestinationResultColumn).Value = longToShortDictionary(longForm)
                Else
                    destinationWorksheet.Cells(j + 1, DestinationResultColumn).Value = longForm
                End If

            Else
                'actions to perform copy action from source cell to destination cell.
                For j = 3 To lastRowSource
                    destinationWorksheet.Cells(j + 1, DestinationResultColumn).Value = sourceWorksheet.Cells(j, sourceResultColumn).Value
                Next j
            End If
        End If
    Next i

    destinationWorkbook.Save
end Sub

Function FindColumnByKeyword(Workbook As Workbook, sheetName As String, searchKey As String, headerRow As Long) As Long
    Dim ws As Worksheet
    Dim col As Range
    
    ' Check if the workbook is not nothing
    If Workbook Is Nothing Then
        MsgBox "Invalid workbook!"
        Exit Function
    End If
    
    ' Check if the sheet name is provided
    If Len(sheetName) = 0 Then
        MsgBox "Sheet name is required!"
        Exit Function
    End If
    
    ' Check if the search key is provided
    If Len(searchKey) = 0 Then
        MsgBox "Search key is required!"
        Exit Function
    End If
    
    ' Set the worksheet reference
    On Error Resume Next
    Set ws = Workbook.Sheets(sheetName)
    On Error GoTo 0
    
    ' Check if the worksheet exists
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found in the workbook!"
        Exit Function
    End If
    
    ' Search for the key in the second row
    Set col = ws.Rows(headerRow).Find(What:=searchKey, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Check if the key is found
    If Not col Is Nothing Then
        ' Return the column number
        FindColumnByKeyword = col.Column
    Else
        ' Indicate that the key is not found
        FindColumnByKeyword = 0
        MsgBox "Search key '" & searchKey & "' not found in the '" & headerRow & "'row of sheet '" & sheetName & "'."
    End If
End Function