Sub YourMacro()
    On Error Resume Next ' Enable error handling

    ' Your code here that may cause errors
    ' For example, filling information that might throw an error

    ' Check for errors
    If Err.Number <> 0 Then
        ' Error occurred, log the error information
        LogError "Error during data filling: " & Err.Description
        ' Reset error object
        Err.Clear
    End If

    ' ... Continue with the rest of your code

    On Error GoTo 0 ' Disable error handling

    ' Display a summary of errors, if any
    If errorLog <> "" Then
        MsgBox "Error Summary:" & vbCrLf & errorLog, vbExclamation, "Errors Occurred"
    End If
End Sub

Dim errorLog As String ' Global variable to store error information

Sub LogError(errorMessage As String)
    ' Log the error information
    errorLog = errorLog & errorMessage & vbCrLf
End Sub
