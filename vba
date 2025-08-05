Sub CreateSheetsBasedOnColumn()
    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim lastRow As Long
    Dim uniqueValues As Object ' Late binding for Dictionary object
    Dim key As Variant
    Dim sourceColumn As Long ' Change this to the column number you're interested in
    Dim predefinedData As Variant ' Adjust the range for your predefined data
    Dim appName As String

    ' Get the Excel application name
    appName = Application.Name

    ' Set the source worksheet and column
    Set wsSource = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your source sheet name
    sourceColumn = 1 ' Assuming the column with unique values is the first column (A)

    ' Define your predefined data range.  This assumes your predefined
    ' data is in a sheet named "TemplateSheet" and starts at "A1".
    ' Change this range as needed.  It's crucial this is correctly set.
    On Error Resume Next ' Handle error if "TemplateSheet" doesn't exist
    predefinedData = ThisWorkbook.Sheets("TemplateSheet").Range("A1:C10").Value ' Example: A1 to C10
    On Error GoTo 0 ' Reset error handling

    ' Check if predefinedData was successfully assigned
    If IsEmpty(predefinedData) Then
        MsgBox "Error: Could not retrieve predefined data.  Ensure the sheet 'TemplateSheet' exists and the range A1:C10 is correct.", vbCritical
        Exit Sub ' Stop the macro if the predefined data is not found
    End If

    ' Create a Dictionary object to store unique values
    Set uniqueValues = CreateObject("Scripting.Dictionary")

    ' Find the last row in the source column
    lastRow = wsSource.Cells(wsSource.Rows.Count, sourceColumn).End(xlUp).Row

    ' Loop through the source column and add unique values to the Dictionary
    For i = 1 To lastRow
        If Not uniqueValues.Exists(wsSource.Cells(i, sourceColumn).Value) Then
            uniqueValues.Add wsSource.Cells(i, sourceColumn).Value, 1
        End If
    Next i

    ' Check if any unique values were found
    If uniqueValues.Count = 0 Then
        MsgBox "No unique values found in the specified column.", vbInformation
        Exit Sub ' Stop the macro if no unique values are found
    End If

    ' Loop through the unique values and create new sheets
    For Each key In uniqueValues.Keys
        ' Create a new sheet.  This adds it *after* the last existing sheet.
        Set wsNew = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        On Error Resume Next ' Handle errors, specifically if the sheet name is invalid
        wsNew.Name = key ' Use the unique value as the sheet name
        On Error GoTo 0

        ' Check if the sheet name was valid and the sheet was created
        If wsNew.Name <> key Then
             MsgBox "Error: Could not create sheet named '" & key & "'.  Sheet name may be invalid.", vbCritical
             ' You could choose to delete the invalid sheet here if you want:
             ' Application.DisplayAlerts = False ' Prevent prompts
             ' wsNew.Delete
             ' Application.DisplayAlerts = True
             ' Continue to the next key (unique value)
             GoTo NextKey
        End If
        ' Paste the predefined data
        wsNew.Range("A1").Resize(UBound(predefinedData, 1), UBound(predefinedData, 2)).Value = predefinedData

NextKey:
    Next key

    MsgBox "Sheets created and populated successfully!", vbInformation

End Sub
