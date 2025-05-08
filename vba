Sub CreateSheetsBasedOnColumn()

    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim lastRow As Long
    Dim uniqueValues As Object ' Late binding for Dictionary object
    Dim key As Variant
    Dim sourceColumn As Long ' Change this to the column number you're interested in
    Dim predefinedData As Variant ' Adjust the range for your predefined data

    ' Set the source worksheet and column
    Set wsSource = ThisWorkbook.Sheets("Sheet1266") ' Change "Sheet1" to your source sheet name
    sourceColumn = 1 ' Assuming the column with unique values is the first column (A)

    ' Define your predefined data range
    predefinedData = ThisWorkbook.Sheets("Sheet1266").Range("B1:N422").Value ' Change "TemplateSheet" and "A1:C10" accordingly

    ' Create a Dictionary object to store unique values
    Set uniqueValues = CreateObject("Scripting.Dictionary")

    ' Find the last row in the source column
    lastRow = wsSource.Cells(Rows.Count, sourceColumn).End(xlUp).Row

    ' Loop through the source column and add unique values to the Dictionary
    For i = 1 To lastRow
        If Not uniqueValues.Exists(wsSource.Cells(i, sourceColumn).Value) Then
            uniqueValues.Add wsSource.Cells(i, sourceColumn).Value, 1
        End If
    Next i

    ' Loop through the unique values and create new sheets
    For Each key In uniqueValues.Keys
        ' Create a new sheet
        Set wsNew = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsNew.Name = key

        ' Paste the predefined data
        wsNew.Range("A1").Resize(UBound(predefinedData, 1), UBound(predefinedData, 2)).Value = predefinedData
    Next key

    MsgBox "Sheets created and populated successfully!", vbInformation

End Sub
