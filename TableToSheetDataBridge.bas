Option Explicit

' === Copy Table Data to Another Sheet Based on Mapping ===
'
' This macro copies specific columns from one or more Excel Tables (ListObjects)
' and pastes them into another sheet, based on a configurable mapping table.
'
' Key Features:
' - Works with structured Excel Tables
' - Uses column index or header name
' - Allows date formatting (e.g., convert to "yyyymmdd")
' - Designed for easy reuse and customization

' -------------------------------
' CONFIGURATION SECTION
' -------------------------------
' Each row in the GetMappings function defines one data transfer rule:
'
' Format:
'   Array( _
'     "SourceSheet",      ' Sheet name containing the table
'     "TableName",        ' Name of the Excel Table (ListObject)
'     ColumnID,           ' Column header (string) or index (number)
'     "TargetSheet",      ' Sheet to paste into
'     "TargetColumn",     ' Column letter in target sheet
'     StartRow,           ' Starting row for pasting
'     Optional Format     ' Optional: Format string like "yyyymmdd"
'   )
'
' Example Use Case:
'   You have a sheet "MainData" with a table called "MainTable",
'   and you want to extract the first column (date), format it, and paste it into column A
'   of another sheet "DataToGo".

' -------------------------------
' Array refrance
' -------------------------------
' 1 is date
' 2 is names
'-------------------------------
Private Function GetMappings() As Variant
    ' === CONFIGURATION ===
    ' Define reusable variables for the mappings
    Dim startFromRow As Long
    Dim sourceSheetName As String
    Dim sourceTableName As String
    Dim targetSheetName As String

    ' Set values once, use everywhere below
    sourceSheetName = "MainData"      ' Name of the worksheet that holds the source table
    sourceTableName = "MainTable"     ' Name of the ListObject (structured table)
    targetSheetName = "DataToGo"      ' Sheet where data will be pasted
    startFromRow = 3                  ' First row on the target sheet to paste data

    ' === MAPPING RULES ===
    ' Each Array() defines one column transfer operation.
    ' Format:
    '   Array(sourceSheet, tableName, columnIndexOrHeader(can be a string or the possion from A), targetSheet, targetColLetter, startRow, optionalFormat)
    ' Example:
    '   Copy column 1, format it as yyyymmdd, and paste to column A of target sheet
    '   Copy column 2, paste to column B without formatting
    GetMappings = Array( _
        Array(sourceSheetName, sourceTableName, 1, targetSheetName, "A", startFromRow, "yyyymmdd"), _
        Array(sourceSheetName, sourceTableName, 2, targetSheetName, "B", startFromRow) _
        ' Add more mappings below as needed
    )
End Function

' -------------------------------
' MAIN COPY ROUTINE
' -------------------------------
' Loops through each mapping entry, copies the data, formats if needed,
' and pastes it into the target location.

Public Sub CopyMappingsData()

    Dim mappingsList   As Variant        ' Array of mappings returned from GetMappings
    Dim mappingItem    As Variant        ' Single mapping entry
    Dim sourceSheet    As Worksheet      ' Sheet with the source table
    Dim targetSheet    As Worksheet      ' Sheet where data will be pasted
    Dim sourceTable    As ListObject     ' Excel Table (ListObject)
    Dim sourceColumn   As ListColumn     ' Specific column in the table to copy
    Dim sourceData     As Variant        ' Raw data values from the table column
    Dim columnID       As Variant        ' Column identifier (header or index)
    Dim columnIndex    As Long           ' If using a column index
    Dim pasteStartRow  As Long           ' Row where pasting starts in target sheet
    Dim pasteEndRow    As Long           ' Row where pasting ends
    Dim targetColumn   As String         ' Column letter in target sheet
    Dim formatString   As String         ' Optional date format string
    Dim i              As Long           ' Loop counter

    mappingsList = GetMappings()         ' Load all mappings

    For Each mappingItem In mappingsList
        On Error GoTo HandleMappingError

        ' --- Extract mapping info ---
        Set sourceSheet = ThisWorkbook.Worksheets(mappingItem(0))
        Set targetSheet = ThisWorkbook.Worksheets(mappingItem(3))
        Set sourceTable = sourceSheet.ListObjects(mappingItem(1))
        columnID = mappingItem(2)
        targetColumn = mappingItem(4)
        pasteStartRow = mappingItem(5)

        ' --- Determine source column ---
        If VarType(columnID) = vbString Then
            Set sourceColumn = sourceTable.ListColumns(CStr(columnID))
        Else
            columnIndex = CLng(columnID)
            Set sourceColumn = sourceTable.ListColumns(columnIndex)
        End If

        ' --- Get data from the column ---
        sourceData = sourceColumn.DataBodyRange.Value

        ' --- Optional: format dates or other values ---
        If UBound(mappingItem) >= 6 Then
            formatString = CStr(mappingItem(6))
            If formatString <> "" Then
                For i = 1 To UBound(sourceData, 1)
                    If IsDate(sourceData(i, 1)) Then
                        sourceData(i, 1) = Format(sourceData(i, 1), formatString)
                    End If
                Next i
            End If
        End If

        ' --- Compute paste range ---
        pasteEndRow = pasteStartRow - 1 + UBound(sourceData, 1)

        ' --- Paste the data block ---
        targetSheet.Range(targetColumn & pasteStartRow & ":" & _
                          targetColumn & pasteEndRow).Value = sourceData

ContinueLoop:
        On Error GoTo 0
        Err.Clear
    Next mappingItem

    MsgBox "All copy operations completed successfully.", vbInformation
    Exit Sub

' -------------------------------
' ERROR HANDLING
' -------------------------------
HandleMappingError:
    MsgBox "Error in mapping:" & vbCrLf & _
           "Source Sheet = '" & mappingItem(0) & "'" & vbCrLf & _
           "Table = '" & mappingItem(1) & "'" & vbCrLf & _
           "ColumnID = '" & mappingItem(2) & "'" & vbCrLf & _
           "Details: " & Err.Description, vbCritical
    Resume ContinueLoop

End Sub
