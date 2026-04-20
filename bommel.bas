'==============================================================================
' BOMMEL - BOM Merger for Excel Lists
'==============================================================================
'
' Author: Daniel Leidner
' License: MIT License
'
' Copyright (c) 2026 Daniel Leidner
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'
'==============================================================================
' USAGE:
' 1. Open BOTH Excel files in Excel Desktop (old and new list)
' 2. Press Alt+F11 to open VBA Editor
' 3. Go to Insert > Module
' 4. Copy this code into the module
' 5. Adjust the filenames in the procedure "MergeProcurementLists" on (line 30-31)
' 6. Press F5 or click "Run" to run the macro
'
' RESULT:
' - A new workbook will be created with:
'   * Sheet 1 = Merged list
'   * Sheet 2 = Items only in old list
'   * Sheet 3 = Summary of changes
'
'==============================================================================

Option Explicit

' Module-level column name constants (used across multiple functions)
Private Const FILE_COL As String = "Datei"

Sub MergeProcurementLists()
    
    '==========================================================================
    ' CONFIGURATION
    '==========================================================================
    
    ' Filenames (filename only, no path, e.g. "file.xlsx")
    Const OLD_WORKBOOK_NAME As String = "POD.02_20260420.xlsm"
    Const NEW_WORKBOOK_NAME As String = "POD.02_20260420.xlsm"
    
    ' Columns to transfer from old workbook (user-maintained data)
    Dim transferCols As Variant
    transferCols = Array("Status", "Angebotsnummer", "Bemerkung", "Lieferzeit", _
                        "Liefertermin", "Link", "Preis Gesamt", "Lieferant")

    ' Column names for assembly row detection and formatting
    Const ASSEMBLY_COL As String = "Baugruppe 2"
    Const CATEGORY_COL As String = "Kategorie"

    ' Assembly columns used for row matching (order matters: most specific first)
    Dim matchingAssemblyCols As Variant
    matchingAssemblyCols = Array("Baugruppe 1", "Baugruppe 2")
    
    '==========================================================================
    ' MAIN LOGIC
    '==========================================================================
    
    Dim oldWb As Workbook, newWb As Workbook, outputWb As Workbook
    Dim oldWs As Worksheet, newWs As Worksheet
    Dim mainWs As Worksheet, deletedWs As Worksheet, logWs As Worksheet
    
    Dim oldData As Variant, newData As Variant
    Dim oldHeaders As Variant, newHeaders As Variant
    Dim allHeaders As Collection
    Dim headerDict As Object
    Dim outputData() As Variant
    
    Dim oldKeys() As String, newKeys() As String
    Dim i As Long, j As Long, k As Long, col As Long
    Dim rowNum As Long, sourceRow As Long, legendeRow As Long
    Dim matchedCount As Long, updatedCount As Long, newCount As Long, deletedCount As Long, quantityWarningCount As Long
    Dim oldIndex As Long
    Dim statusFlag As String
    Dim bg2ColIdx As Long, kategorieColIdx As Long
    Dim hasChanges As Boolean, quantityChanged As Boolean
    Dim startTime As Double
    
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Debug.Print String(80, "=")
    Debug.Print "PROCUREMENT LIST MERGER - VBA Macro"
    Debug.Print String(80, "=")
    
    ' 1. Find workbooks
    Debug.Print vbCrLf & "1. Looking for workbooks..."
    
    On Error Resume Next
    Set oldWb = Workbooks(OLD_WORKBOOK_NAME)
    Set newWb = Workbooks(NEW_WORKBOOK_NAME)
    On Error GoTo 0
    
    If oldWb Is Nothing Then
        MsgBox "ERROR: Workbook '" & OLD_WORKBOOK_NAME & "' not found!" & vbCrLf & vbCrLf & _
               "Please open the file and run the macro again.", vbCritical, "Error"
        GoTo Cleanup
    End If
    
    If newWb Is Nothing Then
        MsgBox "ERROR: Workbook '" & NEW_WORKBOOK_NAME & "' not found!" & vbCrLf & vbCrLf & _
               "Please open the file and run the macro again.", vbCritical, "Error"
        GoTo Cleanup
    End If
    
    Set oldWs = oldWb.Worksheets(1)  ' First sheet (generic)
    Set newWs = newWb.Worksheets(1)  ' First sheet (generic)
    
    Debug.Print "   ✓ Old list: " & oldWb.Name & " / " & oldWs.Name
    Debug.Print "   ✓ New list: " & newWb.Name & " / " & newWs.Name
    
    ' 2. Read data
    Debug.Print vbCrLf & "2. Reading data..."
    
    oldData = oldWs.UsedRange.Value2
    newData = newWs.UsedRange.Value2
    
    Debug.Print "   Old list: " & UBound(oldData, 1) - 1 & " rows"
    Debug.Print "   New list: " & UBound(newData, 1) - 1 & " rows"
    
    ' 3. Extract headers (first row)
    ' Get directly from data, not with Application.Index
    ReDim oldHeaders(1 To 1, 1 To UBound(oldData, 2)) As Variant
    ReDim newHeaders(1 To 1, 1 To UBound(newData, 2)) As Variant
    
    For i = 1 To UBound(oldData, 2)
        oldHeaders(1, i) = oldData(1, i)
    Next i
    
    For i = 1 To UBound(newData, 2)
        newHeaders(1, i) = newData(1, i)
    Next i
    
    ' 4. Merge all columns with smart positioning
    Debug.Print vbCrLf & "3. Creating column mapping..."
    
    Set allHeaders = New Collection
    Set headerDict = CreateObject("Scripting.Dictionary")
    
    ' Add new columns
    For i = 1 To UBound(newHeaders, 2)
        If Not IsEmpty(newHeaders(1, i)) Then
            allHeaders.Add newHeaders(1, i)
            headerDict(CStr(newHeaders(1, i))) = allHeaders.Count
        End If
    Next i
    
    ' Insert columns from old list intelligently (based on left neighbor)
    ' Go from left to right
    For i = 1 To UBound(oldHeaders, 2)
        If Not IsEmpty(oldHeaders(1, i)) Then
            Dim oldColName As String
            oldColName = CStr(oldHeaders(1, i))
            
            ' If column doesn't exist yet
            If Not headerDict.Exists(oldColName) Then
                
                ' Find left neighbor (the column to the left in old list)
                Dim leftNeighbor As String
                leftNeighbor = ""
                
                For j = i - 1 To 1 Step -1
                    If Not IsEmpty(oldHeaders(1, j)) Then
                        Dim candidateName As String
                        candidateName = CStr(oldHeaders(1, j))
                        
                        ' Is this neighbor present in the new list?
                        If headerDict.Exists(candidateName) Then
                            leftNeighbor = candidateName
                            Exit For
                        End If
                    End If
                Next j
                
                ' Insert column
                If leftNeighbor <> "" Then
                    ' Insert after left neighbor
                    Dim insertPos As Long
                    insertPos = headerDict(leftNeighbor)
                    
                    ' Build new collection with inserted column
                    Dim tempHeaders As Collection
                    Set tempHeaders = New Collection
                    
                    For k = 1 To allHeaders.Count
                        tempHeaders.Add allHeaders(k)
                        If k = insertPos Then
                            tempHeaders.Add oldColName
                            Debug.Print "   + Column '" & oldColName & "' inserted after '" & leftNeighbor & "'"
                        End If
                    Next k
                    
                    Set allHeaders = tempHeaders
                    
                    ' Rebuild dictionary
                    Set headerDict = CreateObject("Scripting.Dictionary")
                    For k = 1 To allHeaders.Count
                        headerDict(CStr(allHeaders(k))) = k
                    Next k
                Else
                    ' No left neighbor found -> insert at beginning
                    Set tempHeaders = New Collection
                    tempHeaders.Add oldColName
                    For k = 1 To allHeaders.Count
                        tempHeaders.Add allHeaders(k)
                    Next k
                    Set allHeaders = tempHeaders
                    
                    Set headerDict = CreateObject("Scripting.Dictionary")
                    For k = 1 To allHeaders.Count
                        headerDict(CStr(allHeaders(k))) = k
                    Next k
                    
                    Debug.Print "   + Column '" & oldColName & "' inserted at beginning"
                End If
            End If
        End If
    Next i
    
    ' Status flag column
    allHeaders.Add "_status_flag"
    headerDict("_status_flag") = allHeaders.Count

    ' Quantity changed flag column
    allHeaders.Add "_quantity_changed"
    headerDict("_quantity_changed") = allHeaders.Count
    
    Debug.Print vbCrLf & "✅ Matching key: File + " & Join(matchingAssemblyCols, ", ")

    ' 6. Matching Keys erstellen
    ReDim oldKeys(1 To UBound(oldData, 1) - 1)
    ReDim newKeys(1 To UBound(newData, 1) - 1)

    oldKeys = CreateMatchingKeysWithColumns(oldData, oldHeaders, matchingAssemblyCols)
    newKeys = CreateMatchingKeysWithColumns(newData, newHeaders, matchingAssemblyCols)
    
    ' 5. Prepare output data array
    Debug.Print vbCrLf & "5. Performing matching and transferring data..."
    
    ReDim outputData(1 To UBound(newData, 1), 1 To allHeaders.Count)
    
    ' Header row
    For i = 1 To allHeaders.Count
        outputData(1, i) = allHeaders(i)
    Next i
    
    ' Go through data rows
    matchedCount = 0
    updatedCount = 0
    newCount = 0
    
    For i = 2 To UBound(newData, 1)
        statusFlag = ""
        hasChanges = False
        quantityChanged = False
        
        ' Search for match in old list
        oldIndex = FindInArray(newKeys(i - 1), oldKeys)
        
        If oldIndex > 0 Then
            ' Match found
            matchedCount = matchedCount + 1
            
            ' Check quantity changes (Anzahl or Anzahl gesamt)
            Dim anzahlColIdx_new As Long, anzahlColIdx_old As Long
            Dim anzahlGesamtColIdx_new As Long, anzahlGesamtColIdx_old As Long
            anzahlColIdx_new = FindColumnIndex("Anzahl", newHeaders)
            anzahlColIdx_old = FindColumnIndex("Anzahl", oldHeaders)
            anzahlGesamtColIdx_new = FindColumnIndex("Anzahl gesamt", newHeaders)
            anzahlGesamtColIdx_old = FindColumnIndex("Anzahl gesamt", oldHeaders)
            
            ' Check Anzahl
            If anzahlColIdx_new > 0 And anzahlColIdx_old > 0 Then
                Dim oldAnzahl As Variant, newAnzahl As Variant
                oldAnzahl = oldData(oldIndex + 1, anzahlColIdx_old)
                newAnzahl = newData(i, anzahlColIdx_new)
                
                If Not IsEmpty(oldAnzahl) And Not IsEmpty(newAnzahl) Then
                    On Error Resume Next
                    If CDbl(oldAnzahl) <> CDbl(newAnzahl) Then
                        quantityChanged = True
                    End If
                    On Error GoTo 0
                End If
            End If
            
            ' Check Anzahl gesamt
            If Not quantityChanged And anzahlGesamtColIdx_new > 0 And anzahlGesamtColIdx_old > 0 Then
                Dim oldAnzahlGesamt As Variant, newAnzahlGesamt As Variant
                oldAnzahlGesamt = oldData(oldIndex + 1, anzahlGesamtColIdx_old)
                newAnzahlGesamt = newData(i, anzahlGesamtColIdx_new)
                
                If Not IsEmpty(oldAnzahlGesamt) And Not IsEmpty(newAnzahlGesamt) Then
                    On Error Resume Next
                    If CDbl(oldAnzahlGesamt) <> CDbl(newAnzahlGesamt) Then
                        quantityChanged = True
                    End If
                    On Error GoTo 0
                End If
            End If
            
            ' Check if Status is critical (ordered or higher)
            Dim statusCritical As Boolean
            statusCritical = False
            Dim statusColIdx_old As Long
            statusColIdx_old = FindColumnIndex("Status", oldHeaders)
            
            If statusColIdx_old > 0 Then
                Dim oldStatus As String
                oldStatus = LCase(Trim(CStr(oldData(oldIndex + 1, statusColIdx_old))))
                
                If oldStatus = "ordered" Or oldStatus = "paid" Or _
                   oldStatus = "delivered" Or oldStatus = "completed" Then
                    statusCritical = True
                End If
            End If
            
            ' Warn if quantity changed for ordered items
            If quantityChanged And statusCritical Then
                quantityWarningCount = quantityWarningCount + 1
            End If
            
            ' Go through all columns
            For j = 1 To allHeaders.Count - 2  ' -2 for _status_flag and _quantity_changed
                Dim colName As String
                colName = allHeaders(j)
                
                Dim newColIdx As Long, oldColIdx As Long
                newColIdx = FindColumnIndex(colName, newHeaders)
                oldColIdx = FindColumnIndex(colName, oldHeaders)
                
                ' Is it a transfer column?
                If IsInArray(colName, transferCols) And oldColIdx > 0 Then
                    ' Get value from old list
                    Dim oldValue As Variant
                    oldValue = oldData(oldIndex + 1, oldColIdx)
                    
                    ' Transfer if not empty
                    If Not IsEmpty(oldValue) And Trim(CStr(oldValue)) <> "" Then
                        outputData(i, j) = oldValue
                        hasChanges = True
                    ElseIf newColIdx > 0 Then
                        outputData(i, j) = newData(i, newColIdx)
                    Else
                        outputData(i, j) = ""
                    End If
                ElseIf newColIdx > 0 Then
                    ' Normal column from new list
                    outputData(i, j) = newData(i, newColIdx)
                Else
                    outputData(i, j) = ""
                End If
            Next j
            
            If hasChanges Then
                statusFlag = "UPDATED"
                updatedCount = updatedCount + 1
            End If
        Else
            ' No match - new row
            newCount = newCount + 1
            statusFlag = "NEW"
            
            For j = 1 To allHeaders.Count - 1  ' -1 because of _status_flag
                colName = allHeaders(j)
                newColIdx = FindColumnIndex(colName, newHeaders)
                
                If newColIdx > 0 Then
                    outputData(i, j) = newData(i, newColIdx)
                    
                    ' Set Status to "neu" if it's empty
                    If colName = "Status" Then
                        If IsEmpty(outputData(i, j)) Or Trim(CStr(outputData(i, j))) = "" Then
                            outputData(i, j) = "new"
                        End If
                    End If
                Else
                    outputData(i, j) = ""
                End If
            Next j
        End If
        
        ' Set status flag (second-to-last column)
        outputData(i, allHeaders.Count - 1) = statusFlag
        
        ' Set quantity changed flag (last column)
        outputData(i, allHeaders.Count) = quantityChanged
    Next i
    
    ' 7. Gelöschte Items finden
    Debug.Print vbCrLf & "6. Finding deleted entries..."
    
    Dim deletedData() As Variant
    Dim deletedRows As Collection
    Set deletedRows = New Collection
    
    For i = 1 To UBound(oldKeys)
        If FindInArray(oldKeys(i), newKeys) = 0 Then
            deletedRows.Add i + 1  ' +1 for header row
        End If
    Next i
    
    deletedCount = deletedRows.Count
    Debug.Print "   Found: " & deletedCount & " deleted entries"
    
    ' 8. Output-Arbeitsmappe erstellen
    Debug.Print vbCrLf & "7. Creating output workbook..."
    
    Set outputWb = Workbooks.Add
    Set mainWs = outputWb.Worksheets(1)
    mainWs.Name = "Main_List"
    
    ' Write data
    mainWs.Range(mainWs.Cells(1, 1), mainWs.Cells(UBound(outputData, 1), UBound(outputData, 2))).Value2 = outputData
    
    ' Find Status column index early
    Dim statusColIdx As Long
    statusColIdx = FindColumnIndex("Status", Application.Index(outputData, 1, 0))
    
    ' Set empty Status fields to "-"
    If statusColIdx > 0 Then
        Dim r As Long
        For r = 2 To UBound(outputData, 1)
            If IsEmpty(mainWs.Cells(r, statusColIdx).Value) Or Trim(CStr(mainWs.Cells(r, statusColIdx).Value)) = "" Then
                mainWs.Cells(r, statusColIdx).Value = "-"
            End If
        Next r
    End If

    ' Restore formulas from new workbook.
    ' UsedRange.Cells(r,c) is used so the coordinates are correct even when the
    ' sheet's used range does not start at A1.
    ' Formula cells always win — if the new workbook has a formula, the output
    ' gets that formula regardless of whether a transferred old value was written.
    Debug.Print vbCrLf & "8b. Restoring formulas from new workbook..."

    Dim newToOutCol() As Long
    ReDim newToOutCol(1 To UBound(newHeaders, 2))
    Dim nc As Long
    For nc = 1 To UBound(newHeaders, 2)
        If Not IsEmpty(newHeaders(1, nc)) Then
            Dim nh As String
            nh = CStr(newHeaders(1, nc))
            If headerDict.Exists(nh) Then
                newToOutCol(nc) = headerDict(nh)
            End If
        End If
    Next nc

    Dim formulaCount As Long
    formulaCount = 0
    Dim fRow As Long, fCol As Long
    For fRow = 2 To UBound(newData, 1)
        For fCol = 1 To UBound(newHeaders, 2)
            If newToOutCol(fCol) > 0 Then
                If newWs.UsedRange.Cells(fRow, fCol).HasFormula Then
                    mainWs.Cells(fRow, newToOutCol(fCol)).FormulaR1C1 = _
                        newWs.UsedRange.Cells(fRow, fCol).FormulaR1C1
                    formulaCount = formulaCount + 1
                End If
            End If
        Next fCol
    Next fRow
    Debug.Print "   ✓ " & formulaCount & " formulas restored"

    ' Copy column number formats from new workbook (e.g. accounting, date, number).
    ' Uses the first data row of each column as representative format.
    Dim colFmt As String, fmtColCount As Long
    fmtColCount = 0
    For nc = 1 To UBound(newHeaders, 2)
        If newToOutCol(nc) > 0 Then
            colFmt = newWs.UsedRange.Cells(2, nc).NumberFormat
            If colFmt <> "General" And colFmt <> "" Then
                mainWs.Range(mainWs.Cells(2, newToOutCol(nc)), _
                             mainWs.Cells(UBound(outputData, 1), newToOutCol(nc))).NumberFormat = colFmt
                fmtColCount = fmtColCount + 1
            End If
        End If
    Next nc
    Debug.Print "   ✓ " & fmtColCount & " column number formats applied"

    ' 9. Apply formatting
    Debug.Print vbCrLf & "8. Applying formatting..."
    
    ' Format header row
    With mainWs.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(64, 64, 64)      ' Dark gray
        .Font.Color = RGB(255, 255, 255)       ' White text
    End With
    
    ' Find BG2 and Kategorie columns for Baugruppe detection
    bg2ColIdx = FindColumnIndex(ASSEMBLY_COL, Application.Index(outputData, 1, 0))
    kategorieColIdx = FindColumnIndex(CATEGORY_COL, Application.Index(outputData, 1, 0))
    
    ' Conditional formatting based on Status column
    
    If statusColIdx > 0 Then
        Debug.Print "   Adding conditional formatting (Status column: " & statusColIdx & ")..."
        
        Dim statusCol As String
        Dim filenameCol As String
        Dim lastRow As Long
        Dim dataRange As Range
        
        statusCol = Split(mainWs.Cells(1, statusColIdx).Address, "$")(1)
        lastRow = UBound(outputData, 1)
        Set dataRange = mainWs.Range("A2:" & Split(mainWs.Cells(lastRow, allHeaders.Count).Address, "$")(1) & lastRow)
        
        ' Find quantity_changed column
        Dim quantityChangedColIdx As Long, quantityChangedCol As String
        quantityChangedColIdx = FindColumnIndex("_quantity_changed", Application.Index(outputData, 1, 0))
        
        ' Delete all existing conditional formatting
        dataRange.FormatConditions.Delete
        
        ' Rule 0: Quantity changed for ordered items (LIGHT RED) - HIGHEST PRIORITY
        If quantityChangedColIdx > 0 Then
            quantityChangedCol = Split(mainWs.Cells(1, quantityChangedColIdx).Address, "$")(1)
            With dataRange.FormatConditions.Add(xlExpression, , "=$" & quantityChangedCol & "2")
                .Interior.Color = RGB(255, 179, 179)  ' Light Red (Salmon Pink)
                .Font.Bold = True
                .StopIfTrue = True
            End With
        End If
        
        ' Rule 1: new (Coral) - highest priority
        With dataRange.FormatConditions.Add(xlExpression, , "=$" & statusCol & "2=""new""")
            .Interior.Color = RGB(255, 204, 179)  ' Light Coral
            .StopIfTrue = True
        End With
        
        ' Rule 2: requested (Light Yellow)
        With dataRange.FormatConditions.Add(xlExpression, , "=$" & statusCol & "2=""requested""")
            .Interior.Color = RGB(255, 255, 204)  ' Light Yellow
            .StopIfTrue = True
        End With
        
        ' Rule 3: offered (Lavender)
        With dataRange.FormatConditions.Add(xlExpression, , "=($" & statusCol & "2=""offered"")+($" & statusCol & "2=""anbeboten"")")
            .Interior.Color = RGB(221, 204, 255)  ' Lavender
            .StopIfTrue = True
        End With
        
        ' Rule 4: ordered (Light Blue)
        With dataRange.FormatConditions.Add(xlExpression, , "=$" & statusCol & "2=""ordered""")
            .Interior.Color = RGB(173, 216, 230)  ' Light Blue
            .StopIfTrue = True
        End With
        
        ' Rule 5: paid (Mint Green)
        With dataRange.FormatConditions.Add(xlExpression, , "=$" & statusCol & "2=""paid""")
            .Interior.Color = RGB(198, 239, 206)  ' Mint Green
            .StopIfTrue = True
        End With
        
        ' Rule 6: delivered (Sage Green)
        With dataRange.FormatConditions.Add(xlExpression, , "=$" & statusCol & "2=""delivered""")
            .Interior.Color = RGB(147, 196, 125)  ' Sage Green
            .StopIfTrue = True
        End With
        
        ' Rule 7: completed (Forest Green + Bold)
        With dataRange.FormatConditions.Add(xlExpression, , "=$" & statusCol & "2=""completed""")
            .Interior.Color = RGB(106, 168, 79)   ' Forest Green
            .Font.Bold = True
            .StopIfTrue = True
        End With
        
        ' Rule 8: postponed (Gray + Italic)
        With dataRange.FormatConditions.Add(xlExpression, , "=$" & statusCol & "2=""postponed""")
            .Interior.Color = RGB(217, 217, 217)  ' Light Gray
            .Font.Italic = True
            .StopIfTrue = True
        End With
        
        Debug.Print "   ✓ " & dataRange.FormatConditions.Count & " conditional formatting rules added"
    End If
    
    ' Auto-fit column widths
    mainWs.Cells.EntireColumn.AutoFit
    
    ' 10. Status-Spalte als Dropdown einrichten
    If statusColIdx > 0 Then
        Debug.Print "   Setting up dropdown for Status column..."
        
        Dim statusRange As Range
        Set statusRange = mainWs.Range(mainWs.Cells(2, statusColIdx), mainWs.Cells(UBound(outputData, 1), statusColIdx))
        
        With statusRange.Validation
            .Delete  ' Delete existing validation
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, _
                 Formula1:="-,new,requested,offered,ordered,paid,delivered,completed,postponed"
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = False
        End With
    End If
    
    ' 11. Activate AutoFilter
    Debug.Print "   Activating AutoFilter..."
    mainWs.Range(mainWs.Cells(1, 1), mainWs.Cells(1, allHeaders.Count)).AutoFilter
    
    ' 12. Deleted items sheet
    If deletedCount > 0 Then
        Debug.Print vbCrLf & "9. Creating sheet for deleted entries..."
        
        Set deletedWs = outputWb.Worksheets.Add(After:=mainWs)
        deletedWs.Name = "Deleted_Entries"
        
        ' Write headers
        For i = 1 To UBound(oldHeaders, 2)
            deletedWs.Cells(1, i).Value2 = oldHeaders(1, i)
        Next i
        
        ' Format headers
        With deletedWs.Rows(1)
            .Font.Bold = True
            .Interior.Color = RGB(211, 211, 211)
        End With
        
        ' Write deleted rows
        rowNum = 2
        For i = 1 To deletedRows.Count
            sourceRow = deletedRows(i)
            
            For j = 1 To UBound(oldHeaders, 2)
                deletedWs.Cells(rowNum, j).Value2 = oldData(sourceRow, j)
            Next j
            
            deletedWs.Rows(rowNum).Interior.Color = RGB(255, 182, 193)  ' Red
            rowNum = rowNum + 1
        Next i
        
        deletedWs.Cells.EntireColumn.AutoFit
    End If
    
    ' 13. Create log sheet
    Debug.Print vbCrLf & "10. Creating log..."
    
    Set logWs = outputWb.Worksheets.Add(After:=outputWb.Worksheets(outputWb.Worksheets.Count))
    logWs.Name = "Log"
    
    With logWs
        .Cells(1, 1).Value2 = "PROCUREMENT LIST MERGE LOG"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
        
        .Cells(3, 1).Value2 = "Datum:"
        .Cells(3, 2).Value2 = Format(Now, "dd.mm.yyyy hh:nn:ss")
        
        .Cells(4, 1).Value2 = "Old List:"
        .Cells(4, 2).Value2 = oldWb.Name & " / " & oldWs.Name
        
        .Cells(5, 1).Value2 = "New List:"
        .Cells(5, 2).Value2 = newWb.Name & " / " & newWs.Name
        
        .Cells(7, 1).Value2 = "STATISTICS"
        .Cells(7, 1).Font.Bold = True
        
        .Cells(8, 1).Value2 = "Matched Rows:"
        .Cells(8, 2).Value2 = matchedCount
        
        .Cells(9, 1).Value2 = "Of which updated:"
        .Cells(9, 2).Value2 = updatedCount
        
        .Cells(10, 1).Value2 = "New Entries:"
        .Cells(10, 2).Value2 = newCount
        
        .Cells(11, 1).Value2 = "Deleted Entries:"
        .Cells(11, 2).Value2 = deletedCount
        
        If quantityWarningCount > 0 Then
            .Cells(12, 1).Value2 = "⚠️ Quantity Changes (ordered):"
            .Cells(12, 2).Value2 = quantityWarningCount
            .Cells(12, 1).Font.Bold = True
            .Cells(12, 2).Font.Bold = True
        End If
        
        .Cells(13, 1).Value2 = "Processing Time:"
        .Cells(13, 2).Value2 = Round(Timer - startTime, 2) & " seconds"
        
        .Cells(15, 1).Value2 = "COLOR LEGEND (by Status)"
        .Cells(15, 1).Font.Bold = True
        
        legendeRow = 16
        
        .Cells(legendeRow, 1).Value2 = "⚪ White"
        .Cells(legendeRow, 2).Value2 = "Not yet touched"
        ' White - no formatting
        
        legendeRow = legendeRow + 1
        .Cells(legendeRow, 1).Value2 = "🟠 Coral"
        .Cells(legendeRow, 2).Value2 = "requested"
        .Cells(legendeRow, 1).Interior.Color = RGB(255, 160, 122)
        
        legendeRow = legendeRow + 1
        .Cells(legendeRow, 1).Value2 = "🟣 Light Purple"
        .Cells(legendeRow, 2).Value2 = "offered"
        .Cells(legendeRow, 1).Interior.Color = RGB(221, 204, 255)
        
        legendeRow = legendeRow + 1
        .Cells(legendeRow, 1).Value2 = "🔵 Light Blue"
        .Cells(legendeRow, 2).Value2 = "ordered"
        .Cells(legendeRow, 1).Interior.Color = RGB(173, 216, 230)
        
        legendeRow = legendeRow + 1
        .Cells(legendeRow, 1).Value2 = "🟢 Light Green"
        .Cells(legendeRow, 2).Value2 = "paid"
        .Cells(legendeRow, 1).Interior.Color = RGB(198, 239, 206)
        
        legendeRow = legendeRow + 1
        .Cells(legendeRow, 1).Value2 = "🟢 Medium Green"
        .Cells(legendeRow, 2).Value2 = "delivered"
        .Cells(legendeRow, 1).Interior.Color = RGB(147, 196, 125)
        
        legendeRow = legendeRow + 1
        .Cells(legendeRow, 1).Value2 = "🟢 Dark Green"
        .Cells(legendeRow, 2).Value2 = "completed"
        .Cells(legendeRow, 1).Interior.Color = RGB(106, 168, 79)
        .Cells(legendeRow, 1).Font.Bold = True
        
        legendeRow = legendeRow + 1
        .Cells(legendeRow, 1).Value2 = "⚫ Gray"
        .Cells(legendeRow, 2).Value2 = "postponed"
        .Cells(legendeRow, 1).Interior.Color = RGB(217, 217, 217)
        .Cells(legendeRow, 1).Font.Italic = True
        
        .Columns("A:B").AutoFit
    End With
    
    ' 14. Format Baugruppe rows LAST (overwrites all conditional formatting)
    ' CRITICAL: Do this as the LAST step before activating the sheet!
    If bg2ColIdx > 0 And kategorieColIdx > 0 Then
        Debug.Print vbCrLf & "11. Formatting assembly rows (gray) - FINAL STEP..."
        For rowNum = 2 To UBound(outputData, 1)
            Dim bg2Val As Variant, katVal As String
            bg2Val = mainWs.Cells(rowNum, bg2ColIdx).Value
            katVal = LCase(Trim(CStr(mainWs.Cells(rowNum, kategorieColIdx).Value)))
            
            If (IsEmpty(bg2Val) Or Trim(CStr(bg2Val)) = "") And katVal = "baugruppe" Then
                ' This is a Baugruppe header row - gray all columns + white text
                Dim colNum As Long
                For colNum = 1 To allHeaders.Count
                    mainWs.Cells(rowNum, colNum).Font.Bold = True
                    mainWs.Cells(rowNum, colNum).Font.Color = RGB(255, 255, 255)  ' White text
                    mainWs.Cells(rowNum, colNum).Interior.Color = RGB(165, 165, 165)  ' Gray background
                Next colNum
            End If
        Next rowNum
        Debug.Print "   ✓ Assemblies formatted (overwrites conditional formatting)"
    End If
    
    ' 15. Done!
    Debug.Print vbCrLf & String(80, "=")
    Debug.Print "DONE!"
    Debug.Print String(80, "=")
    Debug.Print "Matched Rows:        " & matchedCount
    Debug.Print "Of which updated:    " & updatedCount
    Debug.Print "New Entries:         " & newCount
    Debug.Print "Deleted Entries:     " & deletedCount
    If quantityWarningCount > 0 Then
        Debug.Print ""
        Debug.Print "⚠️  WARNING: " & quantityWarningCount & " quantity changes in ordered parts!"
        Debug.Print "   → These rows are marked LIGHT RED and need manual review"
    End If
    Debug.Print "Process time:    " & Round(Timer - startTime, 2) & " seconds"
    
    ' Activate main list
    mainWs.Activate
    mainWs.Range("A1").Select
    
    Dim warningMsg As String
    warningMsg = ""
    If quantityWarningCount > 0 Then
        warningMsg = vbCrLf & vbCrLf & _
                     "⚠️  WARNING: " & quantityWarningCount & " quantity changes in ordered parts!" & vbCrLf & _
                     "→ These rows are marked LIGHT RED"
    End If
    
    MsgBox "✅ Merge successful!" & vbCrLf & vbCrLf & _
           "Matched Rows:        " & matchedCount & vbCrLf & _
           "Of which updated:    " & updatedCount & vbCrLf & _
           "New Entries:         " & newCount & vbCrLf & _
           "Deleted Entries:     " & deletedCount & warningMsg & vbCrLf & vbCrLf & _
           "COLOR CODING (by Status column):" & vbCrLf & _
           "⚪ White      = not yet touched" & vbCrLf & _
           "🟠 Coral     = requested" & vbCrLf & _
           "🟣 Purple    = offered" & vbCrLf & _
           "🔵 Light Blue  = ordered" & vbCrLf & _
           "🟢 Green      = paid / delivered / completed" & vbCrLf & _
           "⚫ Gray      = postponed", _
           vbInformation, "Done!"
    
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

'==============================================================================
' HELPER FUNCTIONS
'==============================================================================


Function FindColumnIndex(ByVal colName As String, headers As Variant) As Long
    Dim i As Long
    Dim numDimensions As Long
    
    ' Check if it's a 1D or 2D array
    On Error Resume Next
    numDimensions = 0
    If IsArray(headers) Then
        ' Try to get second dimension
        Dim test As Long
        test = UBound(headers, 2)
        If Err.Number = 0 Then
            numDimensions = 2
        Else
            numDimensions = 1
        End If
    End If
    On Error GoTo 0
    
    If numDimensions = 2 Then
        ' 2D array (row, column)
        For i = 1 To UBound(headers, 2)
            If CStr(headers(1, i)) = colName Then
                FindColumnIndex = i
                Exit Function
            End If
        Next i
    ElseIf numDimensions = 1 Then
        ' 1D array
        For i = LBound(headers) To UBound(headers)
            If CStr(headers(i)) = colName Then
                FindColumnIndex = i - LBound(headers) + 1
                Exit Function
            End If
        Next i
    End If
    
    FindColumnIndex = 0
End Function

Function FindInArray(searchValue As String, arr() As String) As Long
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = searchValue Then
            FindInArray = i
            Exit Function
        End If
    Next i
    FindInArray = 0
End Function

Function IsInArray(searchValue As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If CStr(arr(i)) = searchValue Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function


Function CreateMatchingKeysWithColumns(data As Variant, headers As Variant, columns As Variant) As String()
    Dim keys() As String
    Dim i As Long, j As Long
    Dim dateiIdx As Long
    Dim datei As String
    Dim colIdx As Long
    Dim colValue As String
    
    ' Find file column
    dateiIdx = FindColumnIndex(FILE_COL, headers)
    
    ' Create keys
    ReDim keys(1 To UBound(data, 1) - 1)
    
    For i = 2 To UBound(data, 1)
        ' Start with file
        datei = IIf(dateiIdx > 0, CStr(data(i, dateiIdx)), "NO_FILE")
        If datei = "" Then datei = "NO_FILE"
        
        keys(i - 1) = datei
        
        ' Add specified columns
        If UBound(columns) >= LBound(columns) Then
            For j = LBound(columns) To UBound(columns)
                colIdx = FindColumnIndex(columns(j), headers)
                colValue = ""
                If colIdx > 0 Then
                    If Not IsEmpty(data(i, colIdx)) Then
                        colValue = CStr(data(i, colIdx))
                    End If
                End If
                keys(i - 1) = keys(i - 1) & "|" & colValue
            Next j
        End If
    Next i
    
    CreateMatchingKeysWithColumns = keys
End Function


