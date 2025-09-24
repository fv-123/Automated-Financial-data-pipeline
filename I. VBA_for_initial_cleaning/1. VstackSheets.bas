Attribute VB_Name = "Module2"
Option Explicit

Sub BuildAndCleanMaster_WithCashRule()
    Dim wsMaster As Worksheet
    Dim wbSource As Workbook
    Dim ws As Worksheet
    Dim fDialog As FileDialog
    Dim lastRow As Long, lastCol As Long, pasteRow As Long
    Dim i As Long, j As Long
    Dim selectedCount As Long
    
    Dim statementOrder As Variant
    Dim statementIndex As Long
    Dim cellValue As String
    Dim statementCol As Long
    Dim foundDate As Boolean
    
    ' speed up
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' === 1. Create / Reset Master sheet ===
    On Error Resume Next
    Set wsMaster = ThisWorkbook.Sheets("Master")
    If wsMaster Is Nothing Then
        Set wsMaster = ThisWorkbook.Sheets.Add
        wsMaster.Name = "Master"
    Else
        wsMaster.Cells.Clear
    End If
    On Error GoTo 0
    
    pasteRow = 1
    
    ' === 2. Pick multiple Excel files ===
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .Title = "Select Excel Files to Stack"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm"
    End With
    
    If fDialog.Show <> -1 Then
        GoTo Finish
    End If
    
    selectedCount = fDialog.SelectedItems.Count
    
    ' === 3. Stack all sheets from selected files ===
    For i = 1 To selectedCount
        Set wbSource = Workbooks.Open(fDialog.SelectedItems(i), ReadOnly:=True)
        
        For Each ws In wbSource.Sheets
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            
            If lastRow > 0 And lastCol > 0 Then
                ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Copy
                wsMaster.Cells(pasteRow, 1).PasteSpecial xlPasteValues
                pasteRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row + 1
            End If
        Next ws
        
        wbSource.Close SaveChanges:=False
    Next i
    
    Application.CutCopyMode = False
    
    ' === 4. Set column headers / structure ===
    wsMaster.Cells(1, 1).Value = "Indicator"
    
    ' Delete column B (shift left)
    On Error Resume Next
    wsMaster.Columns("B:B").Delete
    On Error GoTo 0
    
    ' New col B = Unit
    wsMaster.Cells(1, 2).Value = "Unit"
    
    ' Add Statement column at the end and remember its index
    lastCol = wsMaster.Cells(1, wsMaster.Columns.Count).End(xlToLeft).Column
    wsMaster.Cells(1, lastCol + 1).Value = "Statement"
    statementCol = lastCol + 1
    
    ' === 5. Delete junk rows (case-insensitive) ===
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row
    For i = lastRow To 2 Step -1
        cellValue = LCase(Trim(CStr(wsMaster.Cells(i, 1).Value)))
        If cellValue = "" _
        Or cellValue = "period" _
        Or cellValue = "consolidated" _
        Or cellValue = "audited" _
        Or cellValue = "audit firm" _
        Or cellValue = "audit opinion" Then
            wsMaster.Rows(i).Delete
        End If
    Next i
    
    ' recompute bounds before statement fill
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row
    lastCol = wsMaster.Cells(1, wsMaster.Columns.Count).End(xlToLeft).Column
    ' statementCol should still be lastCol (just in case)
    statementCol = lastCol
    
    ' === 6. Fill Statement column with cycling logic (with Cash-flow exception) ===
    statementOrder = Array("Balance Sheet", "Income Statement", "Cash Flow Statement", "Ratios")
    statementIndex = 0  ' Start with Balance Sheet
    
    For i = 2 To lastRow
        foundDate = False
        ' scan across all columns except the Statement column
        For j = 1 To statementCol - 1
            cellValue = LCase(Trim(CStr(wsMaster.Cells(i, j).Value)))
            ' look for quarter pattern like Q1/2020 (case-insensitive; allow surrounding text)
            If cellValue Like "*q[1-4]/2[0-9][0-9][0-9]*" Then
                foundDate = True
                Exit For
            End If
        Next j
        
        If foundDate Then
            ' Cash-flow exception: if Indicator (col A) contains "cash", force/keep Cash Flow
            If InStr(1, LCase(Trim(CStr(wsMaster.Cells(i, 1).Value))), "cash", vbTextCompare) > 0 Then
                statementIndex = 2 ' Cash Flow index in statementOrder
            Else
                ' normal advance
                statementIndex = (statementIndex + 1) Mod (UBound(statementOrder) + 1)
            End If
        End If
        
        wsMaster.Cells(i, statementCol).Value = statementOrder(statementIndex)
    Next i
    
    ' === 7. Delete rows containing quarterly date patterns (except header) ===
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row
    For i = lastRow To 2 Step -1
        For j = 1 To statementCol - 1
            cellValue = LCase(Trim(CStr(wsMaster.Cells(i, j).Value)))
            If cellValue Like "*q[1-4]/2[0-9][0-9][0-9]*" Then
                wsMaster.Rows(i).Delete
                Exit For
            End If
        Next j
    Next i
    
    MsgBox "Master file processed: " & selectedCount & " files stacked, cleaned, and organized (cash-flow rule applied).", vbInformation

Finish:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub


