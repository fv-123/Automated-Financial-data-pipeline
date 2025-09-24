Attribute VB_Name = "Module1"
Option Explicit

' ============================
' Run everything
' ============================
Sub RunAll()
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    DeleteRatiosColumns
    AddTickerColumn
    CreateMasterVStack
    CopyMasterToStatic

Cleanup:
    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "RunAll"
    Resume Cleanup
End Sub

' ============================
' 1) Delete column D cells from the row that CONTAINS "Ratios -" downwards and SHIFT LEFT
' ============================
Sub DeleteRatiosColumns()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim foundRow As Long
    Dim cellVal As String

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Master" And ws.Name <> "Master_Static" Then
            With ws
                lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
                foundRow = 0

                For r = 1 To lastRow
                    If Not IsError(.Cells(r, "A").Value) Then
                        cellVal = CStr(.Cells(r, "A").Value)
                        If InStr(1, cellVal, "Ratios -", vbTextCompare) > 0 Then
                            foundRow = r
                            Exit For
                        End If
                    End If
                Next r

                If foundRow > 0 Then
                    ' Delete the cells in column D from foundRow to lastRow and shift left
                    .Range(.Cells(foundRow, "D"), .Cells(lastRow, "D")).Delete Shift:=xlToLeft
                End If
            End With
        End If
    Next ws
End Sub

' ============================
' 2) Add "Ticker" column at end of each sheet filled with A1
' ============================
Sub AddTickerColumn()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim tickerValue As Variant

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Master" And ws.Name <> "Master_Static" Then
            With ws
                lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
                lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
                tickerValue = .Range("A1").Value

                .Cells(1, lastCol + 1).Value = "Ticker"
                If lastRow >= 2 Then
                    .Range(.Cells(2, lastCol + 1), .Cells(lastRow, lastCol + 1)).Value = tickerValue
                End If
            End With
        End If
    Next ws
End Sub

' ============================
' 3) Create Master sheet with VSTACK of each sheet from A1 to that sheet's last used cell (row & column)
' ============================
Sub CreateMasterVStack()
    Dim ws As Worksheet
    Dim masterWs As Worksheet
    Dim parts As Collection
    Dim formulaStr As String
    Dim sheetRange As String
    Dim lastRow As Long, lastCol As Long
    Dim rngTmp As Range

    ' remove old Master if present
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Master").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set masterWs = ThisWorkbook.Worksheets.Add
    masterWs.Name = "Master"

    Set parts = New Collection

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Master" And ws.Name <> "Master_Static" Then
            With ws
                If Application.WorksheetFunction.CountA(.Cells) = 0 Then
                    ' skip empty sheet
                Else
                    Set rngTmp = .Cells.Find(What:="*", After:=.Cells(1, 1), LookIn:=xlFormulas, _
                                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
                    If rngTmp Is Nothing Then
                        lastRow = 1
                    Else
                        lastRow = rngTmp.Row
                    End If

                    Set rngTmp = .Cells.Find(What:="*", After:=.Cells(1, 1), LookIn:=xlFormulas, _
                                LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
                    If rngTmp Is Nothing Then
                        lastCol = 1
                    Else
                        lastCol = rngTmp.Column
                    End If

                    sheetRange = "'" & Replace(ws.Name, "'", "''") & "'!" & .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    parts.Add sheetRange
                End If
            End With
        End If
    Next ws

    If parts.Count = 0 Then
        masterWs.Range("A1").Value = "No data to VSTACK"
        Exit Sub
    End If

    formulaStr = "=VSTACK(" & Join(CollectionToArray(parts), ",") & ")"

    ' Place the VSTACK formula into A1
    masterWs.Range("A1").Formula2 = formulaStr
End Sub

' helper: convert collection to array for Join()
Private Function CollectionToArray(col As Collection) As Variant
    Dim arr() As String
    Dim i As Long
    If col Is Nothing Then
        CollectionToArray = Array()
        Exit Function
    End If
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count
        arr(i) = col(i)
    Next i
    CollectionToArray = arr
End Function

' ============================
' 4) Copy Master (values only) to Master_Static properly (writes spilled values directly)
' ============================
Sub CopyMasterToStatic()
    Dim masterWs As Worksheet
    Dim staticWs As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rngToCopy As Range
    Dim fileName As String

    ' remove old Master_Static if present
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Master_Static").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' reference Master
    On Error Resume Next
    Set masterWs = ThisWorkbook.Worksheets("Master")
    On Error GoTo 0
    If masterWs Is Nothing Then
        MsgBox "Master sheet not found. Run CreateMasterVStack first.", vbExclamation
        Exit Sub
    End If

    If Application.WorksheetFunction.CountA(masterWs.Cells) = 0 Then
        MsgBox "Master sheet is empty.", vbExclamation
        Exit Sub
    End If

    ' determine last used cell by VALUES (captures spilled results)
    On Error Resume Next
    lastRow = masterWs.Cells.Find(What:="*", After:=masterWs.Cells(1, 1), LookIn:=xlValues, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = masterWs.Cells.Find(What:="*", After:=masterWs.Cells(1, 1), LookIn:=xlValues, _
                LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    On Error GoTo 0

    If lastRow = 0 Or lastCol = 0 Then
        MsgBox "Couldn't determine Master range.", vbExclamation
        Exit Sub
    End If

    Set rngToCopy = masterWs.Range(masterWs.Cells(1, 1), masterWs.Cells(lastRow, lastCol))

    ' create new static sheet
    Set staticWs = ThisWorkbook.Worksheets.Add
    staticWs.Name = "Master_Static"

    ' transfer values directly (no copy/paste)
    staticWs.Range("A1").Resize(rngToCopy.Rows.Count, rngToCopy.Columns.Count).Value = rngToCopy.Value

    ' rename first column header
    staticWs.Cells(1, 1).Value = "Indicator"

    ' add Industry column (file name without extension or "Bank")
    lastCol = staticWs.Cells(1, staticWs.Columns.Count).End(xlToLeft).Column
    staticWs.Cells(1, lastCol + 1).Value = "Industry"

    lastRow = staticWs.Cells(staticWs.Rows.Count, 1).End(xlUp).Row

    If ThisWorkbook.Path <> "" Then
        fileName = ThisWorkbook.Name
        ' strip extension
        If InStrRev(fileName, ".") > 0 Then
            fileName = Left(fileName, InStrRev(fileName, ".") - 1)
        End If
    Else
        fileName = "Bank"
    End If

    If lastRow >= 2 Then
        staticWs.Range(staticWs.Cells(2, lastCol + 1), staticWs.Cells(lastRow, lastCol + 1)).Value = fileName
    End If
End Sub


