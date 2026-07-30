Attribute VB_Name = "modScriptRecycle"
Option Explicit

Private Const RECYCLE_START_ROW As Long = 900
Private Const RECYCLE_MAX_ITEMS As Long = 20
Private Const STORE_SHEET_NAME As String = "_ScriptRecycleStore"

Public Sub InitializeRecycleArea(ByVal ws As Worksheet)
    With ws.Range("A" & RECYCLE_START_ROW & ":G" & RECYCLE_START_ROW)
        .Merge
        .Value = "RECYCLE AREA | Select an item below and run Restore Recycle Item."
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(172, 117, 94)
    End With
    ws.Range("A" & RECYCLE_START_ROW + 1 & ":G" & RECYCLE_START_ROW + RECYCLE_MAX_ITEMS).Clear
    ws.Rows(RECYCLE_START_ROW + 1 & ":" & RECYCLE_START_ROW + RECYCLE_MAX_ITEMS).Rows.Group
    ws.Outline.ShowLevels RowLevels:=1, ColumnLevels:=2
End Sub

Public Sub DeleteSelectedParagraph()
    Dim paragraphRow As Long
    paragraphRow = SelectedParagraphRow(ActiveSheet, ActiveCell.Row)
    If paragraphRow = 0 Then
        ShowUserError "Select a paragraph block before deleting it.", "Delete Paragraph requires a selected paragraph block."
        Exit Sub
    End If
    StoreRecycleItem ActiveSheet, "Paragraph", paragraphRow, 3
    ActiveSheet.Rows(paragraphRow & ":" & paragraphRow + 2).Delete
    RefreshRecycleArea ActiveSheet
End Sub

Public Sub DeleteSelectedChapter()
    Dim chapterRow As Long, nextChapterRow As Long, rowCount As Long
    chapterRow = SelectedChapterRow(ActiveSheet, ActiveCell.Row)
    If chapterRow = 0 Then
        ShowUserError "Select a chapter block before deleting it.", "Delete Chapter requires a selected chapter block."
        Exit Sub
    End If
    nextChapterRow = NextChapterRow(ActiveSheet, chapterRow)
    rowCount = nextChapterRow - chapterRow
    StoreRecycleItem ActiveSheet, "Chapter", chapterRow, rowCount
    ActiveSheet.Rows(chapterRow & ":" & chapterRow + rowCount - 1).Delete
    RefreshRecycleArea ActiveSheet
End Sub

Public Sub RestoreSelectedRecycleItem()
    Dim itemId As String
    itemId = Trim$(CStr(ActiveSheet.Cells(ActiveCell.Row, "A").Value))
    If Len(itemId) = 0 Then
        ShowUserError "Select a recycle item row before restoring it.", "No recycle item ID was found in the selected row."
        Exit Sub
    End If
    If Not RestoreRecycleItem(ActiveSheet, itemId) Then ShowUserError "The selected recycle item could not be restored.", "Recycle restoration failed for " & itemId
End Sub

Public Function RestoreRecycleItem(ByVal ws As Worksheet, ByVal itemId As String) As Boolean
    Dim store As Worksheet, found As Range, originalRow As Long, rowCount As Long
    On Error GoTo Failed
    Set store = GetRecycleStoreSheet()
    Set found = store.Columns("A").Find(What:=itemId, LookIn:=xlValues, LookAt:=xlWhole)
    If found Is Nothing Then GoTo Failed
    originalRow = CLng(store.Cells(found.Row, "D").Value)
    rowCount = CLng(store.Cells(found.Row, "E").Value)
    ws.Rows(originalRow & ":" & originalRow + rowCount - 1).Insert Shift:=xlDown
    store.Range(store.Cells(found.Row, "J"), store.Cells(found.Row + rowCount - 1, "P")).Copy Destination:=ws.Cells(originalRow, "A")
    store.Range(store.Cells(found.Row, "A"), store.Cells(found.Row + rowCount - 1, "P")).Clear
    RefreshRecycleArea ws
    RestoreRecycleItem = True
    Exit Function
Failed:
    RestoreRecycleItem = False
End Function

Public Sub RefreshRecycleArea(ByVal ws As Worksheet)
    Dim store As Worksheet, rowIndex As Long, displayRow As Long, lastRow As Long
    Set store = GetRecycleStoreSheet()
    ws.Range("A" & RECYCLE_START_ROW + 1 & ":G" & RECYCLE_START_ROW + RECYCLE_MAX_ITEMS).Clear
    displayRow = RECYCLE_START_ROW + 1
    lastRow = store.Cells(store.Rows.Count, "A").End(xlUp).Row
    For rowIndex = 2 To lastRow
        If CStr(store.Cells(rowIndex, "B").Value) = ws.Name And Len(CStr(store.Cells(rowIndex, "A").Value)) > 0 Then
            ws.Cells(displayRow, "A").NumberFormat = "@"
            ws.Cells(displayRow, "A").Value = store.Cells(rowIndex, "A").Value
            ws.Cells(displayRow, "B").Value = store.Cells(rowIndex, "C").Value
            ws.Cells(displayRow, "C").Value = store.Cells(rowIndex, "F").Value
            ws.Cells(displayRow, "D").Value = store.Cells(rowIndex, "G").Value
            ws.Cells(displayRow, "E").Value = store.Cells(rowIndex, "H").Value
            ws.Cells(displayRow, "F").Value = "Run Restore Recycle Item"
            displayRow = displayRow + 1
            If displayRow > RECYCLE_START_ROW + RECYCLE_MAX_ITEMS Then Exit For
        End If
    Next rowIndex
End Sub

Public Function NewestRecycleItemId(ByVal ws As Worksheet) As String
    Dim store As Worksheet, rowIndex As Long, lastRow As Long, newestTime As Date
    Set store = GetRecycleStoreSheet()
    lastRow = store.Cells(store.Rows.Count, "A").End(xlUp).Row
    For rowIndex = 2 To lastRow
        If CStr(store.Cells(rowIndex, "B").Value) = ws.Name And Len(CStr(store.Cells(rowIndex, "A").Value)) > 0 Then
            If Len(NewestRecycleItemId) = 0 Or CDate(store.Cells(rowIndex, "H").Value) >= newestTime Then
                NewestRecycleItemId = CStr(store.Cells(rowIndex, "A").Value)
                newestTime = CDate(store.Cells(rowIndex, "H").Value)
            End If
        End If
    Next rowIndex
End Function

Private Sub StoreRecycleItem(ByVal ws As Worksheet, ByVal itemType As String, ByVal originalRow As Long, ByVal rowCount As Long)
    Dim store As Worksheet, storeRow As Long, chapterId As String, paragraphId As String
    Set store = GetRecycleStoreSheet()
    EnforceRecycleLimit ws, store
    storeRow = store.Cells(store.Rows.Count, "A").End(xlUp).Row + 2
    If storeRow < 2 Then storeRow = 2
    chapterId = NearestChapterId(ws, originalRow)
    paragraphId = CStr(ws.Cells(originalRow, "A").Value)
    store.Cells(storeRow, "A").NumberFormat = "@"
    store.Cells(storeRow, "A").Value = "RC-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(storeRow, "0000")
    store.Cells(storeRow, "B").Value = ws.Name
    store.Cells(storeRow, "C").Value = itemType
    store.Cells(storeRow, "D").Value = originalRow
    store.Cells(storeRow, "E").Value = rowCount
    store.Cells(storeRow, "F").Value = chapterId
    store.Cells(storeRow, "G").Value = paragraphId
    store.Cells(storeRow, "H").Value = Now
    ws.Range(ws.Cells(originalRow, "A"), ws.Cells(originalRow + rowCount - 1, "G")).Copy Destination:=store.Cells(storeRow, "J")
End Sub

Private Sub EnforceRecycleLimit(ByVal ws As Worksheet, ByVal store As Worksheet)
    Dim rowIndex As Long, count As Long, oldestRow As Long, oldestTime As Date, lastRow As Long, rowsToClear As Long
    lastRow = store.Cells(store.Rows.Count, "A").End(xlUp).Row
    For rowIndex = 2 To lastRow
        If CStr(store.Cells(rowIndex, "B").Value) = ws.Name And Len(CStr(store.Cells(rowIndex, "A").Value)) > 0 Then
            count = count + 1
            If oldestRow = 0 Or CDate(store.Cells(rowIndex, "H").Value) < oldestTime Then
                oldestRow = rowIndex
                oldestTime = CDate(store.Cells(rowIndex, "H").Value)
            End If
        End If
    Next rowIndex
    If count >= RECYCLE_MAX_ITEMS And oldestRow > 0 Then
        rowsToClear = CLng(store.Cells(oldestRow, "E").Value)
        store.Range(store.Cells(oldestRow, "A"), store.Cells(oldestRow + rowsToClear - 1, "P")).Clear
    End If
End Sub

Private Function GetRecycleStoreSheet() As Worksheet
    On Error Resume Next
    Set GetRecycleStoreSheet = ThisWorkbook.Worksheets(STORE_SHEET_NAME)
    On Error GoTo 0
    If GetRecycleStoreSheet Is Nothing Then
        Set GetRecycleStoreSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetRecycleStoreSheet.Name = STORE_SHEET_NAME
        GetRecycleStoreSheet.Range("A1").Value = "Item ID": GetRecycleStoreSheet.Range("B1").Value = "Script Sheet"
        GetRecycleStoreSheet.Range("C1").Value = "Item Type": GetRecycleStoreSheet.Range("D1").Value = "Original Row"
        GetRecycleStoreSheet.Range("E1").Value = "Row Count": GetRecycleStoreSheet.Range("F1").Value = "Chapter ID"
        GetRecycleStoreSheet.Range("G1").Value = "Paragraph ID": GetRecycleStoreSheet.Range("H1").Value = "Deleted At"
        GetRecycleStoreSheet.Visible = xlSheetVeryHidden
    End If
End Function

Private Function SelectedChapterRow(ByVal ws As Worksheet, ByVal selectedRow As Long) As Long
    Dim rowIndex As Long
    For rowIndex = selectedRow To 12 Step -1
        If Left$(CStr(ws.Cells(rowIndex, "A").Value), 3) = "CH-" Then
            SelectedChapterRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function NextChapterRow(ByVal ws As Worksheet, ByVal chapterRow As Long) As Long
    Dim rowIndex As Long, lastRow As Long
    lastRow = ws.Range("A12:A850").Find(What:="*", After:=ws.Range("A12"), LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    For rowIndex = chapterRow + 1 To lastRow
        If Left$(CStr(ws.Cells(rowIndex, "A").Value), 3) = "CH-" Then
            NextChapterRow = rowIndex
            Exit Function
        End If
    Next rowIndex
    NextChapterRow = lastRow + 1
End Function

Private Function NearestChapterId(ByVal ws As Worksheet, ByVal selectedRow As Long) As String
    Dim chapterRow As Long
    chapterRow = SelectedChapterRow(ws, selectedRow)
    If chapterRow > 0 Then NearestChapterId = CStr(ws.Cells(chapterRow, "A").Value)
End Function
