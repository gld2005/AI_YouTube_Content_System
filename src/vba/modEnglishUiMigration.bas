Attribute VB_Name = "modEnglishUiMigration"
Option Explicit

Public Sub ApplyEnglishUiMigration()
    Application.ScreenUpdating = False
    RenameSheetAtIndex 1, "Instructions"
    RenameSheetAtIndex 2, "Dashboard"
    RenameSheetAtIndex 3, "Content Workspace"
    RenameSheetAtIndex 4, "Topic Pool"
    RenameSheetAtIndex 5, "Video Projects"
    RenameSheetAtIndex 6, "Series Settings"
    RenameSheetAtIndex 7, "Structure Templates"
    RenameSheetAtIndex 8, "Script Library"
    RenameSheetAtIndex 9, "Script Template"
    RenameSheetAtIndex 10, "Source Master"
    RenameSheetAtIndex 11, "Evidence Cards"
    RenameSheetAtIndex 12, "Research Links"
    RenameSheetAtIndex 13, "Packaging Lab"
    RenameSheetAtIndex 14, "Task Plan"
    RenameSheetAtIndex 15, "Gantt"
    RenameSheetAtIndex 16, "Project Assets"
    RenameSheetAtIndex 17, "Shared Assets"
    RenameSheetAtIndex 18, "Tags"
    RenameSheetAtIndex 19, "AI Approval Center"
    RenameSheetAtIndex 20, "Publish Review"
    RenameSheetAtIndex 21, "Settings"
    RenameSheetAtIndex 22, "Migration Audit"
    Application.ScreenUpdating = True
End Sub

Private Sub RenameSheetAtIndex(ByVal index As Long, ByVal newName As String)
    If ThisWorkbook.Worksheets.Count >= index Then ThisWorkbook.Worksheets(index).Name = newName
End Sub
