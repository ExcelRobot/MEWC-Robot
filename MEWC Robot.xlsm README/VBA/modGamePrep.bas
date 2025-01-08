Attribute VB_Name = "modGamePrep"
Option Explicit

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Backup Active Sheet
' Description:            Make a copy of the active sheet with " (Backup)" added to the sheet name.
' Macro Expression:       modGamePrep.BackupActiveSheet()
' Generated:              01/05/2025 07:14 PM
'----------------------------------------------------------------------------------------------------
Sub BackupActiveSheet()
    Dim wsActive As Worksheet
    Dim wsBackup As Worksheet
    
    Set wsActive = ActiveSheet
    wsActive.Copy After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
    Set wsBackup = ActiveSheet
    On Error Resume Next
    wsBackup.Name = wsActive.Name & " (Backup)"
    On Error GoTo 0
    wsActive.Activate
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Backup All Sheets
' Description:            Make a copy of all sheets with " (Backup)" added to each sheet name.
' Macro Expression:       modGamePrep.BackupAllSheets()
' Generated:              01/05/2025 07:18 PM
'----------------------------------------------------------------------------------------------------
Sub BackupAllSheets()
    Dim wsActive As Worksheet
    Dim wsTarget As Worksheet
    Dim wsBackup As Worksheet
    Dim nSheetCount As Integer
    Dim nCtr As Integer
    
    nSheetCount = ActiveWorkbook.Worksheets.Count
    
    Set wsActive = ActiveSheet
    
    Application.ScreenUpdating = False
    
    For nCtr = 1 To nSheetCount
        Set wsTarget = ActiveWorkbook.Worksheets(nCtr)
        wsTarget.Copy After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        Set wsBackup = ActiveSheet
        On Error Resume Next
        wsBackup.Name = wsTarget.Name & " (Backup)"
        On Error GoTo 0
    Next
    
    wsActive.Activate
    
    Application.ScreenUpdating = True
    
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Name Used Ranges On All Sheets
' Description:            Names the used range on each sheet in the workbook using a sanitized version of the sheet name.
' Macro Expression:       modGamePrep.NameUsedRangesOnAllSheets()
' Generated:              01/05/2025 07:20 PM
'----------------------------------------------------------------------------------------------------
Sub NameUsedRangesOnAllSheets()
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        If Right(ws.Name, Len(" (Backup)")) <> " (Backup)" Then
            ws.UsedRange.Name = SanitizeRangeName(ws.Name)
        End If
    Next
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Save Game Answer To Left
' Description:            Saves references to the selected cells in the green answer cells to the left on the same row.
' Macro Expression:       modGamePrep.SaveAnswersToLeft()
' Generated:              01/05/2025 07:23 PM
'----------------------------------------------------------------------------------------------------
Sub SaveAnswersToLeft()
    Dim cell As Range
    Dim greenCol As Integer
    Dim dest As Range
    
    ' MEWC Answer Cell Green
    Const MEWC_GREEN As Long = 3631104
    
    ' Find the green cell
    For Each cell In Intersect(ActiveCell.EntireRow, ActiveSheet.UsedRange)
        If cell.Interior.Color = MEWC_GREEN Then ' mewc answer cell green
            greenCol = cell.Column
            Exit For
        End If
    Next
    
    If greenCol <> 0 Then
        Dim calcMode As Integer
        On Error Resume Next
        calcMode = Application.Calculation
        Application.Calculation = xlCalculationManual
        For Each cell In Selection
            If Cells(cell.row, greenCol).Interior.Color = MEWC_GREEN Then
                Cells(cell.row, greenCol).Formula = "=" & cell.Address(False, False)
                If dest Is Nothing Then
                    Set dest = Cells(cell.row, greenCol)
                Else
                    Set dest = Union(dest, Cells(cell.row, greenCol))
                End If
            End If
        Next
        
        Application.Calculation = calcMode
        
        ' if some green cells were saved to, select them and copy either those answers or the formula below.
        If Not dest Is Nothing Then
            dest.Select
            If Left(dest(1).Offset(dest.Rows.Count + 1).Formula, 1) = "=" Then
                dest(1).Offset(dest.Rows.Count + 1).Copy
            Else
                dest.Copy
            End If
        End If
    End If
End Sub
