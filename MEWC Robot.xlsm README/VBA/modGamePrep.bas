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
        ws.UsedRange.Name = SanitizeRangeName(ws.Name)
    Next
End Sub

Private Function SanitizeRangeName(proposedName As String) As String
    Dim result As String
    Dim i As Long
    Dim char As String
    Dim validFirstChars As String
    Dim validChars As String
    
    ' If empty string, return default name
    If Len(proposedName) = 0 Then
        SanitizeRangeName = "Range1"
        Exit Function
    End If
    
    ' Initialize working string
    result = proposedName
    
    ' Replace spaces with underscores
    result = Replace(result, " ", "_")
    
    ' Define valid characters
    validFirstChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ_\"
    validChars = validFirstChars & "0123456789."
    
    ' Check and fix first character
    char = UCase(Left(result, 1))
    If InStr(validFirstChars, char) = 0 Then
        ' If first char is a number, prepend "N_"
        If IsNumeric(char) Then
            result = "N_" & result
        Else
            ' Replace invalid first char with underscore
            result = "_" & Mid(result, 2)
        End If
    End If
    
    ' Clean remaining characters
    Dim sanitized As String
    sanitized = Left(result, 1)
    For i = 2 To Len(result)
        char = Mid(result, i, 1)
        ' Keep only valid characters
        If InStr(validChars, UCase(char)) > 0 Then
            sanitized = sanitized & char
        Else
            ' Replace invalid chars with underscore
            sanitized = sanitized & "_"
        End If
    Next i
    
    ' Trim to 255 characters if needed
    If Len(sanitized) > 255 Then
        sanitized = Left(sanitized, 255)
    End If
    
    ' Handle reserved words and cell references
    result = sanitized
    If Not IsValidRangeName(result) Then
        ' Add prefix for reserved words or cell references
        result = "RNG_" & result
    End If
    
    ' Verify final result
    If Not IsValidRangeName(result) Then
        ' If still invalid, use a safe default
        result = "Range_" & Format(Now, "yyyymmddhhnnss")
    End If
    
    SanitizeRangeName = result
End Function

Private Function IsValidRangeName(proposedName As String) As Boolean
    Dim i As Long
    Dim char As String
    Dim validFirstChars As String
    Dim validChars As String
    
    ' Initialize result
    IsValidRangeName = True
    
    ' Handle empty string
    If Len(proposedName) = 0 Then
        IsValidRangeName = False
        Exit Function
    End If
    
    ' Check length constraint
    If Len(proposedName) > 255 Then
        IsValidRangeName = False
        Exit Function
    End If
    
    ' Define valid characters
    validFirstChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ_\"
    validChars = validFirstChars & "0123456789."
    
    ' Check first character
    char = UCase(Left(proposedName, 1))
    If InStr(validFirstChars, char) = 0 Then
        IsValidRangeName = False
        Exit Function
    End If
    
    ' Check remaining characters
    For i = 2 To Len(proposedName)
        char = UCase(Mid(proposedName, i, 1))
        If InStr(validChars, char) = 0 Then
            IsValidRangeName = False
            Exit Function
        End If
    Next i
    
    ' Check for spaces
    If InStr(proposedName, " ") > 0 Then
        IsValidRangeName = False
        Exit Function
    End If
    
    ' Check if it's a cell reference
    If IsCellReference(proposedName) Then
        IsValidRangeName = False
        Exit Function
    End If
    
    ' Check reserved words
    Select Case UCase(proposedName)
        Case "C", "R", "PRINT_AREA", "PRINT_TITLES", "CONSOLIDATE_AREA", _
             "DATABASE", "CRITERIA", "TRUE", "FALSE", "ERROR"
            IsValidRangeName = False
            Exit Function
    End Select
    
    ' Check if it's only numbers
    If IsNumeric(proposedName) Then
        IsValidRangeName = False
        Exit Function
    End If
End Function

' Helper function to check if string matches cell reference pattern
Private Function IsCellReference(str As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Pattern matches common Excel cell references like A1, AA123, etc.
    regEx.Pattern = "^[A-Za-z]{1,3}[0-9]+$"
    IsCellReference = regEx.Test(str)
End Function

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
        
        ' if some green cells were saved to, select them and copy either those answers or the formula below.
        If Not dest Is Nothing Then
            dest.Select
            If Left(dest(1).Offset(dest.Rows.Count + 1).Formula, 1) = "=" Then
                dest(1).Offset(dest.Rows.Count + 1).Copy
            Else
                dest.Copy
            End If
        End If
        Application.Calculation = calcMode
    End If
End Sub
