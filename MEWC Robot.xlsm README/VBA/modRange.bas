Attribute VB_Name = "modRange"
Option Explicit

Function ActiveColumnIndexInSpillingRange(targetCell As Range) As String
    Dim spillRange As Range
    
    If targetCell.Cells(1).SpillParent Is Nothing Then Exit Function
    
    Set spillRange = targetCell.Cells(1).SpillParent.SpillingToRange
    
    ActiveColumnIndexInSpillingRange = targetCell.Column - spillRange.Column + 1

End Function


Function IsValidRangeName(proposedName As String) As Boolean
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

Function SanitizeRangeName(proposedName As String) As String
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

