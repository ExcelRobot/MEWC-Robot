Attribute VB_Name = "modGotoSpecial"
Option Explicit

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Goto Similar Background Color
' Description:            Select cells in selection with same background color as active cell.
' Macro Expression:       modGotoSpecial.GotoSimilarBackgroundColor()
' Generated:              01/05/2025 07:27 PM
'----------------------------------------------------------------------------------------------------
Sub GotoSimilarBackgroundColor()

    Dim nColorIndex As Long
    Dim dTintShade As Double
    Dim nCtr As Long
    Dim rngOldActive As Range
    Dim rngOldSelection As Range
    Dim rngNewSelection As Range
    
    Set rngOldActive = ActiveCell
    Set rngOldSelection = SelectionOrUsedRange(Selection)
    nColorIndex = ActiveCell.Interior.ColorIndex
    dTintShade = Round(ActiveCell.Interior.TintAndShade, 3)
    
    Dim rngArea As Range
    For Each rngArea In rngOldSelection.Areas
        For nCtr = 1 To rngArea.Cells.Count
            If rngArea.Cells(nCtr).Interior.ColorIndex = nColorIndex And Round(rngArea.Cells(nCtr).Interior.TintAndShade, 3) = dTintShade Then
                If rngNewSelection Is Nothing Then
                    Set rngNewSelection = rngArea.Cells(nCtr)
                Else
                    Set rngNewSelection = Union(rngNewSelection, rngArea.Cells(nCtr))
                End If
            End If
        Next nCtr
    Next rngArea
    
    If Not rngNewSelection Is Nothing Then
        rngNewSelection.Select
        If Not Intersect(rngNewSelection, rngOldActive) Is Nothing Then
            rngOldActive.Activate
        End If
    End If

End Sub

Private Function SelectionOrUsedRange(vSelection As Variant) As Range
    If TypeName(vSelection) <> "Range" Then
        Set SelectionOrUsedRange = ActiveSheet.UsedRange
    ElseIf vSelection.Cells.Count = 1 Then
        Set SelectionOrUsedRange = vSelection.Parent.UsedRange
    ElseIf Not Intersect(vSelection, vSelection.Parent.UsedRange) Is Nothing Then
        Set SelectionOrUsedRange = Intersect(vSelection, vSelection.Parent.UsedRange)
    Else
        Set SelectionOrUsedRange = vSelection.Parent.UsedRange
    End If
End Function

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Goto Similar Constant Value
' Description:            Select constant cells in selection with similar value as active cell.
' Macro Expression:       modGotoSpecial.GotoSimilarValue()
' Generated:              01/05/2025 07:37 PM
'----------------------------------------------------------------------------------------------------
Sub GotoSimilarValue(Optional CompareType As Integer = 0)
    Const CompareType_CaseInsensitiveMatch As Integer = 0
    Const CompareType_ExactMatch As Integer = 1
    Const CompareType_Contains As Integer = 2
    Const CompareType_StartsWith As Integer = 3
    Const CompareType_EndsWith As Integer = 4
    
    Dim vValue As Variant
    Dim nCtr As Long
    Dim rngOldActive As Range
    Dim rngOldSelection As Range
    Dim rngNewSelection As Range
    Dim bMatch As Boolean
    
    Set rngOldActive = ActiveCell
    Set rngOldSelection = SelectionOrUsedRange(Selection)
    vValue = ActiveCell.Value
    
    Dim rngArea As Range
    For Each rngArea In rngOldSelection.Areas
        For nCtr = 1 To rngArea.Cells.Count
            bMatch = False
            On Error Resume Next
            If rngArea.Cells(nCtr).Value <> "" Then
            Select Case CompareType
                Case CompareType_CaseInsensitiveMatch
                    bMatch = (UCase(rngArea.Cells(nCtr).Value) = UCase(vValue))
                Case CompareType_ExactMatch
                    bMatch = (rngArea.Cells(nCtr).Value = vValue)
                Case CompareType_Contains
                    bMatch = InStr(1, rngArea.Cells(nCtr).Value, vValue) > 0
                Case CompareType_StartsWith
                    bMatch = (Left(rngArea.Cells(nCtr).Value, Len(vValue)) = vValue)
                Case CompareType_EndsWith
                    bMatch = (Right(rngArea.Cells(nCtr).Value, Len(vValue)) = vValue)
            End Select
            End If
            On Error GoTo 0
            If bMatch Then
                If Not Application.IsFormula(rngArea.Cells(nCtr)) _
                    And Not rngArea.Cells(nCtr).HasSpill Then
                    If rngNewSelection Is Nothing Then
                        Set rngNewSelection = rngArea.Cells(nCtr)
                    Else
                        Set rngNewSelection = Union(rngNewSelection, rngArea.Cells(nCtr))
                    End If
                End If
            End If
        Next nCtr
    Next rngArea
    
    If Not rngNewSelection Is Nothing Then
        rngNewSelection.Select
        If Not Intersect(rngNewSelection, rngOldActive) Is Nothing Then
            rngOldActive.Activate
        End If
    End If
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Goto Similar Formulas
' Description:            Select formula cells in selection with similar formula as active cell.
' Macro Expression:       modGotoSpecial.GotoSimilarFormulas()
' Generated:              01/05/2025 07:41 PM
'----------------------------------------------------------------------------------------------------
Sub GotoSimilarFormulas()
    Dim sFormula As String
    Dim nCtr As Long
    Dim rngOldActive As Range
    Dim rngOldSelection As Range
    Dim rngNewSelection As Range
    
    Set rngOldActive = ActiveCell
    Set rngOldSelection = SelectionOrUsedRange(Selection)
    sFormula = ActiveCell.FormulaR1C1
    
    Dim rngArea As Range
    For Each rngArea In rngOldSelection.Areas
        For nCtr = 1 To rngArea.Cells.Count
            If rngArea.Cells(nCtr).FormulaR1C1 = sFormula Then
                If Application.IsFormula(rngArea.Cells(nCtr)) Then
                    If rngNewSelection Is Nothing Then
                        Set rngNewSelection = rngArea.Cells(nCtr)
                    Else
                        Set rngNewSelection = Union(rngNewSelection, rngArea.Cells(nCtr))
                    End If
                End If
            End If
        Next nCtr
    Next rngArea
    
    If Not rngNewSelection Is Nothing Then
        rngNewSelection.Select
        If Not Intersect(rngNewSelection, rngOldActive) Is Nothing Then
            rngOldActive.Activate
        End If
    End If
End Sub
