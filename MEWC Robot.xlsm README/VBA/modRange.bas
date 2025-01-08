Attribute VB_Name = "modRange"
Option Explicit

Function ActiveColumnIndexInSpillingRange(targetCell As Range) As String
    Dim spillRange As Range
    
    If targetCell.Cells(1).SpillParent Is Nothing Then Exit Function
    
    Set spillRange = targetCell.Cells(1).SpillParent.SpillingToRange
    
    ActiveColumnIndexInSpillingRange = targetCell.Column - spillRange.Column + 1

End Function
