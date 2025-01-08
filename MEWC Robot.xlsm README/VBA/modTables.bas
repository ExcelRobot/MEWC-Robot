Attribute VB_Name = "modTables"
Option Explicit

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Create Table With Headers
' Description:            Create table from current region with headers.
' Macro Expression:       modTables.CreateTableWithHeaders(True)
' Generated:              12/29/2024 02:50 PM
'----------------------------------------------------------------------------------------------------
Sub CreateTableWithHeaders(Optional includeFilters As Boolean = False)
    Dim Table As ListObject
        
    If Not ActiveCell.SpillParent Is Nothing Then
        ActiveCell.SpillParent.SpillingToRange.Select
        ActiveCell.SpillParent.Activate
        Selection.Value = Selection.Value
    End If
    
    If Selection.Cells.Count = 1 Then
        Set Table = ActiveSheet.ListObjects.Add(xlSrcRange, Selection.CurrentRegion, , xlYes)
    Else
        Set Table = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    End If
    
    ' Only include filters if requested
    Table.ShowAutoFilter = includeFilters
    
    Dim tableName As String
    tableName = InputBox("Enter table name:", "Create Table With Headers", SanitizeRangeName(Table.Range.Worksheet.Name))
    If tableName <> "" Then
        On Error Resume Next
        Table.Name = tableName
        On Error GoTo 0
    End If
End Sub
