Attribute VB_Name = "modPasteSpecial"
Option Explicit

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Paste Values Over Similar Background Colors
' Description:            Paste the values in the copied cells over all similar background colors on the sheet as the selected cells.
' Macro Expression:       modPasteSpecial.PasteOverSimilarBackgroundColors([[Clipboard]],[[Selection]])
' Generated:              01/05/2025 08:03 PM
'----------------------------------------------------------------------------------------------------
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Paste Formulas Over Similar Background Colors
' Description:            Paste the formulas in the copied cells over all similar background colors on the sheet as the selected cells.
' Macro Expression:       modPasteSpecial.PasteOverSimilarBackgroundColors([[Clipboard]],[[Selection]],True)
' Generated:              01/05/2025 08:07 PM
'----------------------------------------------------------------------------------------------------
Sub PasteOverSimilarBackgroundColors(rngSource As Range, rngTarget As Range, Optional bPasteFormula As Boolean = False)
    Dim nCtr As Long
    Dim nAreaCtr As Long
    Dim nIndex As Long
    Dim rngSearchArea As Range
    Dim rngSimilar As Range
    Dim nCalc As Long
    
    If rngSource.Areas.Count > 1 Or rngTarget.Areas.Count > 1 Then Exit Sub
    
    If rngSource.Cells.Count > 1 And rngTarget.Cells.Count = 1 Then
        Set rngTarget = rngTarget.Resize(rngSource.Rows.Count, rngSource.Columns.Count)
    End If
    
    ' Either rngSource needs to be one cell or it has to be same dimensions as selection
    If rngSource.Cells.Count <> 1 And (rngTarget.Cells.Count = 1 Or rngSource.Rows.Count <> rngTarget.Rows.Count Or rngSource.Columns.Count <> rngTarget.Columns.Count) Then
        Exit Sub
    End If
    
    nCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    If rngTarget.Cells.Count > 1 Then
        For nCtr = 1 To rngTarget.Cells.Count
            If rngSource.Cells.Count = 1 Then
                Call PasteOverSimilarBackgroundColors(rngSource, rngTarget.Cells(nCtr), bPasteFormula)
            Else
                Call PasteOverSimilarBackgroundColors(rngSource.Cells(nCtr), rngTarget.Cells(nCtr), bPasteFormula)
            End If
        Next
    Else
        Dim nColorIndex As Long
        Dim dTintShade As Double
        nColorIndex = rngTarget.Interior.ColorIndex
        dTintShade = Round(rngTarget.Interior.TintAndShade, 3)
        Set rngSearchArea = rngTarget.Parent.UsedRange
        For nAreaCtr = 1 To rngSearchArea.Areas.Count
            For nCtr = 1 To rngSearchArea.Areas(nAreaCtr).Cells.Count
                If rngSearchArea.Areas(nAreaCtr).Cells(nCtr).Interior.ColorIndex = nColorIndex And Round(rngSearchArea.Areas(nAreaCtr).Cells(nCtr).Interior.TintAndShade, 3) = dTintShade Then
                    If rngSimilar Is Nothing Then
                        Set rngSimilar = rngSearchArea.Areas(nAreaCtr).Cells(nCtr)
                    Else
                        Set rngSimilar = Union(rngSimilar, rngSearchArea.Areas(nAreaCtr).Cells(nCtr))
                    End If
                End If
            Next
        Next
        If Not rngSimilar Is Nothing Then
            If bPasteFormula Then
                rngSimilar.Formula2 = rngSource.Formula2
            Else
                rngSimilar.Value = rngSource.Value
            End If
        End If
    End If
    
    If Application.Calculation <> nCalc Then
        Application.Calculation = nCalc
    End If

End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Paste Sum By Background Color
' Description:            Pastes the sum of cells in copied range by background color.
' Macro Expression:       modPasteSpecial.SumByBackgroundColor([[Clipboard]],[[ActiveCell]])
' Generated:              01/05/2025 07:56 PM
'----------------------------------------------------------------------------------------------------
Public Sub SumByBackgroundColor(Source As Range, Destination As Range)
    Dim cell As Range
    Dim ColorDictionary As Object
    Dim Key As Variant
    Dim DestRow As Long
    
    On Error GoTo ErrHandler
    
    Dim nCalc As Integer
    nCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Set ColorDictionary = CreateObject("Scripting.Dictionary")
    
    DestRow = 1
    
    ' Loop through each cell in the source range
    For Each cell In Source
        
        If cell.Text <> "" And TypeName(cell.Value) <> "String" And TypeName(cell.Value) <> "Date" And TypeName(cell.Value) <> "Error" Then
        
            ' Use the RGB color as the key
            Key = cell.Interior.Color
            ' If the color is not yet in the dictionary, add it with the cell's value
            If Not ColorDictionary.Exists(Key) Then
                ColorDictionary.Add Key, cell.Value
            Else
                ' If the color is already in the dictionary, add the current cell's value to the sum
                ColorDictionary(Key) = Application.Sum(ColorDictionary(Key), cell.Value)
            End If
            
        End If
        
    Next cell
    
    On Error Resume Next
    Destination.Worksheet.Names.Add "RGBtoColorName", ThisWorkbook.Names("RGBtoColorName").RefersTo
    On Error GoTo 0
    
    ' Output the results to the destination range
    For Each Key In ColorDictionary.keys
        ' Split the color into its RGB components
        Dim R As Integer, G As Integer, B As Integer
        R = Key Mod 256
        G = (Key \ 256) Mod 256
        B = (Key \ 65536) Mod 256
        ' Write the RGB values and the sum to the destination
        Destination.Cells(DestRow, 1).Value = Evaluate("RGBtoColorName(" & R & ", " & G & ", " & B & ")")
        Destination.Cells(DestRow, 1).Interior.Color = Key
        Destination.Cells(DestRow, 2).Value = ColorDictionary(Key)
        DestRow = DestRow + 1
    Next Key
    
    Call MakeDarkCellsWhiteFont(Destination.Resize(DestRow - 1, 1))
    
    Destination.Resize(DestRow - 1, 2).Sort Destination.Cells(1, 2), xlDescending, Header:=xlNo
    Destination.Resize(DestRow - 1, 2).Select
    
    Set ColorDictionary = Nothing
    
    Application.Calculation = nCalc
    Exit Sub
ErrHandler:
    Application.Calculation = nCalc
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Paste Count By Background Color
' Description:            Pastes the count of cells in copied range by background color.
' Macro Expression:       modPasteSpecial.CountByBackgroundColor([[Clipboard]],[[ActiveCell]])
' Generated:              01/05/2025 07:54 PM
'----------------------------------------------------------------------------------------------------
Public Sub CountByBackgroundColor(Source As Range, Destination As Range)
    Dim cell As Range
    Dim ColorDictionary As Object
    Dim Key As Variant
    Dim DestRow As Long
    
    On Error GoTo ErrHandler
    
    Dim nCalc As Integer
    nCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Set ColorDictionary = CreateObject("Scripting.Dictionary")
    
    DestRow = 1
    
    ' Loop through each cell in the source range
    For Each cell In Source
        
        ' Use the RGB color as the key
        Key = cell.Interior.Color
        ' If the color is not yet in the dictionary, add it with the cell's value
        If Not ColorDictionary.Exists(Key) Then
            ColorDictionary.Add Key, 1
        Else
            ' If the color is already in the dictionary, add the current cell's value to the sum
            ColorDictionary(Key) = ColorDictionary(Key) + 1
        End If
            
    Next cell
    
    On Error Resume Next
    Destination.Worksheet.Names.Add "RGBtoColorName", ThisWorkbook.Names("RGBtoColorName").RefersTo
    On Error GoTo 0
    
    ' Output the results to the destination range
    For Each Key In ColorDictionary.keys
        ' Split the color into its RGB components
        Dim R As Integer, G As Integer, B As Integer
        R = Key Mod 256
        G = (Key \ 256) Mod 256
        B = (Key \ 65536) Mod 256
        ' Write the RGB values and the sum to the destination
        Destination.Cells(DestRow, 1).Value = Evaluate("RGBtoColorName(" & R & ", " & G & ", " & B & ")")
        Destination.Cells(DestRow, 1).Interior.Color = Key
        Destination.Cells(DestRow, 2).Value = ColorDictionary(Key)
        DestRow = DestRow + 1
    Next Key
    
    Call MakeDarkCellsWhiteFont(Destination.Resize(DestRow - 1, 1))
    
    Destination.Resize(DestRow - 1, 2).Sort Destination.Cells(1, 2), xlDescending, Header:=xlNo
    Destination.Resize(DestRow - 1, 2).Select
    
    Set ColorDictionary = Nothing
    
    Application.Calculation = nCalc
    Exit Sub
ErrHandler:
    Application.Calculation = nCalc
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub MakeDarkCellsWhiteFont(Target As Range)
    Dim cell As Range
    Dim RGBValue As Long
    Dim R As Integer, G As Integer, B As Integer
    Dim Luminance As Double
    
    For Each cell In Target
        RGBValue = cell.Interior.Color
        
        ' Calculate R/G/B
        R = RGBValue Mod 256
        G = (RGBValue \ 256) Mod 256
        B = (RGBValue \ 65536) Mod 256
        
        ' Calculate the luminance
        Luminance = 0.299 * R + 0.587 * G + 0.114 * B
        
        ' If cell is considered dark, make font color White
        If Luminance < 128 Then
            cell.Font.Color = RGB(255, 255, 255)
        End If
    Next
End Sub
