# <h1 id="oa-robot-definitions">OA Robot Definitions</h1>

\*\*MEWC Robot.xlsm\*\* contains definitions for:

[11 Robot Commands](#command-definitions)<BR>

<BR>

## Available Robot Commands

[Goto](#goto) | [Name](#name) | [Paste](#paste) | [Prep](#prep) | [Select](#select) | [Other](#other)

### Goto

| Name | Description |
| --- | --- |
| [Goto Similar Background Color](#goto-similar-background-color) | Select cells in selection with same background color as active cell. |
| [Goto Similar Constant Values](#goto-similar-constant-values) | Select constant cells in selection with similar value as active cell. |
| [Goto Similar Formulas](#goto-similar-formulas) | Select formula cells in selection with similar formula as active cell. |

### Name

| Name | Description |
| --- | --- |
| [Name Used Ranges On All Sheets](#name-used-ranges-on-all-sheets) | Names the used range on each sheet in the workbook using a sanitized version of the sheet name. |

### Paste

| Name | Description |
| --- | --- |
| [Paste Count By Background Color](#paste-count-by-background-color) | Pastes the count of cells in copied range by background color. |
| [Paste Formulas Over Similar Background Colors](#paste-formulas-over-similar-background-colors) | Paste the formulas in the copied cells over all similar background colors on the sheet as the selected cells. |
| [Paste Sum By Background Color](#paste-sum-by-background-color) | Pastes the sum of cells in copied range by background color. |
| [Paste Values Over Similar Background Colors](#paste-values-over-similar-background-colors) | Paste the values in the copied cells over all similar background colors on the sheet as the selected cells. |

### Prep

| Name | Description |
| --- | --- |
| [Backup Active Sheet](#backup-active-sheet) | Make a copy of the active sheet with " (Backup)" added to the sheet name. |
| [Backup All Sheets](#backup-all-sheets) | Make a copy of all sheets with " (Backup)" added to each sheet name. |
| [Name Used Ranges On All Sheets](#name-used-ranges-on-all-sheets) | Names the used range on each sheet in the workbook using a sanitized version of the sheet name. |

### Select

| Name | Description |
| --- | --- |
| [Goto Similar Background Color](#goto-similar-background-color) | Select cells in selection with same background color as active cell. |
| [Goto Similar Constant Values](#goto-similar-constant-values) | Select constant cells in selection with similar value as active cell. |
| [Goto Similar Formulas](#goto-similar-formulas) | Select formula cells in selection with similar formula as active cell. |

### Other

| Name | Description |
| --- | --- |
| [Save Game Answer To Left](#save-game-answer-to-left) | Saves references to the selected cells in the green answer cells to the left on the same row. |

<BR>

## Command Definitions

<BR>

### Backup Active Sheet

*Make a copy of the active sheet with " (Backup)" added to the sheet name.*

<sup>`@MEWC Robot.xlsm` `!VBA Macro Command` `#Prep`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modGamePrep.BackupActiveSheet](./VBA/modGamePrep.bas#L10)()</code> |
| Launch Codes | <code>bs</code> |

[^Top](#oa-robot-definitions)

<BR>

### Backup All Sheets

*Make a copy of all sheets with " (Backup)" added to each sheet name.*

<sup>`@MEWC Robot.xlsm` `!VBA Macro Command` `#Prep`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modGamePrep.BackupAllSheets](./VBA/modGamePrep.bas#L29)()</code> |
| Launch Codes | <code>bas</code> |

[^Top](#oa-robot-definitions)

<BR>

### Goto Similar Background Color

*Select cells in selection with same background color as active cell.*

<sup>`@MEWC Robot.xlsm` `!VBA Macro Command` `#Select` `#Goto`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modGotoSpecial.GotoSimilarBackgroundColor](./VBA/modGotoSpecial.bas#L10)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### Goto Similar Constant Values

*Select constant cells in selection with similar value as active cell.*

<sup>`@MEWC Robot.xlsm` `!VBA Macro Command` `#Select` `#Goto`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modGotoSpecial.GotoSimilarValue](./VBA/modGotoSpecial.bas#L64)()</code> |
| User Context Filter | ExcelActiveCellIsNotEmpty |

[^Top](#oa-robot-definitions)

<BR>

### Goto Similar Formulas

*Select formula cells in selection with similar formula as active cell.*

<sup>`@MEWC Robot.xlsm` `!VBA Macro Command` `#Select` `#Goto`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modGotoSpecial.GotoSimilarFormulas](./VBA/modGotoSpecial.bas#L129)()</code> |
| User Context Filter | ExcelActiveCellContainsFormula |

[^Top](#oa-robot-definitions)

<BR>

### Name Used Ranges On All Sheets

*Names the used range on each sheet in the workbook using a sanitized version of the sheet name.*

<sup>`@MEWC Robot.xlsm` `!VBA Macro Command` `#Name` `#Prep`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modGamePrep.NameUsedRangesOnAllSheets](./VBA/modGamePrep.bas#L63)()</code> |
| Launch Codes | <code>nur</code> |

[^Top](#oa-robot-definitions)

<BR>

### Paste Count By Background Color

*Pastes the count of cells in copied range by background color.*

<sup>`@MEWC Robot.xlsm` `!VBA Macro Command` `#Paste`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modPasteSpecial.CountByBackgroundColor](./VBA/modPasteSpecial.bas#L159)([[Clipboard]],[[ActiveCell]])</code> |
| User Context Filter | ClipboardHasExcelData AND ExcelActiveCellIsEmpty |

[^Top](#oa-robot-definitions)

<BR>

### Paste Formulas Over Similar Background Colors

*Paste the formulas in the copied cells over all similar background colors on the sheet as the selected cells.*

<sup>`@MEWC Robot.xlsm` `!VBA Macro Command` `#Paste`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modPasteSpecial.PasteOverSimilarBackgroundColors](./VBA/modPasteSpecial.bas#L16)([[Clipboard]],[[Selection]],True)</code> |
| User Context Filter | ClipboardHasExcelData AND ExcelSelectionIsSingleArea |

[^Top](#oa-robot-definitions)

<BR>

### Paste Sum By Background Color

*Pastes the sum of cells in copied range by background color.*

<sup>`@MEWC Robot.xlsm` `!VBA Macro Command` `#Paste`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modPasteSpecial.SumByBackgroundColor](./VBA/modPasteSpecial.bas#L84)([[Clipboard]],[[ActiveCell]])</code> |
| User Context Filter | ClipboardHasExcelData AND ExcelActiveCellIsEmpty |

[^Top](#oa-robot-definitions)

<BR>

### Paste Values Over Similar Background Colors

*Paste the values in the copied cells over all similar background colors on the sheet as the selected cells.*

<sup>`@MEWC Robot.xlsm` `!VBA Macro Command` `#Paste`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modPasteSpecial.PasteOverSimilarBackgroundColors](./VBA/modPasteSpecial.bas#L16)([[Clipboard]],[[Selection]])</code> |
| User Context Filter | ClipboardHasExcelData AND ExcelSelectionIsSingleArea |

[^Top](#oa-robot-definitions)

<BR>

### Save Game Answer To Left

*Saves references to the selected cells in the green answer cells to the left on the same row.*

<sup>`@MEWC Robot.xlsm` `!VBA Macro Command` </sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modGamePrep.SaveAnswersToLeft](./VBA/modGamePrep.bas#L225)()</code> |
| Keyboard Shortcut | <code>^+s</code> |
| User Context Filter | ExcelSelectionIsMultipleRows AND ExcelSelectionIsSingleColumn |
| Launch Codes | <code>sa</code> |

[^Top](#oa-robot-definitions)
