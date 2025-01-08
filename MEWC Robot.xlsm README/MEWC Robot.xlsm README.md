# <h1 id="oa-robot-definitions">OA Robot Definitions</h1>

\*\*MEWC Robot.xlsm\*\* contains definitions for:

[20 Robot Commands](#command-definitions)<BR>[1 Robot Parameter](#parameter-definitions)<BR>[8 Robot Texts](#text-definitions)<BR>

<BR>

## Available Robot Commands

[Analysis](#analysis) | [Filter](#filter) | [Goto](#goto) | [Group](#group) | [Name](#name) | [Paste](#paste) | [Prep](#prep) | [Save](#save) | [Select](#select)

### Analysis

| Name | Description |
| --- | --- |
| [Least Frequent Value Of Array](#least-frequent-value-of-array) | Wrap with LeastFrequentValue Lambda function. |
| [Most Frequent Value Of Array](#most-frequent-value-of-array) | Wrap with MostFrequentValue Lambda function. |

### Filter

| Name | Description |
| --- | --- |
| [Filter Array By Selected Values](#filter-array-by-selected-values) | Wraps array formula with function returning active array filtered by selected values in a column of active array. |
| [Paste Filter Array By Copied Cell](#paste-filter-array-by-copied-cell) | Wraps array formula with function returning active array filtered by copied cell in a column of active array. |

### Goto

| Name | Description |
| --- | --- |
| [Goto Similar Background Color](#goto-similar-background-color) | Select cells in selection with same background color as active cell. |
| [Goto Similar Constant Values](#goto-similar-constant-values) | Select constant cells in selection with similar value as active cell. |
| [Goto Similar Formulas](#goto-similar-formulas) | Select formula cells in selection with similar formula as active cell. |

### Group

| Name | Description |
| --- | --- |
| [Group By Average](#group-by-average) | Groups by first N column, aggregating remaining columns with AVERAGE function. |
| [Group By Concat With Delimiter](#group-by-concat-with-delimiter) | Groups by first N column, concatenating remaining columns with specified delimiter. |
| [Group By Function\/Lambda](#group-by-functionlambda) | Groups by first N columns, aggregating remaining columns with specified function or lambda. |
| [Group By Sum](#group-by-sum) | Groups by first N column, aggregating remaining columns with SUM function. |
| [Ungroup Column By Delimiter](#ungroup-column-by-delimiter) | Splits the values in the selected column and expands the other columns. |

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

### Save

| Name | Description |
| --- | --- |
| [Save Game Answer To Left](#save-game-answer-to-left) | Saves references to the selected cells in the green answer cells to the left on the same row. |

### Select

| Name | Description |
| --- | --- |
| [Goto Similar Background Color](#goto-similar-background-color) | Select cells in selection with same background color as active cell. |
| [Goto Similar Constant Values](#goto-similar-constant-values) | Select constant cells in selection with similar value as active cell. |
| [Goto Similar Formulas](#goto-similar-formulas) | Select formula cells in selection with similar formula as active cell. |

<BR>

## Available Robot Parameters

| Name | Description |
| --- | --- |
| [Active\_Column\_Index\_In\_Spilling\_Range](#active_column_index_in_spilling_range) | Returns the index of the active cell column within the spill range. |

<BR>

## Available Robot Texts

| Name | Description |
| --- | --- |
| [FilterArray.lambda](#filterarraylambda) | Definition of FilterArray lambda function. |
| [GroupByFirstNColumns.lambda](#groupbyfirstncolumnslambda) | Definition of GroupByFirstNColumns lambda function. |
| [InsertCols.lambda](#insertcolslambda) | Definition of InsertCols lambda function. |
| [IsInList.lambda](#isinlistlambda) | Definition of IsInList lambda function. |
| [LeastFrequentValue.lambda](#leastfrequentvaluelambda) | Definition of LeastFrequentValue lambda function. |
| [MostFrequentValue.lambda](#mostfrequentvaluelambda) | Definition of MostFrequentValue lambda function. |
| [RemoveCols.lambda](#removecolslambda) | Definition of RemoveCols lambda function. |
| [UngroupColumn.lambda](#ungroupcolumnlambda) | Definition of UngroupColumn lambda function. |

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

### Filter Array By Selected Values

*Wraps array formula with function returning active array filtered by selected values in a column of active array.*

<sup>`@MEWC Robot.xlsm` `!Excel Formula Command` `#Filter`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=FilterArray(\[\[ActiveCell.SpillParent::Formula\]\],{{Selected\_Column\_Indexes\_In\_Spilling\_Range}},\[\[Selection::ValueArray\]\])</code> |
| Destination Range Address | <code>\[ActiveCell.SpillParent\]</code> |
| Formula Dependencies | <ol><li>[IsInList.lambda](#isinlistlambda)</li><li>[FilterArray.lambda](#filterarraylambda)</li></ol> |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Launch Codes | <ol><li><code>fa</code></li><li><code>fasv</code></li></ol> |

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

### Group By Average

*Groups by first N column, aggregating remaining columns with AVERAGE function.*

<sup>`@MEWC Robot.xlsm` `!Excel Formula Command` `#Group`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=GroupByFirstNColumns(\[\[ActiveCell.SpillParent::Formula\]\],{{FirstNColumns}},,AVERAGE)</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Formula Dependencies | [GroupByFirstNColumns.lambda](#groupbyfirstncolumnslambda) |
| Parameters | <ol><li>[FirstNColumns](#group-by-average--firstncolumns)</li></ol> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Outputs | <ol></ol> |
| Launch Codes | <code>gb</code> |

<BR>

#### Group By Average \>\> FirstNColumns

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Data will be grouped by the first N columns of the array and remaining columns will be aggregated.</code><br><code></code><br><code>How many columns would you like to group?</code> |
| Default Value | <code>1</code> |

[^Top](#oa-robot-definitions)

<BR>

### Group By Concat With Delimiter

*Groups by first N column, concatenating remaining columns with specified delimiter.*

<sup>`@MEWC Robot.xlsm` `!Excel Formula Command` `#Group`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=GroupByFirstNColumns(\[\[ActiveCell.SpillParent::Formula\]\],{{FirstNColumns}},,LAMBDA(x,TEXTJOIN("{{Text\_Delimiter}}",,x)))</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Formula Dependencies | [GroupByFirstNColumns.lambda](#groupbyfirstncolumnslambda) |
| Parameters | <ol><li>[FirstNColumns](#group-by-concat-with-delimiter--firstncolumns)</li><li>[Text_Delimiter](#group-by-concat-with-delimiter--text_delimiter)</li></ol> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Outputs | <ol></ol> |
| Launch Codes | <code>gb</code> |

<BR>

#### Group By Concat With Delimiter \>\> FirstNColumns

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Data will be grouped by the first N columns of the array and remaining columns will be aggregated.</code><br><code></code><br><code>How many columns would you like to group?</code> |
| Priority | <code>1</code> |
| Default Value | <code>1</code> |

<BR>

#### Group By Concat With Delimiter \>\> Text\_Delimiter

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Enter the delimiter:</code> |
| Show Selector | NotAvailable |
| Data Type | String |
| Priority | <code>2</code> |
| Default Value | <code>,</code> |

[^Top](#oa-robot-definitions)

<BR>

### Group By Function\/Lambda

*Groups by first N columns, aggregating remaining columns with specified function or lambda.*

<sup>`@MEWC Robot.xlsm` `!Excel Formula Command` `#Group`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=GroupByFirstNColumns(\[\[ActiveCell.SpillParent::Formula\]\],{{FirstNColumns}},,{{Function\_Or\_Lambda}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Formula Dependencies | [GroupByFirstNColumns.lambda](#groupbyfirstncolumnslambda) |
| Parameters | <ol><li>[FirstNColumns](#group-by-functionlambda--firstncolumns)</li><li>[Function_Or_Lambda](#group-by-functionlambda--function_or_lambda)</li></ol> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Outputs | <ol></ol> |
| Launch Codes | <code>gb</code> |

<BR>

#### Group By Function\/Lambda \>\> FirstNColumns

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Data will be grouped by the first N columns of the array and remaining columns will be aggregated.</code><br><code></code><br><code>How many columns would you like to group?</code> |
| Show Selector | NotAvailable |
| Priority | <code>1</code> |
| Default Value | <code>1</code> |

<BR>

#### Group By Function\/Lambda \>\> Function\_Or\_Lambda

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Enter function name or define lambda:</code><br><code></code><br><code>Example: MIN or LAMBDA(x,SQRT(x^2))</code> |
| Show Selector | NotAvailable |
| Priority | <code>2</code> |
| Default Value | <code>LAMBDA(x, </code> |

[^Top](#oa-robot-definitions)

<BR>

### Group By Sum

*Groups by first N column, aggregating remaining columns with SUM function.*

<sup>`@MEWC Robot.xlsm` `!Excel Formula Command` `#Group`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=GroupByFirstNColumns(\[\[ActiveCell.SpillParent::Formula\]\],{{FirstNColumns}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Formula Dependencies | [GroupByFirstNColumns.lambda](#groupbyfirstncolumnslambda) |
| Parameters | <ol><li>[FirstNColumns](#group-by-sum--firstncolumns)</li></ol> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange |
| Outputs | <ol></ol> |
| Launch Codes | <code>gb</code> |

<BR>

#### Group By Sum \>\> FirstNColumns

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Data will be grouped by the first N columns of the array and remaining columns will be aggregated.</code><br><code></code><br><code>How many columns would you like to group?</code> |
| Default Value | <code>1</code> |

[^Top](#oa-robot-definitions)

<BR>

### Least Frequent Value Of Array

*Wrap with LeastFrequentValue Lambda function.*

<sup>`@MEWC Robot.xlsm` `!Excel Formula Command` `#Analysis`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=LeastFrequentValue(\[\[ActiveCell::Formula\]\],"")</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [LeastFrequentValue.lambda](#leastfrequentvaluelambda) |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>lf</code></li><li><code>lo</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Most Frequent Value Of Array

*Wrap with MostFrequentValue Lambda function.*

<sup>`@MEWC Robot.xlsm` `!Excel Formula Command` `#Analysis`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=MostFrequentValue(\[\[ActiveCell::Formula\]\],"")</code> |
| Scroll To Destination | ☐Yes ☑No |
| Formula Dependencies | [MostFrequentValue.lambda](#mostfrequentvaluelambda) |
| Update Formula Dependencies | ☑Yes ☐No |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <ol><li><code>mf</code></li><li><code>mo</code></li></ol> |

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

### Paste Filter Array By Copied Cell

*Wraps array formula with function returning active array filtered by copied cell in a column of active array.*

<sup>`@MEWC Robot.xlsm` `!Excel Formula Command` `#Filter`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=FilterArray(\[\[ActiveCell.SpillParent::Formula\]\],{{Selected\_Column\_Indexes\_In\_Spilling\_Range}},\[\[Clipboard\]\])</code> |
| Destination Range Address | <code>\[ActiveCell.SpillParent\]</code> |
| Formula Dependencies | <ol><li>[IsInList.lambda](#isinlistlambda)</li><li>[FilterArray.lambda](#filterarraylambda)</li></ol> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange AND ClipboardHasExcelData AND ExcelCopiedRangeIsSingleCell |
| Launch Codes | <code>pfa</code> |

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

<sup>`@MEWC Robot.xlsm` `!VBA Macro Command` `#Save`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modGamePrep.SaveAnswersToLeft](./VBA/modGamePrep.bas#L225)()</code> |
| Keyboard Shortcut | <code>^+s</code> |
| User Context Filter | ExcelSelectionIsMultipleRows AND ExcelSelectionIsSingleColumn |
| Launch Codes | <code>sa</code> |

[^Top](#oa-robot-definitions)

<BR>

### Ungroup Column By Delimiter

*Splits the values in the selected column and expands the other columns.*

<sup>`@MEWC Robot.xlsm` `!Excel Formula Command` `#Group`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=UngroupColumn(\[\[ActiveCell.SpillParent::Formula\]\],{{Active\_Column\_Index\_In\_Spilling\_Range}},"{{Delimiter}}")</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Formula Dependencies | <ol><li>[UngroupColumn.lambda](#ungroupcolumnlambda)</li><li>[RemoveCols.lambda](#removecolslambda)</li><li>[InsertCols.lambda](#insertcolslambda)</li></ol> |
| Parameters | <ol><li>[Delimiter](#ungroup-column-by-delimiter--delimiter)</li></ol> |
| User Context Filter | ExcelActiveCellIsInSpillingToRange AND ExcelSelectionIsSingleCell |
| Outputs | <ol></ol> |
| Launch Codes | <ol><li><code>ug</code></li><li><code>ugc</code></li></ol> |

<BR>

#### Ungroup Column By Delimiter \>\> Delimiter

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Enter the delimiter to split by:</code> |
| Show Selector | NotAvailable |
| Default Value | <code>,</code> |

[^Top](#oa-robot-definitions)

<BR>

## Parameter Definitions

<BR>

### Active\_Column\_Index\_In\_Spilling\_Range

*Returns the index of the active cell column within the spill range.*

<sup>`@MEWC Robot.xlsm` `!VBA Macro Parameter` </sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modRange.ActiveColumnIndexInSpillingRange](./VBA/modRange.bas#L4)([[ActiveCell]])</code> |
| Data Type | Integer |

[^Top](#oa-robot-definitions)

<BR>

## Text Definitions

<BR>

### FilterArray.lambda

*Definition of FilterArray lambda function.*

<sup>`@MEWC Robot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [FilterArray.lambda](<./Text/FilterArray.lambda.txt>) |
| Value | <code>FilterArray \= LAMBDA(data,column\_indexes,filter\_values,LET(</code><br><code> \\\\LambdaName, "FilterArray",</code><br><code>FILTER(data,BYROW(IsInList(CHOOSECOLS(data,column\_indexes),filter\_values),LAMBDA(x,AND(x)))))</code><br><code>);</code> |
| Location | <code>FilterArray</code> |
| Source Credit | <code>@ExcelRobot</code> |
| Markdown Id | <code>FilterArraylambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### GroupByFirstNColumns.lambda

*Definition of GroupByFirstNColumns lambda function.*

<sup>`@MEWC Robot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [GroupByFirstNColumns.lambda](<./Text/GroupByFirstNColumns.lambda.txt>) |
| Value | <code>GroupByFirstNColumns \= LAMBDA(array,\[n\],\[include\_grand\_total\],\[aggregate\_function\], LET(</code><br><code> \\\\LambdaName, "GroupByFirstNColumns",</code><br><code> n, IF(ISOMITTED(n), 1, n),</code><br><code> include\_grand\_total, IF(</code><br><code> ISOMITTED(include\_grand\_total),</code><br><code> 0,</code><br><code> include\_grand\_total</code><br><code> ),</code><br><code> aggregate\_function, IF(</code><br><code> ISOMITTED(aggregate\_function),</code><br><code> SUM,</code><br><code> aggregate\_function</code><br><code> ),</code><br><code> \_RowFields, TAKE(array, , n),</code><br><code> \_Values, DROP(array, , n),</code><br><code> \_Result, GROUPBY(</code><br><code> \_RowFields,</code><br><code> \_Values,</code><br><code> aggregate\_function,</code><br><code> 0,</code><br><code> include\_grand\_total</code><br><code> ),</code><br><code> \_Result</code><br><code>));</code> |
| Location | <code>GroupByFirstNColumns</code> |
| Source Credit | <code>@ExcelRobot</code> |

[^Top](#oa-robot-definitions)

<BR>

### InsertCols.lambda

*Definition of InsertCols lambda function.*

<sup>`@MEWC Robot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [InsertCols.lambda](<./Text/InsertCols.lambda.txt>) |
| Value | <code>InsertCols \= LAMBDA(array,columns\_to\_insert,\[column\_index\],\[pad\_with\], LET(</code><br><code> \\\\LambdaName, "InsertCols",</code><br><code> \\\\CommandName, "Insert Columns Into Array",</code><br><code> \\\\Description, "Given an array, inserts a second array at a specified column index.",</code><br><code> \\\\Parameters, {"array","the target array within which columns will be inserted";"columns\_to\_insert","an array with one or more columns of data to be inserted in target array";"\[column\_index\]","column index location where columns will be inserted; if negative, counts from the right; if zero (default), columns will be appended to right";"\[pad\_with\]","what to fill blanks with if columns to insert has different number of rows than array; default: \#N\/A"},</code><br><code> \\\\Source, "Excel Robot (@ExcelRobot)",</code><br><code> \_PadWith, IF(ISOMITTED(pad\_with), NA(), pad\_with),</code><br><code> \_Array, IF(</code><br><code> ROWS(columns\_to\_insert) \> ROWS(array),</code><br><code> EXPAND(array, ROWS(columns\_to\_insert), , \_PadWith),</code><br><code> array</code><br><code> ),</code><br><code> \_ColumnsToInsert, IF(</code><br><code> ROWS(\_Array) \> ROWS(columns\_to\_insert),</code><br><code> EXPAND(columns\_to\_insert, ROWS(\_Array), , \_PadWith),</code><br><code> columns\_to\_insert</code><br><code> ),</code><br><code> \_ColumnIndex, IF(@column\_index \< 0, MAX(1, COLUMNS(\_Array) + @column\_index + 1), @column\_index),</code><br><code> IFS(</code><br><code> \_ColumnIndex \= 1,</code><br><code> HSTACK(\_ColumnsToInsert, \_Array),</code><br><code> OR(\_ColumnIndex \<\= 0, \_ColumnIndex \> COLUMNS(\_Array)),</code><br><code> HSTACK(\_Array, \_ColumnsToInsert),</code><br><code> TRUE,</code><br><code> HSTACK(TAKE(\_Array, , \_ColumnIndex \- 1), \_ColumnsToInsert, DROP(\_Array, , \_ColumnIndex \- 1))</code><br><code> )</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>InsertCols</code> |
| Source Credit | <code>@ExcelRobot</code> |

[^Top](#oa-robot-definitions)

<BR>

### IsInList.lambda

*Definition of IsInList lambda function.*

<sup>`@MEWC Robot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [IsInList.lambda](<./Text/IsInList.lambda.txt>) |
| Value | <code>IsInList \= LAMBDA(array,list,LET(</code><br><code> \\\\LambdaName, "IsInList",</code><br><code>MAP(array,LAMBDA(x,OR(list\=x))))</code><br><code>);</code> |
| Location | <code>IsInList</code> |
| Source Credit | <code>@ExcelRobot</code> |
| Markdown Id | <code>IsInListlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### LeastFrequentValue.lambda

*Definition of LeastFrequentValue lambda function.*

<sup>`@MEWC Robot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [LeastFrequentValue.lambda](<./Text/LeastFrequentValue.lambda.txt>) |
| Value | <code>LeastFrequentValue \= LAMBDA(array,\[exclude\], LET(</code><br><code> \\\\LambdaName, "LeastFrequent",</code><br><code> \_unique, UNIQUE(TOCOL(IF(array \= "", "", array))),</code><br><code> \_include, BYROW(\_unique \<\> TOROW(exclude), LAMBDA(x, PRODUCT(N(x)))),</code><br><code> \_filtered, IF(ISOMITTED(exclude), \_unique, FILTER(\_unique, \_include)),</code><br><code> \_countif, MAP(\_filtered, LAMBDA(x, SUM(N(array \= x)))),</code><br><code> Result, TAKE(SORTBY(\_filtered, \_countif, 1), 1),</code><br><code> Result</code><br><code>));</code> |
| Content Type | ExcelLambda |
| Location | <code>LeastFrequentValue</code> |
| Source Credit | <code>@ExcelRobot</code> |
| Markdown Id | <code>MostFrequentValuelambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### MostFrequentValue.lambda

*Definition of MostFrequentValue lambda function.*

<sup>`@MEWC Robot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [MostFrequentValue.lambda](<./Text/MostFrequentValue.lambda.txt>) |
| Value | <code>MostFrequentValue \= LAMBDA(array,\[exclude\], LET(</code><br><code> \\\\LambdaName, "MostFrequentValue",</code><br><code> \\\\CommandName, "Most Frequent Value",</code><br><code> \\\\Description, "Returns the most frequent value in an array.",</code><br><code> \_unique, UNIQUE(TOCOL(IF(array \= "", "", array))),</code><br><code> \_include, BYROW(\_unique \<\> TOROW(exclude), LAMBDA(x, PRODUCT(N(x)))),</code><br><code> \_filtered, IF(ISOMITTED(exclude), \_unique, FILTER(\_unique, \_include)),</code><br><code> \_countif, MAP(\_filtered, LAMBDA(x, SUM(N(array \= x)))),</code><br><code> Result, TAKE(SORTBY(\_filtered, \_countif, \-1), 1),</code><br><code> Result</code><br><code>));</code> |
| Content Type | ExcelLambda |
| Location | <code>MostFrequentValue</code> |
| Source Credit | <code>@ExcelRobot</code> |
| Markdown Id | <code>MostFrequentValuelambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### RemoveCols.lambda

*Definition of RemoveCols lambda function.*

<sup>`@MEWC Robot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [RemoveCols.lambda](<./Text/RemoveCols.lambda.txt>) |
| Value | <code>RemoveCols \= LAMBDA(array,column\_indexes, LET(</code><br><code> \\\\LambdaName, "RemoveCols",</code><br><code> \\\\CommandName, "Remove Columns Of Array",</code><br><code> \\\\Description, "Removes specified columns of array using RemoveCols lambda.",</code><br><code> \_Seq, SEQUENCE(COLUMNS(array)),</code><br><code> \_Keep, ISERROR(MATCH(\_Seq, TOROW(column\_indexes), 0)),</code><br><code> \_Included, FILTER(\_Seq, \_Keep, TRUE),</code><br><code> \_Result, CHOOSECOLS(array, \_Included),</code><br><code> \_Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>RemoveCols</code> |
| Source Credit | <code>@ExcelRobot</code> |

[^Top](#oa-robot-definitions)

<BR>

### UngroupColumn.lambda

*Definition of UngroupColumn lambda function.*

<sup>`@MEWC Robot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [UngroupColumn.lambda](<./Text/UngroupColumn.lambda.txt>) |
| Value | <code>UngroupColumn \= LAMBDA(data,\[column\_index\],\[delimiter\], LET(</code><br><code> \\\\LambdaName, "UngroupColumn",</code><br><code> \_ColIndex, IF(</code><br><code> OR(ISOMITTED(column\_index), ISBLANK(column\_index), column\_index \= ""),</code><br><code> COLUMNS(data),</code><br><code> column\_index</code><br><code> ),</code><br><code> \_Delim, IF(OR(ISOMITTED(delimiter), ISBLANK(delimiter)), ",", delimiter),</code><br><code> \_Rows, ROWS(data),</code><br><code> \_Cols, COLUMNS(data),</code><br><code> \_Keys, IF(\_ColIndex \= \_Cols, DROP(data, , \-1), RemoveCols(data, \_ColIndex)),</code><br><code> \_Values, CHOOSECOLS(data, \_ColIndex),</code><br><code> \_Result, REDUCE(</code><br><code> "",</code><br><code> SEQUENCE(\_Rows),</code><br><code> LAMBDA(acc,r, LET(</code><br><code> key, CHOOSEROWS(\_Keys, r),</code><br><code> vals, INDEX(\_Values, r, 1),</code><br><code> valueList, IF(\_Delim \= "", MID(vals, SEQUENCE(LEN(vals)), 1), TEXTSPLIT(vals, , \_Delim)),</code><br><code> trimmed, TRIM(valueList),</code><br><code> keyCols, MAKEARRAY(ROWS(trimmed), COLUMNS(key), LAMBDA(row,col, INDEX(key, 1, col))),</code><br><code> allCols, IF(</code><br><code> \_ColIndex \= \_Cols,</code><br><code> HSTACK(keyCols, trimmed),</code><br><code> InsertCols(keyCols, trimmed, \_ColIndex)</code><br><code> ),</code><br><code> IF(@acc \= "", allCols, VSTACK(acc, allCols))</code><br><code> ))</code><br><code> ),</code><br><code> \_Result</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>UngroupColumn</code> |
| Source Credit | <code>@ExcelRobot</code> |

[^Top](#oa-robot-definitions)
