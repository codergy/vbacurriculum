# VBA Curriculum
VBA curriculum for training co-workers.

## Table of Contents

- [Module 1](https://github.com/codergy/vbacurriculum#module-1) - Introduction, variables, cells and ranges
- [Module 2](https://github.com/codergy/vbacurriculum#module-2) - Workbooks, sheets, numbers, strings
- [Module 3](https://github.com/codergy/vbacurriculum#module-3) - If, logical operators, for next loop
- [Intermediate VBA](https://github.com/codergy/vbacurriculum#intermediate-vba) - Arrays, workbook manipulation, file operations, more loops
- [Advanced VBA](https://github.com/codergy/vbacurriculum#advanced-vba) - Special tasks and fine tuning
- [Tips and tricks](https://github.com/codergy/vbacurriculum#tips-and-tricks) - Snippets and optimization

## Module 1

- Visual Basic Editor introduction
  - Macro-enabled filetypes
  - Developer ribbon
  - Macro recording, macro start, macro stop, step into, break point
  - Subs
  - Comments
- Variables
  - Naming
  - [Declaration](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/declaring-variables)
  - Data types (Integer, String, Boolean, Range, Object) - [All types](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/data-types/data-type-summary)
  - [Variable scope](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/understanding-scope-and-visibility), public and private variables
  - [Option Explicit](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/option-explicit-statement)
- Cells(), Range(), Columns(), Rows()
- Variables to Cells, Cells to variables

## Module 2

- Workbooks and Sheets
  - [ActiveWorkbook vs ThisWorkbook](http://analystcave.com/vba-tip-day-activeworkbook-vs-thisworkbook/)
  - [Worksheets vs Sheet vs Activesheet](http://analystcave.com/excel-vba-worksheets-tutorial-vba-activesheet-vs-worksheets/)
  - [.Select vs .Activate](https://stackoverflow.com/questions/7180008/excel-select-vs-activate)
  - .Add, .Count, .Name, .Delete
- Numbers
  - [Arithmetic operators](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/operators/arithmetic-operators) (+ - * /, mod, \\)
  - [Comparison operators](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/operators/comparison-operators) (=, <, >, >=, <=, <>)
  - Rounding: Round(), Int(), Cint()
  - Random numbers
- Strings
    - [Concatenation](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/operators-and-expressions/concatenation-operators)
    - [Left()](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/left-function), [Right()](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/right-function), [Mid()](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/mid-function)
    - [Format()](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/format-function-visual-basic-for-applications), [LCase()](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/lcase-function), [UCase](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/ucase-function)
    - [Len()](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/len-function)
- Number to String: [CStr()](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/functions/type-conversion-functions), String to Number: [Val()](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/val-function)
- Msgbox

**Practice:**
1. Collect two numbers from two different cells, then multiply them, then give a third cell this calculated value.
2. Collect first name and last name from two different cells, msgbox "Greetings Mr. Firstname Lastname!"

## Module 3

- [If, Then, Else, ElseIf, nested If](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/if-then-else-statement)
- Logical operators ([And, Or, Not](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/operators-and-expressions/logical-and-bitwise-operators))
- [For Next loop](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/for-next-statement)
- Nested For Next loops

**Practice:**
1. Collect name, age and sex from a table. If the person's name is more than 7 characters long, check if the person is a she, if yes, check if she's between 20 and 30 years old, then msgbox true or false.
2. Loop through a table (Header: product name, quantity, unit price, discount%, VAT), calculate the total price for each product, and add the amount to the next cell.

## Intermediate VBA

- Arrays
- Workbook manipulation
  - Last row
  - Copy and Paste, Range copy with destination, Range copy with .Value and .Formula
  - Delete columns, rows and Clearcontents
  - Column autofit
  - Text to columns
- Autofilter, copy filtered data
- Modules
- File operations
  - Open, Close
  - Save, SaveAs
  - Set wb, Set ws
  - Dir, MkDir
- More Msgbox
- Excel formulas in vba (native and recorded formulas)
- Constants
- Do While loop
- For Each loop

## Advanced VBA

- Error handling
- Function
- SAP scripts
  - Recording
  - Add variables
  - Optimization
- Email sending
- With statement
- Iif()
- Read data from another workbook without opening it
- Optimization, bad practices (loops, select, activate, deleting columns)

## Tips and tricks

- My favorite VBA snippets
- Speed improvement [(1)](http://analystcave.com/excel-improve-vba-performance/) [(2)](http://www.ozgrid.com/VBA/SpeedingUpVBACode.htm)
