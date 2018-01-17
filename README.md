# VBA Curriculum
VBA curriculum for training co-workers.

## Table of Contents

1. [Beginner VBA](https://github.com/codergy/vbacurriculum#beginner-vba-grinning)
   1. [Module 1](https://github.com/codergy/vbacurriculum#module-1) - Introduction, variables, cells and ranges
   1. [Module 2](https://github.com/codergy/vbacurriculum#module-2) - Workbooks, sheets, numbers, strings
   1. [Module 3](https://github.com/codergy/vbacurriculum#module-3) - If, logical operators, for next loop
1. [Intermediate VBA](https://github.com/codergy/vbacurriculum#intermediate-vba-metal) - Arrays, workbook manipulation, file operations, more loops
1. [Advanced VBA](https://github.com/codergy/vbacurriculum#advanced-vba-trollface) - Special tasks and fine tuning
1. [Tips and tricks](https://github.com/codergy/vbacurriculum#tips-and-tricks-bowtie) - Snippets and optimization

---

## Beginner VBA :grinning:

## Module 1
- [ ] **Visual Basic Editor introduction**
  - Macro-enabled filetypes
  - Developer ribbon
  - Macro recording, macro start, macro stop, step into, break point
  - Subs
  - Comments
- [ ] **Variables**
  - [Naming](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/visual-basic-naming-rules)
  - [Declaration](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/declaring-variables)
  - [Data types](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/data-types/data-type-summary) (Integer, String, Boolean, Range, Object)
  - [Variable scope](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/understanding-scope-and-visibility), public and private variables
  - [Option Explicit](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/option-explicit-statement)
- [ ] Cells(), Range(), Columns(), Rows()
- [ ] Variables to Cells, Cells to variables

## Module 2

- [ ] **Workbooks and Sheets**
  - [ActiveWorkbook vs ThisWorkbook](http://analystcave.com/vba-tip-day-activeworkbook-vs-thisworkbook/)
  - [Worksheets vs Sheet vs Activesheet](http://analystcave.com/excel-vba-worksheets-tutorial-vba-activesheet-vs-worksheets/)
  - [.Select vs .Activate](https://stackoverflow.com/questions/7180008/excel-select-vs-activate)
  - .Add, .Count, .Name, .Delete
- [ ] **Numbers**
  - [Arithmetic operators](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/operators/arithmetic-operators) (+ - * /, mod, \\)
  - [Comparison operators](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/operators/comparison-operators) (=, <, >, >=, <=, <>)
  - Rounding: [Round()](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/worksheetfunction-round-method-excel), Int(), Fix(), Cint()
  - Random numbers
- [ ] **Strings**
    - [Concatenation](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/operators-and-expressions/concatenation-operators)
    - [Left()](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/left-function), [Right()](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/right-function), [Mid()](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/mid-function)
    - [Format()](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/format-function-visual-basic-for-applications), [LCase()](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/lcase-function), [UCase](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/ucase-function)
    - [Len()](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/len-function)
- [ ] Number to String: [CStr()](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/functions/type-conversion-functions), String to Number: [Val()](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/val-function)
- [ ] Msgbox

**Practice:**
1. Collect two numbers from two different cells, then multiply them, then give a third cell this calculated value.
1. Collect first name and last name from two different cells, msgbox "Greetings Mr. Firstname Lastname!"

## Module 3

- [ ] Logical operators ([And, Or, Not](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/operators-and-expressions/logical-and-bitwise-operators))
- [ ] [If, Then, Else, ElseIf, nested If](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/if-then-else-statement)
- [ ] [For Next loop](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/for-next-statement)
- [ ] Nested For Next loops

**Practice:**
1. Collect name, age and sex from a table. If the person's name is more than 7 characters long, check if the person is a she, if yes, check if she's between 20 and 30 years old, then msgbox true or false.
2. Loop through a table (Header: product name, quantity, unit price, discount%, VAT), calculate the total price for each product, and add the amount to the next cell.

## Intermediate VBA :metal:

- [ ] Arrays
- [ ] **Workbook manipulation**
  - Last row
  - Copy and Paste, Range copy with destination, Range copy with .Value and .Formula
  - Delete columns, rows and Clearcontents
  - Column autofit
  - Text to columns
- [ ] Autofilter, copy filtered data
- [ ] Modules
- [ ] **File operations**
  - Open, Close
  - Save, SaveAs
  - Set wb, Set ws
  - Dir, MkDir
- [ ] More Msgbox
- [ ] Excel formulas in vba (native and recorded formulas)
- [ ] Constants
- [ ] Do While loop
- [ ] For Each loop

## Advanced VBA :trollface:

- [ ] Error handling [(1)](http://analystcave.com/vba-proper-vba-error-handling/) [(2)](https://excelmacromastery.com/vba-error-handling/)
- [ ] Functions
- [ ] **SAP scripts**
  - Recording
  - Add variables
  - Optimization
- [ ] [Email sending](https://www.rondebruin.nl/win/s1/outlook/mail.htm)
- [ ] [With statement](https://www.homeandlearn.org/with_end_with.html)
- [ ] [Iif()](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/iif-function)
- [ ] [Read data from another workbook without opening it](https://github.com/codergy/vba-snippets/blob/master/README.md#read-data-from-another-workbook-without-opening-it)

## Tips and tricks :bowtie:

- [My favorite VBA snippets](https://github.com/codergy/vba-snippets/blob/master/README.md)
- Speed improvement [(1)](http://analystcave.com/excel-improve-vba-performance/) [(2)](http://www.ozgrid.com/VBA/SpeedingUpVBACode.htm)
