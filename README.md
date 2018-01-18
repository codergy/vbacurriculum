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
  - [x] Macro-enabled filetypes
  - [x] Developer ribbon
  - [ ] Macro recording, macro start, macro stop, [step through](https://www.wiseowl.co.uk/blog/s196/step-through-code.htm), [break point](https://www.wiseowl.co.uk/blog/s196/breakpoints.htm)
  - [x] Subs
  - [x] Modules
  - [x] Comments
- [ ] **Variables**
  - [ ] [Naming](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/visual-basic-naming-rules)
  - [x] [Declaration](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/declaring-variables)
  - [x] [Data types](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/data-types/data-type-summary) (Integer, String, Boolean, Range, Object, [all numeric data types](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/data-types/numeric-data-types))
  - [ ] [Variable scope](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/understanding-scope-and-visibility), public and private variables
  - [ ] [Option Explicit](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/option-explicit-statement) - [why to use it](http://www.excelkey.com/forum/viewtopic.php?f=7&t=417)
- [ ] Cells(), Range(), Columns(), Rows() - [short tutorial](http://www.excel-easy.com/vba/range-object.html)
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
    - [Trim()](https://github.com/codergy/vba-snippets/blob/master/README.md#trim-a-whole-range)
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

- [ ] [Arrays](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/arrays/)
- [ ] **Workbook manipulation**
  - [Last row](https://github.com/codergy/vba-snippets/blob/master/README.md#last-row)
  - Copy and Paste, [Range copy with destination, Range copy with .Value and .Formula](https://github.com/codergy/vba-snippets/blob/master/README.md#fast-range-copy)
  - Delete columns, rows and Clearcontents
  - Columns.AutoFit
  - Text to columns, [reset delimiter](https://github.com/codergy/vba-snippets/blob/master/README.md#reset-text-to-column-delimiter)
- [ ] Autofilter, [copy filtered rows](https://github.com/codergy/vba-snippets/blob/master/README.md#copy-filtered-rows)
- [ ] **File operations**
  - [Open](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/workbooks-open-method-excel), [Close](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/workbook-close-method-excel)
  - [Save](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/workbook-save-method-excel), [SaveAs](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/workbook-saveas-method-excel)
  - Set wb, Set ws
  - [Dir()](http://analystcave.com/vba-reference-functions/vba-file-functions/vba-dir-function/), [MkDir()](http://analystcave.com/vba-reference-functions/vba-file-functions/vba-mkdir-function/), [create folder if it doesn't exist](https://www.mrexcel.com/forum/excel-questions/575970-determine-if-directory-exists-if-not-create.html)
- [ ] Functions
- [ ] [Excel formulas in vba](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/worksheetfunction-object-excel)
- [ ] [VBA functions as Excel formulas](http://www.fontstuff.com/vba/vbatut01.htm)
- [ ] [Constants](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/declaring-constants)
- [ ] [Do While/Until loop](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/using-doloop-statements)
- [ ] [For Each loop](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/for-each-next-statement)
- [ ] [More Msgbox](https://www.mrexcel.com/forum/excel-questions/492894-msgbox-vbyesno.html)

## Advanced VBA :trollface:

- [ ] Error handling [(1)](http://analystcave.com/vba-proper-vba-error-handling/) [(2)](https://excelmacromastery.com/vba-error-handling/)
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
