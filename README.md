# VBA Curriculum
VBA curriculum for training co-workers.

## Modules 1-2

- Visual Basic Editor introduction
  - Macro-enabled filetypes
  - Developer ribbon
  - Macro recording, macro start, macro stop
  - Subs
  - Comments
- Variables
  - Naming
  - Declaration (Integer, String, Boolean, Object)
  - Variable scope, public variables
  - Option Explicit
- Cells, Range, Columns, Rows
- Variables to Cells, Cells to variables
- Workbooks and Sheets
  - ActiveWorkbook, ThisWorkbook
  - Activesheet
  - Sheets(), .Add, .Count, .Name
- Arithmetic operators (+ - * /, =, <, >, >=, <=, <>, mod, \)
- Rounding: Round, Int, Cint
- Random numbers
- String manipulation (concatenation, left, right, mid, len, format)
- Number to String: CStr(), String to Number: Val()
- Msgbox

**Practice:**
1. Collect two numbers from two different cells, then multiply them, then give a third cell this calculated value.
2. Collect first name and last name from two different cells, msgbox "Greetings Mr. Firstname Lastname!"

## Module 3

- IF THEN ELSE ELSEIF, nested IFS
- Logical operators (And, Or, Not)
- FOR NEXT loop
- Nested FOR NEXT loops

**Practice:**
1. Collect name, age and sex from a table. If the person's name is more than 7 characters long, check if the person is a she, if yes, check if she's between 20 and 30 years old, then msgbox true or false.
2. Loop through a table (Header: product name, quantity, unit price, discount%, VAT), calculate the total price for each product, and add the amount to the next cell.

## Intermediate VBA

- Arrays
- Workbook manipulation
  - Last row
  - Copy and Paste, Range copy with destination, Range copy with .Value and .Formula
  - Delete Columns, Rows and Clearcontents
  - Column autofit
  - Text to columns
- Autofilter, copy filtered data
- Modules
- File operations
  - Open, Close
  - Save, SaveAs
  - Set wb, Set ws
- More Msgbox
- Excel formulas in vba (native and recorded formulas)
- Constants
- DO WHILE loop

## Advanced VBA

- Error handling
- Functions
- SAP scripts
  - Recording
  - Add variables
  - Optimization
- IIF
- With statement
- Optimization, bad practices (loops, select, activate)
