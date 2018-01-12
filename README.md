# VBA Curriculum
VBA curriculum for training co-workers.

## Module 1

- Visual Basic Editor introduction
  - Macro-enabled filetypes
  - Developer ribbon
  - Macro recording, macro start, macro stop, step into, break point
  - Subs
  - Comments
- Variables
  - Naming
  - Declaration (Integer, String, Boolean, Range, Object)
  - Variable scope, public variables
  - Option Explicit
- Cells(), Range(), Columns(), Rows()
- Variables to Cells, Cells to variables

## Module 2

- Workbooks and Sheets
  - ActiveWorkbook, ThisWorkbook
  - Activesheet
  - Sheets(), .Add, .Count, .Name
- Numbers
  - Arithmetic operators (+ - * /, =, <, >, >=, <=, <>, mod, \)
  - Rounding: Round(), Int(), Cint()
  - Random numbers
- Strings
    - Concatenation
    - Left(), Right(), Mid()
    - Format(), LCase(), UCase
    - Len()
- Number to String: CStr(), String to Number: Val()
- Msgbox

**Practice:**
1. Collect two numbers from two different cells, then multiply them, then give a third cell this calculated value.
2. Collect first name and last name from two different cells, msgbox "Greetings Mr. Firstname Lastname!"

## Module 3

- If, Then, Else, ElseIf, nested If
- Logical operators (And, Or, Not)
- For Next loop
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
