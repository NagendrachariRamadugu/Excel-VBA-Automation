#  Excel File Automation with VBA (AI-Assisted)

This repository contains a simple but powerful **VBA automation** solution to:
1. Create 100 Excel files
2. Rename them programmatically
3. Delete them when needed

The code was generated with the help of **ChatGPT** (AI), showcasing how artificial intelligence can assist in automating repetitive office tasks and improving productivity — even with tools like Microsoft Excel.

## Why This Project?

Organizations today encourage professionals to use AI tools to **boost productivity**, not to replace them.  
This project demonstrates a practical use case:  
> "How can AI help me reduce hours of manual Excel file handling to just seconds?"

##  What's Inside?

| File                  | Description                                  |
|-----------------------|----------------------------------------------|
| `CreateFiles.bas`     | VBA module to create 100 Excel workbooks     |
| `RenameFiles.bas`     | VBA module to rename all `.xlsx` files in a folder in order (`Book1`, `Book2`, etc.) |
| `DeleteFiles.bas`     | VBA module to delete all `.xlsx` files in a folder |
| `ExcelAutomation.xlsm`| A macro-enabled workbook containing the code, ready to run |

##  How to Use

### Option 1: Import `.bas` files into Excel

1. Open Excel → Press `Alt + F11` to open the **VBA Editor**
2. Right-click your VBA project → **Import File...**
3. Choose `CreateFiles.bas`, `RenameFiles.bas`, or `DeleteFiles.bas`
4. Run the macro using `F5` or assign it to a button

###  Option 2: Use the ready-made `.xlsm` file (included)

1. Open `ExcelAutomation.xlsm`
2. Enable macros when prompted
3. Run the required macro from `Developer → Macros`
