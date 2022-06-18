'BASIC: APPLICATION PROPERTY AND METHOD

'1. WHEN TO USE? 
'   If you want to manage Excel file (ex. workbook) as a whole, use application object.

'2. WORKSHEETS PROPERTY
'   (1) Path: 
        Application.defaultFilePath ="c://folder"
'   (2) Display error: 
'       Below example will retult in Excel alert to be surpressed when closing the file.
        Application.DisplayAlerts = False
        ActiveWorkbook.close 
'   (3) Create workbook: 
'       Below example will create 3 new workbooks
        Application.SheetInNewWorkBook = "3"
'   (4) Apply Excel-wise fotn style/size: 
'       Same result can be obtained by going into Preference > General (in MacOS)
        Application.StandardFont = "Arial"
        Application.StandardFontSize = 12

'3. WORKSHEETS METHOD
'   (1) Quit: 
        Application.Quit 
'   (2) Use Excel function: 
'       Syntax: Application.WorksheetFunction."function"
        avg = Application.WorksheetFunction.average(range("A1", Range("A1").End(XlDown)))
        Range("A1").End(XlDown).Offset(1,0).Select
        Selection = Avg
'   (3) SendKeys: 
'       Syntax: Application.SendKeys("keycode", wait)
        Application.SendKeys("^c", true)
'       If wait-->true: wait until key is sent VS wait-->false: do not wait, just do next code.

'       Types of "keycode" (only working in Windows, not in MacOS):
'       Cntrl  --> ^
'       Alt    --> %
'       Shift  --> +
'       F1~F15 --> {F1}~{F15}

'4. EXERCISE: 
'   (1) COUNT ROW NUMBER FOR GIVEN WORKSHEET
'       Scenario: You have a worksheet with long data. 
'                 Instead of scrolling down till the end, you would like to know the row number with one click. 
'       Goal    : Create a button which will tell you the number of rows upon clicking it.

        Sub CountRow_Click()
          Dim cnt As Integer
              cnt = Application.WorksheetFunction.CountA(Range("A2:" & Range("A2").End(xlDown).Address(0#)))
              'cnt = Application.WorksheetFunction.CountA(Range("A2", Range("A2").End(xlDown).Address(0.0)))
              'cnt = Application.WorksheetFunction.CountA(Range("A2", Range("A2").End(xlDown)))
              'Tip1: All three lines are working fine
              'Tip2: Be careful to differenciate Count vs CountA, Count function will return different result
          
           MsgBox "There are " & cnt & " rows in this worksheet"

        End Sub

'   (2) AUTO-SAVE USER DEFINED RANGE IN A SEPARATE EXCEL FILE
'       Scenario: You have worksheet containing raw/big data. You need to copy certain section of this data and paste in notepad.
'                 Imagine you may have to do this task multiple times. Recording steps would reduce time.
'       Goal    : Instead of manually repeating tasks, create a button which will allow you to select range and auto-save files.
'       Tip     : Below has been created for Windows environment. Please refer to exercise005_CopyAndCreateFile.bas for MacOS version.


        Sub CopyPasteCreate_Click()
          Dim rng As Range
          Set rng = Application.InputBox("Select the range", Type:=8)
              rng.Copy

          'Step 1. Open notepad.
          Call Shell("notepad.exe", vbNormalFocus)
          'Step 2. Paste the copied data (Cntl v). "Application" can be omitted from Application.SendKeys 
          SendKeys "^v", True
          'Step 3. Save files
          SendKeys "%fs", True
          'Step 4. Name the file, enter and close.
          SendKeys "sales summar.txt", True
          SendKeys "{enter}", True
          SendKeys "%{f4}", True
          'Step 5. (Optional) Dis-select the area from original file
          Application.CutCopyMode = False

        End Sub

