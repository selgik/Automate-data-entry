'Refer to > note020_Application_Property_&_Method.bas > 4.(2)
'Example on 4.(2) was created for Windows environment with SendKeys, this section has been created for MacOS.

' Problem: You have worksheet containing raw/big data. There are 5+ files you need to make, form this original file.
'          For each new files, you just need to copy and paste certain part such as certain rows, columns or range from original data.
' Goal   : Instead of manually repeating tasks 5+ times, create a button which will allow you to select range and auto-save files.
'          Additionally, let the system auto-assign the file name with number.

  Sub CopyPasteCreateMac_Click()
    Static i As Integer  '<--- Dim instead of Static will give an error before saving 2nd file (duplicate name). 
           i = i + 1     '     Static will allow to retain i value even after procedure is ended.

    Dim rng As Range
    Set rng = Application.InputBox("Select the range", Type:=8)

        rng.Copy
        Workbooks.Add
        Range("A1").PasteSpecial xlPasteAll
        ActiveWorkbook.SaveAs FileName:="Segment_Sales" & i & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
        ActiveWorkbook.Save
        ActiveWorkbook.Close SaveChanges:=True 
        Application.CutCopyMode = False

  End Sub
  
  'Tip: Created files from above will be saved under users/"name"/
