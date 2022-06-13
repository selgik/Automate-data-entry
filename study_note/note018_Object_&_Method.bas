'BASIC: OBJECTS AND METHOD

'1. CLEAR METHOD
'   This method allow to clear (remove) objects
'   (1) Range("A1:A10").Clear --> clear everything including contents, format or memo etc.
'   (2) Range("A1:A10").ClearContents --> only clear contents
'   (3) Range("A1:A10").ClearFormat --> only clear formats
'   (4) Range("A1:A10").ClearComments --> only clear comments


'2. COPY METHOD
'   This method allows to copy and paste object
'   Syntax: [FromWhere].Copy [ToWhere]
    Range("A1").CurrentRegion.Select
    Selection.Copy Sheets("Sheet2").Range("B1")

  
'3. PASTESPECIAL METHOD
'   Syntax: [TargetLocation].PasteSpecial [Option]
'   (1) Option1: xlPasteAll --> paste all
'   (2) Option2: xlPasteFormats --> paste formats only
'   (3) Option3: xlPastealues --> paste values only
'   (4) Option4: xlPasteSpecialOperationAdd --> add up and paste
'   (5) Option5: xlPasteSpecialOperationMultiply --> multipy and paste
'   (6) Option6: Transpose:=true --> paste by swapping column with row
        Range("A1:A10").Copy
        Range("C1").PasteSpecial Transpose:=True

  
'4. ROW, COLUMN / ROWS, COLUMNS
'   (1) Row, Column 
'       This method will return the number of row or column.
        MsgBox Range("B2").Row 
       'Messagebox will show 2
        MsgBox Range("F5").Column 
       'Messagebox will show 6
  
'   (2) Rows, Columns
'       This method will return the value corresponding the last cell of rows/columns
'       Ex. From B3 to B10 we have alphabet listed from A. 
        MsgBox Range("B4").End(xlDown).Rows
        'Messagebox will show H

'   (3) Rows, Columns with =
        Range("A1:F5").Rows = "1234"
        'From A1 to F5, 1234 will be inserted. Same for Range("A1:F5").Columns = "1234"
        Range("B2:F5").Rows(1) = "Hello"
        'Starting from the 1st row of the range (which will be 2) B2 to F2, "Hello" will be inserted.
        Range("B2:F5").EntireRow = "Surprise!"
        'Entire row 2 to 5 will be inserted with "Surprise!".


'5. COUNT
    Range("A1:B7").Select
    Selection.Rows.Count
    'or, Selection.Columns.Count


'6. EXERCISE: CREATE AN ENTRY FORM TO ADD DATA TO THE EXISTING TABLE
'   Scenario: You have a table with code name and price list, from B4 to C13.
'             You have an entry form. User can write code name in C2 and price in E2.
'   Goal    : Upon clicking a button, transfer entered data(C2 and E2) to the next line of the table (B14:C14).
'             Each time user enters data and clicks a button, those will be transffered to next line (then B15:C15). 
'          
 
Sub EntryForm_Click()
  Dim r As Integer
      r = Range("B4").Row + Range("B4").CurrentRegion.Rows.Count
  
  Cells(r, 2) = Cells(2,3).Value
  Cells(r, 3) = Val(Cells(2, 5).Value)
               'Val() will allow contents in the cells(2,5) to be transformed into number format.
  Cells(2, 3).ClearContents
  Cells(2, 5).ClearContents

End Sub

