'Problem: You have a list of motivational quotes, daily challenge idea and fortune cookies (100 items each). 
'         You want a message box to pop up to show a random quote. How to do that?
'Solution: create a blank page, build 3 buttons for each section linked with VBA


'[1] 1st Attempt:

Sub failed_quote_generator()
    MsgBox "=INDEX(B1:B100,RANDBETWEEN(1,ROWS(B1:B100)),1)"
End Sub  

'Result: When I clicked a button, msgbox showed a code chunk, not a result of the code. 
'        I realized, it might be because I created a button in a different sheet than the sheet where quotes sit in.
'        Hence, I decided to generate a quote in random cell and have msgbox read out that cell's value.


'[2] 2nd Attemtp:

Sub test()
  Dim rng As Range
  Set rng = Range("A1")
      rng.Value = "=sum(B1:B4)"
      MsgBox rng
      Range("A1").ClearContents
End Sub
  
'Result: I created simplified msgbox which can read the calculated A1 value out. It worked out.
'        I now, will try to edit above to read cell's value outside the current worksheet.
  
  
'[3] 3rd Attempt:
Sub random_quote_generator()
  Dim rng As Range
  Set rng = Worksheets("Motivational_Q").Range("A2")
      rng.Value = "=INDEX(B1:B100,RANDBETWEEN(1,ROWS(B1:B100)),1)"
      MsgBox rng
      Worksheets("Motivational_Q").Range("A2").ClearContents
End Sub

  
'Result: (1) Pick a random quote in A2 
'        (2) Use message box to pop up the quote 
'        (3) Clear contents in A2 after that. 
  
'Tip: If you have daily challenges ideas and fortuen cookies quotes saved in different sheets, 
'     You would have to create two more buttons and similar codes as above. 
  
'Failure Note: I tried below but I received an error. Besides, code is too long and misleading.
    
Sub failed_code_generator2()
    MsgBox "=INDEX(Worksheets("Motivational_Q").Range("B1:B100"),RANDBETWEEN(1,ROWS(Worksheets("Motivational_Q").Range ("B1:B100"))),1)"
End Sub
    
    
