'BASIC: FOR ... NEXT (LOOP PART 1)

'1. FOR NEXT
'   (Q) I want Excel to repeat certain task on the cells. (ex. filling in with texts or numbers)
'   (A) Use FOR NEXT 
'       Workflow: User clicks button -> program will auto-fill in the text/numbers on defined range.

'1) Let's start with repeating message boxes.
Sub Repeat_Msgbox()
  Dim i As Integer
  For i = 1 To 5 
  'same as For i = 1 to 5 step 1
      MsgBox "What a Wonderful Trick!"
  Next i
End Sub

'Similarly, we can fill-in the cells A1 to A5 with the texts.
Sub Repeat_Input()
  Dim i As Integer
  For i = 1 to 5
      Cells(i, 1).Value = "This is a test"
  Next i
End Sub


'2) Now, let's fill in the cells A1 to A5 numbering from 1 to 5.
Sub Repeat_Numb()
  Dim i As Integer
  For i = 1 to 5
      Cells(i, 1) = i
  Next i
End Sub

'What if you need to fill in from B3 to B7 instead? 
Sub Repeat_Numb2()
  Dim i, start As Integer
  start = 2
  For i = 1 to 5
      Cells(start + i, 2) = i
  Next i
End Sub

'Instead, you could also write as below:
'Sub Repeat_Numb2()
'  Dim i As Integer
'  For i = 1 to 5
'      Cells(i + 2, 2) = i
'  Next i
'End Sub


'3) Next, imagine user wants to fill in the number starting from 260. 
'InputBox can be made for user to decide from which number they want to start filling-in (and end too.)
'Imagine user needs to fill in from B3, which is Cells(3,1).

Sub Repeat_Ask()
  Dim i, start, a, z, incr As Integer
  start = 2
  a = InputBox("Enter starting number")
  z = InputBox("Enter ending number")
  For i = a To z
      incr = incr + 1
      Cells(start + incr, 2) = i
  Next i
End Sub


'4) Finally, imagine user wants to fill-in with dates next to the column B from 3rd scenario.
'Example: From B3 to B130, we have certain numbers filled in per exercise 3. 
'         User now wants to fill in with the current date from C3 to C130

Sub Fill_Date()
  Dim i, rowcount As Integer
  rowcount = Range("B3").End(xlDown).Row
  'Count the number of rows starting from B3 till the data ends
  For i = 3 to rowcount
    Cells(i, 3) = Date
  Next
End Sub

