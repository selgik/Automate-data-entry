'BASIC: DO WHILE, DO UNTIL(LOOP PART 3)

'1. DO WHILE
'   (Q) Instead of telling when to start and end the loop, I want to tell the condition
'   (A) Write with Do While or Do Until. For example, For i = 1 To 5 can be re-written as:
        Do While i <= 5 
        Do Until i > 5
    
'1) Simple sum calculation from 1 to 10, can be written with For/Next as well as Do While/Loop:

'For/Next:
Sub for_next_ex()
  Dim i, sum As Integer
    sum = 0
      For i = 1 To 10
       sum = sum + i
      next i
  MsgBox "The result is " & sum
End Sub

'Do While/Loop:
Sub do_while_loop_ex()
  Dim ii, sum2 As Integer
    ii = 1
    sum2 = 0
      Do While ii <= 10
        sume2 = sum2 + ii
        ii = ii + 1
      Loop
  MsgBox "The result is " & sum
End Sub 


'2) Let's apply Do While logic. Think about the task where user has to enter 15 sales code.
'   It will be easier if Input box appears and all user has to do is typing name - enter.
  
  Sub example()
    Dim i As Integer
    Dim code As String
      i = 1
        Do While i <= 15
          code = InputBox("Enter the sales code")
           Cells(i, 1) = code
          i = i +1
        Loop
  End Sub
  
        
'3) TIP: Understanding Loop
' What if I change the order from 2nd example?
        
  Sub example()
    Dim i As Integer
    Dim code As String
      i = 1
      code = InputBox("Enter the sales code")
      'If I put this part here, this won't be the part of the loop. Inputbox will only appear once.
      'That means, if I put "A123", loop will repeat the tast by filling cells till 15th row with "A123"
        Do While i <= 15
           Cells(i, 1) = code
          i = i +1
        Loop
  End Sub
        
        
