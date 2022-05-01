'BASIC: NESTED FOR (LOOP PART 2)

'1. FOR I, FOR J
'   (Q) Imagine user needs to do Job A, Job B and Job C.
'       For each Job, there are Tasks that need to be done, before moving to the next Job.

'     [A]  -->   [B]   --> [C]
'    Task 1    Task 1'   Task 1''
'    Task 2    Task 2'   Task 2''
'    Task 3    Task 3'   Task 3''
'    Task 4    Task 4'   Task 4''
'     ...       ...       ...

'   (A) Use nested for. Outer FOR for Job A, B and C and inner FOR for Task 1, Task 2, Task 3 etc. 


'1) Let's create a procedure for defined range to be filled in with numbers and colors with one click.

Sub Nested_ex()
  Dim i, j As Integer
  For i = 3 to 10
    For j = 3 to 10
      Cells(i, j).Value = i * j
      Cells(i, j).Interior.Color = RGB(100 + i * 10, 150 + i * 15, 200 + i * 20)
    Next j
  Next i
End Sub


'2) Imagine user has sales data from B4 to G30. With one click, user wants to see the cell that exceeded target number.

Sub Target_check()
 Dim i, j As Integer
  For i = 4 To 30
    For j = 2 To 7
       If Cells(i, j) >= 2500 Then
          Cells(i, j).Interior.Color = RGB(20, 200, 40)
      End If
      'Otherwise, end if (no need else)
    Next j
  Next i
End Sub


