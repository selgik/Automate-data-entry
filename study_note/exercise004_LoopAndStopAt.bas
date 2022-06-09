'Same scenario as exercise 003 but let's think about more useful solution.

'Problem: I have target grams to sell. (Example: 831g)
'         I want to know how many sets of red/yellow/green/blue balls I have to sell (loop) and,
'         I want to know the color of last ball I am selling (sequence).

'Solution: Use Do while + For codes.

Sub calculator_Click()
  Dim i, t, count, order, sum As Integer
  Dim arr As Variant
  
  arr = Array(60, 55, 75, 80)
    t = InputBox("Enter the target weight")
    i = 0
  sum = 0
count = 0

Do While sum < t    
     For i = 1 To 4
      If sum < t Then
         sum = sum + arr(i)
      Else: Exit For
      End If  
     Next i
  count = count + 1
Loop
    
 If i = 2 Then
    order = 4
 Else
    order = i - 2
 End If
 
MsgBox "Sell a set of " & (count - 1) & " ball(s). Last ball you are selling is " & order & "th one."

End Sub
