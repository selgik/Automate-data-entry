'BASIC: SELECT CASE & FOR STATEMENT

'1. SELECT CASE
'   (Q) How can I tell program to do specific action based on certain criteria?
'   (A) Use SELECT CASE 
'       Scenario: Teacher is reviewing students' test restuls (column D). Test results are ranging from 5 to 10. 
'                 Score above 9 is pass, score between 6-8 is conditional pass and below 6 is a fail.
'                 Teacher wants to have a color remark (pass=green, conditional pass=yellow, fail=red)
'                 under the cell and have message box popped up with how many students have failed.
'       Workflow: teacher clicks button -> cells under coulmn D get colored -> msg box pops up


'1) Let's begin with coloring Cell D3.
Sub conditional_coloring()
  Select Case Range("D3") 
    Case Is > 9
     Range("D3").Interior.Color = RGB(0 ,0, 255)
    Case 6 To 9  
      Range("D3").Interior.Color = RGB(250, 200, 30)
    Case Else
    'Same as Case Is < 6
      Range("D3").Interior.Color = RGB(255, 0, 0)
  End Select
End Sub

'2) What we really need to do is, to repeat the process from D3 to D10. We need FOR ~ NEXT statement.

