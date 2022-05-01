'BASIC: ABSOLUTE VS RELATIVE REFERENCE 

'1. ABSOLUTE VS RELATIVE
'   (Q) What is the difference between absolute vs relative reference? And when to use?
'   (A) By using absolute reference, you are using exact cell location. Ex. B4
'       By using relative reference, you are using direction based on your current cell. Ex. +2 rows and -2 columns from current cell.

'2. CALCULATE SUM OF TWO CELLS USING ABSOLUTE VS RELATIVE REFERENCE
'   Scenario: calculate sum of two cells and fill down the value
'   Workflow: user click button -> system calculates sum -> sum will display in Cell A4

'1) Use absolute reference
Sub absolute_calculation()
  Dim rng as Range
  'Why Set? Because we are dealing with an object variable
  Set rng = Range("D2:D10")
    rng.Value = "=sum(A2:A3)"
    rng.FillDown
End Sub
  
'2) Use relative reference
Sub relative_calculation()
  Dim rng as Range
  Set rng = Range("D2:D10")
    rng.FormulaR1C1 = "=sum(RC[-2]:RC[-1])"
    rng.FillDown
End Sub  
  
'For above example, RC[-2]:RC[-1] is used because row number didn't change. What if row number need to change as well?
    
    Sub test()
      ActiveCell.FormulaR1C1="=sum(R[-5]C[0]:R[-5]C[1])"
      'This is same as ActiveCell.Value="=sum(B3:C3)" in absoulte reference.   
    End Sub
    
   
