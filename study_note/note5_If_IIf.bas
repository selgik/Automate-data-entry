'BASIC: IF VS IIF

'1. IF, ELSEIF, ELSE
'   (Q) How can I build the program to update user with different answers based on the criteria?
'   (A) Build codes with IF. 
'       Workflow: user input data "x" -> 
'                 IF x meets condition 1 then output 1, else -> 
'                 ELSEIF x meets condition 2 then output 2 -> ELSE output 3 -> END IF
'                 (ELSEIF is optional. In such case, flow is IF -> ELSE -> END IF)

'Ex 1) User enter number -> system tells whether the number is odd or even number.
Sub is_it_even_number_or_odd_number()
  Dim i As Integer
  i = InputBox("Enter the number")
  
  If i < 0 Then
    MsgBox "Enter positive number"
  Elseif i Mod 2 <> 0 Then
    MsgBox i & " is odd number"
  Else 
    MsgBox i & " is even number"
  End If 
End Sub

'Ex 2) User clicks button -> system will ask whether to open Sheet#2
Sub wsheet_index()
  If MsgBox("Go to Sheet 2?", vbYesNo) = vbYes Then
     Sheet("Sheet_2").Activate
  Else  
    Exit Sub
  End If
End Sub
  
  
'2. IIF
'   (Q) What is the difference between IF vs IFF
'   (A) Different syntax and performance (to be slower than IF). Similar to Excel's IF function.
'       Syntax: IFF(condition, true value, falsevalue)
  
Sub is_it_even_number_or_odd_number_v2()
  Dim i As Integer
  Dim rs As String
    
  i = InputBox("Enter the number")
  rs = IIf(i Mod 2 <>0, "odd number", "even number")
    MsgBox i & " is " & rs & "!" 
End Sub
  
