'BASIC: FOR EACH (LOOP PART 4)

'1. FOR EACH
'   (Q) I would like to use loop reffering to the object (such as worksheet or range.)
'   (A) Use For Each In syntax

'1) Let's say user has 10 worksheets and wants to have worksheet's name pasted by clicking one button.

Sub ws_name()
  Dim ws As Worksheet
  Dim i As Integer
  i = 2
    For Each ws In Worksheets
      Cells(i, 2) = ws.Name 
        i = i + 1
    Next ws
End Sub


'2) Now, let's build another button where by clicking it, user can search and go to the entered worksheet.

Sub ws_search()
Dim ws As Worksheet
Dim keywd As String
keywd = InputBox("Enter the worksheet's name")
'keywd will be case-sensitive
 
  For Each ws In Worksheets
     If keywd = ws.Name Then
       Sheets(keywd).Activate
       End For
       'If this line comes after End if, then, no condition met => End For. It does not reach to Next ws.
       'No loop is created.
     End If  
  Next ws
End Sub


'3) Imagine user has a list of student ID (A2:A20) with their latest test score in column (B2:B20)
'   While the list may grow up, user wants to selectively sum up the scores which exceeded the passing score (ex. 80) 

Sub sum_up()
Dim rng As Range
Dim total, i As Integer

For Each rng In Range("B2:B20")
  If rng.Value >= 80 Then
    total = total + rng.Value
  End If
Next

i = Range("B2").End(xlDown).Row
'Same as Range("B21").Value = total
Cells(i + 1, 2).Value = total

End Sub

