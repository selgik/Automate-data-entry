'Problem: You have worksheet with 2,000 rows. There are empty rows in the worksheet and you need to remove them.
'Solution: run VBA code with loop to find the empty cell and remove corresponding rows.

Sub delete_rows()
  Dim i As Integer
  For i = 2000 To 1 Step -1
      If Range("A" & i) = "" Then
         Rows(i).EntireRow.Delete
      End If
  Next 
End Sub

'Tip: For adding/removing task, it is recommended to start loop backward. 
'Failure Note: IsEmpty did not work out --> If IsEmpty(Range("B" & i)) Then Rows([i]).EntireRow.Delete
