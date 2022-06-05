'Problem: You have 4 type of ball of red, yellow, green and blue. Each ball weights as follows,

'         Red   : 60g 
'         Yellow: 55g
'         Green : 75g
'         Blue  : 80g

'         You are selling balls to customers per designated weight. You canot however change the sequence of ball being sold.
'         That means, if customer X wanted to buy 130g of ball, you should cell red + yellow (together weighting 115g)
'         And you cannot add green ball as it will make the weight 115+75=190g in total. 

'         Make a program that will calculate the sequence number. 
'         In above example of 130g, sequence is 2 (red as 1, yellow as 2).

'Solution: Use Do while + For codes.

Sub calculator_click()
  Dim i, t, count, sum As Integer
  Dim arr As Variant
  
  arr = Array(60, 55, 75, 80)
    t =  InputBox("Enter the target weight")
  
    i = 0
  sum = 0
count = 0

  Do While sum < t
    For i = 1 To 4
       If sum < t Then
          sum = sum + arr(i)
          count = count + 1
        Else: Exit For
       End If
    Next i
  Loop

MsgBox count-1

End Sub
