'BASIC: ARRAY

'1. WHAT IS AN ARRAY?
'   Array is a variable holding a set of values within the same data type. Used for memory efficienty.
'   Example: [1, 2, 5, 14, 36] or [apple, orange, melon, mango]
'   Syntax : Dim a(number) As data_type
'   Remark : Array will start from 0. To start from 1, declaration is needed: Option Base 1
'   Comparison: 

'   Without Array, codes will look like below:
    Private Sub Worksheet_Activate() 
      Dim a, b, c As Integer
          a = 14
          b = 24
          c = 42
      Range("A5").Value = a
      Range("A6").Value = b
      Range("A7").Value = c
    End Sub
  
'   With Array, codes will look like below:  
    Option Base 1
    Private Sub Worksheet_Activate() 
      Dim a(3), i, rownb As Integer 'Same as: Dim a(0 To 2) As Integer
          a(1) = 14
          a(2) = 24
          a(3) = 42
          rownb = 5
      For i = 1 To 3
        Cells(rownb, 1).Value = a(i)
        rownb = rownb + 1
      Next i
    End Sub

'Comment: With 3 items as above, benefit of using array may not seem obvious. 
'         But if we have more items? If we have uncertainty with the size of the sets? Array could be the only option.


'2. WHAT IS AN ARRAY AS FUNCTION?
'   With Array function, codes can be even shorter and clearer.
'   Comparison:

'   Without Array functions:
    Option Base 1
    Private Sub Worksheet_Activate() 
      Dim a(4) As String 
        Range("A1").Value = "Apple"
        Range("A2").Value = "Orange"
        Range("A3").Value = "Melon"
        Range("A4").Value = "Mango"
    End Sub

'   With Array function:
    Option Base 1
    Private Sub Worksheet_Activate() 
      Dim arr As Variant 'Variant: not defining data type
         arr = Array("Apple", "Orange", "Melong" , "Mango")
         Range("A1").Resize(1, 4).Value = arr
    End Sub

'Comment: But above code will print values in A1, B1, C1 and D1. 
'         To print in from A1 to A4, codes should be re-written as below.
'----->>> 

    Option Base 1
    Private Sub Worksheet_Activate() 
      Dim i As Integer
      Dim arr As Variant 
          arr = Array("Apple", "Orange", "Melong" , "Mango")
        For i = 1 To 4
            Cells(i, 1).Value = arr(i)
        Next i
     End Sub


'3. WHAT IS A DYNAMIC ARRAY?
'   You do not know how big the array will be. Therefore, you are telling it later. 
'   Syntax:
    Dim Arr() As Integer
    Redim Arr(10) 


'4. LBOUND / UBOUND
'   Syntax : Lbound(array_name, dimension_number)
'   Comment: These functions calculates the first / last number of the array.
'   Example: Latest codes from 2nd point can be re-written as below:

    Option Base 1
    Private Sub Worksheet_Activate() 
      Dim i As Integer
      Dim arr As Variant 
          arr = Array("Apple", "Orange", "Melong" , "Mango")
        For i = 1 To Ubound(arr, 1) '1 can be omitted. ==> For i = 1 To Ubound(arr)
            Cells(i, 1).Value = arr(i)
        Next i
     End Sub

