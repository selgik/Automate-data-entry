'BASIC: VBA NUMERIC AND DATE FUNCTIONS

'1. NUMERIC FUNCTIONS
'   Assume a = 2.6443

'   (1) Int(a) returns integer. Result will be 2.
'   (2) Round(a) returns rounded value. Result will be 3. 
'   (3) Abs(a) returns absolute value. 
'   (4) Isnumeric(a) validates whether the value is numeric or not. If numeric, result will show TRUE. Otherwise, FALSE.
'   (5) Val(string) transfoms string value to numeric.
'   (6) Rnd(a) returns random value. See example below.
'       Scenario: With one click, you would like to generate random number from cell A1 to A20 (within the ranage of 1 to 100)

Sub rnd_test_type1()
    Randomize
    'Why Randomize is needed? See this blog:
    'https://www.techonthenet.com/excel/formulas/randomize.php#:~:text=The%20Randomize%20function%20would%20ensure,equivalent%20to%20the%20system%20timer.
  
    Dim i As Integer
    For i = 1 to 50
        Cells(i, 1).Value = Int(Rnd * 100) + 1
        '*100 means, till 100 (number range between 1 to 100) 
        'Without Int(), function will return double value.
    Next
End Sub

'Same code, but used with Range
Sub rnd_test_type2()
Dim rng As Range  
    For Each rng In Range("A1:A20")
        rng.Value = Int(Rnd * 100) + 1  
    Next
End Sub


'2. DATE FUNCTIONS
'   (1) Time: returns current time.
'   (2) Hour(time), Minute(time): returns each particular. 
'   (3) Date: returns today's date.
'   (4) Year(date), Month(date), Day(date): returns each particular.
'   (5) Weekday(date): returns the sequence of weekday. Ex) 1, 2, ...7 reffering to Sun, Mon, ...Sat.
'   (6) IsDate(value): validates whether the value is date or not. If date, result will show TRUE. Otherwise, FALSE.
'   (7) DateAdd("d, m", number, date): calculates date after number of days/months.
'   (8) DateDiff("d, m", start_date, ref_date): calculates number of days/months between two dates. 
'   (9) DatePart("d, m, q", date): returns value for particular. Ex. DatePart("q", Date) returns 2 where Date is 28/5/22.
'   (10) DateSerial allows user to modify date information. 
'        Scenario: Let's calcualte the last day of the month.

Sub cal_last_day()
Dim d as Date
    MsgBox "The last day of this month is " & DateSerial(Year(d), Month(d)+1, 0)
    'When you put DateSerial(Year(d), Month(d), Day(d)), result will be today's date.
    'DateSerial(Year(d), Month(d)+1, 0) <- There is no 0st day of next month. It will show the last day of current month.
End Sub


'3. REAL LIFE SCENARIO
'Scenario: Imagine you have requested various departments to send out next quarter's budget. Column C is where budget amount is inserted.
'          Unfortunately you have found out some errors. Some cells are empty, other cells shoinwg error message such as "REF!"
'Goal    : With one click, highlight cells and change its value to "0", where value is non-numeric or empty.

Sub cleaning()
Dim rng As Range
  For Each rng In Range("C1:C100")
      If IsEmpty(rng) or IsNumeric(rng) = False Then
         rng.Value = 0
         rng.Interior.Color = RGB(250, 100, 100)
      End if
  Next
End Sub

'Fail note for my reference: I wrote as below in my first attempt. It worked out but code is longer. 

Sub cleaning2()
Dim rng As Range
  For Each rng In Range("C1:C100")
      If IsEmpty(rng) = True Then
         rng.Value = 0
         rng.Interior.Color = RGB(250, 100, 100)
      ElseIf IsNumeric(rng) = False Then
         rng.Value = 0
         rng.Interior.Color = RGB(250, 100, 100)
      End If
  Next
End Sub

