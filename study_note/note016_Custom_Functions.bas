'BASIC: CUSTOM FUNCTIONS

'1. WHAT IS A CUSTOM FUNCTION?
'   Custom function is a user defined function that can be created/saved in VBA and used in Excel. 
'   Imagine a long and complicated function in Excel. Instead of writing long argument, use custom function.
'   It will be easier to manage (ex. edit) function and minimize the errors.

'2. WHAT IS A WORK FLOW OF CUSTOM FUNCTION?
'   (1) Create custom function in VBA: Go to Insert > Module 
'   (2) Use function in Excel. There is no need for a button. Use it as other Excel functions: =custom_f(A1, A2)

'3. WHAT IS THE SYNTAX?
    Public Function function_name(argument1, argument2...)
    'Use Private Function if the function is sheet specific. Public will allow function to be used in all sheets. 
        function_name = "define function here"
    End Function
  
  
'4. EXAMPLE-1
  ' Scenario: Imagine you have a table containing employee's name, department, dependent number and salary. 
  '           As a HR/Finance officer, you need to calculate salary tax. Tax rate depends on the number of dependent. 
  '           Example: If John has 4 dependents, his tax rate equals to 6%. For 1-3 dependents, 9% and if none, then 12%.
  
  '(1) Let's create custom function
  
  Public Function tax_cal(dependent As Integer, salary As Currency)
    Dim tax As Currency
    
    Select Case dependent
    Case Is >= 4
      tax = salary * 0.06
    Case 1 To 3
      tax = salary * 0.09
    Case Else
      tax = salary * 0.12
    End Select
    
    tax_calc = tax
  End Function
  
  '(2) Let's use function in Excel
  '    In the table, dependent is on columm C and salary is on column H. You'd like to calculate tax on coulmn I.
  '    In I2, write below function:
  
  =tax_cal(C2, H2)
  
  
  '5. EXAMPLE-2
  ' Scenario: Imagine you have to transform date into day of the week. 
  '           You can use =WEEKDAY function but it will return the number (ex. 1 instead of Sunday, 2 instead of Monday etc.)
  '           You want to use certain function that will convert directly into day of week.
  
  '(1) Let's create custom function
  
  Option Base 1
  'Remember to add above. Otherwise, array will start from 0 while weekday function start from 1.
  
  Public Function dow_conv(raw_d As Date)
    Dim arr As Variant
    Dim i As Integer
    
      arr = Array("Sun", "Mon", "Tue", "Wed", "Thr", "Fri", "Sat")
        i = Weekday(raw_d)
 dow_conv = arr(i)
    
  End Function
  
  '(2) Let's use function in Excel
  '    In the table, date is on columm A. You'd like to obtain day of week on column B.
  '    In B2, write below function:
  
  = dow_conv(A2)

  
