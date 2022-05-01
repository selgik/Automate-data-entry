'BASIC: INTRO TO VARIABLES

'1. SYNTAX
'   (Q) How to declare variable in VBA?
'   (A) Dim variable_name as variable_type (ex) Dim name As String, Dim ws_1 As Worksheet

'2. EXERCISE
'   (Q) Create a button which will ask user to input number for calculation.  
'   (A) Pre-requisite: B2="name" | C2="purchase_price"   
'                      B3="Adam" | C3="200,000,000  

'       Developer workflow: create button -> write module -> test
'            user workflow: user enter name on B3, price on C3 -> user clicks button 
'                           -> system asks user to enter number in the inputbox -> user enters number
'                           -> system will calculate (ex. add, substract, multiply, divide etc.)
'                           -> system will show calculated result on the message box 

Sub calculator()

'declare variables
Dim name As String
Dim purchase_price As Currency
Dim fee_rate As Double

'declare worksheet object. it will be useful if we have multiple worksheets
Dim ws As Worksheet
Set ws = Worksheets("purchase_cal")

'tell system what are the variables and what to do with variables
name = ws.Range("B3")
purchase_price = ws.Range("C3")
fee_rate = InputBox("Enter the fee ratio and we will calculate for you!")
MsgBox name & "s fee rate is = " & purchase_price * fee_rate

End Sub
