'BASIC: INTRO TO VBA CONCEPTS

'1. DEFINITION
'   (Q1) What is object, property and method?
'   (A1) Kiki is beautiful cat. Kiki can meow. --> "Kiki" is object, "beautiful" is property and "can meow" is method.
'        In VBA, Range("A1:A5").Select or Range("A1:A5").Font.Bold=True --> range is object, select is method, bold is property.
'
'   (Q2) What is code, procedure, module and project?
'   (A2) Hirarchy: code < procedure < module < project
'        - Code: object, property and method makes code. Do something or make effects to something.
'        - Procedure: codes make procedure. It is a step-by-step, a complete set of action. Procedure creates event (result).
'        - Module: Procedures make modules. It is a group/set of procedures.
'        - Project: modules make project. It is group of modules within the worksheet.

'2. CREATE EVENT UPON USER ACTION
'   (Q1) When I click the worksheet, I want certain cell/range to be highlighted.
'   (A1) Use Worksheet_Activate() --> Need to insert below code to the target sheet. 

Private Sub Worksheet_Activate()
  Range("A1:E5").Select
    Selection.Interior.Color = RGB(200, 200, 100)
    Selection.Font.Color = RGB(100, 100, 200)
    'The more the number is close to 0, it gets darker
End Sub
  
'   (Q2) When I click the worksheet, I want certain msg to be printed in targeted cell(s)
'   (A2) Change to Selection.Value = "msg"
  
Private Sub Worksheet_Activate()
    Range("F1:F5").Select
     Selection.Value = "Happy Day!"
     Selection.Font.Name = "Gothic"
     Rows("1:5").RowHeight = 50
     'this will enlarge the row size. Optional to add.
End Sub
    
'   (Q3) When I click the on cell (move mouse), I want certain msg to be printed out.
'   (A3) Change to Worksheet_SelectionChange(ByVal Taget As Range)

Private Sub Worksheet_SelectionChange(ByVal Taget As Range)
  Selection.Value = Date
  'In this case, I am printing out today's date.
End Sub

    
'3. HOW TO ORGANIZE PROCEDURE AND MODULE
'   (Q1) How to create procuedure an module?
'   (A1) Developer > VB > Insert procedure 
'        Developer > VB > Insert module > Sub > Public (if you need to use for given project) or Private 

'   (Q2) How to organize procedures and modules?
'   (A2) In the module, write codes (ex. Action()) -> create procedure in the targetted worksheet -> call procedure Action

'4. USE MODULE'S PROCEDURE
'Case1: Workfolw 1. create module: show message box ex. "today is 01/01/2022 yay!"
Public Sub msg_info()
      MsgBox "today is " & Date & "yay!"
      'error msg will apper if & is not inserted, due to string + function being mixed up.
End Sub

'Case1: Workfolw 2. in the targetted worksheet, create procedure and insert module from workflow 1.
Private Sub worksheet_selectionchange(ByVal target As Range)
      msg_info
End Sub

'Case2: Workflow 1. create module: user types their name. Message will show up as "Hello Sylvia you will be redired"    
Sub authorize_msg()
    Dim name As String
    name = InputBox("Enter you name", "Authorize")
    MsgBox "Hello " + name + ", you will be redirected!"
End Sub

'Case2: Workflow 2. create button to order series of action
'       User clicks button --> user types name --> message box will appear --> user is redirected to result worksheet.            

Sub Button2_Click()
 authorize_msg
 Sheets("result").Visible = True
 'True/False usually appears when property is being defined.
 Sheets("result").Activate
End Sub
    
    
