'BASIC: INTRO TO VBA CONCEPTS

'1. DEFINITION
'   (Q1) What is object, property and method?
'   (A1) Kiki is beautiful cat. Kiki can meow. --> "Kiki" is object, "beautiful" is property and "can meow" is method.
'        In VBA, Range("A1:A5").Select or Range("A1:A5").Font.Bold=True --> range is object, select is method, bold is property.
'
'   (Q2) What is code, procedure, module and project?
'   (A2) Hirarchy: code < procedure < module < project
'        - Code: object, property and method makes code. Do something or make effects to something.
'        - Procedure: codes make procedure. It is a step-by-step, complete set of action. Procedure creates event (result).
'        - Module: Procedures make modules. It is a group/set of procedures.
'        - Project: modules make project. It is group of modules within the worksheet.

'2. CREATE EVENT UPON USER ACTION
'   (Q1) When I click the worksheet, I want certain cell/range to be highlighted.
'   (A1) Use Worksheet_Activate() --> Need to insert below code to the target sheet. 

'Below is a complete procedure
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

    
