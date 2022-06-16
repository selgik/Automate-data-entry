'BASIC: WORKSHEETS PROPERTY AND METHOD

'1. WHEN TO USE? 
'   If you need to manage (create, move etc) worksheets regularly, use VBA codes to autoamte the process.
'   Example: You are cleaning data and you want to transform data and auto-organize in new worksheets.

'2. WORKSHEETS PROPERTY
'   (1) ActiveSheet
'       ActiveSheet.Name = "Sales" --> assign name to the sheet
'       ActiveSheet.Range("A1:A10").Select --> select range from active(current) sheet.

'   (2) Worksheets
'       Worksheets("Sales").select/activate
'       Worksheets(2).select/activate --> 2 is worksheet number. 
'       Example:
        Worksheet(2).activate
        ActiveSheet.name = "Summary"

'3. WORKSHEETS METHOD
'   (1) Worksheets.Add --> Add new worksheet before the current sheet
'   (2) Worksheets.Add Before/After:= --> Add worksheet before or after the sheet
'       Example:
        Worksheets.Add After:=Worksheets("Sales")
        ActiveSheet.Name = "Profit"
'   (3) Worksheets.Add Count:= --> Add multiple worksheets
'       Example: 
        Worksheets.Add After:=Worksheets("Sales"), Count:=2
'   (4) Worksheets.Count --> Count the number of worksheets
'   (5) Worksheets.Move Before,After:=
'       Example:
        Worksheets("Sales").Move After:= Worksheets("Profit")
'   (6) Worksheets.Copy Before,After:=
'   (7) Worksheets.visible = xlSheetHidden (=false), xlSheetVisible (=true), xlSheetVeryHidden
        Worksheets("Sales").visible = false

'4. EXERCISE: 
'   (1) CREATE PROCEDURE WHERE BY CLICKING A BUTTON, NEW WORKSHEET WITH CERTAIN FORMAT WILL BE GENERATED.
'       Scenario: You are managing sales data in Excel. Everyday, you have to create worksheets with such format below:
'                 Sales-BcDo1, Sales-BcDo2, Sales-BcDo3..., etc.
'       Goal    : Upon clicking a button, you will have sheet created with the file name as above, auto-increased on the last digit. 

        Sub create_wksh()
          Static i As Integer
          'Dim i As Integer --> won't work as system will give an error message.
          'This is because procedure will end and i value will not be stored with DIM.
          'So if you click button again, system will try to create duplicate sheet name using the same i. Static must be hence used.
          i = i + 1
          Worksheets.Add
          ActiveSheet.Name = "Sales-BcDo" & i
        End Sub

'   (2) CREATE PROCEDURE WHERE BY CLICKING A BUTTON, USER WILL BE ASKED TO TYPE FROM/TO WORKSHEET TO MOVE AROUND.
'       Scenario: You have 10+ worksheets and you are constantly moving those worksheets.
'                 While number of worksheets are increasing, it is inefficient to find and move around those manually.
'       Goal    : Upon clicking a button, you will be asekd to type two sheet names. (From_sheet and To_sheet names)
'                 From_sheet will be moved after To_sheet. 

        Sub move_wksh()
          Dim fromsh, tosh As String
              fromsh = InputBox("Enter the sheet name you would like to move")
              tosh   = InputBox("Where do you want to place after?")
          Worksheets(fromsh).Move After:=Worksheets(tosh)
        End Sub

'        TIP!! What if you have mis-spelled the worksheet name? You will receive an error box.
'              For end-user, error box will be scary as it will not tell what went wrong.
'              So let's turn error message into more user-friendly notice

        Sub move_wksh()
          Dim fromsh, tosh As String
          On Error GoTo Err_handle
              fromsh = InputBox("Enter the sheet name you would like to move")
              tosh   = InputBox("Where do you want to place after?")
          Worksheets(fromsh).Move After:=Worksheets(tosh)
        Exit Sub
        'Why this one has been added?
        'Without it, system will show Err_handle msgbox even if user entered correct worksheet name
        'To stop system from showing error msgbox incorrectly, make system STOP if everything worked out.
          Err_handle:
              MsgBox "Please enter correct worksheet name"
        End Sub

'       TIP!! If want the error message to not show up, use below instead.
'             On Error Resume Next
