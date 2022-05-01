'BASIC: INPUTBOX AND MSGBOX

'1. INPUT DATA AND OUTPUT RESULT
'   (Q) How can I build input and output message box which will run certain calculation?
'   (A) Create button to input number -> message box to output result.

Sub GST_calculator()
  Dim total as double
  Dim gst as double
  
  total = InputBox("Enter the total amount", "Calculator")
  'InputBox("Enter the total amount", "Calculator", 100) <-- 100 will pre-appear in the inputbox.
  gst = total * 0.07
  
  MsgBox "The GST is " & gst
End Sub

'If I want to show result in certain cell instead of MsgBox, I can write as:
  Range("A1").Select
  Selection.Value = gst


'2. MSGBOX OPTION: CALL USER ACTION WITH IF
'   (Q) Can I have MsgBox carry certain task depending on user's answer?
'   (A) Yes. There are options to change alert icon or user's answers. For latter one, usually If function comes together.
'       To do so, use MsgBox("your message", <-- as soon as comma is entered, option will appear
'       Ex. MsgBox("Finish this project?", vbYesNo)
  
Sub button1()
  If MsgBox("Finish this project?", vbYesNo) = vbYes Then
  Exit Sub
  End if
End Sub

    
'3. INPUTBOX OPTION: USE AS METHOD FOR AN APPLICATION OBJECT
'   (Q) Can I drag cells for system to auto-calculate the sum?
'   (A) Use Application.InputBox to enable user to select/drag cells.
'       Workflow: user clicks button -> system asks user to drag cells -> sum of selected fields to be shown in msgbox.
'       Type:8 is select/drag type in application.inputbox
'       VBA function is different than Excel function. WorksheetFunction will enable to use Excel function.  

Sub button2()    
  Dim rng as Range 
  Set rng = Application.InputBox("Drag the range for calculation", "Sum", Type:=8)
  MsgBox "Sum of selected area is " & WorksheetFunction.Sum(rng)
End Sub
    
