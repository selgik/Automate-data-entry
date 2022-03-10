'BASIC: INPUTBOX AND MSGBOX

'1. INPUT DATA AND OUTPUT RESULT
'   (Q) How can I build input and output message box which will run certain calculation?
'   (A) Create button to input number -> message box to output result.

Sub GST_calculator()
  Dim total as double
  Dim gst as double
  
  total = Inputbox("Enter the total amount", "Calculator")
  'I can add ("Enter the total amount", "Calculator", 100) <-- 100 will pre-appear in the inputbox.
  gst = total * 0.07
  
  MsgBox "The GST is " & gst
  'If I want to show result in certain cell instead, I can write as:
  'Range("A1").Select
  'Selection.Value = gst
End Sub


'2. EXERCISE
