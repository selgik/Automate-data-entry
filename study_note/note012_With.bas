'BASIC: FOR EACH (WITH)

'1. WITH
'   (Q) I need to refer certain object repetitively. But I do not want my codes getting longer and cluttered. 
'   (A) Use WIHT to organize the codes.

'1) Let's say user needs to set various properties to an object. In ex. user needs to change font to ariel, make font size 20 and make it italic.

'Instead of below:
Sub before_with()
  Selection.Font.Name = "ariel"
  Selection.Font.Size = 20
  Selection.Font.italic = true
End Sub

'Write as below:
Sub after_with()
  With Selection.Font
    .Name = "ariel"
    .Size = 20
    .italic = true
  End With
End Sub

