'BASIC: LET'S START FROM MACRO

'1. (Q) Where can I find macro I recorded? 
'   (A) Developer > Visual Basic ("VB") > Modules

'2. (Q) How to use macro?
'   (A1) By creating a shortcut: Developer > Macros
'   (A2) By creating a button: Deveoper > Button > Create button AND THEN assign macro: 
'        Automatically by right clicking > Assign macros OR
'        Manually by VB > create a module (you will see recorded macro's vba) > call macro under button's module
'        Example below: I created macro under the name "Total_sum" so I add macro name under my button.

        Sub Button_Click()
         Call Total_sum
        End Sub

'3. 
