'BASIC: LET'S START FROM MACRO

'1. FIND MACRO
'   (Q) Where can I find macro I recorded? 
'   (A) Developer > Visual Basic ("VB") > Modules

'2. USE MACRO
'   (Q) How to use macro?
'   (A1) By creating a shortcut: Developer > Macros
'   (A2) By creating Quick Access Toolbar: Go to QAT ... > More Commands > Choose Commands From: Macros > Add
'   (A3) By creating a button: Deveoper > Button > Create button AND THEN assign macro: 
'        Automatically by right clicking > Assign macros OR
'        Manually by VB > create a module (you will see recorded macro's vba) > call macro under button's module
'        Example below: I created macro under the name "Total_sum" so I add macro name under my button.

        Sub Button_Click()
         Call Total_sum
        End Sub

'3. FILTER OPTIONS 
'   (Q1) Filter where column 2 contains words "kor" or "sg"
'   (A1) Use Operator:=xlOr 
'   (A2) Simplify the codes
'   The problme is, this code will not work for more than 2 criteria

Sub Filter()

    Range("B3").Select
    Selection.AutoFilter
    
    'A1: descriptive version
    ActiveSheet.Range("$B$3:$E$8").AutoFilter Field:=2, Criteria1:="*kor*", _
    Operator:=xlOr, Criteria2:="*sg*"
    
    'A2: same as A1, simplified version
    ActiveSheet.Range("$B$3:$E$8").AutoFilter 2, "*kor*", xlOr, "*sg*"
                
End Sub

'   (Q2) Filter where column 2 matches with keywords (multiple or)
'   (A2) Criteria1:=Array("keyword 1","keyword 2","keyword 3")
        
Sub Filter()    
    ActiveSheet.Range("$B$3:$E$9").AutoFilter Field:=2, Criteria1:=Array("korean", "singaporean", _
                                                        "japanese", "vietnamese"), _
                                                        Operator:=xlFilterValues
End Sub
        
        
'4.
