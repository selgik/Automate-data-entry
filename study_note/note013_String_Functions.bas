'BASIC: VBA STRING FUNCTIONS

'1. STRING FUNCTIONS
'   Assume that A1.Value=astronout-in-the-ocean
'   Assume that B1.Value= astronout 

'   (1) Instr(start_point, string, keyword) will return the place number where keyword appears.
'       Example: Instr(1, A1, -) will return 10.
'       Exmplain: first - appears at 10th place of A1.

'   (2) InStrRev(string, keyword, [starting_point]) will work similarly as Instr but in reverse order.
'       Exmaple: InStrRev(A1, -) will return 17.
'       Exmplain: from the back, - appear 6th place "-ocean". It is 17th place from the beginning.

'   (3) Lcase(string) will return strings in lower case. 

'   (4) Ucase(string) will return strings in upper case.

'   (5) Ltrim / Rtrim / Trim(string) will trim the space on the left side (Ltrim), right side (Rtrim) or in both sides (Trim).
'       Exmaple: Trim(B1) will return astronout (without space in the beginning)
'       Tip: Trim will not remove space within the string. In ex, Trim(as tronout) will strill be "as tronout".

'   (6) Left(string, number) / Right(string, number) / Mid(string, starting_point, number) will extract string as requested.
'       Example: Right(A1, 5) = ocean

'   (7) Len(string) will return number of characters in string.

'   (8) Str(string) will transform number into string 

'   (9) Strcomp(string1, string2, [vbtextcompare]) will verify whether string 1 and 2 are the same or not.
'       Example: 

        Sub strcomp_test()
          Dim a, b As String
          Dim result1, result2 As Integer
          a = "HAPPY"
          b = "happy"

          result1 = Strcomp(a, b, vbtextcompare)
            MsgBox "result1 is " & result1
            'Result will show 0. vbtextcompare will merely compare the strings. It will not consider lower/upper case. 

          result2 = Strcomp(a, b)
            MsgBox "result2 " & result2
            'Result will show -1. System will compare in binary and b will be considered to have bigger value.
        End Sub
  

'2. EXERCISE
'   (1) Situation: You have a table containing employee data. 
'       In column A, you have employee name, in column B - their team name and in column C -  manager's name.
'   (2) Problem: Each team has sent their own list and the format is not unified.
'       In example, team name may appear in many format such as "amr-sales" or "sales(emea)". 
'   (4) Task: With one click, you want to highlight the sell containing certain keyword in team name.
'       In example, you want to see the employee name highlighted with team name containing a keyword "sales".


  
