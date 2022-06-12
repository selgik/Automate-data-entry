'BASIC: OBJECTS AND PROPERTY

'1. OBJECTS
'   Object is an entity to which you can define its characteritic(property) or action(method)
'   Examples: range, worksheet, workbook, application etc. 
'             You can tell VBA to select range(object) and change its colour(property).
'   Collection is group of obejct within the same types. 

'2. SYNTAX
'   (1) Range        : Range("A1") or Range("A1:C2")
'   (2) Cell         : Cells(3, 2) or Cells(4,"D")
'   (3) Value        : Cells(1, 1).Value = 2
'   (4) ActiveCell   : ActiveCell.Value = "Test"
'   (5) CurrentRegion: Range("A1").CurrentRegion
'   (6) Formula      : Range("A1").Formula = "=Sum(A1:A10)"

'3. EXERCISES
'   (1) 3 ways to tell VBA to select the range:

     Range("A1:C10").Select
     Range("A1", "C10").Select
     Range(Range("A1"), Range("C10")).Select

    '(2) If user does not know where the data ends?
    '    Use may work on a project where data is being added regularly. 
    '    Scorlling down and finding out the ending cell number everytime, will be time-consuming.
    '    Solution: Use below two cods.
    
     Range([A1],[C1].End(xlDown)).Select
     Range("A1").CurrentRegion.Select
     'For seconde code however, user will have to make sure there is no missing row.
     'Otherwise, region will be incorrectly selected.
    
     'Applying the code:
      Range("A1").CurrentRegion.Select
      Selection.Borders.LineStyle = xlContinuous 
      'or xldot, xldouble
      Selection.Borders.Weight = xlThick
      'or xlmedium, xlthin
      
    '(3) Difference between Cells(3,2) vs Range("B5:D10").Cells(3,2)
    '    Cells(3,2) is B3 whereas, Range("B5:D10").Cells(3,2) is C7. This is because,
    '    We told system that cells(3,2) will be checked from the range of B5:D10.
    
    '(4) Assigning property for the Range:
    
      Range("B5:D10").Cells(3,2).Value = "Surprise"
      Range("B5:D10").Cells(3,2).Font.Size = 15
      Range("B5:D10").Cells(3,2).Font.Bold = True
      'Of course, above three codes can be neatly organized as:
    
      With Range("B5:D10").Cells(3,2)
        .Value = "Surprise"
        .Font.Size = 15
        .Font.Bold = True
      End With
    
    '(5) Difference between Select vs Active
      Range("A1:B10").Select
      Selection.Value = "Test1"
      'Result: Test1 will be inserted from cell A1 to B10
      
      Range("A1:B10").Activate
      ActiveCell.Value = "Test2"
      'Result: Test2 will be inserted on A1 only
    
    '(6) Using Active
     Sub highligh_Click()
       Dim rng As Range
        For Each rng In ActiveSheet.Range([A2],[B2].End(xlDown))
          If rng.Value >= 1000 Then
             rng.Interior.Color = RGB(230, 100, 30)
          End If
        Next
    End Sub
    
