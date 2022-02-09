'Tast1: User fills in the form, click the button and data will be transffered.
'Data transfer route: Main_Dashboard -> sheet_2

Sub Update_Prd()

Dim a As Integer

a = Worksheets("sheet_2").Cells(1, 11).Value

'Step1: If cell C6 is NOT empty, transfer data from Main_Dashboard to sheet_2
If IsEmpty(Worksheets("Main_Dashbaord").Range("C6")) = False Then

Worksheets("sheet_2").Cells(a, 2) = Worksheets("Main_Dashbaord").Cells(5, 3)
Worksheets("sheet_2").Cells(a, 3) = Worksheets("Main_Dashbaord").Cells(6, 3)
Worksheets("sheet_2").Cells(a, 4) = Worksheets("Main_Dashbaord").Cells(7, 3)
Worksheets("sheet_2").Cells(a, 5) = Worksheets("Main_Dashbaord").Cells(8, 3)
Worksheets("sheet_2").Cells(a, 6) = Worksheets("Main_Dashbaord").Cells(9, 3)
Worksheets("sheet_2").Cells(a, 7) = Worksheets("Main_Dashbaord").Cells(10, 3)
Worksheets("sheet_2").Cells(a, 8) = Worksheets("Main_Dashbaord").Cells(11, 3)
Worksheets("sheet_2").Cells(a, 9) = Worksheets("Main_Dashbaord").Cells(12, 3)
Worksheets("sheet_2").Cells(a, 10) = Worksheets("Main_Dashbaord").Cells(13, 3)
Worksheets("sheet_2").Cells(a, 11) = Worksheets("Main_Dashbaord").Cells(14, 3)
Worksheets("sheet_2").Cells(a, 12) = Worksheets("Main_Dashbaord").Cells(15, 3)

'And then, clear the contents for following range.     
Worksheets("Main_Dashbaord").Range("C5:C15").ClearContents

'Otherwise (if C6 is empty) alert user to fill in the date.
Else
    MsgBox "Date is missing"
End If
'end

'Step2: Mark the day as default "full day" instead of "half day" so that user clicks less.
Range("C5").Select
ActiveCell.FormulaR1C1 = "Full"
Range("C6").Select

            
'Step3: refresh pivot so that dashboard can be auto-refreshed upon clicking button.
ActiveWorkbook.RefreshAll
    
'Step 4: sort horizontal bar chart analyzing offline activity, in ASC order.
    ActiveWorkbook.Worksheets("sheet_3").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("sheet_3").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("AG2:AG9"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
      With ActiveWorkbook.Worksheets("sheet_3").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


End Sub

