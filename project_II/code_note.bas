Sub Update_Prd()

Dim a As Integer

a = Worksheets("sheet_2").Cells(1, 11).Value

'Alert msg added
If IsEmpty(Worksheets("Main_Dashbaord").Range("C5")) = False Then
'end

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

    
Worksheets("Main_Dashbaord").Range("C5:C15").ClearContents

'Alert msg added
Else
    MsgBox "Date is missing"
End If
'end

Range("C5").Select
ActiveCell.FormulaR1C1 = "Full"
Range("C6").Select

'Refresh pivot
ActiveWorkbook.RefreshAll
    
    
'Sort bar chart (offline activity) asc
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
'End Sort

End Sub

