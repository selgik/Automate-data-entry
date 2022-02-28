'-------------------------------------------------------------------------------------------------
' GENERAL NOTE
' Back tick ' in front indicate it is a note.
' Cells(x, y) in VBA indicates Cells(row#, column#). Ex: Cell D3 = Cells(3, 4) 
' To add VBA as sub query, add code name without () in right position. See example under #2 below. 
'-------------------------------------------------------------------------------------------------


' #1. Insert_buttonA() 
'     Note: This is Type 4(1) macro mentioned in the report. 

Sub Insert_buttonA()

Dim a As Integer
a = Worksheets("tracker_worksheet").Cells(1, 9).Value
'Number 4 put in I1. This number will increase by +1 upon adding record in tracker_table

Worksheets("tracker_worksheet").Cells(a, 4) = Worksheets("entry_form_worksheet").Cells(2, 7)
Worksheets("tracker_worksheet").Cells(a, 7) = Worksheets("entry_form_worksheet").Cells(3, 7)
Worksheets("tracker_worksheet").Cells(a, 9) = Worksheets("entry_form_worksheet").Cells(4, 7)
'Add as much needed. Currently, reference ID, error type and comment need to be tracked. Hence 3 lines.

Worksheets("entry_form_worksheet").Range("x:y").ClearContents
'x, y refers to Cell where error data is being input.

End Sub


'-------------------------------------------------------------------------------------------------
' #2. Insert_buttonB() 
'     Note: This is Type 4(2) macro mentioned in the report. 
'           By using SplitMacro, selected items from the list will auto-assign to one and each cells. See #4.
'           After transffering items, Dup_Insert will move items left over as result of SplitMacro.

Sub Insert_buttonB()

Dim a As Integer

a = Worksheets("tracker_worksheet").Cells(x,y).Value
'where count number will be sitted

'1. Do a split first
SplitMacro1 '<-- splitting selected error types 
SplitMacro2 '<-- splitting added comments per types
    
'2. Insert 1st splitted item 
Worksheets("tracker_worksheet").Cells(a, y) = Worksheets("entry_form_worksheet").Cells(x, y6)
Worksheets("tracker_worksheet").Cells(a, y3) = Worksheets("entry_form_worksheet").Cells(x2, y6)
Worksheets("tracker_worksheet").Cells(a, y4) = Worksheets("entry_form_worksheet").Cells(x3, y6)

'3. Repeat insert for rest items
Dup_Insert2 
Dup_Insert3
Dup_Insert4
   
Worksheets("entry_form_worksheet").Range("x:y").ClearContents
End Sub


'------------------------------------------------
' #3. Split()
'     NOTE: This is required for using Type 4(2) macro mentioned in the report. 
'           Macro recorded by going through Data > Text To Columns > Delimited > commas

Sub Split()

If ActiveSheet.Range("A1").Value = "" Then Exit Sub
'Stop macro when/if cell is empty

    Range("A1").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, OtherChar _
        :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
        
        
End Sub


'------------------------------------------------
' #4. MultipleSelection()
'     Note   : This is required for using Type 4(2) macro mentioned in the report. 
'              To allow multiple selections in a Drop Down List in Excel (without repetition)
'     Credit : https://trumpexcel.com/select-multiple-items-drop-down-list-excel/
      
      
Private Sub Worksheet_Change(ByVal Target As Range)

Dim Oldvalue As String
Dim Newvalue As String
Application.EnableEvents = True
On Error GoTo Exitsub
If Target.Address = "$A$1" Then 
  If Target.SpecialCells(xlCellTypeAllValidation) Is Nothing Then
    GoTo Exitsub
  Else: If Target.Value = "" Then GoTo Exitsub Else
    Application.EnableEvents = False
    Newvalue = Target.Value
    Application.Undo
    Oldvalue = Target.Value
      If Oldvalue = "" Then
        Target.Value = Newvalue
      Else
        If InStr(1, Oldvalue, Newvalue) = 0 Then
            Target.Value = Oldvalue & "," & Newvalue 
      Else:
        Target.Value = Oldvalue
      End If
    End If
  End If
End If
Application.EnableEvents = True
Exitsub:
Application.EnableEvents = True
End Sub


'-------------------------------------------------------------------------------------------------
' #5. Clear_button()

Sub Clear_button()

CarryOn = MsgBox("Are you sure to clear them all?", vbYesNo, "ALERT")
If CarryOn = vbYes Then

Worksheets("tracker_worksheet").Range("range1, range2").ClearContents

End If
End Sub


'-------------------------------------------------------------------------------------------------
' #5. If cell is empty, alert and stop

If IsEmpty(Range) = False Then
     'execute code
Else
    MsgBox "Cell is empty"
End If
End Sub


'- END -
