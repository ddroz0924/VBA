Attribute VB_Name = "Carryover"
Option Compare Text 'so carryover check is not case sensitive
Sub List_Carryover_Items()
'
' List_Carryover_Items Macro
' 1) Copy Carryover item to Savings area 2) copy amount 3) insert a new row
'

Dim sort As Long
sort = WorksheetFunction.CountIf(ActiveSheet.Range("H23:H85"), "Carryover") 'find number of carryover items

If IsEmpty(Range("B107").Value = True) Then 'True that it is empty and has not been run
    Copy_Carryover_Items
Else 'If its not empty, boolean registers as false
    MsgBox ("Already run. Reseting to run again.")
    Range("A107", Cells(106 + sort, 1)).EntireRow.Select 'highlight rows that were inserted by Copy_Carryover the first time
    Selection.EntireRow.Delete 'delete previous records
    Copy_Carryover_Items 'copy the items
End If
MsgBox ("Done!")
End Sub

Sub Copy_Carryover_Items()
Dim counter As Integer
Dim sort As Long
sort = WorksheetFunction.CountIf(ActiveSheet.Range("H23:H85"), "Carryover") 'find number of carryover items

For counter = 23 To 85
    'is it Carryover?
    If ActiveSheet.Cells(counter, 8).Value = "Carryover" Then
        Range("M107").EntireRow.Insert xlShiftDown 'adds new line if needed
        Range("C106").Copy Range("C107") 'copies the line total formula in savings area
        Cells(counter, 3).Copy Range("B107") 'copies the item description to savings area
        
        'copying the value
        If IsEmpty(Cells(counter, 5).Value) = False Then 'is it TO or FROM?
           Cells(counter, 5).Copy Range("M107") 'copy TO value as positive
        Else: Cells(counter, 6).Copy Range("M107") 'copy FROM value
            Range("M107").Value = Range("M107") * -1 'make copied FROM value as negative
        End If
    Else
    End If
Next counter

'sort the carryover copies to group FROMS and TOS
Range("A107", Cells(106 + sort, 13)).sort Key1:=Range("M98", Cells(106 + sort, 13)), order1:=xlDescending

End Sub

