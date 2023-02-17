Attribute VB_Name = "Merge_Macro"
Option Compare Text

Sub Merge_Macro()

Dim path As String
Dim label As String
Dim rerun As Integer

'Attempts to speed up macro
On Error GoTo Fallthrough


'loop so can run macro multiple times
Do Until rerun = vbNo

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

    path = InputBox("Would you like to make the ""Join Key"", ""Enrollment"", ""Profile"", or ""Data Quality""?")

    If path = "Join Key" Or path = "J" Then
        label = "JOIN KEY"
        Unique_Key label
    ElseIf path = "Enrollment" Or path = "E" Then
        label = "ENROLLMENT ID"
        VLOOKUP label, path
    ElseIf path = "Profile" Or path = "P" Then
        label = "PROFILE ID"
        VLOOKUP label, path
    ElseIf path = "Data Quality" Or path = "Q" Then
        label = "VLOOKUP"
        Quality label
    End If
    
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.ScreenUpdating = True

    'ask if you want to start over
    rerun = MsgBox("Do you want to create another merge tool?", vbYesNo, "Rerun?")
Loop

Fallthrough:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


Sub Unique_Key(label As String)

Dim nCol As Integer
Dim Key As Range
Dim BotRow As Long
Dim eDate As Range
Dim HMIS As Range
Dim visible As Range

On Error GoTo Fallthrough


'NEW COLUMN?
'

'Ask if you want new column
nCol = MsgBox("Do you want a new column?", vbQuestion + vbYesNo, "New Column?")

If nCol = vbYes Then
    Set Key = Range(InputBox("Type column where you want it", "Location") & 2)
    Range(Key.Address).EntireColumn.Insert
    Set Key = Range(Key.Offset(0, -1).Address) 'reset Key for new column location

    'label only if you added a new column
    Key.Offset(-1, 0).Value = label
    Key.Offset(-1, 0).Font.Bold = True

Else
    Set Key = Range(InputBox("Type column where you want it to start", "Location") & 2)

End If

'
'SET VARIABLES
'
     
'find last non-empty row
BotRow = Cells.Find(What:="*", _
                After:=Range("A1"), _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).Row

'Input box to ask for column of Entry Date
Set eDate = Range(InputBox("Type column letter of Entry/Exit Date:", "Entry/Exit Date") & 2)


'Input box to ask for column of HMIS ID
Set HMIS = Range(InputBox("Type column letter of HMIS ID:", "HMIS ID") & 2)


'Offset references if you added a new column
If nCol = vbYes Then
    Set eDate = eDate.Offset(0, 1)
    Set HMIS = HMIS.Offset(0, 1)
End If


'
'FORMULA to generate unique key
'
    'RowAbsolute arguement uses relative references to copy formula for the range
    'I think TODAY() is a volitile function, so I don't need to put calculate arguement yet

'Set General format in Key column
Range(Key.Address, Cells(BotRow, Key.Column)).NumberFormat = General
    
'Check if date is already formatted
If eDate.NumberFormat = "mm/dd/yyyy;@" Then 'needs DATEVALUE
    For Each visible In Range(Key.Address, Cells(BotRow, Key.Column)) 'checks if row is visible/filtered and only applies formula to visible rows
        If visible.EntireRow.Hidden = False Then
            Cells(visible.Row, Key.Column).Formula = "=IF(ISBLANK(" & Cells(visible.Row, eDate.Column).Address(RowAbsolute:=False) & ")=FALSE," & _
            "TODAY()-DATEVALUE(" & Cells(visible.Row, eDate.Column).Address(RowAbsolute:=False) & ")&"" ""&" & Cells(visible.Row, HMIS.Column).Address(RowAbsolute:=False) & _
            ", ""not exited""&"" ""&" & Cells(visible.Row, HMIS.Column).Address(RowAbsolute:=False) & ")"
        End If
    Next

Else 'already Excel date serial number format, or some other format

    For Each visible In Range(Key.Address, Cells(BotRow, Key.Column))
        If visible.EntireRow.Hidden = False Then
                       Cells(visible.Row, Key.Column).Formula = "=IF(ISBLANK(" & Cells(visible.Row, eDate.Column).Address(RowAbsolute:=False) & ")=FALSE," & _
            "TODAY()-" & Cells(visible.Row, eDate.Column).Address(RowAbsolute:=False) & "&"" ""&" & Cells(visible.Row, HMIS.Column).Address(RowAbsolute:=False) & _
            ", ""not exited""&"" ""&" & Cells(visible.Row, HMIS.Column).Address(RowAbsolute:=False) & ")"
        End If
    Next
End If

Fallthrough:
End Sub

Sub VLOOKUP(label As String, path As String)

Dim nCol As Integer
Dim Result As Range
Dim ref As Range
Dim BotRow As Long
Dim nAnswer As Integer
Dim shName As String
Dim visible As Range

On Error GoTo Fallthrough

'
'NEW COLUMN?
'

'Ask if you want new column
nCol = MsgBox("Do you want a new column?", vbQuestion + vbYesNo, "New Column?")

If nCol = vbYes Then
    Set Result = Range(InputBox("Type column where you want it", "Location") & 2)
    Range(Result.Address).EntireColumn.Insert
    Set Result = Range(Result.Offset(0, -1).Address) 'reset Result for new column location
    
    'label only if you added a new column
    Result.Offset(-1, 0).Value = label
    Result.Offset(-1, 0).Font.Bold = True
Else
    Set Result = Range(InputBox("Type column where you want it to start", "Location") & 2)

End If

'
'SET VARIABLES
'

'find last non-empty row
BotRow = Cells.Find(What:="*", _
                After:=Range("A1"), _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).Row

'Input box to ask for name of Apricot sheet
nAnswer = MsgBox("Is your Apricot sheet called ""Apricot""?", vbYesNoCancel, "Where is Apricot info?")

If nAnswer = vbNo Then
    shName = InputBox("What is Apricot sheet called?", "Apricot name?")
Else
    shName = "Apricot"
End If

'Set General format in Key column
Range(Result.Address, Cells(BotRow, Result.Column)).NumberFormat = General

'Is this Profile or Enrollment?
If path = "Enrollment" Or path = "E" Then

    'Input box to ask for column of Join Key/refernce
    Set ref = Range(InputBox("Type column letter of Join Key:", "Join Key") & 2)
           
    'VLOOKUP for Profile
    For Each visible In Range(Result.Address, Cells(BotRow, Result.Column))
        If visible.EntireRow.Hidden = False Then
            Cells(visible.Row, Result.Column).Formula = "=VLOOKUP(" & Cells(visible.Row, ref.Column).Address(RowAbsolute:=False) & ",'" & shName & "'!$A:$Z,9,FALSE)"
        End If
    Next
    
ElseIf path = "Profile" Or path = "P" Then

    'Input box to ask for column of Join Key/refernce
    Set ref = Range(InputBox("Type column letter of HMIS ID:", "HMIS ID") & 2)
    
        'Offset references if you added a new column
        If nCol = vbYes Then
        
            Set ref = ref.Offset(0, 1)
        End If
    
    'VLOOKUP for Profile
    For Each visible In Range(Result.Address, Cells(BotRow, Result.Column))
        If visible.EntireRow.Hidden = False Then
            Cells(visible.Row, Result.Column).Formula = "=VLOOKUP(" & Cells(visible.Row, ref.Column).Address(RowAbsolute:=False) & ",'" & shName & "'!$D:$M,4,FALSE)"
        End If
    Next
End If

Fallthrough:
End Sub


Sub Quality(label As String)

Dim nCol As Integer
Dim Result As Range
Dim ref As Range
Dim BotRow As Long
Dim nAnswer As Integer
Dim shName As String
Dim visible As Range

On Error GoTo Fallthrough

'
'NEW COLUMN?
'

'Ask if you want new column
nCol = MsgBox("Do you want a new column?", vbQuestion + vbYesNo, "New Column?")

If nCol = vbYes Then
    Set Result = Range(InputBox("Type column where you want it", "Location") & 2)
    Range(Result.Address).EntireColumn.Insert
    Set Result = Range(Result.Offset(0, -1).Address) 'reset Result for new column location
    
    'label only if you added a new column
    Result.Offset(-1, 0).Value = label
    Result.Offset(-1, 0).Font.Bold = True
Else
    Set Result = Range(InputBox("Type column where you want it to start", "Location") & 2)

End If

'
'SET VARIABLES
'

'Input box to ask for column of Join Key/refernce
Set ref = Range(InputBox("Type column letter of Join Key:", "Join Key") & 2)

'Input box to ask for name of HMIS sheet
nAnswer = MsgBox("Is your HMIS sheet called ""HMIS""?", vbYesNoCancel, "Where is HMIS info?")

If nAnswer = vbNo Then
    shName = InputBox("What is HMIS sheet called?", "HMIS name?")
Else
    shName = "HMIS"
End If


'find last non-empty row
BotRow = Cells.Find(What:="*", _
                After:=Range("A1"), _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).Row

'Set General Cell format for result column
Range(Result.Address, Cells(BotRow, Result.Column)).NumberFormat = General

'VLOOKUP formula to look for Join Key in HMIS sheet
For Each visible In Range(Result.Address, Cells(BotRow, Result.Column))
    If visible.EntireRow.Hidden = False Then
        Cells(visible.Row, Result.Column).Formula = "=VLOOKUP(" & Cells(visible.Row, ref.Column).Address(RowAbsolute:=False) & ",'" & shName & "'!$A:$Z,1,FALSE)"
    End If
Next

'Screen Update so you can see result
Application.ScreenUpdating = True

Fallthrough:

End Sub


