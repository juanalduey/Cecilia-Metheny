Attribute VB_Name = "NoteButtons"
Public NotesField As Object
Public Notes As String



'NOTE FUNCTIONS
' These functions are used in the note macros, which add text to the 'notes' field in either the income or
'expenses tab.

Function SetNotesFieldObject()
Set NotesField = ActiveSheet.Range("d" & ActiveCell.Row)
End Function

Function Note(Notes)
    If ActiveCell.Row = 1 Then
    criticalMsg ("Please choose a cell outside the header row")
    Exit Function
    Else
    Call SetNotesFieldObject
    End If
    If Not NotesField.Value = "" Then
        a = MsgBox("This cell has information in it. Do you want to override it?", vbYesNo, "Replace Content")
        If a = vbYes Then
         NotesField.Value = Notes
         Else: Exit Function
         End If
       Else
         NotesField.Value = Notes
    End If
End Function

Function noteadd(Notes, append)
    'enter user customized string
    Call SetNotesFieldObject
    If ActiveCell.Row > 1 And Not NotesField.Value = "" Then
        NotesField.Value = NotesField.Value & " - " & Notes & " $" & append
    Else
        Call criticalMsg("You are either in the header row, or you have not yet entered your payment method. Please address one or both issues.")
   End If
End Function

Function NoteAddDetails(question As String, title As String, amount, leadString As String, trailString As String)
        Call SetNotesFieldObject
        If NotesField.Value = "" Or NotesField.Value = "Cash" Then GoTo ErrorHandler
AskForAmount:         amount = InputBox(question, title)
        Call noteadd(leadString, amount + trailString)
        Exit Function
ErrorHandler:
       If NotesField.Value = "Cash" Then
       criticalMsg ("No such thing as a debit/credit charge, or cash back, when you pay cash.")
       Else
       If NotesField.Value = "" Then
       criticalMsg ("Please enter a payment method first")
       Else
       If a = MsgBox("This cell has information in it. Do you want to override it?", vbYesNo, "Replace Content") = vbYes Then GoTo AskForAmount
       End If
       End If
End Function

'NOTE MACROS
'These macros are used for the buttons that add text to the 'notes' field in the income and expense
'tabs.

Sub tdb()
'Add TDB- 4978 to the notes field
    Call Note("TDB - 4978")
End Sub

Sub AmazonPrimeVisa()
' Add Amazon Prime Visa - 5474
    Call Note("Amazon Prime Visa - 5474")
End Sub

Sub cash()
    Call Note("Cash")
End Sub

Sub totalCharge()
    Call NoteAddDetails("What was the total charge made on the card?", "Total Charge", amount, "Total charge on card:", "")
End Sub

Sub cashBack()
    Call NoteAddDetails("How much cash back did you get during the transaction?", "Cash Back", amount, "including", " cash back")
End Sub

Sub test()
    ActiveSheet.Range("d" & ActiveCell.Row).Value = Notes
End Sub
