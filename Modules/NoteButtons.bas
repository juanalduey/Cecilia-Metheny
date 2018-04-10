Attribute VB_Name = "NoteButtons"

'NOTE FUNCTIONS
' These functions are used in the note macros, which add text to the 'notes' field in either the income or
'expenses tab.

Function note(notes As String)
'enter string in notes column
    If ActiveCell.Row = 1 And ActiveCell.Value = "" Then
        Call criticalMsg("Choose a blank row or clear this cell")
        Else
        ActiveSheet.Range("d" & ActiveCell.Row).Value = notes
        ActiveWorkbook.Save
    End If
End Function

Function noteadd(notes, append)
    ActiveSheet.Range("d" & ActiveCell.Row).Value = ActiveSheet.Range("d" & ActiveCell.Row).Value & " - " & notes & " $" & append
End Function

Function NoteAddDetails(question As String, title As String, amount, leadString As String, trailString As String)
    If ActiveCell.Row > 1 Then
        amount = InputBox(question, title)
        Call noteadd(leadString, amount + trailString)
        Else
        Call criticalMsg("Please choosse a blank row or clear the current one")
    End If
End Function

'NOTE MACROS
'These macros are used for the buttons that add text to the 'notes' field in the income and expense
'tabs.

Sub tdb()
'Add TDB- 4978 to the notes field
    Call note("TDB - 4978")
End Sub

Sub cash()
    Call note("Cash")
End Sub

Sub totalCharge()
    Call NoteAddDetails("What was the total charge made on the card?", "Total Charge", amount, "Total charge on card:", "")
End Sub

Sub cashBack()
    Call NoteAddDetails("How much cash back did you get during the transaction?", "Cash Back", amount, "including", " cash back")
End Sub

