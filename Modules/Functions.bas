Attribute VB_Name = "Functions"
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

Function category(categoryt As String)
'enter string in category column
If ActiveCell.Row = 1 And ActiveCell.Value = "" Then
Call criticalMsg("Choose a blank row")
Else
ActiveSheet.Range("b" & ActiveCell.Row) = categoryt
End If
End Function

Function criticalMsg(message As String)
'pop up message with alert icon
a = MsgBox(message, vbCritical, "Heads up")
ActiveWorkbook.Save
End Function

Function infoMsg(message As String)
'pop up message with info icon
a = MsgBox(message, vbInformation, "Quick Note")
End Function

Function Dates(dater As Date)
If ActiveCell.Row = 1 Then
Call criticalMsg("Choose a blank row")
Else
ActiveSheet.Range("a" & ActiveCell.Row) = dater
End If
End Function

Function GetDataForNextSheet()

 ActiveSheet.Range("a" & ActiveCell.Row, "c" & ActiveCell.Row).Copy
 cmCat = ActiveSheet.Range("b" & ActiveCell.Row).Value
 cmNote = ActiveSheet.Range("d" & ActiveCell.Row).Value
 Worksheets("Income").Activate

End Function

Function NoteAddDetails(question As String, title As String, amount, leadString As String, trailString As String)
 If ActiveCell.Row > 1 Then
amount = InputBox(question, title)
Call noteadd(leadString, amount + trailString)
Else
Call criticalMsg("Please choosse a blank row or clear the current one")
End If
End Function
