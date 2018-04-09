Attribute VB_Name = "LucilleMethenyButtons"
 Sub fvb()
 'Add FVB - 1380 to notes field, then instruct on next step for recording transaction
 
 Call note("FVB - 1380")
  
 If ActiveCell.Row > 1 Then
 
 Call criticalMsg("Go to the income tab, select the next availabe row, then press the Lucille Metheny button.")
 Call GetDataForNextSheet
 Else: Exit Sub
 End If
 
 End Sub

Sub lmftb()
'Add 53B - 4896 to notes field, then instruct on next step for recording transaction
 
Call note("53B - 4896")

If ActiveCell.Row > 1 Then

Call criticalMsg("Go to the income tab, select the next available row, and press the Lucille Metheny button")
Call GetDataForNextSheet

Else: Exit Sub
End If

End Sub

Sub lmincome()
'add income from Lucille Metheny, then give further instruction on recording transaction

If ActiveCell.Row > 1 And ActiveCell.Value = "" Then

ActiveSheet.Range("a" & ActiveCell.Row).Select
ActiveSheet.Paste
Call category("Lucille Metheny")
Call note("for " & cmCat & " - " & cmNote)
Call criticalMsg("Now enter this transaction as an expense in Lucille's spreadsheet. Select the next available row in that sheet, then click the CECILIA METHENY button.")
ActiveSheet.Range("a" & ActiveCell.Row, "c" & ActiveCell.Row).Copy
Workbooks.Open "/Users/ceciliametheny/Desktop/Cecilia Numbers with Macros/LM Sheet (by Juan Alduey).xlsx"
Else
Call criticalMsg("Please choose a blank row.")
Exit Sub
End If

End Sub

