Attribute VB_Name = "LucilleMethenyButtons"
Public cmCat As String
Public cmNote As String

'LUCILLE METHENY FUNCTION
'This function is used to set the cmCat and cmNote variables, which determine the category and price of a
'transaction that involves Lucille Metheny's accounts.


Function GetDataForNextSheet()
    ActiveSheet.Range("a" & ActiveCell.Row, "c" & ActiveCell.Row).Copy
    cmCat = ActiveSheet.Range("b" & ActiveCell.Row).Value
    cmNote = ActiveSheet.Range("d" & ActiveCell.Row).Value
    Worksheets("Income").Activate
End Function
 
'LUCILLE METHENY MACROS
'These macros are for the buttons that facilitate the process of entering a transaction
'that involves Lucille Metheny's accounts.
 
 
 Sub fvb()
 'Add FVB - 1380 to notes field, then instruct on next step for recording transaction
    Call Note("FVB - 1380")
    If ActiveCell.Row > 1 Then
        Call criticalMsg("Go to the income tab, select the next availabe row, then press the Lucille Metheny button.")
        Call GetDataForNextSheet
        Else: Exit Sub
    End If
 End Sub


Sub lmftb()
'Add 53B - 4896 to notes field, then instruct on next step for recording transaction
    Call Note("53B - 4896")
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
        Selection.PasteSpecial Paste:=xlPasteValues
        Call category("Lucille Metheny")
        Call Note("for " & cmCat & " - " & cmNote)
        Call criticalMsg("Now enter this transaction as an expense in Lucille's spreadsheet. Select the next available row in that sheet, then click the CECILIA METHENY button.")
        ActiveSheet.Range("a" & ActiveCell.Row, "c" & ActiveCell.Row).Copy
        Workbooks.Open FileName:=ThisWorkbook.Path & "/LM Sheet (by Juan Alduey).xlsx"
        Worksheets("Expense").Activate
        Else
        Call criticalMsg("Please choose a blank row.")
        Exit Sub
    End If
End Sub


Sub cmExp()
'insert 'Cecile Metheny' in 'expense' field of active row in Lucille Metheny spreadsheet, along with transaction details.
    If ActiveCell.Row > 1 And ActiveCell.Value = "" Then
        ActiveSheet.Range("a" & ActiveCell.Row).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Call category("Cecilia Metheny")
        Call Note("for " & cmCat & " - " & cmNote)
        ActiveSheet.Range("a" & ActiveCell.Row).Select
        Call infoMsg("Great job! You finished recording this transaction.")
        With ActiveWorkbook
            .Save
            .Close
        End With
        Workbooks(1).Activate
        Worksheets("Expenses").Activate
        Else: Call criticalMsg("Please select a blank row.")
        Exit Sub
    End If
End Sub
