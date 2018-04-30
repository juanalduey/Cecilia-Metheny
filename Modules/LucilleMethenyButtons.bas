Attribute VB_Name = "LucilleMethenyButtons"
Public cmCat As String
Public cmNote As String

'LUCILLE METHENY FUNCTIONS
'These functions are used in the macros that record transactions related to Lucille Metheny

Function NextBlankRow()
'This function selects the next blank row in the active sheet
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
End Function

Function GetDataForNextSheet()
'This function is used to set the cmCat and cmNote variables, which determine the category and price of a
'transaction that involves Lucille Metheny's accounts.
    ActiveSheet.Range("a" & ActiveCell.Row, "c" & ActiveCell.Row).Copy
    cmCat = ActiveSheet.Range("b" & ActiveCell.Row).Value
    cmNote = ActiveSheet.Range("d" & ActiveCell.Row).Value
    Worksheets("Income").Activate
    Call lmincome
End Function

 
'LUCILLE METHENY MACROS
'These macros are for the buttons that facilitate the process of entering a transaction
'that involves Lucille Metheny's accounts.
 
 
 Sub fvb()
 'Add FVB - 1380 to notes field, then instruct on next step for recording transaction
    Call Note("FVB - 1380")
    Call GetDataForNextSheet
    End Sub

Sub lmftb()
'Add 53B - 4896 to notes field
    Call Note("53B - 4896")
    Call GetDataForNextSheet
End Sub


Sub lmincome()
'add income from Lucille Metheny, then open Lucille Metheny spreadsheet
        Call NextBlankRow
        ActiveSheet.Range("a" & ActiveCell.Row).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        ActiveSheet.Range("b" & ActiveCell.Row).Value = ""
        Call category("Lucille Metheny")
        Call Note("for " & cmCat & " - " & cmNote)
        ActiveSheet.Range("a" & ActiveCell.Row, "c" & ActiveCell.Row).Copy
        Workbooks.Open FileName:=ThisWorkbook.Path & "/LM Sheet (by Juan Alduey).xlsx"
        ActiveWindow.Visible = True
        Worksheets("Expense").Activate
        Call NextBlankRow
        Call cmExp
End Sub


Sub cmExp()
'insert 'Cecile Metheny' in 'expense' field of active row in Lucille Metheny spreadsheet, along with transaction details.
    Call NextBlankRow
    ActiveSheet.Range("a" & ActiveCell.Row).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Range("b" & ActiveCell.Row).Value = ""
    Call category("Cecilia Metheny")
    Call Note("for " & cmCat & " - " & cmNote)
    ActiveSheet.Range("a" & ActiveCell.Row).Select
    With ActiveWorkbook
           .Save
           .Close
    End With
        Workbooks(1).Activate
        Worksheets("Expenses").Activate
        Call infoMsg("The corresponding transactions have been recorded in both your income tab and in Lucille Metheny's expense tab.")
End Sub
