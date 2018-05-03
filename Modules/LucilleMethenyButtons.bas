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
    Call SetNotesFieldObject
    If Not NotesField.Value = "" Then GoTo errorhandler
continue:    Call Note("FVB - 1380")
    Call GetDataForNextSheet
    Exit Sub
errorhandler:
    If NotesField.Value = "FVB - 1380" Or NotesField.Value = "53B - 4896" Then
        criticalMsg ("To override this note, delete the corresponding transactions in your income tab and in Lucille's expense tab, then manually clear the notes field of this row.")
        Else
        a = MsgBox("This cell has information in it. Do you want to override it?", vbYesNo, "Replace Content")
        If a = vbYes Then
        NotesField.Value = ""
        GoTo continue
        Else: Exit Sub
        End If
        End If
    
End Sub

Sub lmftb()
'Add 53B - 4896 to notes field
Call SetNotesFieldObject
    If Not NotesField.Value = "" Then GoTo errorhandler
continue:    Call Note("53B - 4896")
    Call GetDataForNextSheet
    Exit Sub
errorhandler:
    If NotesField.Value = "FVB - 1380" Or NotesField.Value = "53B - 4896" Then
        criticalMsg ("To override this note, delete the corresponding transactions in your income tab and in Lucille's expense tab, then manually clear the notes field of this row.")
        Else
        a = MsgBox("This cell has information in it. Do you want to override it?", vbYesNo, "Replace Content")
        If a = vbYes Then
        NotesField.Value = ""
        GoTo continue
        Else: Exit Sub
        End If
        End If
    
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
        Worksheets("Expense").Activate
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
        Workbooks("CM Sheet (by Juan Alduey).xlsm").Activate
        Sheets("Expenses").Select
        Call infoMsg("The corresponding transactions have been recorded in both your income tab and in Lucille Metheny's expense tab.")
End Sub
