Attribute VB_Name = "DateButtons"

'DATE FUNCTION
'This function is used in the macros that insert text in the 'date' field of the income and expense tabs.

Function Dates(dater As Date)
    If ActiveCell.Row = 1 Or Not ActiveCell.Value = "" Then
        Call criticalMsg("Choose a blank row")
        Else
        With ActiveSheet
            .Range("a" & ActiveCell.Row) = dater
            .Range("b" & ActiveCell.Row).Select
    End With
    End If
End Function

'DATE MACROS
'These macros insert dates in the 'dates' field of the income and expense tabs.

Sub todaydate()
    Call Dates(Date)
End Sub

Sub yestdate()
    Call Dates(Date - 1)
End Sub
 
Sub daybfryest()
    Call Dates(Date - 2)
End Sub
