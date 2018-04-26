Attribute VB_Name = "CategotyButtons"


'CATEGORY FUNCTION
'This function is used in all macros that enter text in the category field, in the income and expense tabs.

Function category(categoryt As String)
'enter string in category field of active row
    If ActiveCell.Row = 1 Or Not ActiveCell.Value = "" Then
        Call criticalMsg("Choose a blank row")
        Else
        With ActiveSheet
            .Range("b" & ActiveCell.Row) = categoryt
            .Range("c" & ActiveCell.Row).Select
        End With
    End If
End Function

'CATEGORY MACROS
'These macros are used for the buttons that enter text in the 'category' field of the income and
'expense tabs.


Sub CeMe()
    Call category("CeMe")
End Sub

Sub NewSchool()
    Call category("The New School")
End Sub

Sub SocialSecurity()
    Call category("Social Security")
    ActiveSheet.Range("c" & ActiveCell.Row).Value = "928"
End Sub
 
 Sub Foodout()
    Call category("Food Out")
 End Sub
 
Sub BizTravel()
    Call category("Business Travel")
End Sub

Sub OfficeSupplies()
    Call category("Office Supplies")
End Sub

Sub Laundry()
    Call category("Laundry")
End Sub

Sub Taxi()
    category ("Taxi")
End Sub
 
 Sub Publictrans()
    Call category("Public Transit")
 End Sub
 
 Sub Grocerystore()
    Call category("Grocery Store")
 End Sub
