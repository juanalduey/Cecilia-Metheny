Attribute VB_Name = "CategotyButtons"
'CATEGORY FUNCTION
'This function is used in all macros that enter text in the category field, in the income and expense tabs.

Function CategorySet(Categoryt As String, CategoryColum As String, NextCategoryColumn As String)
'enter string in category field of active row
    If ActiveCell.Row = 1 Or Not Range("b" & ActiveCell.Row).Value = "" Then
        Call criticalMsg("Choose a blank row")
        Else
        With ActiveSheet
            .Range("b" & ActiveCell.Row) = Categoryt
            .Range("c" & ActiveCell.Row).Select
        End With
    End If
End Function
Function CategoryPlusSubCategoryTemplate(Category As String, SubCategory As String, CategoryColumn As String, SubCategoryColumn As String, NextColumn As String)
If ActiveCell.Row = 1 Or Not Range(CategoryColumn & ActiveCell.Row).Value = "" Or Not Range(SubCategoryColumn & ActiveCell.Row).Value = "" Then
        Call criticalMsg("Choose a blank row")
        Else
        With ActiveSheet
            .Range(CategoryColumn & ActiveCell.Row) = Category
            .Range(SubCategoryColumn & ActiveCell.Row) = SubCategory
            .Range(NextColumn & ActiveCell.Row).Select
        End With
    End If
End Function
Function CategoryPlusSubCategory(Categoryt As String, SubCategoryt As String)
  Call CategoryPlusSubCategoryTemplate(Categoryt, SubCategoryt, "b", "c", "d")
End Function

'CATEGORY MACROS
'These macros are used for the buttons that enter text in the 'category' field of the income and
'expense tabs.

'EXPENSES

Sub FoodFoodOut()
    Call CategoryPlusSubCategory("Food", "Food Out")
End Sub

Sub FoodGroceryStore()
    Call CategoryPlusSubCategory("Food", "Grocery Store")
End Sub

Sub TransportPublicTransit()
    Call CategoryPlusSubCategory("Transport", "Public Transit")
End Sub
 
 Sub TaxDeductBizTravel()
    Call CategoryPlusSubCategory("Tax Deductible", "Business Travel")
 End Sub

Sub TaxDeductOfficeSupplies()
    Call CategoryPlusSubCategory("Tax Deductible", "Office Supplies")
End Sub

Sub HouseholdLaundry()
    Call CategoryPlusSubCategory("Household", "Laundry")
End Sub

'INCOME

Sub CeMe()
    Call Category("CeMe")
End Sub

Sub NewSchool()
    Call Category("The New School")
End Sub

Sub SocialSecurity()
    Call Category("Social Security")
    ActiveSheet.Range("c" & ActiveCell.Row).Value = "928"
End Sub
 
