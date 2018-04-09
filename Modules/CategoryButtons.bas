Attribute VB_Name = "CategotyButtons"
Public cmCat As String
Public cmNote As String

Sub clearfilters()
ActiveSheet.ShowAllData
End Sub

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

Sub cmExp()

If ActiveCell.Row > 1 Then

ActiveSheet.Range("a" & ActiveCell.Row).Select
ActiveSheet.Paste
Call category("Cecilia Metheny")
Call note("for " & cmCat & " - " & cmNote)
ActiveSheet.Range("a" & ActiveCell.Row).Select
'Application.CutCopyMode = False
infoMsg ("Great job! You finished recording this transaction.")
ActiveWorkbook.Save
ActiveWorkbook.Close
Workbooks(1).Activate
Worksheets("Expenses").Activate

Else: criticalMsg ("Please select a blank row.")
Exit Sub

End If

End Sub
