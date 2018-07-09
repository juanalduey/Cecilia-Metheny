Attribute VB_Name = "CreditCardButtons"
'CREDIT CARD TEMPLATE FUNCTION
'This function defines the behavior of the buttons in the "credit card" column of the "credit card purchases" tab.
'It generates an error message when pressed if the selected cell is either in the first row, or not blank, otherwise
'it populates a cell in the "credit card" column of the selected row.

Function CreditCardTemplate(Cardt As String, CreditCardColumn As String)
'enter string in category field of active row
    If ActiveCell.Row = 1 Or Not Range(CreditCardColumn & ActiveCell.Row).Value = "" Then
        Call criticalMsg("Choose a blank row")
        Else
        With ActiveSheet
            .Range(CreditCardColumn & ActiveCell.Row) = Cardt
            .Range(CreditCardColumn & ActiveCell.Row).Select
        End With
    End If
End Function

'CREDIT CARD FUNCTION
'Tells the Credit Card Template Function what the "Credit Card" column is.

Function CreditCard(Cardt As String)
Call CreditCardTemplate(Cardt, "e")
End Function

'CREDIT CARD BUTTON MACROS
'These subroutines tell the Credit Card and Credit Card Template functions what text to enter in the "Credit Card" column when
'buttons are pressed.

Sub AmazonChase()
Call CreditCard("Amazon Chase - 5474")
End Sub

Sub HomeDepot()
Call CreditCard("Home Depot - 3655")
End Sub
