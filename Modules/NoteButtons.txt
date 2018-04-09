Attribute VB_Name = "NoteButtons"
Sub tdb()
'Add TDB- 4978 to the notes field
Call note("TDB - 4978")
End Sub

Sub cash()
Call note("Cash")
End Sub

Sub totalCharge()
Call NoteAddDetails("What was the total charge made on the card?", "Total Charge", amount, "Total charge on card:", "")
End Sub

Sub cashBack()
Call NoteAddDetails("How much cash back did you get during the transaction?", "Cash Back", amount, "including", " cash back")
End Sub

