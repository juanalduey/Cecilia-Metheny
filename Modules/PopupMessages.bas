Attribute VB_Name = "PopupMessages"

'POP-UP MESSAGE FUNCTIONS
'These functions are used in varius macros to warn or instruct the user.

Function criticalMsg(message As String)
'pop up message with alert icon
    a = MsgBox(message, vbCritical, "Heads up")
    ActiveWorkbook.Save
End Function

Function infoMsg(message As String)
'pop up message with info icon
    a = MsgBox(message, vbInformation, "Quick Note")
End Function

