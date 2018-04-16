Attribute VB_Name = "ClearFilters"
'CLEAR FILTERS
' This macro clears filters while keeping filter buttons

Function clearthefilters()
    On Error GoTo ErrorHandler
    ActiveSheet.ShowAllData
Exit Function
ErrorHandler: infoMsg ("All filters are already cleared.")
    If Not ActiveSheet.AutoFilterMode Then
        ActiveSheet.Range("A1:d1").AutoFilter
    End If
End Function

Sub ClearFilters()
    Call clearthefilters
End Sub

