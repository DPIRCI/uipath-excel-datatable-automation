Sub Macro1()
'
' Macro1 Macro
'

'
    ActiveSheet.PivotTables("positionString").PivotFields("Office Office"). _
        AutoSort xlAscending, "Count of Office Office"
End Sub