# NursingOrdersProject

Sub Macro1()
'
' Macro1 Macro
'

'
    Dim CountDup As Integer
    
    CountDup = 1
    Range("C2", Range("C2").End(xlDown)).Select
    For Each x In Selection
      If x.Value = x.Offset(1, 0).Value Then
        CountDup = CountDup + 1
        x.Offset(0, 1).Value = "delete"
      End If
      If x.Value <> x.Offset(1, 0).Value Then
        x.Offset(0, 1).Value = CountDup
        CountDup = 1
      End If
    Next x
    
End Sub
