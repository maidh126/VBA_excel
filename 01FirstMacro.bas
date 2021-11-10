Sub Macro1()
'
' Macro1 Macro
'

'
    Columns("A:A").Select
    Selection.ClearContents
    Columns("C:C").Select
    Selection.ClearContents
    Columns("B:B").Select
    Selection.Cut Destination:=Columns("A:A")
    Columns("D:D").Select
    Selection.Cut Destination:=Columns("C:C")
    Columns("C:C").Select
End Sub
