Private Sub Workbook_BeforeClose(Cancel As Boolean)
    'If the user responds NO, the Cancel variable will have the value TRUE (which will cancel the closing of the workbook)
    If MsgBox("Are you sure that you want to close this workbook ?", 36, "Confirm") = vbNo Then
        Cancel = True
    End If
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If Sh.Name = "Sheet1" Then
        Target.Interior.Color = RGB(255, 108, 0) 'Orange color
    Else
        Target.Interior.Color = RGB(136, 255, 0) 'Green color
    End If
End Sub

