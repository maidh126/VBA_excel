Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Static previous_selection As String

    If previous_selection <> "" Then
        'Removing background color from previous selection:
        Range(previous_selection).Interior.ColorIndex = xlColorIndexNone
    End If

    'Adding background color to current selection:
    Target.Interior.Color = RGB(181, 244, 0)

    'Saving the address of the current selection:
    previous_selection = Target.Address
End Sub

Private Sub Worksheet_Activate()

End Sub

Private Sub Worksheet_Deactivate()

End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)

End Sub

Private Sub Worksheet_Calculate()

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

End Sub

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

End Sub

Application.EnableEvents = False ' => deactivate events
'Instructions
Application.EnableEvents = True ' => reactivate events
