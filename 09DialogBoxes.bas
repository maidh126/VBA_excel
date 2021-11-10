Sub delete_B2()
    Range("B2").ClearContents
    MsgBox "The contents of B2 have been deleted !"
End Sub

Sub delete_B2()
    If MsgBox("Are you sure that you wish to delete the contents of B2 ?", vbYesNo, "Confirm") = vbYes Then
        Range("B2").ClearContents
        MsgBox "The contents of B2 have been deleted !"
    End If
End Sub
  

Sub humor()
    Do
        If MsgBox("Do you like the Excel-Pratique site ?", vbYesNo, "Survey") = vbYes Then
            Exit Do ' => Yes response = Yes we exit the loop
        End If
    Loop While 1 = 1 ' => Infinite loop
    MsgBox ";-)"
End Sub
      
Sub example()
    Dim result As String
    
    result = InputBox("Text ?", "Title") 'The variable is assigned the value entered in the InputBox
    
    If result <> "" Then 'If the value anything but "" the result is displayed
        MsgBox result
    End If
End Sub
    
    
