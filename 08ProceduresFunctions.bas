Private Sub warning()
    MsgBox "Caution !!!"
End Sub

Sub macro_test()
    If Range("A1") = "" Then
        warning ' <= execute the procedure "warning"
    End If
    'etc ...
End Sub

Private Sub warning(var_text As String)
    MsgBox "Caution: " & var_text & " !"
End Sub

Sub macro_test()
    If Range("A1") = "" Then
        warning "empty cell"
    ElseIf Not IsNumeric(Range("A1")) Then
        warning "non-numerical value"
    End If
End Sub

'Example 1: the last name is displayed:
dialog_boxes last_name1
    
'Example 2: last name and first name are displayed:
dialog_boxes last_name1, first_name1
    
'Example 3: last name and age are displayed:
dialog_boxes last_name1, , age1
    
'Example 4: last name, first name, and age are displayed:
dialog_boxes last_name1, first_name1, age1

Sub macro_test()

    Dim last_name1 As String, first_name1 As String, age1 As Integer
    
    last_name1 = Range("A1")
    first_name1 = Range("B1")
    age1 = Range("C1")

    'Example 1: the last name is displayed:
    dialog_boxes last_name1
    
    'Example 2: last name and first name are displayed:
    dialog_boxes last_name1, first_name1
    
    'Example 3: last name and age are displayed:
    dialog_boxes last_name1, , age1
    
    'Example 4: last name, first name, and age are displayed:
    dialog_boxes last_name1, first_name1, age1

End Sub

Private Sub dialog_boxes(last_name As String, Optional first_name, Optional age)
    
    If IsMissing(age) Then 'If the age variable is missing ...
        
        If IsMissing(first_name) Then 'If the first_name variable is missing, only the last name will be displayed
            MsgBox last_name
        Else 'Otherwise, last name and first name will be displayed
            MsgBox last_name & " " & first_name
        End If
        
    Else 'If the age variable is present ...

        If IsMissing(first_name) Then 'If the first_name variable is missing, last name and age will be displayed
            MsgBox last_name & ", " & age & " years old"
        Else 'Otherwise, last name, first name, and age will be displayed
            MsgBox last_name & " " & first_name & ", " & age & " years old"
        End If
    
    End If
       
End Sub

Sub macro_test()
    Dim var_number As Integer
    var_number = 30
    
    calcul_square var_number
    
    MsgBox var_number
End Sub

Private Sub calcul_square(ByRef var_value As Integer) 'ByRef does not need to be specified (because it is the default)
    var_value = var_value * var_value
End Sub

var_number = 30
'The initial value of the "var_number" variable is 30

calcul_square var_number
'The sub procedure is launched with "var_number" as an argument

Private Sub calcul_square(ByRef var_value As Integer)
'The "var_value" variable is in some way a shortcut to "var_number", which means that if the "var_value" variable is modified, the "var_number" variable will also be modified (and they don't have to have the same name)

var_value = var_value * var_value
'The value of the "var_value" variable is modified (and therefore the "var_number" is modified as well)

End Sub
'End of sub procedure

MsgBox var_number
'The "var_number" variable was modified, so 900 will now be displayed in the dialog box

var_number = 30
'The initial value of the variable "var_number" is 30

calcul_square var_number
'The sub procedure is launched with the variable "var_number" as an argument

Private Sub calcul_square(ByVal var_value As Integer)
'The variable "var_value" copies the value of the variable "var_number" (the 2 variables are not linked)

var_value = var_value * var_value
'The value of the variable "var_value" is modified

End Sub
'End of sub procedure (the sub procedure in this example doesn't have any effect at all)

MsgBox var_number
'The variable "var_number" has not been modified, and so 30 will be displayed in the dialog box

Function square(var_number)
    square = var_number ^ 2 'The function "square" returns the value of "square"
End Function

Sub macro_test()
    Dim result As Double
    result = square(9.876) 'The variable result is assigned the value returned by the fonction
    MsgBox result 'Displays the result (the square of 9.876, in this case)
End Sub

