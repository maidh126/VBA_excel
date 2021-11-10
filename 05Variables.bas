'Display the value of the variable in a dialog box
Sub variables()
    'Declaring the variable
    Dim my_variable As Integer
    'Assigning a value to the variable
    my_variable = 12
    'Displaying the value of my_variable in a MsgBox
    MsgBox my_variable
End Sub

Sub variables()
    'Declaring variables
    Dim last_name As String, first_name As String, age As Integer
    
End Sub


Sub variables()
    'Declaring variables
    Dim last_name As String, first_name As String, age As Integer
    
    'Variable values
    last_name = Cells(2, 1)
    first_name = Cells(2, 2)
    age = Cells(2, 3)
    
End Sub

Sub variables()
    'Declaring variables
    Dim last_name As String, first_name As String, age As Integer
    
    'Variable values
    last_name = Cells(2, 1)
    first_name = Cells(2, 2)
    age = Cells(2, 3)
    
    'Dialog box
    MsgBox last_name & " " & first_name & ", " & age & " years old"
End Sub

Sub variables()
    'Declaring variables
    Dim last_name As String, first_name As String, age As Integer, row_number As Integer
        
    'Variable values
    row_number = Range("F5") + 1
    last_name = Cells(row_number, 1)
    first_name = Cells(row_number, 2)
    age = Cells(row_number, 3)
    
    'Dialog box
    MsgBox last_name & " " & first_name & ", " & age & " years old"
End Sub

'Declaring variables
Dim last_name As String, first_name As String, age As Integer, row_number As Integer

row_number = Range("F5") + 1

last_name = Cells(row_number, 1)
first_name = Cells(row_number, 2)
age = Cells(row_number, 3)

Sub variables()
    MsgBox Cells(Range("F5")+1,1) & " " & Cells(Range("F5")+1,2) & ", " & Cells(Range("F5")+1,3) & " years old"
End Sub
