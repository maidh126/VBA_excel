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


'Sample variable declaration
Dim var1 As String
    
'Sample 1 dimensional array declaration
Dim array1(4) As String
    
'Sample 2 dimensional array declaration
Dim array2(4, 3) As String
    
'Sample 3 dimensional array declaration
Dim array3(4, 3, 2) As String

'Assigning values to three colored cells
array2(0, 0) = "Red cell value"
array2(4, 1) = "Green cell value"
array2(2, 3) = "Blue cell value"

Sub const_example()
    Cells(1, 1) = Cells(1, 2) * 6.87236476641
    Cells(2, 1) = Cells(2, 2) * 6.87236476641
    Cells(3, 1) = Cells(3, 2) * 6.87236476641
    Cells(4, 1) = Cells(4, 2) * 6.87236476641
    Cells(5, 1) = Cells(5, 2) * 6.87236476641
End Sub

Sub const_example()
   'Declaration of a constant + assignment of value
    Const ANNUAL_RATE As Double = 6.87236476641
    
    Cells(1, 1) = Cells(1, 2) * ANNUAL_RATE
    Cells(2, 1) = Cells(2, 2) * ANNUAL_RATE
    Cells(3, 1) = Cells(3, 2) * ANNUAL_RATE
    Cells(4, 1) = Cells(4, 2) * ANNUAL_RATE
    Cells(5, 1) = Cells(5, 2) * ANNUAL_RATE
End Sub

Sub procedure1()
   Dim var1 As Integer
   ' => Use of a variable only within a procedure
End Sub

Sub procedure2()
   ' => var1 cannot be used here
End Sub

Dim var1 As Integer

Sub procedure1()
   ' => var1 can be used here
End Sub

Sub procedure2()
   ' => var1 can also be used here
End Sub

'Creation of a variable type
Type guests
    last_name As String
    first_name As String
End Type
    
Sub variables()
    'Declaration
    Dim p1 As guests
    
    'Assigning values to p1
    p1.last_name = "Smith"
    p1.first_name = "John"
    
    'Example of use
    MsgBox p1.last_name & " " & p1.first_name
End Sub
