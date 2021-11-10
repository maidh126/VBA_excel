If [CONDITION HERE] Then ' => IF condition is validated, THEN
    'Instructions if true
Else ' => OTHERWISE
    'Instructions if false
End If

Sub variables()
    'Declaring variables
    Dim last_name As String, first_name As String, age As Integer, row_number As Integer
        
    'Assigning values to variables
    row_number = Range("F5") + 1
    last_name = Cells(row_number, 1)
    first_name = Cells(row_number, 2)
    age = Cells(row_number, 3)
    
    'Dialog box
    MsgBox last_name & " " & first_name & ", " & age & " years old"
End Sub

Sub variables()

    'If the value in parentheses (cell F5) is numerical (AND THEREFORE IF THE CONDITION IS TRUE) then
    'execute the instructions that follow THEN
    If IsNumeric(Range("F5")) Then
    
        'Declaring variables
        Dim last_name As String, first_name As String, age As Integer, row_number As Integer
        'Values of variables
        row_number = Range("F5") + 1
        last_name = Cells(row_number, 1)
        first_name = Cells(row_number, 2)
        age = Cells(row_number, 3)
        'Dialog Box
        MsgBox last_name & " " & first_name & ", " & age & " years old"
        
    End If
    
End Sub

Sub variables()

    If IsNumeric(Range("F5")) Then 'IF CONDITION TRUE
    
        'Declaring variables
        Dim last_name As String, first_name As String, age As Integer, row_number As Integer
        'Values of variables
        row_number = Range("F5") + 1
        last_name = Cells(row_number, 1)
        first_name = Cells(row_number, 2)
        age = Cells(row_number, 3)
        'Dialog box
        MsgBox last_name & " " & first_name & ", " & age & " years old"
        
    Else 'IF CONDITION FALSE
    
        'Dialog box: warning
        MsgBox "Your entry" & Range("F5") & " is not valid !"
        'Deleting the contents of cell F5
        Range("F5").ClearContents
    
    End If
    
End Sub


Sub variables()
    If IsNumeric(Range("F5")) Then 'IF NUMERICAL
        Dim last_name As String, first_name As String, age As Integer, row_number As Integer
        row_number = Range("F5") + 1

        If row_number >= 2 And row_number <= 17 Then 'IF CORRECT NUMBER
            last_name = Cells(row_number, 1)
            first_name = Cells(row_number, 2)
            age = Cells(row_number, 3)
            MsgBox last_name & " " & first_name & ", " & age & " years old"
        Else 'IF NUMBER IS INCORRECT
            MsgBox "Your entry " & Range("F5") & " is not a valid number !"
            Range("F5").ClearContents
        End If
        
    Else 'IF NOT NUMERICAL
        MsgBox "Your entry " & Range("F5") & " is not valid !"
        Range("F5").ClearContents
    End If
End Sub


Sub variables()
    If IsNumeric(Range("F5")) Then 'IF NUMERICAL
        Dim last_name As String, first_name As String, age As Integer, row_number As Integer
        Dim nb_rows As Integer
        
        row_number = Range("F5") + 1
        nb_rows = WorksheetFunction.CountA(Range("A:A")) 'NBVAL Function
        
        If row_number >= 2 And row_number <= nb_rows Then 'IF CORRECT NUMBER 
            last_name = Cells(row_number, 1)
            first_name = Cells(row_number, 2)
            age = Cells(row_number, 3)
            MsgBox last_name & " " & first_name & ", " & age & " years old"
        Else 'IF NUMBER IS INCORRECT
            MsgBox "Your entry " & Range("F5") & " is not a valid number !"
            Range("F5").ClearContents
        End If

    Else 'IF NOT NUMERICAL
        MsgBox "Your entry " & Range("F5") & " is not valid !"
        Range("F5").ClearContents
    End If
End Sub

If [CONDITION 1] Then ' => IF condition 1 is true, THEN
    'Instructions 1
ElseIf [CONDITION 2] Then ' => IF condition 1 is false, but condition 2 is true, THEN
    'Instructions 2
Else ' => OTHERWISE
    'Instructions 3
End If

Sub scores_comment()
    'Variables
    Dim note As Integer, score_comment As String
    note = Range("A1")
    
    'Comments based on the score
    If note = 6 Then
        score_comment = "Excellent score !"
    ElseIf note = 5 Then
        score_comment = "Good score"
    ElseIf note = 4 Then
        score_comment = "Satisfactory score"
    ElseIf note = 3 Then
        score_comment = "Unsatisfactory score"
    ElseIf note = 2 Then
        score_comment = "Bad score"
    ElseIf note = 1 Then
        score_comment = "Terrible score"
    Else
        score_comment = "Zero score"
    End If
    
    'Comments in B1
    Range("B1") = score_comment
End Sub

Sub scores_comment()
    'Variables
    Dim note As Integer, score_comment As String
    note = Range("A1")
    
    'Comments based on the score
    Select Case note    ' <= the value to test (the score, in this case)
    Case Is = 6         ' <= if the value = 6
        score_comment = "Excellent score !"
    Case Is = 5         ' <= if the value = 5
        score_comment = "Good score"
    Case Is = 4         ' <= if the value = 4
        score_comment = "Satisfactory score"
    Case Is = 3         ' <= if the value = 3
        score_comment = "Unsatisfactory score"
    Case Is = 2         ' <= if the value = 2
        score_comment = "Bad score"
    Case Is = 1         ' <= if the value = 1
        score_comment = "Terrible score"
    Case Else           ' <= if the value isn't equal to any of the above values
        score_comment = "Zero score"
    End Select
    
    'Comment in B1
    Range("B1") = score_comment
End Sub

