Sub while_loop()

    Cells(1, 1) = 1
    Cells(2, 1) = 2
    Cells(3, 1) = 3
    Cells(4, 1) = 4
    Cells(5, 1) = 5
    Cells(6, 1) = 6
    Cells(7, 1) = 7
    Cells(8, 1) = 8
    Cells(9, 1) = 9
    Cells(10, 1) = 10
    Cells(11, 1) = 11
    Cells(12, 1) = 12

End Sub

Sub while_loop()

    While [condition]
        'Instructions
    Wend

End Sub
  
Sub while_loop()

    Dim num As Integer
    num = 1 'Starting number (in this case, this is both the row number and the number that will be placed in each cell)

    While num <= 12 'As long as the num variable is <= 12, the instructions will loop
        Cells(num, 1) = num 'Numbering
        num = num + 1 'The number is increased by 1 each time the instructions loop
    Wend
	
End Sub
    
    
    
Sub do_while_loop()

    Do While [condition]
        'Instructions
    Loop

End Sub
    
Sub do_while_loop()

    Do
        'Instructions
    Loop While [condition]

End Sub
      
Sub do_while_loop()

    Do Until [condition]
        'Instructions
    Loop

End Sub
      
Sub for_loop()

    For i = 1 To 5
        'Instructions
    Next

End Sub
    
Sub for_loop()

    For i = 1 To 5
        MsgBox i
    Next

End Sub
  
Sub for_loop()

    Dim max_loops As Integer
    max_loops = Range("A1") 'In A1: we have defined a limit to the number of repetitions

    For i = 1 To 7 'Number of loops expected: 7
        If i > max_loops Then 'If A1 is empty or contains a number < 7, decrease the number of loops
            Exit For 'If the condition is true, we exit the For loop
        End If
    
        MsgBox i
    Next

End Sub

Sub loops_exercise()

    Const NB_CELLS As Integer = 10 'Number of cells to which we want to add background colors

    '...
    
End Sub

Sub loops_exercise()

    Const NB_CELLS As Integer = 10 'Number of cells to which we want to add background colors

    For r = 1 To NB_CELLS 'r => row number
    
        Cells(r, 1).Interior.Color = RGB(0, 0, 0) 'Black

    Next
    
End Sub

Sub loops_exercise()

    Const NB_CELLS As Integer = 10 'Number of cells to which we want to add background colors

    For r = 1 To NB_CELLS 'r => row number
    
        If r Mod 2 = 0 Then 'Mod => is the remainder from division
            Cells(r, 1).Interior.Color = RGB(200, 0, 0) 'Red
        Else
            Cells(r, 1).Interior.Color = RGB(0, 0, 0) 'Black
        End If

    Next
    
End Sub

Sub loops_exercise()

    Const NB_CELLS As Integer = 10 '10x10 checkerboard of cells

    For r = 1 To NB_CELLS 'r => row number
    
        For c = 1 To NB_CELLS 'c => column number
        
            If r Mod 2 = 0 Then
                Cells(r, c).Interior.Color = RGB(200, 0, 0) 'Red
            Else
                Cells(r, c).Interior.Color = RGB(0, 0, 0) 'Black
            End If
            
        Next
    Next
    
End Sub

Sub loops_exercise()

    Const NB_CELLS As Integer = 10 '10x10 checkerboard of cells
    Dim offset_row As Integer, offset_col As Integer ' => adding 2 variables
    
    'Shift (rows) starting from the first cell = the row number of the active cell - 1
    offset_row = ActiveCell.Row - 1
    'Shift (columns) starting from the first cell = the column number of the active cell - 1
    offset_col = ActiveCell.Column - 1
    
    For r = 1 To NB_CELLS 'Row number
    
        For c = 1 To NB_CELLS 'Column number
        
            If (r + c) Mod 2 = 0 Then
            'Cells(row number + number of rows to shift, column number + number of columns to shift)
                Cells(r + offset_row, c + offset_col).Interior.Color = RGB(200, 0, 0) 'Red
            Else
                Cells(r + offset_row, c + offset_col).Interior.Color = RGB(0, 0, 0) 'Black
            End If
            
        Next
    Next
    
End Sub

