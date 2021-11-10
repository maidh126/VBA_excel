Sub example()
    'Select cell A8
    Range("A8").Select
End Sub
        
Sub example()        
    'Activating of Sheet 2
    Sheets("Sheet2").Activate
End Sub
    
Sub example()    
    'Selecting of Cell A8
    Range("A8").Select
End Sub
    
Sub example()    
    'Selecting A8 and C5
    Range("A8, C5").Select
End Sub
    
Sub example()    
    'Selecting cells A1 to A8
    Range("A1:A8").Select
End Sub
    
Sub example()    
    'Selecting cells from the "my_range" range
    Range("my_range").Select
End Sub  
    
Sub example()    
    'Selecting the cell in row 8 and column 1
    Cells(8, 1).Select
End Sub

Sub example()
    'Random selection of a cell from row 1 to 10 and column 1
    Cells(Int(Rnd * 10) + 1, 1).Select
    'Translation:
    'Cells([random_number_between_1_and_10], 1).Select
End Sub
    
Sub example()
    'Selecting a cell (described in relation to the cell that is currently active)
    ActiveCell.Offset(2, 1).Select
              
Sub example()
    'Selecting rows 2 to 6
    Range("2:6").Select
End Sub
        
Sub example()
    'Selecting rows 2 to 6
    Rows("2:6").Select
End Sub
        
Sub example()
    'Selecting columns B to G
    Range("B:G").Select
End Sub
        
Sub example()
    'Selecting columns B to G
    Columns("B:G").Select
End Sub
  

