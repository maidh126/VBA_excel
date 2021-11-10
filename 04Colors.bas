Sub example()
    'Text color for A1: green (Color num. 10)
    Range("A1").Font.ColorIndex = 10
End Sub

Sub example()
    'Text color for A1: RGB(50, 200, 100)
    Range("A1").Font.Color = RGB(50, 200, 100)
End Sub

Sub example()
    'Text color for A1: RGB(192, 32, 255)
    Range("A1").Font.Color = RGB(192, 32, 255)
End Sub

Sub example()
    'Border weight
    ActiveCell.Borders.Weight = 4
    'Border color: red
    ActiveCell.Borders.Color = RGB(255, 0, 0)
End Sub

Sub example()
    'Border weight
    Selection.Borders.Weight = 4
    'Border color: red
    Selection.Borders.Color = RGB(255, 0, 0)
End Sub

Sub example()
    'Add background color to the selected cells
    Selection.Interior.Color = RGB(174, 240, 194)
End Sub

Sub example()
    'Add color to the tab for "Sheet1"
    Sheets("Sheet1").Tab.Color = RGB(255, 0, 0)
End Sub

