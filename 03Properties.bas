Sub properties()
    'Incomplete Macro
    Range("A8").Value
End Sub

Sub properties()
    'A8 = 48
    Range("A8").Value = 48
    'Translation:
    'The value of cell A8 is equal to 48
End Sub


Sub properties()
    'A8 = Sample text
    Range("A8").Value = "Sample text"
End Sub


Sub properties()
    'A8 on sheet 2 = Sample text
    Sheets("Sheet2").Range("A8").Value = "Sample text"
    'Or:
    'Sheets(2).Range("A8").Value = "Sample text"
End Sub


Sub properties()
    'A8 on sheet 2 of workbork 2 = Sample text
    Workbooks("Book2.xlsx").Sheets("Sheet2").Range("A8").Value = "Sample text"
End Sub

Range("A8").Value = 48
Range("A8") = 48


Sub properties()
    'Erase the contents of column A
    Range("A:A").ClearContents
End Sub


Sub properties()
    'Edit the size of text in cells A1 through A8
    Range("A1:A8").Font.Size = 18
End Sub

Sub properties()
    'Make cells A1 through A8 bold
    Range("A1:A8").Font.Bold = True
End Sub

Sub properties()
    'Remove "bold" formatting from cells A1 through A8
    Range("A1:A8").Font.Bold = False
End Sub

Sub properties()
    'Italicize cells A1 through A8
    Range("A1:A8").Font.Italic = True
End Sub

Sub properties()
    'Underline cells A1 through A8
    Range("A1:A8").Font.Underline = True
End Sub

Sub properties()
    'Edit font in cells A1 through A8
    Range("A1:A8").Font.Name = "Arial"
End Sub

Sub properties()
    'Add a border to cells A1 to A8
    Range("A1:A8").Borders.Value = 1
    'Value = 0    => no border
End Sub


Sub properties()
    'Add a border to selected cells
    Selection.Borders.Value = 1
End Sub

Sub properties()
    'Hide a worksheet
    Sheets("Sheet3").Visible = 0
    'Visible = -1     => cancels the effect
End Sub

Sub properties()
    'A7 = A1
    Range("A7") = Range("A1")
    'Or:
    'Range("A7").Value = Range("A1").Value
End Sub


Sub properties()
    Range("A7").Font.Size = Range("A1").Font.Size
End Sub

Sub properties()
    'Click counter in A1
    Range("A1") = Range("A1") + 1
End Sub

'For example: before the code is executed, A1 has the value 0

Sub properties()

    'The button has been clicked, so the procedure is starting
    'For the moment, A1 still has the value 0
    
    'DURING the execution of the line immediately below, A1 still has the value 0
    Range("A1") = Range("A1") + 1 'And now the calculation is: New_value_of_A1 = 0 + 1
    'A1 has the value 1 only AFTER the execution of the line of code
    
End Sub

Sub properties()
    ActiveCell.Borders.Weight = 3
    ActiveCell.Font.Bold = True
    ActiveCell.Font.Size = 18
    ActiveCell.Font.Italic = True
    ActiveCell.Font.Name = "Arial"
End Sub

Sub properties()
    'Beginning of instructions using command: WITH
    With ActiveCell
        .Borders.Weight = 3
        .Font.Bold = True
        .Font.Size = 18
        .Font.Italic = True
        .Font.Name = "Arial"
    'End of instructions using command: END WITH
    End With
End Sub

Sub properties()
    With ActiveCell
        .Borders.Weight = 3
        With .Font
            .Bold = True
            .Size = 18
            .Italic = True
            .Name = "Arial"
        End With
    End With
End Sub




