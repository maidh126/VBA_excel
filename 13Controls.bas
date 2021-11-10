Private Sub CommandButton_validate_Click()

    Range("A1") = Textbox_number.Value
    'Textbox_number is the name of the text box
    'Value is the property that contains the value of the text box
    
    Unload Me
    'Unload closes the UserForm
    'We are using Me in place of the name of the UserForm (because this code is within the UserForm that we want to close)
End Sub

Private Sub Textbox_number_Change()
    If IsNumeric(Textbox_number.Value) Then 'IF numerical value ...
        Label_error.Visible = False 'Label hidden
    Else 'OTHERWISE ...
        Label_error.Visible = True 'Label shown
    End If
End Sub

Private Sub CommandButton_validate_Click()
    If IsNumeric(Textbox_number.Value) Then 'IF numerical value ...
        Range("A1") = Textbox_number.Value 'Copy to A1
        Unload Me 'Closing
    Else 'OTHERWISE ...
        MsgBox "Incorrect value"
    End If
End Sub

Private Sub Textbox_number_Change()
    If IsNumeric(Textbox_number.Value) Then 'IF numerical value ...
        Label_error.Visible = False 'Label hidden
        Me.Width = 156 'UserForm Width
    Else 'OTHERWISE ...
        Label_error.Visible = True 'Label shown
        Me.Width = 244 'UserForm Width
    End If
End Sub

Private Sub CheckBox1_Click() 'Number 1
    If CheckBox1.Value = True Then 'If checked ...
        Range("A2") = "Checked"
    Else 'If not checked ...
        Range("A2") = "Unchecked"
    End If
End Sub

Private Sub CheckBox2_Click() 'Number 2
    If CheckBox2.Value = True Then 'If checked ...
        Range("B2") = "Checked"
    Else 'If not checked ...
        Range("B2") = "Unchecked"
    End If
End Sub

Private Sub CheckBox3_Click() 'Number 3
    If CheckBox3.Value = True Then 'If checked ...
        Range("C2") = "Checked"
    Else 'If not checked ...
        Range("C2") = "Unchecked"
    End If
End Sub

Private Sub UserForm_Initialize() 'Check box if "Checked"
    If Range("A2") = "Checked" Then
        CheckBox1.Value = True
    End If
    
    If Range("B2") = "Checked" Then
        CheckBox2.Value = True
    End If
    
    If Range("C2") = "Checked" Then
        CheckBox3.Value = True
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim column_value As String, row_value As String
    
    'Loop for each Frame_column control
    For Each column_button In Frame_column.Controls
        'If the value of the control  = True (then, if checked) ...
        If column_button.Value Then
            'The variable "column_value" takes the value of the button text
            column_value = column_button.Caption
        End If
    Next
    
    'Loop for the other frame
    For Each row_button In Frame_row.Controls
        If row_button.Value Then
            row_value = row_button.Caption
        End If
    Next

    Range(column_value & row_value) = "Cell chosen !"
    Unload Me
End Sub

Private Function column_value()
'The function returns the value of the text for the button chosen (column_value)
    For Each column_button In Frame_column.Controls
        If column_button.Value Then
            column_value = column_button.Caption
        End If
    Next
End Function

Private Function row_value()
'The function returns the value of the text for the button chosen (row_value)
    For Each row_button In Frame_row.Controls
        If row_button.Value Then
            row_value = row_button.Caption
        End If
    Next
End Function

Private Sub CommandButton1_Click() 'Action that is taken when you click "Confirm your selection"
    Range(column_value & row_value) = "Cell chosen !"
    'column_value and row_value are the values returned by the functions
    Unload Me
End Sub

Private Sub activate_button()
'Activating the button if the condition is verified
    If column_value <> "" And row_value <> "" Then
    'column_value and row_value are the values returned by the functions
        CommandButton1.Enabled = True
        CommandButton1.Caption = "Confirm your selection"
    End If
End Sub

Private Sub OptionButton11_Click()
    activate_button 'Run the "activate_button" procedure
End Sub
Private Sub OptionButton12_Click()
    activate_button
End Sub
Private Sub OptionButton13_Click()
    activate_button
End Sub
Private Sub OptionButton14_Click()
    activate_button
End Sub
Private Sub OptionButton15_Click()
    activate_button
End Sub
Private Sub OptionButton16_Click()
    activate_button
End Sub
Private Sub OptionButton17_Click()
    activate_button
End Sub
Private Sub OptionButton18_Click()
    activate_button
End Sub
Private Sub OptionButton19_Click()
    activate_button
End Sub
Private Sub OptionButton20_Click()
    activate_button
End Sub

'Gray background color in the cells
Cells.Interior.Color = RGB(240, 240, 240)

'Applying color and selecting the cell
With Cells(ScrollBar_vertical.Value, ActiveCell.Column) 'Identifying the cell using Value
    .Interior.Color = RGB(255, 220, 100) 'Applying Orange Color
    .Select 'Selecting the cell
End With

Private Sub vertical_bar()
    'Applying gray background color to the cells
    Cells.Interior.Color = RGB(240, 240, 240)
    
    'Applying background color and selecting the cell
    With Cells(ScrollBar_vertical.Value, ActiveCell.Column)
        .Interior.Color = RGB(255, 220, 100) 'Orange
        .Select 'Selecting the cell
    End With
End Sub

Private Sub ScrollBar_vertical_Change()
    vertical_bar
End Sub

Private Sub ScrollBar_vertical_Scroll()
    vertical_bar
End Sub

Private Sub horizontal_bar()
    'Applying gray background color to the cells
    Cells.Interior.Color = RGB(240, 240, 240)

    'Applying background color and selecting cell
    With Cells(ActiveCell.Row, ScrollBar_horizontal.Value)
        .Interior.Color = RGB(255, 220, 100) 'Orange
        .Select 'Selecting the cell
    End With
End Sub

Private Sub ScrollBar_horizontal_Change()
    horizontal_bar
End Sub

Private Sub ScrollBar_horizontal_Scroll()
    horizontal_bar
End Sub

Private Sub UserForm_Initialize()
    For i = 1 To 4 ' => to list the 4 countries
        ComboBox_Country.AddItem Cells(1, i) 'Add the values of cells A1 through D1 using the loop
    Next
End Sub

Private Sub ComboBox_Country_Change()
    'Emptied list area (otherwise the cities are added immediately)
    ListBox_Cities.Clear
    
    Dim column_number As Integer, rows_number As Integer
    
    'The number of the selection (ListIndex starts with 0):
    column_number = ComboBox_Country.ListIndex + 1
    'Number of rows in the chosen country's column:
    rows_number = Cells(1, column_number).End(xlDown).Row

    For i = 2 To rows_number ' => to list cities
        ListBox_Cities.AddItem Cells(i, column_number)
    Next
End Sub

Private Sub ComboBox_Country_Change()
    ListBox_Cities.Clear
    For i = 2 To Cells(1, ComboBox_Country.ListIndex + 1).End(xlDown).Row
        ListBox_Cities.AddItem Cells(i, ComboBox_Country.ListIndex + 1)
    Next
End Sub

Private Sub ListBox_Cities_Click()
    TextBox_Choice.Value = ListBox_Cities.Value
End Sub
