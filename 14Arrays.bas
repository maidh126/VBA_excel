'Sample declaration of a 1 dimensional array 
Dim array1(4)

'Sample declaration of a 2 dimensional array 
Dim array2(6, 1)

'Sample declaration of a dynamic array 
Dim array3()

'Storing values in the array
array_example(0) = Range("A2")
array_example(1) = Range("A3")
array_example(2) = Range("A4")
array_example(3) = Range("A5")
array_example(4) = Range("A6")
array_example(5) = Range("A7")
array_example(6) = Range("A8")
array_example(7) = Range("A9")
array_example(8) = Range("A10")
array_example(9) = Range("A11")
array_example(10) = Range("A12")

Sub example()
    'Declaration
    Dim array_example(10)
    
    'Storing values in the array
    array_example(0) = Range("A2")
    array_example(1) = Range("A3")
    array_example(2) = Range("A4")
    array_example(3) = Range("A5")
    array_example(4) = Range("A6")
    array_example(5) = Range("A7")
    array_example(6) = Range("A8")
    array_example(7) = Range("A9")
    array_example(8) = Range("A10")
    array_example(9) = Range("A11")
    array_example(10) = Range("A12")
    
    'Test 1
    MsgBox array_example(8) '=> returns: 02.04.2016
    
    'Changing one of the values
    array_example(8) = Year(array_example(8))
    
    'Test 2
    MsgBox array_example(8) '=> returns: 2016
End Sub

'Declaration
Dim array_example(10)

'Storing values in the array
For i = 0 To 10
    array_example(i) = Range("A" & i + 2)
Next



'Declaration
Dim array_example(10, 2) '11 x 3 "element" array

'Storing values in the array
For i = 0 To 10
    array_example(i, 0) = Range("A" & i + 2)
    array_example(i, 1) = Range("B" & i + 2)
    array_example(i, 2) = Range("C" & i + 2)
Next

MsgBox array_example(0, 0) '=> returns: 03.11.2026
MsgBox array_example(0, 1) '=> returns: 24
MsgBox array_example(9, 2) '=> returns: NO
MsgBox array_example(10, 2) '=> returns: YES

last_row = Range("A1").End(xlDown).Row

Dim array_example()
ReDim array_example(last_row - 2, 2)

Sub example()
    last_row = Range("A1").End(xlDown).Row 'Last row of the data set

    Dim array_example()
    ReDim array_example(last_row - 2, 2)
    
    'Storing values in the array
    For i = 0 To last_row - 2
        array_example(i, 0) = Range("A" & i + 2)
        array_example(i, 1) = Range("B" & i + 2)
        array_example(i, 2) = Range("C" & i + 2)
    Next
End Sub

For i = 0 To last_row - 2

For i = 0 To UBound(array_example)

Sub example()
    Dim array_example(10, 2)
    
    MsgBox UBound(array_example) '=> returns: 10
    MsgBox UBound(array_example, 1) '=> returns: 10
    MsgBox UBound(array_example, 2) '=> returns: 2
End Sub

'Declaration
Dim array_example(10, 2) '11 x 3 "element" array

'Storing values in the array
For i = 0 To 10
    array_example(i, 0) = Range("A" & i + 2)
    array_example(i, 1) = Range("B" & i + 2)
    array_example(i, 2) = Range("C" & i + 2)
Next

'Declaration
Dim array_example()

'Storing values in the array
array_example = Range("A2:C12").Value

Dim en(5)

en(0) = "IF"
en(1) = "VLOOKUP"
en(2) = "SUM"
en(3) = "COUNT"
en(4) = "ISNUMBER"
en(5) = "MID"

en = Array("IF", "VLOOKUP", "SUM", "COUNT", "ISNUMBER", "MID")

Sub replace_example()
    Dim var_translate As String

    'A string for this example
    var_translate = "Hello World !"
    
    'Replacement of "World" with "you" in the character string
    var_translate = Replace(var_translate, "World", "you")

    'The string after replacement
    MsgBox var_translate '=> returns "Hello you !"
End Sub

Sub translate() 'Simplified example of EN-FR translation for formulas
    Dim var_translate As String

    'A string for this example
    var_translate = "Formula to translate: SUM(IF(ISNUMBER(A1:E1),A1:E1,0))"
    
    'The two series of values
    en = Array("IF", "VLOOKUP", "SUM", "COUNT", "ISNUMBER", "MID")
    fr = Array("SI", "RECHERCHEV", "SOMME", "NB", "ESTNUM", "STXT")
    
    'Replacing "SI" with "IF", and "RECHERVEV" with "VLOOKUP", etc.
    For i = 0 To UBound(en)
        var_translate = Replace(var_translate, en(i), fr(i))
    Next

    'The string after the replacements
    MsgBox var_translate '=> returns "Formula to translate: SOMME(SI(ESTNUM(A1:E1),A1:E1,0))"
End Sub

variable = "IF/VLOOKUP/SUM/COUNT/ISNUMBER/MID"

en = Split(variable, "/")

MsgBox en(0) '=> returns: IF
MsgBox en(1) '=> returns: VLOOKUP
MsgBox en(2) '=> returns: SUM
MsgBox en(3) '=> returns: COUNT
MsgBox en(4) '=> returns: ISNUMBER
MsgBox en(5) '=> returns: MID

en = Array("IF", "VLOOKUP", "SUM", "COUNT", "ISNUMBER", "MID")
en = Split("IF,VLOOKUP,SUM,COUNT,ISNUMBER,MID", ",")
en = Split("IF VLOOKUP SUM COUNT ISNUMBER MID", " ")

MsgBox Split("IF,VLOOKUP,SUM,COUNT,ISNUMBER,MID", ",")(2) '=> returns: SUM

MsgBox Join(Array(1, 2, 3, 4, 5), "") '=> returns: 12345
