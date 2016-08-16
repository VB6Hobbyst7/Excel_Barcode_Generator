# Excel_Barcode_Generator
A barcode generator developped purly in excel VBA

## V1.0 Log
The file only generates Code 128 barcode standard.
The second sheet (named "DB") has the lookup table needed for Code 128 encoding.
Code 128 has three start characters to choose from (Start A, Start B, and Start C). The file is set to use Start A by default as a global variable in the workbook object

## How does it work
The sheet basically monitors changes in the first column and treats any entry as input string. This string will be encoded into Code 128 barcode and represented it in the sheet's cells themselves.
It deos that by first reducing the column widthes (of columns "C" and beyond) to 0.15. Then it draws the barcode by changing the background color to white and black.
For example, the string "1234" is entered in cell(5,1). After leaving the cell, the sheet will represent this string in the 5th row starting from the thrid column, and it keeps expanding depending on the input string length.

## How can you change the default behaviour
The code is written into functions, so the most efficient way is to copy the barcode generation functions into your own sheet and adjust their functionality (don't forget the lookup table in sheet(2) called "DB"). The function that you may need to edit the most is the "encode" function.

You may also adjust the following variables in your worksheet. These variables live in the workbook object. To access them use ThisWOrkbook.<variableName>:
```VB
startColumn As Integer: column the barcode will start appearing in (default is 3)
start As String:		one of three values of start characters: "A", "B", "C" (default is "A")
inputColumn As Integer:	input column (default is 1 or column "A")
```

*if anyone has a better and more convenient way to implement this, please contribute to this small project
