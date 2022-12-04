' What does this code do?
' This code allows you to select a range of cells with formulas
' And copy the formulas with no changes (e.g. rows or columns changing) to a destination range
' This could should ONLY be used for a selection of cells that is one row and multiple columns
' Code for multiple rows with one column and multiple rows with multiple columns will be added at a later date


Sub copyformularange()
Dim wb As Workbook
Dim ws As Worksheet
Dim crange As Range
Dim prange As Range
Dim i As Integer

'Once the workbook and worksheet variables are declared, you will need to set the objects
'If the code errors in the VB editor, it is likely due to the code below
'In order for the code to run properly, the workbook where you are copying and pasting MUST be selected
'Note: this code will likely not work correctly if the copy and paste ranges are in different workbooks
Set wb = ActiveWorkbook
Set ws = wb.ActiveSheet

'set the range with formulas that you would like to copy
' type=8 refers to an input box that lets you select a range
Set crange = Application.InputBox(Title:="Please select the copy range", _
    Prompt:="Select range", Type:=8)

'set the range with formulas that you would like to paste from the copy range
' type=8 refers to an input box that lets you select a range
Set prange = Application.InputBox(Title:="Please select the paste range", _
    Prompt:="Select range", Type:=8)

For i = 1 To crange.Columns.Count

    prange.Cells(RowIndex:=1, ColumnIndex:=i).Formula = _
        crange.Cells(RowIndex:=1, ColumnIndex:=i).Formula
        
Next i

End Sub
